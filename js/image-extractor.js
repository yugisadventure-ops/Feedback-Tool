/**
 * Extract embedded images from Excel (.xlsx) files and map them to row numbers.
 */

const MIME_BY_EXT = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  webp: "image/webp",
  bmp: "image/bmp",
};

function mimeForPath(path) {
  const ext = path.split(".").pop()?.toLowerCase() || "png";
  return MIME_BY_EXT[ext] || "image/png";
}

async function readZipText(zip, path) {
  const entry = zip.file(path);
  if (!entry) return null;
  return entry.async("string");
}

async function loadMediaDataUrls(zip) {
  const media = {};
  for (const [path, entry] of Object.entries(zip.files)) {
    if (!path.startsWith("xl/media/") || entry.dir) continue;
    const base64 = await entry.async("base64");
    media[path] = `data:${mimeForPath(path)};base64,${base64}`;
  }
  return media;
}

function parseDrawingRels(xml) {
  const map = {};
  if (!xml) return map;
  const relRegex = /Relationship[^>]+Id="([^"]+)"[^>]+Target="([^"]+)"/g;
  let match;
  while ((match = relRegex.exec(xml)) !== null) {
    const target = match[2].replace(/^\.\.\//, "xl/");
    map[match[1]] = target.startsWith("xl/") ? target : `xl/${target}`;
  }
  return map;
}

function parseDrawingAnchors(xml) {
  const anchors = [];
  if (!xml) return anchors;

  const anchorBlocks = xml.split(/<xdr:(?:twoCellAnchor|oneCellAnchor)/).slice(1);
  for (const block of anchorBlocks) {
    const rowMatch = block.match(/<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/);
    const embedMatch = block.match(/r:embed="([^"]+)"/);
    if (rowMatch && embedMatch) {
      anchors.push({
        row: Number(rowMatch[1]),
        rId: embedMatch[1],
      });
    }
  }
  return anchors;
}

async function mapImagesToRows(zip) {
  const media = await loadMediaDataUrls(zip);
  const rowImages = {};

  const drawingPaths = Object.keys(zip.files).filter(
    (p) => /^xl\/drawings\/drawing\d+\.xml$/.test(p)
  );

  for (const drawingPath of drawingPaths) {
    const drawingXml = await readZipText(zip, drawingPath);
    const relsPath = drawingPath.replace("drawings/", "drawings/_rels/") + ".rels";
    const relsXml = await readZipText(zip, relsPath);
    const relMap = parseDrawingRels(relsXml);
    const anchors = parseDrawingAnchors(drawingXml);

    for (const anchor of anchors) {
      const mediaPath = relMap[anchor.rId];
      if (mediaPath && media[mediaPath]) {
        // Excel rows are 0-based in drawing XML; sheet data row 1 = header, row 2 = first question (index 0)
        const dataRowIndex = Math.max(anchor.row - 1, 0);
        rowImages[dataRowIndex] = media[mediaPath];
      }
    }
  }

  // Fallback: assign unattached media to rows in order if no anchors found
  if (!Object.keys(rowImages).length) {
    const mediaPaths = Object.keys(media).sort();
    mediaPaths.forEach((path, i) => {
      rowImages[i + 1] = media[path];
    });
  }

  return rowImages;
}

export async function extractEmbeddedImages(arrayBuffer) {
  if (typeof JSZip === "undefined") return {};
  try {
    const zip = await JSZip.loadAsync(arrayBuffer);
    return await mapImagesToRows(zip);
  } catch (e) {
    console.warn("Could not extract embedded images:", e);
    return {};
  }
}

export async function resolveImagePath(imageRef, baseUrl = "") {
  if (!imageRef) return null;
  const ref = String(imageRef).trim();
  if (!ref || ref === "true" || ref === "false") return null;
  if (/^data:image\//i.test(ref) || /^https?:\/\//i.test(ref)) return ref;

  const normalized = ref.replace(/^\.\//, "");
  const candidates = [
    normalized,
    `assets/images/${normalized}`,
    `assets/${normalized}`,
    `templates/${normalized}`,
  ];

  for (const path of candidates) {
    try {
      const resp = await fetch(`${baseUrl}${path}`);
      if (resp.ok) {
        const blob = await resp.blob();
        return await blobToDataUrl(blob);
      }
    } catch {
      /* try next */
    }
  }

  return ref;
}

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}
