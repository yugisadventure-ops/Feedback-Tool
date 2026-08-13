#!/usr/bin/env python3
"""Build assets/images/manifest.json from files in this directory."""

from pathlib import Path
import json

ROOT = Path(__file__).resolve().parents[1]
IMG_DIR = ROOT / "assets" / "images"
EXTS = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg"}

def main():
    images = sorted(
        p.name for p in IMG_DIR.iterdir()
        if p.is_file() and p.suffix.lower() in EXTS
    )
    manifest = {"images": [f"assets/images/{name}" for name in images]}
    out = IMG_DIR / "manifest.json"
    out.write_text(json.dumps(manifest, indent=2) + "\n")
    print(f"Wrote {len(images)} images to {out}")

if __name__ == "__main__":
    main()
