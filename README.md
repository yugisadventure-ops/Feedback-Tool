# S2 Launch Feedback Tool

A premium, interactive feedback collection tool for the S2 Launch Event. Built with HTML/CSS/JS and Firebase Firestore for cloud data storage.

## Features

- **Dynamic Excel-based forms** — Upload `.xlsx` templates to create/update feedback questions
- **Image support** — Questions can include images via URL in the `image` column
- **Firebase Firestore** — Responses and form config stored in the cloud
- **Interactive dashboard** — KPIs, charts (NPS, ratings, response trends), recent submissions
- **Export** — Download responses (CSV/JSON) and dashboard summaries (CSV/JSON)
- **Premium UI** — Dark glass-morphism design with smooth animations

## Quick Start

### 1. Serve the app locally

```bash
# Using Python
python3 -m http.server 8080

# Or using Node.js
npx serve .
```

Open `http://localhost:8080` in your browser.

### 2. Configure Firebase

1. Create a project at [Firebase Console](https://console.firebase.google.com)
2. Enable **Firestore Database**
3. Register a **Web app** and copy the config
4. Paste values into `js/firebase-config.js`:

```js
export const firebaseConfig = {
  apiKey: "your-api-key",
  authDomain: "your-project.firebaseapp.com",
  projectId: "your-project-id",
  storageBucket: "your-project.appspot.com",
  messagingSenderId: "123456789",
  appId: "your-app-id",
};
```

5. Deploy Firestore rules from `firestore.rules` (or use test mode initially)

### 3. Upload your Excel template

Go to **Admin → Upload Form Template** and upload your Excel file, or use the included sample at `templates/s2-launch-feedback-template.xlsx`.

## Excel Template Format

| Column | Description |
|--------|-------------|
| `id` | Unique question ID (e.g. `overall_rating`) |
| `section` | Section name for grouping |
| `type` | `text`, `email`, `textarea`, `select`, `radio`, `checkbox`, `rating`, `scale`, `nps`, `info` |
| `question` | Question text |
| `options` | Pipe-separated options: `A\|B\|C` |
| `image` | Image URL for question illustration |
| `required` | `true` or `false` |
| `min` / `max` | Range for rating/scale/nps |
| `placeholder` | Input placeholder |

Optional **Meta** sheet: `event_name`, `event_date`, `version`, `description`.

## Data Structure (Firestore)

```
s2_feedback/
  config/                    ← Form template JSON
  config/responses/          ← Participant responses
    {responseId}/
      answers: { ... }
      participantName: "..."
      submittedAt: "..."
```

## Fallback Mode

Without Firebase configuration, the tool stores data in `localStorage` so you can test locally.

## Deploy

Host the folder on any static hosting (Firebase Hosting, Netlify, Vercel, GitHub Pages). Ensure CORS and Firebase domain authorization are configured for your domain.
