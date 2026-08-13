/**
 * Firebase configuration for S2 Launch Feedback Tool.
 * Replace these values with your Firebase project credentials.
 * Get them from: Firebase Console → Project Settings → Your apps → Web app
 */
export const firebaseConfig = {
  apiKey: "YOUR_API_KEY",
  authDomain: "YOUR_PROJECT.firebaseapp.com",
  projectId: "YOUR_PROJECT_ID",
  storageBucket: "YOUR_PROJECT.appspot.com",
  messagingSenderId: "YOUR_SENDER_ID",
  appId: "YOUR_APP_ID",
};

/** Firestore collection names */
export const COLLECTIONS = {
  ROOT: "s2_feedback",
  CONFIG: "config",
  RESPONSES: "responses",
};
