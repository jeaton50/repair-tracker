import { initializeApp } from "firebase/app";
import {
  initializeFirestore,
  persistentLocalCache,
  persistentMultipleTabManager,
} from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyGCtJKb9tfptKFml0BDcCUeWNMOU_L6uSDs",
  authDomain: "repair-tracker-add33.firebaseapp.com",
  projectId: "repair-tracker-add33",
  storageBucket: "repair-tracker-add33.firebasestorage.app",
  messagingSenderId: "105779807489",
  appId: "1:105779807489:web:83bc4a86c67535f624c700",
  measurementId: "G-MR3RVKELD7"
};

const app = initializeApp(firebaseConfig);

// More robust Firestore transport (works behind proxies/VPN)
export const db = initializeFirestore(app, {
  // If streaming is blocked, SDK will fall back to long-polling
  experimentalAutoDetectLongPolling: true,
  // This combo avoids fetch-based streams some networks block
  useFetchStreams: false,

  // Optional but recommended: local cache + multi-tab
  localCache: persistentLocalCache({
    tabManager: persistentMultipleTabManager(),
  }),
});