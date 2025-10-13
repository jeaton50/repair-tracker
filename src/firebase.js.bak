import { initializeApp } from 'firebase/app';
import { initializeFirestore, persistentLocalCache, persistentMultipleTabManager } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyGCtJKb9tfptKFml0BDcCUeWNMOU_L6uSDs",
  authDomain: "repair-tracker-add33.firebaseapp.com",
  projectId: "repair-tracker-add33",
  storageBucket: "repair-tracker-add33.firebasestorage.app",
  messagingSenderId: "105779807489",
  appId: "1:105779807489:web:83bc4a86c67535f624c700",
  measurementId: "G-MR3RVKELD7"
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);

// Initialize Firestore with offline persistence (20-40% cost savings!)
const db = initializeFirestore(app, {
  localCache: persistentLocalCache({
    tabManager: persistentMultipleTabManager()
  })
});

export { db };