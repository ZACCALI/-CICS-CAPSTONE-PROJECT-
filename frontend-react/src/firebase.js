import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { initializeFirestore, persistentLocalCache, persistentMultipleTabManager, CACHE_SIZE_UNLIMITED } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyBTPcqqCQgKc5Sfyr9elbbCONiINMlZRiM",
  authDomain: "cics-pa-system-v2.firebaseapp.com",
  projectId: "cics-pa-system-v2",
  storageBucket: "cics-pa-system-v2.firebasestorage.app",
  messagingSenderId: "577784873184",
  appId: "1:577784873184:web:1c4e6d87a5178ef9961085"
};

const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);

// FIX: Correct usage of persistentMultipleTabManager
export const db = initializeFirestore(app, {
  localCache: persistentLocalCache({
    tabManager: persistentMultipleTabManager(),
    cacheSizeBytes: CACHE_SIZE_UNLIMITED
  })
});

export default app;
