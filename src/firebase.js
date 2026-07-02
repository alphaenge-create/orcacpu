import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyBXeXk4T0IlANh1Y8SHX5x_q6Fx9PaCyrY",
  authDomain: "alphaorc-d3667.firebaseapp.com",
  projectId: "alphaorc-d3667",
  storageBucket: "alphaorc-d3667.firebasestorage.app",
  messagingSenderId: "686963355753",
  appId: "1:686963355753:web:a9c0568c2ca486b9c932f8"
};

export const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
