import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

// Import the functions you need from the SDKs you need
import { initializeApp } from "firebase/app";
import { getAnalytics } from "firebase/analytics";
// TODO: Add SDKs for Firebase products that you want to use
// https://firebase.google.com/docs/web/setup#available-libraries

// Your web app's Firebase configuration
// For Firebase JS SDK v7.20.0 and later, measurementId is optional
const firebaseConfig = {
  apiKey: "AIzaSyCynbEhLjpqlLg8TQtFN4KhOpZmXXnwEH8",
  authDomain: "alpha-orc.firebaseapp.com",
  projectId: "alpha-orc",
  storageBucket: "alpha-orc.firebasestorage.app",
  messagingSenderId: "846678862133",
  appId: "1:846678862133:web:507af3f988ea9e29df447c",
  measurementId: "G-LZLJQL4W51"
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);
const analytics = getAnalytics(app);
// (Configurações do Projeto > Geral > Seus apps > ícone "</>")
const firebaseConfig = {
  apiKey: "COLE_AQUI",
  authDomain: "SEU-PROJETO.firebaseapp.com",
  projectId: "SEU-PROJETO",
  storageBucket: "SEU-PROJETO.appspot.com",
  messagingSenderId: "COLE_AQUI",
  appId: "COLE_AQUI",
};

export const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
