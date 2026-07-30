import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyALlkl1rum7pXo8XV900wLL26HBoEHke-E",
  authDomain: "controle-financeiro-cd4c2.firebaseapp.com",
  projectId: "controle-financeiro-cd4c2",
  storageBucket: "controle-financeiro-cd4c2.firebasestorage.app",
  messagingSenderId: "302580715652",
  appId: "1:302580715652:web:dd760c4430a4d78f6ef2d1",
};

export const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);
export const db = getFirestore(app);
