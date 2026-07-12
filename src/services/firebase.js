// src/services/firebase.js
//
// Inicialização única do Firebase — MESMO projeto do MiContas
// (appmicontas). Não crie um segundo initializeApp em outro arquivo;
// todo mundo importa `auth` e `db` daqui.

import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { getDatabase } from "firebase/database";

const firebaseConfig = {
  apiKey: "AIzaSyAmifhwuFquGldRB8YjbWUErqaaUF1-wNA",
  authDomain: "appmicontas.firebaseapp.com",
  databaseURL: "https://appmicontas-default-rtdb.firebaseio.com",
  projectId: "appmicontas",
  storageBucket: "appmicontas.firebasestorage.app",
  messagingSenderId: "969588602634",
  appId: "1:969588602634:web:87707e38effeeb1c75ecf3",
};

const app = initializeApp(firebaseConfig);

export const auth = getAuth(app);
export const db = getDatabase(app);
