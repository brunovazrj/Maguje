import { initializeApp } from 'firebase/app';
import { getFirestore, doc, getDoc, setDoc } from 'firebase/firestore';

const firebaseConfig = {
  apiKey: "AIzaSyDpTdAsI8WAguOOWeZ6yxWDFY8Vtls0hFQ",
  authDomain: "maguje-comissao.firebaseapp.com",
  projectId: "maguje-comissao",
  storageBucket: "maguje-comissao.firebasestorage.app",
  messagingSenderId: "612605949049",
  appId: "1:612605949049:web:792dbff7906fb2aaec10a8"
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);

// ── Helpers genéricos ─────────────────────────────────────────
async function fsGet(path) {
  try {
    const snap = await getDoc(doc(db, ...path.split('/')));
    return snap.exists() ? snap.data().d : null;
  } catch (e) { console.error('fsGet', path, e); return null; }
}
async function fsSave(path, value) {
  try {
    await setDoc(doc(db, ...path.split('/')), { d: value });
  } catch (e) { console.error('fsSave', path, e); }
}

// ── API pública ───────────────────────────────────────────────
export const fsGetRevenue   = month => fsGet(`revenue/${month}`);
export const fsSaveRevenue  = (month, v) => fsSave(`revenue/${month}`, v);
export const fsGetAbsences  = month => fsGet(`absences/${month}`);
export const fsSaveAbsences = (month, v) => fsSave(`absences/${month}`, v);
export const fsGetEmployees = () => fsGet('app/employees');
export const fsSaveEmployees = v => fsSave('app/employees', v);
export const fsGetHistory   = () => fsGet('app/history');
export const fsSaveHistory  = v => fsSave('app/history', v);
