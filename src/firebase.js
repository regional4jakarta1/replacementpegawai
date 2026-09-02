import { initializeApp } from 'firebase/app';
import { getFirestore } from 'firebase/firestore';

// GANTI semua nilai di bawah ini dengan config Firebase project lo sendiri.
// Cara ambil: Firebase Console -> gear icon (Project settings) -> scroll ke
// "Your apps" -> klik ikon web </> (atau app web yang udah ada) -> config
// muncul di situ, tinggal copy-paste.
const firebaseConfig = {
  apiKey: 'AIzaSyDtZNNejA52muzGYVJRHo_knCFs9aoLLTA',
  authDomain: 'replacement-pegawai.firebaseapp.com',
  projectId: 'replacement-pegawai',
  storageBucket: 'replacement-pegawai.firebasestorage.app',
  messagingSenderId: '660839095669',
  appId: '1:660839095669:web:bac3867c6ad5765dcc873c',
};

export const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);

// Path dokumen Firestore tempat data dashboard disimpan.
// Semua orang yang buka web ini baca & tulis ke dokumen yang sama persis ini.
export const DASHBOARD_DOC_PATH = ['headcount', 'current'];
