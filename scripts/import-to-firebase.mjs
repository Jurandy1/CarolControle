import { existsSync } from 'fs';
import { resolve } from 'path';
import { initializeApp } from 'firebase/app';
import { getAuth, signInAnonymously } from 'firebase/auth';
import { getFirestore, doc, setDoc } from 'firebase/firestore';
import { parseWorkbookFromFile } from './xlsx-import.mjs';

const firebaseConfig = {
  apiKey: 'AIzaSyAY9Xy04_cz_1x7t54EzHa70LIOHGVA8dM',
  authDomain: 'carolcontrole.firebaseapp.com',
  projectId: 'carolcontrole',
  storageBucket: 'carolcontrole.firebasestorage.app',
  messagingSenderId: '591231038569',
  appId: '1:591231038569:web:a9d1cd4a780fe78031def8'
};

const RECORDS_PATH = 'artifacts/carolcontrole/public/data/records/_all_records';

async function main() {
  const input = process.argv[2];
  if (!input) {
    console.error('');
    console.error('Uso: npm run import:firebase -- "C:\\caminho\\N. PROCESSO MAE.xlsx"');
    console.error('');
    process.exit(1);
  }

  const filePath = resolve(input);
  if (!existsSync(filePath)) {
    console.error('Arquivo não encontrado:', filePath);
    process.exit(1);
  }

  console.log('Lendo planilha:', filePath);
  const { records, sheetNames } = parseWorkbookFromFile(filePath);

  if (!records.length) {
    console.error('Nenhum registro reconhecido. Confira se é o arquivo N. PROCESSO MAE.xlsx.');
    process.exit(1);
  }

  const byTab = {};
  records.forEach(r => { byTab[r.tab] = (byTab[r.tab] || 0) + 1; });

  console.log('');
  console.log(`Total: ${records.length} registros em ${sheetNames.length} abas`);
  sheetNames.forEach(s => console.log(`  - ${s}: ${byTab[s] || 0}`));
  console.log('');
  console.log('Conectando ao Firebase...');

  const app = initializeApp(firebaseConfig);
  const auth = getAuth(app);
  const db = getFirestore(app);

  await signInAnonymously(auth);
  console.log('Autenticação anônima OK');

  const updatedAt = new Date().toISOString();
  await setDoc(doc(db, RECORDS_PATH), { records, updatedAt });

  console.log('');
  console.log('Importação concluída e salva na nuvem.');
  console.log(`Documento: ${RECORDS_PATH}`);
  console.log(`Atualizado em: ${updatedAt}`);
}

main().catch(err => {
  console.error('');
  console.error('Erro:', err.message || err);
  if (String(err).includes('permission') || err.code === 'permission-denied') {
    console.error('');
    console.error('Verifique no Firebase Console:');
    console.error('  1. Authentication > Sign-in method > Anonymous (ativado)');
    console.error('  2. Firestore > Rules permitindo escrita para usuários autenticados');
  }
  process.exit(1);
});
