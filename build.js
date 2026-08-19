#!/usr/bin/env node
// Build estático simples: copia os arquivos do site pra dist/ e injeta a chave do
// Firebase (env var FIREBASE_API_KEY do Vercel) no lugar do placeholder __FIREBASE_API_KEY__.
const fs = require('fs');
const path = require('path');

const SRC  = __dirname;
const DIST = path.join(__dirname, 'dist');

const FIREBASE_API_KEY = process.env.FIREBASE_API_KEY;
if (!FIREBASE_API_KEY) {
  console.error('ERRO: variável de ambiente FIREBASE_API_KEY não definida (configure em Vercel → Settings → Environment Variables).');
  process.exit(1);
}

fs.rmSync(DIST, { recursive: true, force: true });
fs.mkdirSync(DIST, { recursive: true });

const STATIC_FILES = ['sw.js', 'manifest.json', 'logo.png', 'card-evento.html'];
for (const file of STATIC_FILES) {
  const from = path.join(SRC, file);
  if (fs.existsSync(from)) fs.copyFileSync(from, path.join(DIST, file));
}

const html = fs.readFileSync(path.join(SRC, 'index.html'), 'utf8')
  .split('__FIREBASE_API_KEY__').join(FIREBASE_API_KEY);
fs.writeFileSync(path.join(DIST, 'index.html'), html);

console.log('Build concluído em dist/');
