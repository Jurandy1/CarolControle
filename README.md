# Controle Carol

Sistema web para monitoramento de processos SEI, baseado na planilha **N. PROCESSO MAE.xlsx**.

## Como usar

1. Abra `index.html` no navegador (Chrome ou Edge)
2. Clique em **Importar Planilha (.xlsx)** e selecione o arquivo da planilha
3. Navegue pelas abas, edite registros e cole os links diretos do SEI manualmente

## Abas suportadas

- PCA 2026 / PCA 2025
- 2025 (pagamentos mensais)
- MEMOS CAP
- STARTGOV
- FISCAIS DE CONTRATO
- SEI 2024
- PROCESSOS CONECTA
- MEMOS DIVERSOS

## Preview local

```powershell
npm run preview
```

Acesse: http://localhost:8765/index.html

## Importar planilha para o Firebase (todas as abas)

Requisitos: Node.js instalado e planilha `N. PROCESSO MAE.xlsx` no computador.

```powershell
cd "C:\Users\PC\Desktop\CarolControle-main"
npm install
npm run import:firebase -- "C:\caminho\completo\N. PROCESSO MAE.xlsx"
```

Ou use o atalho:

```powershell
powershell -ExecutionPolicy Bypass -File importar-planilha.ps1 -Planilha "C:\caminho\N. PROCESSO MAE.xlsx"
```

O script lê **todas as abas** da planilha e grava em um único documento no Firestore (`_all_records`), igual ao app no navegador.

### Firebase (console)

No [Firebase Console](https://console.firebase.google.com/) do projeto `carolcontrole`:

1. **Authentication** → Sign-in method → **Anonymous** (ativado)
2. **Firestore Database** → criado e com regras que permitam leitura/escrita para usuários autenticados

## GitHub

O código pode ser publicado sem a planilha (`.xlsx` está no `.gitignore`).

```powershell
git init
git add .
git commit -m "Controle Carol - sistema SEI"
gh auth login
gh repo create CarolControle --public --source=. --remote=origin --push
```
