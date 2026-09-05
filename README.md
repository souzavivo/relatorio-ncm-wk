# Diagnóstico NCM — Reforma Tributária

## Rodar localmente
npm install
npm run dev

## Publicar na Vercel
Veja o passo a passo que o Claude te mandou no chat.

## Importante
- Login e contador de uso ficam salvos em localStorage (no navegador de quem acessa).
  Isso é por dispositivo/navegador — se o mesmo cliente abrir de outro computador,
  vai aparecer como conta nova. Suficiente para validar a ideia; para produção de
  verdade (contas de verdade, multi-dispositivo), migrar para um backend com banco
  de dados real.
- Troque o UNLOCK_CODE em src/App.jsx por um código só seu antes de publicar.
