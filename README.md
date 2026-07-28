# Controle Financeiro

App de controle financeiro pessoal (Custos, Receitas, Resumo Anual e Demandas), com exportação para Excel. Os dados ficam salvos no `localStorage` do navegador — ou seja, no dispositivo/navegador onde você abrir o app.

## Rodar localmente

Pré-requisito: [Node.js](https://nodejs.org) instalado (versão 18 ou mais recente).

```bash
npm install
npm run dev
```

Abra o endereço mostrado no terminal (geralmente `http://localhost:5173`).

## Publicar online (grátis)

### Opção A — Vercel
1. Crie uma conta em [vercel.com](https://vercel.com)
2. Instale a CLI: `npm install -g vercel`
3. Na pasta do projeto, rode: `vercel`
4. Siga as instruções na tela (aceite as opções padrão)

### Opção B — Netlify
1. Rode `npm run build` (gera a pasta `dist`)
2. Crie uma conta em [netlify.com](https://netlify.com)
3. Arraste a pasta `dist` para o painel do Netlify (drag & drop)

Depois de publicado, você recebe um link (ex: `seu-app.vercel.app`) que pode acessar de qualquer navegador — mas atenção: como os dados ficam no `localStorage`, cada navegador/dispositivo terá seus próprios dados, não sincronizados entre si.

## Estrutura

- `src/App.jsx` — todo o app (abas, formulários, cálculos, exportação Excel)
- `src/main.jsx` — ponto de entrada
- Dados iniciais (seed) já vêm carregados na primeira vez que o app roda; depois disso, tudo que você adicionar fica salvo no navegador.

## Limitações desta versão

- Os dados **não sincronizam** entre dispositivos diferentes ou navegadores diferentes.
- Limpar o cache/dados do navegador apaga o conteúdo salvo.
- Se quiser sincronização entre aparelhos, é possível evoluir para um banco de dados online (Firebase ou Supabase) — é só pedir.
