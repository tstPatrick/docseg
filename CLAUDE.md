# DocSeg Digital — instruções para o Claude

App web (PWA) de Permissão de Trabalho e checklists. Tudo do app fica em
`index.html`; o servidor é o Google Apps Script em `apps-script/Codigo.gs`.
Fale sempre em português com o dono do projeto.

## Versão — a cada alteração no app

Toda mudança que vai para o app sobe **os dois números juntos**, no mesmo commit:

1. `APP_VERSION` em `index.html` (aparece na tela de Suporte e vai nos chamados):
   - mudança pequena ou correção: `5.0` → `5.1` → `5.2`…
   - funcionalidade grande: `5.x` → `6.0`
2. `CACHE_NAME` em `service-worker.js` (`docseg-shell-vNN` → `vNN+1`), senão os
   celulares continuam com a versão antiga.

Informe a nova versão ao dono do projeto ao terminar.

## Jeito de trabalhar

- Mudança que altera o visual: mandar **mockups antes** (de preferência em cima
  do app real) e só fazer depois da aprovação.
- Trabalhar numa branch, testar no navegador e só abrir Pull Request quando o
  dono pedir. O merge é feito por ele.
- Mudança em `apps-script/Codigo.gs` precisa ser colada no editor do Apps Script
  e publicada como **Nova versão** (passo a passo em `apps-script/LEIA-ME.md`):
  sempre avisar e dar o passo a passo, e publicar o servidor antes do app.
- Documento novo segue o padrão de liberação descrito no comentário acima de
  `LIB_DOCS_ROTULO` em `index.html` (e em `DOC_TIPOS` no servidor).
- Nunca colocar chaves do Firebase no código: elas ficam nas Propriedades do script.
