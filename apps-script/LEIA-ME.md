# Servidor do DocSeg Digital (Google Apps Script)

`Codigo.gs` é o código do servidor. Ele **não** é publicado sozinho: depois de
alterar aqui, é preciso copiar para o editor do Apps Script.

## Como atualizar o servidor

1. Abra https://script.google.com com a conta dona do DocSeg e entre no projeto.
2. Abra o arquivo `Código.gs`, apague todo o conteúdo e cole o de `Codigo.gs`. Salve (Ctrl+S).
3. **Implantar → Gerenciar implantações →** ✏️ (editar) **→ Versão: Nova versão → Implantar.**
   Assim o endereço `/exec` continua o mesmo e o app não precisa mudar nada.
4. Na barra de cima, escolha a função `instalarGatilhos` e clique em ▶ **Executar** (uma vez).
   Isso liga o lembrete dos documentos que estão esperando liberação.

As chaves do Firebase (`fcm_client_email`, `fcm_private_key`) ficam nas
**Propriedades do script** e nunca devem ser coladas neste arquivo.

## Liberação de documentos (Apto / Não apto)

- Ligada por empresa em **Empresas → Liberação de documentos**.
- Os PDFs esperando avaliação ficam em `DocSeg Digital - PDFs / _Pendentes` (o cliente não vê).
- Depois da avaliação: liberados na pasta do mês da empresa; recusados em
  `<mês> / Não Aptas` (PT) ou `<mês> / Não Aptos` (checklists).
- Quem libera: o **Técnico de Segurança** (perfil `tst`, cadastrado pelo gestor). Se a
  empresa não tiver nenhum, o gestor libera.
- Tudo fica registrado na aba **Liberações** da planilha; o registro final entra
  em **Registros** com as colunas Liberação, Liberado por, Liberado em e Motivo.
