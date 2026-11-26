# Instruções do Daniel

## Escrita

- Não usar o travessão "—". Usar sempre o meia-risca: –

## E-mail e Microsoft 365: sempre pelo Microsoft Graph

- **Tudo do Microsoft 365** (e-mail, calendário, OneDrive, SharePoint, To Do, contatos)
  é feito pelo **Microsoft Graph API**, direto da sessão.
- **Não usar o Cofre (Cofre_RAA)** nem **conectores** (Outlook/Microsoft 365) para isso.
- Credenciais já estão no ambiente de nuvem: `MS_GRAPH_CLIENT_ID`, `MS_GRAPH_TENANT_ID`,
  `MS_GRAPH_REFRESH_TOKEN`. Gerar o token com `grant_type=refresh_token` em
  `https://login.microsoftonline.com/<tenant>/oauth2/v2.0/token`. Nunca imprimir os valores.
- Conta: daniel.aquino@raa.adv.br (Rocha, Araújo & Arraes Advogados – domínio `raa.adv.br`).
- Para achar o e-mail de alguém: `GET /me/people?$search=<nome>`.
- Script pronto: `ferramentas/graph_email.py` (função `enviar`), que já aplica o padrão abaixo.

## Padrão de formatação dos e-mails

Tirado dos e-mails que o Daniel envia pelo Outlook:

- Corpo em **HTML**, fonte **Century Gothic** em todo o texto.
- **Abertura com menção (arroba)**: `@Nome Sobrenome,` como link `mailto:` sem sublinhado.
  Quem é mencionado também entra como destinatário.
- Para ser **menção de verdade** (e não só link), são necessárias duas coisas:
  o link com `id="OWAAM<GUID em hex maiúsculo>"` e a mensagem criada na API **beta**
  (`POST /beta/me/messages`) com a lista `mentions`, enviada depois com `/send`.
  O `ferramentas/graph_email.py` já faz isso.
- No `style`, escrever a fonte com aspas simples (`font-family:'Century Gothic',sans-serif`);
  aspas duplas quebram o atributo e a fonte some.
- **Uma linha em branco entre cada parágrafo** (inclusive após a menção e entre tópicos).
- Frases curtas e diretas. Rótulos em negrito no início do parágrafo quando ajudam
  (ex.: **Tese:**, **Números:**, **Chance:**).
- Listas com marcadores (•) para itens.
- Tratamento "o senhor" para sócios/superiores quando for o caso.
- Fecho: `Atenciosamente,` + linha em branco + assinatura.
- Assunto de resposta/encaminhamento no padrão do Outlook em português: `RES:` e `ENC:`.

## Assinatura

- A assinatura é a **imagem** `ferramentas/assinatura.png` (1350×450), inserida **inline**
  (anexo com `isInline: true` e `contentId`, referenciada por `cid:`) e exibida em 589×196 px.
- Conteúdo da imagem, caso precise da versão em texto:

  **DANIEL AQUINO MOREIRA** – Coordenador – OAB/CE 45.717
  +55 85 99222 5399 · daniel.aquino@raa.adv.br
  **ROCHA, ARAÚJO & ARRAES ADVOGADOS**
  Fortaleza – CE: Av. Santos Dumont – 1687, 7º andar – Aldeota / 60150-160 – +55 85 3264-1659
  São Paulo – SP: Av. Pres. Juscelino Kubitschek – 1327, 4º andar – Itaim Bibi / 04543-011 – +55 11 92062-1244
  @raa.advogados | raa.adv.br
