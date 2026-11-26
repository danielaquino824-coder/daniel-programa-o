"""Envio de e-mail pelo Microsoft Graph no padrão do Daniel.

Uso como módulo:
    from graph_email import enviar, Citacao
    enviar(
        para=[("Victor Morais", "victor.morais@raa.adv.br")],
        assunto="Assunto",
        mencoes=[("Victor Morais", "victor.morais@raa.adv.br")],
        paragrafos=["Primeiro parágrafo.", "<b>Rótulo:</b> texto.",
                    Citacao("Art. 724. O corretor terá direito ..."),
                    ["tópico 1", "tópico 2"]],
    )

Credenciais: variáveis MS_GRAPH_CLIENT_ID, MS_GRAPH_TENANT_ID e
MS_GRAPH_REFRESH_TOKEN do ambiente de nuvem. Nunca imprimir os valores.
"""
import base64, html, json, os, urllib.parse, urllib.request, uuid

ASSINATURA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assinatura.png")
FONTE = "font-family:'Century Gothic',sans-serif"
CID = "assinatura_raa"


def token():
    t = os.environ["MS_GRAPH_TENANT_ID"]
    data = urllib.parse.urlencode({
        "client_id": os.environ["MS_GRAPH_CLIENT_ID"],
        "grant_type": "refresh_token",
        "refresh_token": os.environ["MS_GRAPH_REFRESH_TOKEN"],
        "scope": "https://graph.microsoft.com/.default offline_access",
    }).encode()
    url = f"https://login.microsoftonline.com/{t}/oauth2/v2.0/token"
    return json.load(urllib.request.urlopen(url, data))["access_token"]


def call(method, path, body=None, tk=None):
    url = path if path.startswith("http") else "https://graph.microsoft.com/v1.0" + path.replace(" ", "%20")
    url = urllib.parse.quote(url, safe=":/?&=$,()'\"%@+")
    req = urllib.request.Request(
        url, method=method,
        data=json.dumps(body).encode() if body is not None else None,
        headers={"Authorization": "Bearer " + (tk or token()), "Content-Type": "application/json"},
    )
    r = urllib.request.urlopen(req)
    b = r.read()
    return r.status, (json.loads(b) if b else None)


ESTILO = (
    "<style>"
    "p.MsoNormal, li.MsoNormal, div.MsoNormal"
    "{margin:0cm; font-size:11.0pt; font-family:'Century Gothic',sans-serif}"
    "p.MsoListParagraph, li.MsoListParagraph"
    "{margin:0cm 0cm 0cm 36.0pt; font-size:11.0pt; font-family:'Century Gothic',sans-serif}"
    "</style>"
)


RECUO_CITACAO = "141.6pt"  # 4 tabs do Outlook (4 x 35,4 pt)


class Citacao(str):
    """Trecho citado (jurisprudência, artigo de lei, doutrina, trecho de documento...).
    Vai recuado 4 tabs, justificado, mesma fonte e tamanho."""


def _p(conteudo, recuo=None):
    # margem zero também inline, para leitores que ignoram o <style>; sempre justificado
    estilo = (f"margin:0cm 0cm 0cm {recuo}" if recuo else "margin:0cm") + "; text-align:justify"
    return (f'<p class="MsoNormal" style="{estilo}"><span style="{FONTE}; font-size:11.0pt">'
            f'{conteudo}</span></p>')


def _vazio():
    return _p("&nbsp;")


def mencao(nome, email):
    # o id "OWAAM" + GUID é o que faz o Outlook tratar o link como menção (@)
    mid = "OWAAM" + uuid.uuid4().hex.upper()
    return (f'<a id="{mid}" href="mailto:{email}"><span style="{FONTE}; text-decoration:none">'
            f'@{html.escape(nome)}</span></a>')


def corpo_html(paragrafos, mencoes=None):
    """paragrafos: lista de itens, cada um:
      - str (HTML permitido): parágrafo normal;
      - Citacao("..."): citação recuada 4 tabs;
      - list/tuple de str: tópicos com marcador.

    Padrão do Daniel: tudo justificado, parágrafos sem margem, 11 pt,
    UMA linha vazia entre eles.
    """
    partes = []
    if mencoes:
        partes += [_p(", ".join(mencao(n, e) for n, e in mencoes) + ","), _vazio()]
    for item in paragrafos:
        if isinstance(item, (list, tuple)):
            for li in item:
                partes.append(f'<ul type="disc" style="margin-top:0cm; margin-bottom:0cm">'
                              f'<li class="MsoListParagraph" style="margin:0cm; text-align:justify">'
                              f'<span style="{FONTE}; font-size:11.0pt">{li}</span></li></ul>')
                partes.append(_vazio())
        elif isinstance(item, Citacao):
            partes += [_p(item, RECUO_CITACAO), _vazio()]
        else:
            partes += [_p(item), _vazio()]
    partes.append(_p("Atenciosamente,"))
    partes.append(_vazio())
    partes.append(f'<p class="MsoNormal" style="margin:0cm"><img border="0" width="589" height="196" '
                  f'src="cid:{CID}" style="width:6.1354in; height:2.0416in"></p>')
    return ('<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8">'
            + ESTILO + '</head><body lang="PT-BR"><div class="WordSection1">'
            + "".join(partes) + "</div></body></html>")


def enviar(para, assunto, paragrafos, mencoes=None, cc=None):
    def rec(lista):
        return [{"emailAddress": {"name": n, "address": e}} for n, e in (lista or [])]
    # quem é mencionado também entra como destinatário (igual ao Outlook)
    para = list(para) + [m for m in (mencoes or []) if m[1] not in {e for _, e in para}]
    msg = {
        "subject": assunto,
        "body": {"contentType": "HTML", "content": corpo_html(paragrafos, mencoes)},
        "toRecipients": rec(para),
        "ccRecipients": rec(cc),
        "attachments": [{
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": "assinatura.png", "contentType": "image/png",
            "contentId": CID, "isInline": True,
            "contentBytes": base64.b64encode(open(ASSINATURA, "rb").read()).decode(),
        }],
    }
    if not mencoes:
        return call("POST", "/me/sendMail", {"message": msg, "saveToSentItems": True})
    # menção de verdade: cria o rascunho na API beta com a lista "mentions" e envia
    msg["mentions"] = [{"mentioned": {"name": n, "address": e}} for n, e in mencoes]
    tk = token()
    beta = "https://graph.microsoft.com/beta"
    _, rasc = call("POST", beta + "/me/messages", msg, tk)
    return call("POST", f"{beta}/me/messages/{rasc['id']}/send", None, tk)
