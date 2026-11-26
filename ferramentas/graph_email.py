"""Envio de e-mail pelo Microsoft Graph no padrão do Daniel.

Uso como módulo:
    from graph_email import enviar
    enviar(
        para=[("Victor Morais", "victor.morais@raa.adv.br")],
        assunto="Assunto",
        mencoes=[("Victor Morais", "victor.morais@raa.adv.br")],
        paragrafos=["Primeiro parágrafo.", "<b>Rótulo:</b> texto."],
    )

Credenciais: variáveis MS_GRAPH_CLIENT_ID, MS_GRAPH_TENANT_ID e
MS_GRAPH_REFRESH_TOKEN do ambiente de nuvem. Nunca imprimir os valores.
"""
import base64, html, json, os, urllib.parse, urllib.request

ASSINATURA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assinatura.png")
FONTE = 'font-family:"Century Gothic",sans-serif'
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
    req = urllib.request.Request(
        url, method=method,
        data=json.dumps(body).encode() if body is not None else None,
        headers={"Authorization": "Bearer " + (tk or token()), "Content-Type": "application/json"},
    )
    r = urllib.request.urlopen(req)
    b = r.read()
    return r.status, (json.loads(b) if b else None)


def _p(conteudo):
    return f'<p class="MsoNormal"><span style="{FONTE}">{conteudo}</span></p>'


def _vazio():
    return _p("&nbsp;")


def mencao(nome, email):
    return (f'<a href="mailto:{email}"><span style="{FONTE}; text-decoration:none">'
            f'@{html.escape(nome)}</span></a>')


def corpo_html(paragrafos, mencoes=None):
    """paragrafos: lista de strings (HTML permitido) ou listas (viram tópicos)."""
    partes = []
    if mencoes:
        partes += [_p(", ".join(mencao(n, e) for n, e in mencoes) + ","), _vazio()]
    for item in paragrafos:
        if isinstance(item, (list, tuple)):
            for li in item:
                partes.append(f'<ul type="disc" style="margin-top:0cm"><li class="MsoListParagraph" '
                              f'style="margin-left:0cm"><span style="{FONTE}">{li}</span></li></ul>')
                partes.append(_vazio())
        else:
            partes += [_p(item), _vazio()]
    partes.append(_p("Atenciosamente,"))
    partes.append(_vazio())
    partes.append(f'<p class="MsoNormal"><img border="0" width="589" height="196" '
                  f'src="cid:{CID}" style="width:6.1354in; height:2.0416in"></p>')
    return ('<html><body lang="PT-BR"><div class="WordSection1">' + "".join(partes)
            + "</div></body></html>")


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
    return call("POST", "/me/sendMail", {"message": msg, "saveToSentItems": True})
