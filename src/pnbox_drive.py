"""
Google Drive de cada usuário (v9.1): login com o Google, pasta "ViabilizaLab" no Drive do aluno,
salvamento automático dos projetos e envio do relatório e das planilhas para o Google Documentos
e o Google Planilhas.

Segurança
- Permissões pedidas: openid, e-mail e drive.file (o programa só enxerga os arquivos que ele criou
  ou que o usuário escolheu abrir nele). Nenhum escopo sensível.
- O login acontece no navegador (Google Identity Services, fluxo de token). O token fica só na
  memória da sessão; nunca é gravado em arquivo. O segredo do cliente não é usado.
- O ID do cliente, a chave da Picker e o número do projeto abaixo são públicos por natureza:
  o Google só os aceita vindos dos endereços cadastrados no projeto do Google Cloud.
"""
import json
import sys
import urllib.parse
import uuid

GOOGLE_CLIENT_ID = "724365524589-gf4iv0op3mkdr4umbh85asdi78ar6u50.apps.googleusercontent.com"
GOOGLE_PICKER_KEY = "AIzaSyBsfjri6BaiiHt-Lu9PnFMz4Um9gyF_WAk"
GOOGLE_APP_ID = "724365524589"
ESCOPOS = "openid https://www.googleapis.com/auth/userinfo.email https://www.googleapis.com/auth/drive.file"

DRIVE_API = "https://www.googleapis.com/drive/v3"
DRIVE_UPLOAD = "https://www.googleapis.com/upload/drive/v3"
MIME_PASTA = "application/vnd.google-apps.folder"
MIME_PNBOX = "application/octet-stream"
NOME_PASTA = "ViabilizaLab"
CONVERSOES = {
    "doc": ("application/vnd.openxmlformats-officedocument.wordprocessingml.document", "application/vnd.google-apps.document"),
    "planilha": ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.google-apps.spreadsheet"),
}


class ErroDrive(Exception):
    def __init__(self, status: int, detalhe: str = ""):
        super().__init__(f"Google Drive respondeu {status}: {detalhe[:200]}")
        self.status = status


def _drive_http(metodo: str, url: str, token: str, corpo: bytes | None = None, tipo: str | None = None) -> bytes:
    """Requisição síncrona: XMLHttpRequest no navegador (Pyodide, em web worker) ou urllib no computador."""
    cab = {"Authorization": f"Bearer {token}"}
    if tipo:
        cab["Content-Type"] = tipo
    if sys.platform == "emscripten":
        import js
        xhr = js.XMLHttpRequest.new()
        xhr.open(metodo, url, False)
        xhr.responseType = "arraybuffer"
        for k, v in cab.items():
            xhr.setRequestHeader(k, v)
        if corpo is None:
            xhr.send()
        else:
            arr = js.Uint8Array.new(len(corpo))
            arr.assign(corpo)
            xhr.send(arr)
        status = int(xhr.status)
        dados = js.Uint8Array.new(xhr.response).to_bytes() if xhr.response else b""
    else:
        import urllib.error
        import urllib.request
        req = urllib.request.Request(url, data=corpo, method=metodo, headers=cab)
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                status, dados = r.status, r.read()
        except urllib.error.HTTPError as e:
            status, dados = e.code, e.read()
    if status >= 400 or status == 0:
        raise ErroDrive(status, dados.decode("utf-8", "replace") if dados else "")
    return dados


def _drive_json(metodo, url, token, corpo=None, tipo=None) -> dict:
    d = _drive_http(metodo, url, token, corpo, tipo)
    return json.loads(d.decode("utf-8")) if d else {}


def _drive_multipart(meta: dict, conteudo: bytes, mime: str) -> tuple[bytes, str]:
    fronteira = "vl" + uuid.uuid4().hex
    corpo = (f"--{fronteira}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n{json.dumps(meta)}\r\n"
             f"--{fronteira}\r\nContent-Type: {mime}\r\n\r\n").encode("utf-8") + conteudo + f"\r\n--{fronteira}--".encode()
    return corpo, f"multipart/related; boundary={fronteira}"


def drive_usuario(token: str) -> str:
    return _drive_json("GET", "https://www.googleapis.com/oauth2/v3/userinfo", token).get("email", "")


def drive_pasta(token: str) -> str:
    """Id da pasta ViabilizaLab do usuário; cria se não existir."""
    q = urllib.parse.quote(f"name='{NOME_PASTA}' and mimeType='{MIME_PASTA}' and trashed=false")
    r = _drive_json("GET", f"{DRIVE_API}/files?q={q}&fields=files(id)&spaces=drive", token)
    if r.get("files"):
        return r["files"][0]["id"]
    meta = json.dumps({"name": NOME_PASTA, "mimeType": MIME_PASTA}).encode()
    return _drive_json("POST", f"{DRIVE_API}/files?fields=id", token, meta, "application/json; charset=UTF-8")["id"]


def drive_listar(token: str, id_pasta: str) -> list[dict]:
    q = urllib.parse.quote(f"'{id_pasta}' in parents and trashed=false and name contains '.pnbox'")
    r = _drive_json("GET", f"{DRIVE_API}/files?q={q}&orderBy=modifiedTime desc&pageSize=100"
                     "&fields=files(id,name,modifiedTime,size)", token)
    return r.get("files", [])


def drive_metadados(token: str, id_arq: str) -> dict:
    return _drive_json("GET", f"{DRIVE_API}/files/{id_arq}?fields=id,name,modifiedTime,trashed", token)


def drive_baixar(token: str, id_arq: str) -> bytes:
    return _drive_http("GET", f"{DRIVE_API}/files/{id_arq}?alt=media", token)


def drive_criar(token: str, nome: str, conteudo: bytes, id_pasta: str) -> dict:
    corpo, tipo = _drive_multipart({"name": nome, "parents": [id_pasta], "mimeType": MIME_PNBOX}, conteudo, MIME_PNBOX)
    return _drive_json("POST", f"{DRIVE_UPLOAD}/files?uploadType=multipart&fields=id,name,modifiedTime", token, corpo, tipo)


def drive_atualizar(token: str, id_arq: str, conteudo: bytes) -> dict:
    return _drive_json("PATCH", f"{DRIVE_UPLOAD}/files/{id_arq}?uploadType=media&fields=id,name,modifiedTime", token,
                 conteudo, MIME_PNBOX)


def drive_enviar_convertido(token: str, nome: str, conteudo: bytes, tipo: str, id_pasta: str) -> dict:
    """Envia .docx ou .xlsx pedindo a conversão para Google Documentos ou Google Planilhas."""
    origem, destino = CONVERSOES[tipo]
    corpo, ct = _drive_multipart({"name": nome, "parents": [id_pasta], "mimeType": destino}, conteudo, origem)
    return _drive_json("POST", f"{DRIVE_UPLOAD}/files?uploadType=multipart&fields=id,name,webViewLink", token, corpo, ct)


def drive_mensagem_erro(e: Exception) -> str:
    st_ = getattr(e, "status", None)
    if st_ == 401:
        return "A autorização do Google expirou. Clique em \"Entrar com Google\" de novo."
    if st_ == 403:
        return "O Google recusou o pedido (permissão ou limite de uso). Tente de novo em alguns minutos."
    if st_ == 404:
        return "O arquivo não foi encontrado no Drive (pode ter sido apagado ou o acesso foi retirado)."
    if st_ in (0, None):
        return "Sem conexão com o Google Drive. As alterações continuam salvas aqui e serão enviadas depois."
    return f"O Google Drive respondeu com erro {st_}. Tente de novo."


# ------------------------------------------------------------------------------
# CÓDIGO DA PÁGINA (navegador): login, Picker e canal de mensagens para o Python
# ------------------------------------------------------------------------------
# Executado no contexto da página principal (não no quadro do componente, que o Streamlit remove).
JS_GOOGLE = r"""
(function () {
  if (window.__vlGoogle) return;
  var CLIENT = "%(client)s", CHAVE = "%(chave)s", APP = "%(app)s", ESCOPOS = "%(escopos)s";
  function carregar(src) {
    return new Promise(function (ok, erro) {
      var ja = document.querySelector('script[src="' + src + '"]');
      if (ja) { if (ja.dataset.ok) return ok(); ja.addEventListener("load", ok); return; }
      var s = document.createElement("script"); s.src = src; s.async = true;
      s.onload = function () { s.dataset.ok = "1"; ok(); }; s.onerror = erro; document.head.appendChild(s);
    });
  }
  var gis = carregar("https://accounts.google.com/gsi/client");
  var cliente = null, token = null;
  function enviar(msg) {
    msg.n = Date.now() + "-" + Math.random();
    var el = document.querySelector(".st-key-gcanal input");
    if (!el) return;
    var set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value").set;
    el.focus(); set.call(el, JSON.stringify(msg));
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", code: "Enter", keyCode: 13, which: 13, bubbles: true }));
    el.blur();
  }
  function iniciar() {
    if (cliente) return cliente;
    cliente = google.accounts.oauth2.initTokenClient({
      client_id: CLIENT, scope: ESCOPOS,
      callback: function (r) {
        if (r.error) { enviar({ tipo: "erro", erro: r.error }); return; }
        token = r.access_token;
        enviar({ tipo: "token", token: r.access_token, expira: r.expires_in });
      },
      error_callback: function (e) { enviar({ tipo: "erro", erro: (e && e.type) || "janela" }); }
    });
    return cliente;
  }
  window.__vlGoogle = {
    enviar: enviar,
    entrar: function () {
      if (window.google && google.accounts && google.accounts.oauth2) { iniciar().requestAccessToken(); return; }
      gis.then(function () { iniciar().requestAccessToken(); });
    },
    sair: function () {
      gis.then(function () { try { if (token) google.accounts.oauth2.revoke(token, function () {}); } catch (e) {} token = null; enviar({ tipo: "sair" }); });
    },
    escolher: function () {
      if (!token) { enviar({ tipo: "erro", erro: "sem_login" }); return; }
      carregar("https://apis.google.com/js/api.js")
        .then(function () { return new Promise(function (ok) { gapi.load("picker", ok); }); })
        .then(function () {
          var vis = new google.picker.DocsView(google.picker.ViewId.DOCS).setMode(google.picker.DocsViewMode.LIST)
            .setQuery(".pnbox").setIncludeFolders(true);
          var comp = new google.picker.DocsView(google.picker.ViewId.DOCS).setOwnedByMe(false).setQuery(".pnbox");
          new google.picker.PickerBuilder().addView(vis).addView(comp).setOAuthToken(token).setDeveloperKey(CHAVE)
            .setAppId(APP).setLocale("pt-BR").setTitle("Escolha um projeto do ViabilizaLab (.pnbox)")
            .setCallback(function (d) {
              if (d.action === google.picker.Action.PICKED) { var f = d.docs[0]; enviar({ tipo: "picker", id: f.id, nome: f.name }); }
            }).build().setVisible(true);
        }).catch(function () { enviar({ tipo: "erro", erro: "picker" }); });
    }
  };
})();
""" % {"client": GOOGLE_CLIENT_ID, "chave": GOOGLE_PICKER_KEY, "app": GOOGLE_APP_ID, "escopos": ESCOPOS}


def drive_botao(acao: str, rotulo: str, cor: str = "#1a73e8", texto: str = "#fff") -> str:
    """Botão num quadro do componente que injeta o código acima na página principal e chama a ação.
    O clique do usuário no quadro conta como gesto também para a página, o que libera a janela do Google."""
    return ("<script>try{var w=window.parent;if(!w.__vlGoogle){new w.Function(" + json.dumps(JS_GOOGLE) + ")();}}catch(e){}</script>"
            "<button onclick='window.parent.__vlGoogle." + acao + "()' style='width:100%;padding:8px 10px;border-radius:8px;"
            f"border:1px solid {texto if cor.lower() == '#ffffff' else cor};background:{cor};color:{texto};font:600 14px system-ui,Segoe UI,sans-serif;cursor:pointer'>"
            + rotulo + "</button>")
