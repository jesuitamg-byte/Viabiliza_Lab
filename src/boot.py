
# ==============================================================================
# INICIALIZAÇÃO AUTOMÁTICA (executada antes de tudo)
# ==============================================================================
import importlib.util as _ilu
import subprocess as _sp
import sys as _sys
from pathlib import Path as _Path


def _no_streamlit() -> bool:
    """True quando o arquivo já está rodando dentro do Streamlit."""
    try:
        from streamlit.runtime import exists
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return exists() or get_script_run_ctx(suppress_warning=True) is not None
    except Exception:
        return False


def _preparar_ambiente() -> None:
    pacotes = {"streamlit": "streamlit", "pandas": "pandas", "openpyxl": "openpyxl",
               "matplotlib": "matplotlib", "altair": "altair", "docx": "python-docx", "pptx": "python-pptx"}
    faltando = [pip for mod, pip in pacotes.items() if _ilu.find_spec(mod) is None]
    try:
        import fpdf
        if int(fpdf.__version__.split(".")[0]) < 2 or tuple(map(int, fpdf.__version__.split(".")[:2])) < (2, 7):
            faltando.append("fpdf2")
    except Exception:
        faltando.append("fpdf2")
    if faltando:
        print("Instalando bibliotecas necessárias:", ", ".join(faltando))
        _sp.check_call([_sys.executable, "-m", "pip", "install", "-U", *faltando])
    # Evita a pergunta "Email:" do Streamlit na primeira execução
    cred = _Path.home() / ".streamlit" / "credentials.toml"
    if not cred.exists():
        cred.parent.mkdir(parents=True, exist_ok=True)
        cred.write_text('[general]\nemail = ""\n', encoding="utf-8")


if __name__ == "__main__" and not _no_streamlit():
    _preparar_ambiente()
    import os as _os
    import socket as _socket
    # Porta fixa 8765: o login do Google só aceita endereços cadastrados (http://localhost:8765).
    # Se ela estiver ocupada por um ViabilizaLab já aberto, só abre o navegador nele; se estiver
    # ocupada por outro programa, usa uma porta livre (o programa funciona, mas sem o Google Drive).
    _porta = 8765
    try:
        with _socket.socket() as _s:
            _s.bind(("127.0.0.1", _porta))
    except OSError:
        try:
            import urllib.request as _ur
            _ur.urlopen(f"http://localhost:{_porta}/_stcore/health", timeout=2)
            import webbrowser as _wb
            print(f"O ViabilizaLab já está aberto. Abrindo http://localhost:{_porta}")
            _wb.open(f"http://localhost:{_porta}")
            _sys.exit(0)
        except Exception:
            with _socket.socket() as _s:
                _s.bind(("127.0.0.1", 0))
                _porta = _s.getsockname()[1]
            print("Aviso: a porta 8765 está ocupada por outro programa; o login do Google Drive não vai funcionar.")
    _os.environ["PNBOX_AUTOFECHAR"] = "1"
    print(f"Abrindo o ViabilizaLab em http://localhost:{_porta}  (feche a aba do navegador para encerrar)")
    from streamlit.web import cli as _stcli
    _sys.argv = ["streamlit", "run", str(_Path(__file__).resolve()),
                 f"--server.port={_porta}", "--server.address=localhost",
                 "--browser.gatherUsageStats=false", "--server.headless=false"]
    _sys.exit(_stcli.main())
