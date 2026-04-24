import sys
import os
from streamlit.web import cli as stcli

def resolve_path(path):
    # Função para o executável achar os arquivos dentro dele mesmo
    if getattr(sys, "frozen", False):
        basedir = sys._MEIPASS
    else:
        basedir = os.path.dirname(__file__)
    return os.path.join(basedir, path)

if __name__ == "__main__":
    # Aponta para o seu arquivo principal
    sys.argv = [
        "streamlit",
        "run",
        resolve_path("dashboard.py"),
        "--global.developmentMode=false"
        "--browser.gatherUsageStats=false",
        "--server.headless=true"
    ]
    sys.exit(stcli.main())