#!/usr/bin/env bash
# Instalador Linux — mesma lógica do LogSystem.bat
cd "$(dirname "${BASH_SOURCE[0]}")"

echo ""
echo "  LogSystem Pro v28 — Instalador"
echo "  =============================="
echo ""

command -v python3 >/dev/null || {
    echo "  [ERRO] Python 3 nao encontrado."
    exit 1
}

if [ ! -d ".venv" ]; then
    echo "  Criando ambiente virtual..."
    python3 -m venv .venv --system-site-packages
    echo "  Instalando dependencias..."
    .venv/bin/pip install --upgrade pip -q
    .venv/bin/pip install -r requirements.txt -q
    .venv/bin/python -c "from modules import database, auth; database.configure('.'); database.inicializar_banco(); auth.load_users('data/users.json')"
    echo "  Instalacao concluida!"
    echo ""
fi

PY=".venv/bin/python"
$PY -c "import webview" 2>/dev/null || $PY -m pip install -r requirements.txt -q

echo "  Iniciando LogSystem Pro..."
exec "$PY" main.py
