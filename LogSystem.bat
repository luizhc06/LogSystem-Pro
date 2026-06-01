@echo off
:: Instalador Windows — cria .venv, instala deps, abre main.py
title LogSystem Pro v28 - Instalador
cd /d "%~dp0"

echo.
echo   LogSystem Pro v28
echo   Instalador Windows
echo   ==================
echo.

set FIRST=0
if not exist ".venv\Scripts\python.exe" set FIRST=1

set SYS_PY=
where py >nul 2>&1 && set SYS_PY=py -3
if not defined SYS_PY where python >nul 2>&1 && set SYS_PY=python

if not defined SYS_PY (
    echo   [ERRO] Python 3 nao encontrado.
    echo   Instale em https://www.python.org/downloads/
    echo   Marque "Add Python to PATH" na instalacao.
    pause
    exit /b 1
)

if %FIRST%==1 (
    echo   Criando ambiente virtual...
    %SYS_PY% -m venv .venv
    if errorlevel 1 (
        echo   [ERRO] Falha ao criar ambiente virtual.
        pause
        exit /b 1
    )
    echo   Instalando dependencias...
    ".venv\Scripts\python.exe" -m pip install --upgrade pip -q
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt -q
    ".venv\Scripts\python.exe" -c "from modules import database, auth; database.configure('.'); database.inicializar_banco(); auth.load_users('data/users.json')"
    echo   Instalacao concluida!
    echo.
) else (
    ".venv\Scripts\python.exe" -c "import webview" 2>nul
    if errorlevel 1 (
        echo   Atualizando dependencias...
        ".venv\Scripts\python.exe" -m pip install -r requirements.txt -q
    )
)

echo   Iniciando LogSystem Pro...
if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" main.py
) else (
    start "" /b ".venv\Scripts\python.exe" main.py
)
