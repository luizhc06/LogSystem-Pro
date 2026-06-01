@echo off
:: Gera LogSystemPro.exe em release\
cd /d "%~dp0\.."

set PY=
if exist ".venv\Scripts\python.exe" set PY=.venv\Scripts\python.exe
if not defined PY set PY=python

echo Instalando PyInstaller...
"%PY%" -m pip install pyinstaller -q

if not exist "release" mkdir release

echo Gerando executavel...
"%PY%" -m PyInstaller --noconfirm --distpath release --workpath build --specpath scripts scripts\LogSystem.spec

if exist "release\LogSystemPro.exe" (
    echo.
    echo Build concluido: release\LogSystemPro.exe
) else (
    echo [ERRO] Build falhou.
    exit /b 1
)
