"""
Checa dependências na 1ª execução e instala via pip se faltar algo.
"""

from __future__ import annotations

import importlib
import subprocess
import sys

REQUIRED_PACKAGES = [
    ("webview", "pywebview"),
    ("openpyxl", "openpyxl"),
    ("docx", "python-docx"),
]

OPTIONAL_PACKAGES = [
    ("PyQt6", "PyQt6"),
    ("PyQt6.QtWebEngineWidgets", "PyQt6-WebEngine"),
]


def _missing_packages():
    missing = []
    for module_name, pip_name in REQUIRED_PACKAGES:
        try:
            importlib.import_module(module_name)
        except ImportError:
            missing.append(pip_name)
    return missing


def ensure_dependencies(auto_install=True):
    """
    Verifica dependências obrigatórias.
    Retorna (ok, mensagem).
    """
    missing = _missing_packages()
    if not missing:
        return True, "Dependências OK"

    if not auto_install:
        return False, (
            "Dependências ausentes: "
            + ", ".join(missing)
            + f". Instale com: {sys.executable} -m pip install "
            + " ".join(missing)
        )

    print(f"[LogSystem] Instalando dependências: {', '.join(missing)}")
    cmd = [sys.executable, "-m", "pip", "install", "--upgrade", *missing]
    try:
        subprocess.check_call(cmd)
    except subprocess.CalledProcessError as exc:
        return False, f"Falha ao instalar dependências: {exc}"

    still_missing = _missing_packages()
    if still_missing:
        return False, f"Ainda faltam dependências: {', '.join(still_missing)}"

    return True, "Dependências instaladas com sucesso"


def check_optional_backend():
    """Informa se há backend gráfico disponível para pywebview."""
    for module_name, _ in OPTIONAL_PACKAGES:
        try:
            importlib.import_module(module_name)
            return True, module_name
        except ImportError:
            continue
    return False, "Nenhum backend Qt/WebEngine detectado (pywebview usará fallback do sistema)"
