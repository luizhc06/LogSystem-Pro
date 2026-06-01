"""
SQLite em data/logsystem.db — cada pedido é um JSON na coluna data_json.

Migra automaticamente de data.json na primeira vez que o banco está vazio.
"""

import json
import os
import shutil
import sqlite3
from datetime import datetime

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH = os.path.join(DATA_DIR, "logsystem.db")
JSON_PATH = os.path.join(DATA_DIR, "data.json")
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
BACKUP_DIR = os.path.join(DATA_DIR, "backups")

SCHEMA_VERSION = 1


def configure(base_dir=None):
    """Aponta caminhos pro diretório da instalação e move arquivos legados p/ data/."""
    global BASE_DIR, DATA_DIR, DB_PATH, JSON_PATH, CONFIG_PATH, BACKUP_DIR
    if base_dir:
        BASE_DIR = os.path.abspath(base_dir)
    DATA_DIR = os.path.join(BASE_DIR, "data")
    DB_PATH = os.path.join(DATA_DIR, "logsystem.db")
    JSON_PATH = os.path.join(DATA_DIR, "data.json")
    CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
    BACKUP_DIR = os.path.join(DATA_DIR, "backups")
    _migrate_legacy_layout()


def _migrate_legacy_layout():
    """Versões antigas guardavam data.json na raiz — move tudo pra data/."""
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(BACKUP_DIR, exist_ok=True)

    for name in ("logsystem.db", "data.json", "users.json"):
        old = os.path.join(BASE_DIR, name)
        new = os.path.join(DATA_DIR, name)
        if os.path.exists(old) and not os.path.exists(new):
            shutil.move(old, new)

    old_backups = os.path.join(BASE_DIR, "backups")
    if os.path.isdir(old_backups):
        for item in os.listdir(old_backups):
            src = os.path.join(old_backups, item)
            dst = os.path.join(BACKUP_DIR, item)
            if not os.path.exists(dst):
                shutil.move(src, dst)
        try:
            os.rmdir(old_backups)
        except OSError:
            pass


def conectar(timeout=30):
    """Abre conexão SQLite com WAL (menos trava entre leitura/escrita)."""
    os.makedirs(DATA_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, timeout=timeout)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def verificar_conexao():
    """Verifica se o banco SQLite está acessível. Retorna (ok, mensagem)."""
    try:
        conn = conectar()
        conn.execute("SELECT 1")
        conn.close()
        return True, f"Banco SQLite OK: {DB_PATH}"
    except sqlite3.OperationalError as exc:
        return False, (
            f"Erro ao acessar o banco SQLite em '{DB_PATH}'. "
            f"Verifique permissões de escrita no diretório. Detalhe: {exc}"
        )
    except Exception as exc:
        return False, f"Falha na conexão com o banco de dados: {exc}"


def _init_schema(conn):
    cursor = conn.cursor()
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS meta (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
        """
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS orders (
            id TEXT PRIMARY KEY,
            nf TEXT,
            chav_nfe TEXT,
            data_json TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_orders_nf ON orders(nf)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_orders_chav ON orders(chav_nfe)")
    cursor.execute(
        "INSERT OR IGNORE INTO meta (key, value) VALUES (?, ?)",
        ("schema_version", str(SCHEMA_VERSION)),
    )
    conn.commit()


def _carregar_json(path):
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, list):
        raise ValueError(f"Formato inválido em {path}: esperado array JSON")
    return data


def _migrar_de_json(conn):
    """Importa data.json quando o banco ainda não possui pedidos."""
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM orders")
    if cursor.fetchone()[0] > 0:
        return 0
    if not os.path.exists(JSON_PATH):
        return 0

    orders = _carregar_json(JSON_PATH)
    if not orders:
        return 0

    now = datetime.now().isoformat()
    for order in orders:
        oid = str(order.get("id") or "").strip()
        if not oid:
            continue
        nf = str(order.get("nf") or "")
        chav = str(order.get("chavNfe") or "")
        cursor.execute(
            """
            INSERT OR REPLACE INTO orders (id, nf, chav_nfe, data_json, updated_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            (oid, nf, chav, json.dumps(order, ensure_ascii=False), now),
        )

    conn.commit()
    os.makedirs(BACKUP_DIR, exist_ok=True)
    backup_name = f"data.json.migrated_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    shutil.copy2(JSON_PATH, os.path.join(BACKUP_DIR, backup_name))
    return len(orders)


def inicializar_banco():
    """Cria estruturas necessárias e migra data.json na primeira execução."""
    ok, msg = verificar_conexao()
    if not ok:
        raise RuntimeError(msg)

    os.makedirs(BACKUP_DIR, exist_ok=True)
    conn = conectar()
    try:
        _init_schema(conn)
        return _migrar_de_json(conn)
    finally:
        conn.close()


# --- CRUD de pedidos ---

def listar_pedidos():
    """Retorna todos os pedidos do banco (lista de dicts)."""
    inicializar_banco()
    conn = conectar()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT data_json FROM orders ORDER BY id")
        orders = []
        for row in cursor.fetchall():
            try:
                orders.append(json.loads(row[0]))
            except json.JSONDecodeError:
                continue

        if not orders and os.path.exists(JSON_PATH):
            orders = _carregar_json(JSON_PATH)
            if orders:
                salvar_pedidos(orders)
        return orders
    finally:
        conn.close()


def salvar_pedidos(orders):
    """Persiste todos os pedidos no SQLite e sincroniza data.json."""
    if not isinstance(orders, list):
        raise ValueError("orders deve ser uma lista")

    inicializar_banco()
    conn = conectar()
    now = datetime.now().isoformat()
    try:
        cursor = conn.cursor()
        cursor.execute("BEGIN")
        cursor.execute("DELETE FROM orders")
        for order in orders:
            oid = str(order.get("id") or "").strip()
            if not oid:
                continue
            cursor.execute(
                """
                INSERT INTO orders (id, nf, chav_nfe, data_json, updated_at)
                VALUES (?, ?, ?, ?, ?)
                """,
                (
                    oid,
                    str(order.get("nf") or ""),
                    str(order.get("chavNfe") or ""),
                    json.dumps(order, ensure_ascii=False),
                    now,
                ),
            )
        conn.commit()
        _sync_json_backup(orders)
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def _sync_json_backup(orders):
    tmp_path = JSON_PATH + ".tmp"
    try:
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(orders, f, ensure_ascii=False, indent=2)
        os.replace(tmp_path, JSON_PATH)
    except Exception as exc:
        print(f"Aviso: falha ao sincronizar data.json: {exc}")
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass


# Snapshot completo (pedidos + config) → data/backups/backup_*.json
def backup_para_json():
    """Gera backup completo (pedidos + config) em backups/."""
    orders = listar_pedidos()
    cfg = {}
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        except Exception:
            cfg = {}

    os.makedirs(BACKUP_DIR, exist_ok=True)
    nome = f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    caminho = os.path.join(BACKUP_DIR, nome)
    payload = {
        "orders": orders,
        "config": cfg,
        "exported_at": datetime.now().isoformat(),
    }
    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    return caminho


def restaurar_de_json(caminho_json):
    """Restaura pedidos de um arquivo JSON de backup."""
    with open(caminho_json, "r", encoding="utf-8") as f:
        raw = json.load(f)

    if isinstance(raw, list):
        orders = raw
    elif isinstance(raw, dict):
        orders = raw.get("orders", [])
    else:
        raise ValueError("Formato de backup inválido")

    if not isinstance(orders, list):
        raise ValueError("Backup não contém lista de pedidos")

    salvar_pedidos(orders)
    return len(orders)
