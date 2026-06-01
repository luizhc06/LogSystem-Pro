# LogSystem Pro v28

Sistema desktop para acompanhar entregas: pedidos, status logístico, chamados, SLA, mapas e relatórios.  
Roda localmente no Windows ou Linux — interface web dentro de uma janela nativa (PyWebView).

---

## Como funciona (visão geral)

```
LogSystem.bat / LogSystem.sh     →  instala deps e abre o app
         │
         ▼
    main.py                        →  backend Python (API + arquivos + banco)
         │
         ├── app.html               →  toda a interface (HTML/CSS/JS)
         ├── config.json            →  listas editáveis (transportadoras, etc.)
         └── data/                  →  tudo que é gerado na máquina do usuário
```

O JavaScript chama métodos do Python via `window.pywebview.api.*`.  
Exemplo: salvar pedidos → `api.save_orders()` → grava no SQLite em `data/`.

---

## Estrutura de pastas e arquivos

### Raiz do projeto

| Arquivo | O que é |
|---------|---------|
| **`main.py`** | Ponto de entrada. Abre a janela, expõe a classe `API` pro front-end, importa planilhas, gera relatórios e gerencia backup. |
| **`app.html`** | Interface completa: layout, estilos, lógica de telas (dashboard, pedidos, mapas, chamados, SLA). Arquivo único — não tem build separado. |
| **`config.json`** | Preferências compartilháveis: transportadoras, plataformas, tamanho de página. **Pode ir pro Git.** |
| **`requirements.txt`** | Dependências Python (`pywebview`, `openpyxl`, `python-docx`, PyQt6). |
| **`LogSystem.bat`** | Instalador Windows: cria `.venv`, instala deps, inicializa banco e abre o sistema. |
| **`LogSystem.sh`** | Mesma coisa no Linux. |
| **`LEIA-ME.txt`** | Instruções rápidas para quem baixa o pacote sem abrir o README. |
| **`.gitignore`** | Impede subir dados sensíveis (ver seção abaixo). |

### `modules/` — código Python modular

| Arquivo | Responsabilidade |
|---------|------------------|
| **`auth.py`** | Login, hash de senha (PBKDF2), leitura/gravação de `users.json`. |
| **`database.py`** | SQLite (`logsystem.db`): pedidos serializados em JSON por linha, migração automática do antigo `data.json`. |
| **`bootstrap.py`** | Checa e instala dependências faltantes na primeira execução. |

### `scripts/` — ferramentas auxiliares (não rodam no dia a dia)

| Arquivo | Para quê |
|---------|----------|
| **`build_exe.bat`** | Gera executável com PyInstaller. |
| **`LogSystem.spec`** | Config do PyInstaller. |
| **`gerar_historico_versoes.py`** | Gera o PDF de histórico de versões. |

### `data/` — **NÃO versionar** (local de cada instalação)

| Caminho | Conteúdo |
|---------|----------|
| **`data/data.json`** | Cópia legada / fallback; pedidos reais ficam no SQLite. |
| **`data/logsystem.db`** | Banco SQLite com todos os pedidos. |
| **`data/users.json`** | E-mails e **hashes** de senha dos operadores. |
| **`data/.session`** | Token de sessão (se usado em disco). |
| **`data/backups/`** | Snapshots automáticos a cada ~5 min + backups manuais. |

> **Importante:** nunca faça `git add data/` nem envie backups/senhas pro GitHub. O `.gitignore` já bloqueia isso.

### `.venv/` — ambiente Python

Criado automaticamente pelo instalador. Também ignorado pelo Git.

---

## Telas do sistema (`app.html`)

| View (`S.view`) | Função |
|-----------------|--------|
| `dashboard` | KPIs, alertas, resumo do mês. |
| `mapa` | Mapa do Brasil por volume de pedidos por UF. |
| `transMapa` | Distribuição por transportadora e estado. |
| `pedidos` | Lista/tabela com filtros, paginação e modal de edição. |
| `analiticos` | Gráficos e exportação analítica. |
| `chamados` | Tickets abertos por pedido (atraso, extravio, etc.). |
| `sla` | Acompanhamento mensal de cumprimento de prazo. |

Ações na sidebar (importar XLSX/CSV/XML, backup, deduplicar NF, relatório DOCX) chamam métodos da classe `API` em `main.py`.

---

## Requisitos

- Python **3.10+** com `pip`
- Windows 10+ ou Linux com `python3` no PATH

No Windows: [python.org/downloads](https://www.python.org/downloads/) — marque **"Add Python to PATH"**.

---

## Instalação

### Windows
Duplo-clique em **`LogSystem.bat`**.

### Linux
```bash
chmod +x LogSystem.sh
./LogSystem.sh
```

### Desenvolvimento manual
```bash
python -m venv .venv
# Windows:  .venv\Scripts\activate
# Linux:    source .venv/bin/activate
pip install -r requirements.txt
python -c "from modules import database, auth; database.configure('.'); database.inicializar_banco(); auth.load_users('data/users.json')"
python main.py
```

---

## Login

Use e-mail e senha cadastrados em `data/users.json` (arquivo local, criado na 1ª execução).  
O administrador pode alterar senhas pela tela de configurações.

---

## Mapas de status (planilha ↔ sistema)

O backend traduz rótulos de planilha para IDs internos:

| Planilha (exemplos) | ID interno |
|---------------------|------------|
| NO PRAZO, EM TRÂNSITO | `no_prazo`, `em_transito` |
| ATRASADO | `em_atraso` |
| ENTREGUE NO PRAZO | `entregue_prazo` |
| EXTRAVIADO | `extraviado` |
| RETORNO | `retorno` |

Definições completas: dicionários `STATUS_*` no topo de `main.py` e objeto `STATUS` em `app.html`.

---

## Dependências

| Pacote | Uso |
|--------|-----|
| pywebview | Janela desktop embutindo o HTML |
| openpyxl | Import/export Excel |
| python-docx | Relatório de chamados em Word |
| PyQt6 + WebEngine | Backend gráfico no Linux |

---

## Versão

**v28** — histórico detalhado no PDF `LogSystem_Pro_Historico_Versoes.pdf` (opcional, gerado por `scripts/gerar_historico_versoes.py`).

**Autor:** [Luiz Henrique](https://github.com/luizhc06)
