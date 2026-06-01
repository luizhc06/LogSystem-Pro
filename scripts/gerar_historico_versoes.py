#!/usr/bin/env python3
"""Gera LogSystem_Pro_Historico_Versoes.pdf — documento formal de histórico de versões."""

import os
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.platypus.doctemplate import BaseDocTemplate, PageTemplate
from reportlab.platypus.frames import Frame

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
OUTPUT = os.path.join(ROOT, "LogSystem_Pro_Historico_Versoes.pdf")

# Tipografia: duas famílias — Times (corpo) e Helvetica (títulos e ênfases estruturais)
FONT_BODY = "Times-Roman"
FONT_BODY_BOLD = "Times-Bold"
FONT_HEADING = "Helvetica-Bold"
FONT_HEADING_LIGHT = "Helvetica"

FOOTER = "LogSystem Pro — Histórico de Versões"
AUTHOR = "Luiz Henrique · github.com/luizhc06 · 2026"


def footer(canvas, doc):
    canvas.saveState()
    canvas.setFont(FONT_HEADING_LIGHT, 8)
    canvas.setFillColor(colors.HexColor("#444444"))
    canvas.drawString(2 * cm, 1.2 * cm, f"{FOOTER}  |  Pág. {canvas.getPageNumber()}")
    canvas.drawRightString(A4[0] - 2 * cm, 1.2 * cm, AUTHOR)
    canvas.restoreState()


def build_styles():
    base = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "title",
            parent=base["Title"],
            fontName=FONT_HEADING,
            fontSize=20,
            leading=24,
            spaceAfter=8,
            textColor=colors.HexColor("#1a1a1a"),
        ),
        "subtitle": ParagraphStyle(
            "subtitle",
            parent=base["Normal"],
            fontName=FONT_BODY,
            fontSize=12,
            leading=16,
            textColor=colors.HexColor("#333333"),
            spaceAfter=6,
        ),
        "meta": ParagraphStyle(
            "meta",
            parent=base["Normal"],
            fontName=FONT_BODY,
            fontSize=10,
            leading=14,
            textColor=colors.HexColor("#555555"),
            spaceAfter=4,
        ),
        "h2": ParagraphStyle(
            "h2",
            parent=base["Heading2"],
            fontName=FONT_HEADING,
            fontSize=12,
            leading=15,
            spaceBefore=16,
            spaceAfter=8,
            textColor=colors.HexColor("#1a1a1a"),
        ),
        "h3": ParagraphStyle(
            "h3",
            parent=base["Heading3"],
            fontName=FONT_BODY_BOLD,
            fontSize=10.5,
            leading=14,
            spaceBefore=10,
            spaceAfter=4,
            textColor=colors.HexColor("#1a1a1a"),
        ),
        "body": ParagraphStyle(
            "body",
            parent=base["Normal"],
            fontName=FONT_BODY,
            fontSize=10,
            leading=14,
            spaceAfter=4,
            alignment=4,  # JUSTIFY
        ),
        "bullet": ParagraphStyle(
            "bullet",
            parent=base["Normal"],
            fontName=FONT_BODY,
            fontSize=10,
            leading=14,
            leftIndent=16,
            bulletIndent=0,
            spaceAfter=3,
            alignment=4,
        ),
    }


def bullet(text, styles):
    return Paragraph(f"• {text}", styles["bullet"])


def version_block(title, items, styles):
    parts = [Paragraph(title, styles["h3"])]
    for item in items:
        parts.append(bullet(item, styles))
    parts.append(Spacer(1, 4))
    return parts


def main():
    styles = build_styles()
    story = []
    now = datetime.now().strftime("%d/%m/%Y")

    # Capa
    story.append(Spacer(1, 3 * cm))
    story.append(Paragraph("LogSystem Pro", styles["title"]))
    story.append(
        Paragraph(
            "Documento de Histórico de Versões e Alterações Técnicas",
            styles["subtitle"],
        )
    )
    story.append(Spacer(1, 0.5 * cm))
    story.append(
        Paragraph(
            "Sistema Desktop de Gestão Logística e Rastreamento de Pedidos",
            styles["meta"],
        )
    )
    story.append(Spacer(1, 0.8 * cm))
    story.append(Paragraph("Autor: Luiz Henrique", styles["meta"]))
    story.append(Paragraph("Repositório: github.com/luizhc06", styles["meta"]))
    story.append(Spacer(1, 0.4 * cm))
    story.append(Paragraph(f"Data de emissão: {now}", styles["meta"]))
    story.append(Spacer(1, 1 * cm))
    story.append(
        Paragraph(
            "<b>Versão vigente:</b> v28 (Maio/2026) — consolidação de persistência de dados, "
            "instalação multiplataforma, segurança de acesso e estabilidade operacional.",
            styles["body"],
        )
    )

    # Visão geral
    story.append(Paragraph("1. Visão Geral", styles["h2"]))
    story.append(
        Paragraph(
            "O LogSystem Pro é uma aplicação desktop para gestão logística, desenvolvida "
            "de forma iterativa ao longo de aproximadamente três meses. O presente documento "
            "registra vinte e nove versões documentadas (v10.4 a v28), descrevendo evolução "
            "funcional, correções de defeitos e melhorias de infraestrutura identificadas "
            "durante o ciclo de desenvolvimento e validação do sistema.",
            styles["body"],
        )
    )
    story.append(Spacer(1, 8))
    timeline = [
        ["Período", "Versões", "Escopo principal"],
        ["Março/2026", "v10.4 → v11.9", "Fundação, correções críticas, SLA, painel e alertas"],
        ["Março/Abril/2026", "v12.0 → v12.6", "Temas visuais, identificadores, retorno e status"],
        ["Abril/Maio/2026", "v12.7 → v12.15", "Visão SLA, exportação, cards, reenvios, CPF/CEP"],
        ["Maio/2026", "v28", "SQLite, instalação, XML NF-e, autenticação PBKDF2, temas"],
    ]
    t = Table(timeline, colWidths=[3.2 * cm, 3.5 * cm, 9.5 * cm])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2c2c2c")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), FONT_HEADING),
                ("FONTNAME", (0, 1), (-1, -1), FONT_BODY),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f5f5f5")]),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cccccc")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(t)
    story.append(Spacer(1, 10))
    story.append(
        Paragraph(
            "<b>Stack tecnológico:</b> Python 3, pywebview, HTML/CSS/JavaScript, SQLite, "
            "openpyxl, python-docx e JSON (backup e migração).",
            styles["body"],
        )
    )

    # v10
    story.append(Paragraph("2. Versões v10 — Fundação do Sistema", styles["h2"]))
    for title, items in [
        (
            "v10.4 · Março/2026 · Estrutura inicial",
            [
                "Normalização de números de nota fiscal mediante expressão regular /^NF-?/i.",
                "Deduplicação cruzada por chave NF-e de quarenta e quatro dígitos.",
                "Preservação de registros manuais e importados via XML durante importação de planilha.",
            ],
        ),
        (
            "v10.5 · Março/2026 · Correção do painel principal",
            [
                "Correção crítica: indicadores carrierStats/ufStats retornavam vazio após remoção do seletor mensal.",
                "Reestruturação de estilos CSS com propriedades customizadas para suporte a temas.",
            ],
        ),
        (
            "v10.7 · Março/2026 · Estabilização de inicialização",
            [
                "Correção de falha em init() causada por referência a elemento #monthBtn removido.",
                "Correção em openOrder() sem invocação de calcOrder(), resultando em status final incorreto.",
            ],
        ),
    ]:
        story.extend(version_block(title, items, styles))

    story.append(Paragraph("3. Versões v11 — Funcionalidades Operacionais", styles["h2"]))
    for title, items in [
        (
            "v11.2 · Março/2026 · Cálculo de trânsito",
            [
                "Parser de datas DD/MM/AAAA corrigido (evita interpretação MM/DD).",
                "Recálculo de diasEntregaReal mesmo na presença de valor pré-existente.",
            ],
        ),
        (
            "v11.4 · Março/2026 · Painel analítico",
            [
                "Rankings por estado, cidade e desempenho por transportadora.",
                "Funções fmtDate()/nowStr() independentes de toLocaleDateString (compatibilidade Linux/pywebview).",
            ],
        ),
        (
            "v11.9 · Março/2026 · Indicadores em tempo real",
            [
                "Campo atrasadosAbertos — mapa não exibe alertas por histórico já resolvido.",
            ],
        ),
    ]:
        story.extend(version_block(title, items, styles))

    story.append(Paragraph("4. Versões v12 — Refinamento e Consolidação", styles["h2"]))
    for title, items in [
        (
            "v12.0 · Março/2026 · Personalização visual",
            [
                "Temas configuráveis via propriedades CSS; seletor disponível em Configurações.",
                "Edição inline de observações e comentários de chamados.",
            ],
        ),
        (
            "v12.1 · Março/2026 · Identificadores e instaladores",
            [
                "Função safe_str() corrige conversão indevida de identificadores numéricos para float.",
                "Scripts iniciais de instalação: install.sh, LogSystem.sh e LogSystem.bat.",
            ],
        ),
        (
            "v12.8 · Abril/2026 · Identificador de pedido externo",
            [
                "Regra consolidada: utilizar compra/xPed (identificador completo) quando disponível na NF-e.",
            ],
        ),
        (
            "v12.10 · Abril/2026 · Exportação estruturada",
            [
                "Planilha exportada com três abas: Pedidos formatados, Por Transportadora e Atrasos e Riscos.",
            ],
        ),
        (
            "v12.12 · Maio/2026 · Visualização e status de retorno",
            [
                "Alternância entre visualização em tabela e cards.",
                "Status retorno_entregue para fluxo de devolução pós-entrega.",
            ],
        ),
        (
            "v12.15 · Maio/2026 · Campos operacionais",
            [
                "Campo 'Data de Envio' renomeado para 'Data de Postagem'.",
                "Formatação fmtCpf() para CPF/CNPJ; dataContato como texto DD/MM/AAAA HH:MM.",
            ],
        ),
    ]:
        story.extend(version_block(title, items, styles))

    # v28
    story.append(
        Paragraph("5. Versão v28 — Estabilização e Versão Vigente", styles["h2"])
    )
    story.append(
        Paragraph(
            "A versão v28 consolida correções críticas de persistência, instalação "
            "multiplataforma, segurança de autenticação e melhorias operacionais "
            "identificadas em auditoria técnica integral do sistema.",
            styles["body"],
        )
    )
    story.append(Spacer(1, 6))

    v28_sections = [
        (
            "5.1 Banco de dados SQLite",
            [
                "Correção crítica: load_data/save_data acessava tabela 'logs' vazia em vez dos pedidos — "
                "o sistema iniciava sem dados e corrompia salvamentos subsequentes.",
                "Implementação de SQLite com tabela orders (JSON) e migração automática de data.json na primeira execução.",
                "Verificação de conexão e criação de tabelas na inicialização.",
                "Organização de dados em pasta data/ (logsystem.db, backups/, users.json).",
                "Rotina de backup completo incluindo pedidos e configurações.",
            ],
        ),
        (
            "5.2 Instalação Windows e Linux",
            [
                "LogSystem.bat: instalador para Windows (ambiente virtual, dependências e execução do sistema).",
                "LogSystem.sh: instalador equivalente para Linux.",
                "main.py: ponto de entrada da aplicação, invocado automaticamente pelo instalador.",
                "requirements.txt e instalação automática de dependências ausentes (bootstrap).",
                "Build de executável via PyInstaller (scripts/build_exe.bat) → release/LogSystemPro.exe.",
                "Estrutura de diretórios: modules/, data/, scripts/, release/.",
            ],
        ),
        (
            "5.3 Importação de XML NF-e",
            [
                "Suporte a XML autorizado (nfeProc / infProt/chNFe).",
                "Extração de cidade e UF do destinatário (enderDest), não do emitente.",
                "Leitura de volumes e peso em transp/vol.",
                "Extração de chave NF-e de múltiplas fontes (chNFe, infProt, atributo Id).",
                "Campos nNF, vNF e dhEmi obtidos de ide/total dentro de infNFe.",
            ],
        ),
        (
            "5.4 Autenticação e segurança",
            [
                "Login obrigatório — remoção de bypass em modo de desenvolvimento.",
                "Senha protegida com PBKDF2-SHA256 (260.000 iterações e salt único).",
                "Usuário administrador padrão: luizhcastro06@gmail.com.",
                "Armazenamento de credenciais em data/users.json.",
                "Interface de login redesenhada (tema escuro corporativo).",
            ],
        ),
        (
            "5.5 Backup e restauração",
            [
                "Modal Backup/Restaurar com zona de arrastar e soltar.",
                "Restauração mediante arquivo .json de backup de qualquer diretório.",
                "API restore_backup_content — mescla pedidos, chamados e histórico.",
            ],
        ),
        (
            "5.6 Temas visuais",
            [
                "Tema Escuro (padrão): alto contraste, adequado a uso prolongado.",
                "Tema Azul: interface clara e acessível.",
                "Temas anteriores removidos em favor de consistência visual.",
            ],
        ),
        (
            "5.7 Demais correções",
            [
                "Inclusão de import re — importação XLSX falhava com NameError.",
                "Correção de DATA_DIR indefinido na geração de relatório DOCX.",
                "Caminhos compatíveis com executável PyInstaller (sys.frozen).",
                "Numeração de versão unificada (v28) em toda a interface.",
            ],
        ),
    ]
    for title, items in v28_sections:
        story.append(Paragraph(title, styles["h3"]))
        for item in items:
            story.append(bullet(item, styles))
        story.append(Spacer(1, 4))

    # Sumário executivo
    story.append(
        Paragraph("6. Sumário Executivo — Capacidades da Versão v28", styles["h2"])
    )
    story.append(Paragraph("<b>6.1 Gestão de pedidos</b>", styles["body"]))
    for item in [
        "Listagem paginada, filtros rápidos e avançados, visualização em tabela ou cards.",
        "Alertas visuais por urgência, independentes por coluna de prazo.",
        "Persistência em SQLite com backup automático a cada cinco minutos.",
    ]:
        story.append(bullet(item, styles))

    story.append(Paragraph("<b>6.2 Importação de dados</b>", styles["body"]))
    for item in [
        "XML NF-e (SEFAZ) com parser tolerante a múltiplos formatos e layouts.",
        "Planilha XLSX com detecção automática de colunas e deduplicação.",
        "Suporte a identificador externo completo (compra/xPed) quando presente na nota fiscal.",
    ]:
        story.append(bullet(item, styles))

    story.append(Paragraph("<b>6.3 Infraestrutura</b>", styles["body"]))
    for item in [
        "Instaladores LogSystem.bat (Windows) e LogSystem.sh (Linux).",
        "Autenticação PBKDF2-SHA256 com sessão por token.",
        "Backup automático, manual e restauração por arquivo .json.",
        "Dois temas visuais: Escuro (padrão) e Azul.",
        "Executável Windows opcional: release/LogSystemPro.exe.",
    ]:
        story.append(bullet(item, styles))

    story.append(Spacer(1, 1.5 * cm))
    story.append(
        Paragraph(
            "<b>LogSystem Pro v28</b> — vinte e nove versões documentadas. "
            "Luiz Henrique · github.com/luizhc06 · 2026.",
            styles["meta"],
        )
    )

    def on_page(canvas, doc_obj):
        footer(canvas, doc_obj)

    frame = Frame(2 * cm, 1.8 * cm, A4[0] - 4 * cm, A4[1] - 3.8 * cm, id="normal")
    template = PageTemplate(id="main", frames=[frame], onPage=on_page)

    class NumberedDoc(BaseDocTemplate):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self.addPageTemplates([template])

    ndoc = NumberedDoc(
        OUTPUT,
        pagesize=A4,
        leftMargin=2 * cm,
        rightMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=1.8 * cm,
        title="LogSystem Pro — Histórico de Versões",
        author="Luiz Henrique",
    )
    ndoc.build(story)
    print(f"PDF gerado: {OUTPUT}")


if __name__ == "__main__":
    main()
