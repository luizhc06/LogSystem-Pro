# LogSystem Pro

> **⚠️ Projeto encerrado (julho/2026)**
>
> Este repositório está arquivado e não recebe mais atualizações.
>
> O LogSystem Pro foi um sistema desktop de gestão logística (Python +
> PyWebView) que cumpriu seu papel e chegou ao limite da arquitetura —
> principalmente por não suportar múltiplos usuários simultâneos.
>
> **Ele deu origem ao CarbonLog**, uma reconstrução completa como sistema
> web multiusuário, que é onde o desenvolvimento continua.

## Sobre o projeto (histórico)

Sistema de gestão de pedidos e logística com controle de status, risco e
prazo de envio, integração com transportadoras e plataformas de venda,
relatórios e importação de XML de NF-e.

- **Stack**: Python 3, PyWebView, SQLite
- **Período ativo**: 2025 — 2026
- Histórico completo de versões em `LogSystem_Pro_Historico_Versoes.pdf`

## Lições que levaram à reescrita

- Persistência por blob JSON com rewrite completo a cada save inviabilizou
  o uso multiusuário
- UI inteira num único HTML de 350KB dificultava manutenção
- Distribuição desktop exigia instalação máquina a máquina

O sucessor resolve esses pontos com PHP/MySQL, escrita por registro e
acesso via navegador.
