# neo-finance-app

Sistema de gestao financeira com backend em Google Apps Script (microservicos) e frontend web em HTML/JS com Bootstrap.

## Objetivo
- Centralizar contas a pagar/receber, fluxo de caixa, DRE e relatorios em uma arquitetura de microservicos.
- Fornecer um frontend responsivo para interacao com dados e configuracoes do sistema.

## Estrutura de pastas
```
neo-finance-app/
|-- backend/
|   |-- auth-service/
|   |-- config-service/
|   |-- ap-service/            # contas a pagar
|   |-- ar-service/            # contas a receber
|   |-- cashflow-service/      # fluxo de caixa
|   |-- dre-service/           # demonstracao de resultados
|   |-- reporting-service/
|   |-- userprefs-service/
|   |-- integration-service/   # integracoes externas (ex.: ERP, bancos)
|   |-- shared-libs/           # utilitarios e contratos comuns
|-- frontend/
|   |-- src/                   # codigo-fonte do app (JS/HTML)
|   |-- public/                # arquivos estaticos e index.html
|   |-- assets/                # imagens, estilos e fontes
|-- docs/                      # documentacao funcional e tecnica
|-- infra/                     # scripts de deploy/automacao
```

## Proximos passos sugeridos
- Definir contratos de API (JSON) para cada servico em `backend/`.
- Criar layout inicial do frontend em `frontend/public/index.html` e organizar componentes em `frontend/src/`.
- Documentar fluxos criticos e integracoes em `docs/`.
