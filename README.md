# Neoformula Finance App

Aplicativo de gestão financeira e contábil desenvolvido em Google Apps Script, usando Google Sheets como base de dados e TypeScript para desenvolvimento.

## Arquitetura

O projeto segue uma arquitetura modular de microserviços, com separação clara de responsabilidades:

- **10 microserviços** independentes e testáveis
- **Módulo centralizado** de acesso ao Google Sheets (`sheets-client`)
- **Frontend** em HTML/CSS/JS com tema customizado Neoformula
- **Cache** via CacheService para otimização de performance

## Estrutura do Projeto

```
/src
  /config          # Configurações e mapeamentos
  /shared          # Utilitários compartilhados
  /services        # Microserviços (10)
  /frontend        # Views, components, styles, scripts
  main.ts          # Entry point
```

## Serviços

1. **config-service** - Parâmetros globais e cache
2. **reference-data-service** - Dados de referência (plano de contas, filiais, etc.)
3. **ledger-service** - Lançamentos financeiros (CRUD)
4. **reconciliation-service** - Conciliação bancária automática
5. **cashflow-service** - Fluxo de caixa realizado e projetado
6. **dre-service** - DRE gerencial
7. **kpi-analytics-service** - KPIs e indicadores
8. **reporting-service** - Relatórios para comitê
9. **ui-service** - Interface web
10. **scheduler-service** - Jobs automatizados

## Setup Inicial

### 1. Instalação de dependências

```bash
npm install
```

### 2. Configuração do clasp

```bash
npx clasp login
npx clasp create --type sheets --title "Neoformula Finance App"
```

### 3. Build e Deploy

```bash
npm run deploy
```

### 4. Configurar planilha

No Google Sheets, crie as seguintes abas conforme especificação:

**Configuração (CFG_*):**
- CFG_CONFIG
- CFG_BENCHMARKS
- CFG_LABELS
- CFG_THEME
- CFG_DFC
- CFG_VALIDATION

**Referência (REF_*):**
- REF_PLANO_CONTAS
- REF_FILIAIS
- REF_CANAIS
- REF_CCUSTO
- REF_NATUREZAS

**Transacional (TB_*):**
- TB_LANCAMENTOS
- TB_EXTRATOS
- TB_DRE_MENSAL
- TB_DRE_RESUMO
- TB_DFC_REAL
- TB_DFC_PROJ
- TB_KPI_RESUMO
- TB_KPI_DETALHE

**Relatórios (RPT_*):**
- RPT_COMITE_FATURAMENTO
- RPT_COMITE_DRE
- RPT_COMITE_DFC
- RPT_COMITE_KPIS

### 5. Instalar triggers

No menu da planilha: **Neoformula Finance > Administração > Instalar Triggers**

## Desenvolvimento

### Build em modo watch

```bash
npm run watch
```

### Push para Apps Script

```bash
npm run push
```

### Ver logs

```bash
npm run logs
```

## Próximos Passos

Esta estrutura inicial contém:
- ✅ Arquitetura completa e modular
- ✅ Tipos e contratos definidos
- ✅ Stubs de todos os serviços
- ✅ Frontend básico funcional
- ✅ Sistema de cache
- ✅ Validações estruturadas

**Pendente de implementação:**
- 🔲 Lógica de negócio completa de DRE, DFC, KPIs
- 🔲 Algoritmos de conciliação bancária
- 🔲 Importação de extratos (OFX, CSV)
- 🔲 Validações cruzadas de dados
- 🔲 Exportação para PDF/Slides
- 🔲 Testes unitários e de integração

## Especificação

Consulte o arquivo `neoformula-finance-app-spec-v1.md` para detalhes completos da arquitetura, regras de negócio e estrutura de dados.

## Licença

UNLICENSED - Uso interno Neoformula
