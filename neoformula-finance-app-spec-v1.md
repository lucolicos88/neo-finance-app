# Neoformula Finance App – Especificação para Claude Code (v1)

> Arquivo de especificação para desenvolvimento de um aplicativo de gestão financeira e contábil em Google Apps Script, com arquitetura modular em microserviços e base de dados em Google Sheets.

---

## 1. Objetivo deste arquivo

Este arquivo será usado pelo **Claude Code** para gerar o código do projeto em Google Apps Script (via VS Code).  
Ele descreve **arquitetura**, **responsabilidades de cada microserviço** e **estrutura completa da base de dados (Google Sheets)**.

### O que Claude deve fazer

1. Criar a estrutura de pastas e arquivos conforme especificado na seção 2.
2. Implementar os microserviços com base nas interfaces e contratos definidos:
   - Cada serviço em seu próprio arquivo, **sem criar um monólito**.
   - Cada serviço deve ser **testável isoladamente**.
3. Implementar o acesso ao Google Sheets **apenas via um módulo compartilhado** (`sheets-client`).
4. Preparar o código para ser empacotado em `/dist` e enviado para Apps Script via `clasp`.

### O que Claude **não** deve fazer neste momento

- Não implementar lógica de negócio complexa em detalhes (cálculos completos de DRE, DFC, KPIs).
- Não acoplar serviços entre si de forma rígida (uso apenas via interfaces).
- Não adicionar bibliotecas externas não compatíveis com Google Apps Script.
- Não fazer chamadas HTTP externas (sem integrações com bancos, ERPs ou APIs de terceiros).

---

## 2. Estrutura de pastas do projeto

### 2.1. Visão geral

O repositório deve seguir a seguinte organização:

```text
neoformula-finance-app/
  appsscript.json
  .clasp.json
  /dist                        # Saída final em JS/HTML/CSS (para Apps Script)
  /src
    /config
      config.schema.ts
      sheet-mapping.ts
      benchmarks.ts
    /shared
      date-utils.ts
      money-utils.ts
      validation.ts
      sheets-client.ts
      cache.ts
      types.ts
    /services
      config-service.ts
      reference-data-service.ts
      ledger-service.ts
      reconciliation-service.ts
      cashflow-service.ts
      dre-service.ts
      kpi-analytics-service.ts
      reporting-service.ts
      ui-service.ts
      scheduler-service.ts
    /frontend
      /views
        index.html
        dashboard.html
        lancamentos.html
        conciliacao.html
        dre.html
        dfc.html
        kpi.html
        configuracoes.html
      /components
        header.html
        sidebar.html
        kpi-card.html
      /styles
        base.css
        theme-neoformula.css
      /scripts
        main.js
        forms.js
        dashboard.js
        conciliacao.js
        configuracoes.js
  /tests
    /unit
    /integration
  /docs
    architecture.md
    services.md
```

### 2.2. Entrypoint Apps Script

No build final (`/dist`) deve existir um arquivo principal, por exemplo `main.js`, contendo apenas:

- `function doGet(e)` – para servir a interface web.
- Registros de menus (se necessário, via `onOpen`).
- Roteamento básico para os serviços.

Implementação da lógica fica nos serviços em arquivos separados.

---

## 3. Arquitetura de microserviços

### 3.1. Lista de serviços

1. `config-service` – parâmetros globais, IDs de planilhas, benchmarks, toggles.
2. `reference-data-service` – tabelas mestras (plano de contas, filiais, canais, centros de custo, naturezas).
3. `ledger-service` – lançamentos financeiros/contábeis (pagar/receber, provisões, ajustes).
4. `reconciliation-service` – conciliação entre lançamentos e extratos bancários.
5. `cashflow-service` – fluxo de caixa realizado e projetado, contas futuras.
6. `dre-service` – DRE gerencial, por mês, filial, grupo.
7. `kpi-analytics-service` – KPIs e indicadores financeiros críticos.
8. `reporting-service` – geração de relatórios em abas do Sheets e exportação (PDF/Slides).
9. `ui-service` – camada HTML/CSS/JS, validação de formulários e navegação.
10. `scheduler-service` – triggers (jobs) diários/mensais e orquestração.

### 3.2. Contratos gerais (TypeScript – apenas interfaces)

```ts
// src/shared/types.ts

export type Period = {
  year: number;
  month: number; // 1-12
};

export type BranchId = string;
export type ChannelId = string;
export type CostCenterId = string;
export type AccountCode = string;

export type Money = number; // sempre em moeda local (ex.: BRL), 2 casas decimais

export interface LedgerEntry {
  id: string;
  competencia: Date;
  vencimento: Date | null;
  pagamento: Date | null;
  tipo: 'PAGAR' | 'RECEBER' | 'TRANSFERENCIA' | 'AJUSTE';
  filial: BranchId;
  centroCusto: CostCenterId | null;
  contaGerencial: AccountCode;
  contaContabil: AccountCode | null;
  grupoReceita: 'SERVICOS' | 'REVENDA' | null;
  canal: ChannelId | null;
  descricao: string;
  valorBruto: Money;
  desconto: Money;
  juros: Money;
  multa: Money;
  valorLiquido: Money;
  status: 'PREVISTO' | 'REALIZADO' | 'CANCELADO';
  idExtratoBanco: string | null;
  origem: 'MANUAL' | 'IMPORTADO';
}

export interface DRELine {
  period: Period;
  branchId: BranchId | null;
  group: string;     // ex.: RECEITA_BRUTA, IMPOSTOS, DESPESAS_FIXAS
  subGroup: string;  // ex.: PESSOAL, MARKETING
  value: Money;
}

export interface CashflowLine {
  date: Date;
  type: 'ENTRADA' | 'SAIDA';
  category: string; // OPERACIONAL, INVESTIMENTO, FINANCIAMENTO
  description: string;
  value: Money;
  projected: boolean;
}
```

> **Importante:** Claude deve manter estes contratos e apenas estendê-los de forma compatível (sem quebrar tipos).

---

## 4. Base de Dados – Estrutura Google Sheets

### 4.1. Visão geral

O sistema usará **uma ou mais planilhas** no Google Sheets.  
No mínimo:

1. **Planilha Master**: `NF_FINANCE_APP_MASTER`
   - Contém todas as abas de **configuração**, **referência** e **transacionais**.
2. **Planilha de Relatórios** (opcional, pode ser a mesma): `NF_FINANCE_REPORTS`
   - Contém abas de **DRE**, **DFC**, **KPIs** e **relatórios formatados** para comitê.

Claude deve assumir que os nomes de planilhas e abas virão do módulo `config-service` / `sheet-mapping.ts`, não hard-coded dentro dos serviços.

---

### 4.2. Abas de Configuração (prefixo `CFG_`)

#### 4.2.1. `CFG_CONFIG` – Parâmetros globais

| Coluna          | Tipo    | Obrigatório | Exemplo                        | Observações                                                    |
|-----------------|---------|-------------|--------------------------------|----------------------------------------------------------------|
| chave           | string  | sim         | `MASTER_SPREADSHEET_ID`        | Nome do parâmetro                                              |
| valor           | string  | sim         | `1AbcDe123...`                 | Valor em string (parseado conforme `tipo`)                     |
| tipo            | string  | sim         | `STRING`, `NUMBER`, `BOOLEAN`  | Define como será interpretado                                  |
| descricao       | string  | não         | `ID da planilha master`        | Ajuda para o usuário                                           |
| ativo           | boolean | não         | `TRUE`                         | Permite desativar parâmetros sem removê-los                    |

#### 4.2.2. `CFG_BENCHMARKS` – Faixas de indicadores

| Coluna             | Tipo   | Obrigatório | Exemplo                        |
|--------------------|--------|-------------|--------------------------------|
| metric             | string | sim         | `DESCONTO_MEDIO`, `CMA`, `CMV` |
| unidade            | string | sim         | `%`, `R$/UNID`                 |
| sensacional_min    | number | sim         | `0`                            |
| sensacional_max    | number | sim         | `3`                            |
| excelente_min      | number | sim         | `3`                            |
| excelente_max      | number | sim         | `5`                            |
| bom_min            | number | sim         | `5`                            |
| bom_max            | number | sim         | `7`                            |
| ruim_min           | number | sim         | `7`                            |
| ruim_max           | number | sim         | `10`                           |
| pessimo_min        | number | sim         | `10`                           |
| pessimo_max        | number | sim         | `100`                          |

#### 4.2.3. `CFG_LABELS` – Textos e rótulos da interface

| Coluna      | Tipo   | Obrigatório | Exemplo                    |
|-------------|--------|-------------|----------------------------|
| chave       | string | sim         | `MENU_DASHBOARD`           |
| pt_br       | string | sim         | `Dashboard`                |
| pt_br_long  | string | não         | `Visão geral dos resultados` |

#### 4.2.4. `CFG_THEME` – Tema visual

| Coluna        | Tipo   | Obrigatório | Exemplo      |
|---------------|--------|-------------|--------------|
| chave         | string | sim         | `primary`    |
| valor_hex     | string | sim         | `#009f9a`    |
| descricao     | string | não         | `Cor primária` |

#### 4.2.5. `CFG_DFC` – Parâmetros de fluxo de caixa

| Coluna                 | Tipo   | Obrigatório | Exemplo           |
|------------------------|--------|-------------|-------------------|
| chave                  | string | sim         | `HORIZONTE_MESES` |
| valor_numero           | number | sim         | `6`               |
| descricao              | string | não         | `Meses a projetar` |

#### 4.2.6. `CFG_VALIDATION` – Regras de validação customizáveis

| Coluna        | Tipo   | Obrigatório | Exemplo                             |
|---------------|--------|-------------|-------------------------------------|
| regra         | string | sim         | `MAX_DIAS_RETROATIVO`               |
| valor         | number | sim         | `7`                                 |
| descricao     | string | não         | `Máx. dias para lançamento retroativo` |

---

### 4.3. Abas de Referência (prefixo `REF_`)

#### 4.3.1. `REF_PLANO_CONTAS` – Plano de contas gerencial

| Coluna          | Tipo   | Obrigatório | Exemplo              | Observações                                      |
|-----------------|--------|-------------|----------------------|--------------------------------------------------|
| codigo          | string | sim         | `3.01.01`            |
| descricao       | string | sim         | `Receita Serviços`   |
| tipo            | string | sim         | `RECEITA`, `DESPESA`, `CUSTO` |
| grupo_dre       | string | sim         | `RECEITA_BRUTA`      |
| subgrupo_dre    | string | não         | `SERVICOS`           |
| grupo_dfc       | string | não         | `OPERACIONAL`        |
| variavel_fixa   | string | não         | `VARIAVEL` ou `FIXA` |
| cma_cmv         | string | não         | `CMA`, `CMV`, vazio  |

#### 4.3.2. `REF_FILIAIS`

| Coluna     | Tipo   | Obrigatório | Exemplo                 |
|------------|--------|-------------|-------------------------|
| id         | string | sim         | `BOSQUE`                |
| nome       | string | sim         | `Bosque`                |
| ativa      | bool   | sim         | `TRUE`                  |

#### 4.3.3. `REF_CANAIS`

| Coluna     | Tipo   | Obrigatório | Exemplo           |
|------------|--------|-------------|-------------------|
| id         | string | sim         | `LOJA_FISICA`     |
| nome       | string | sim         | `Lojas Físicas`   |
| grupo      | string | não         | `SERVICOS`        |

#### 4.3.4. `REF_CCUSTO` – Centros de custo

| Coluna     | Tipo   | Obrigatório | Exemplo           |
|------------|--------|-------------|-------------------|
| id         | string | sim         | `ADM`             |
| nome       | string | sim         | `Administrativo`  |

#### 4.3.5. `REF_NATUREZAS` – Naturezas de despesa/receita

| Coluna     | Tipo   | Obrigatório | Exemplo        |
|------------|--------|-------------|----------------|
| id         | string | sim         | `MARKETING`    |
| nome       | string | sim         | `Marketing`    |
| grupo_dre  | string | sim         | `DESPESAS_VAR` |

---

### 4.4. Abas Transacionais (prefixo `TB_`)

#### 4.4.1. `TB_LANCAMENTOS` – Lançamentos financeiros

| Coluna             | Tipo      | Obrig. | Exemplo                 | Observações                                                   |
|--------------------|-----------|--------|-------------------------|---------------------------------------------------------------|
| id                 | string    | sim    | `L2025-000001`          | Gerado pelo sistema                                           |
| data_competencia   | date      | sim    | `01/10/2025`           | Competência contábil                                         |
| data_vencimento    | date      | não    | `10/10/2025`           |
| data_pagamento     | date      | não    | `09/10/2025`           |
| tipo               | string    | sim    | `PAGAR`, `RECEBER`     |
| filial             | string    | sim    | `BOSQUE`               |
| centro_custo       | string    | não    | `ADM`                  |
| conta_gerencial    | string    | sim    | `3.01.01`              |
| conta_contabil     | string    | não    | `3.01.01.001`          |
| grupo_receita      | string    | não    | `SERVICOS`, `REVENDA`  |
| canal              | string    | não    | `LOJA_FISICA`          |
| descricao          | string    | sim    | `Venda balcão`         |
| valor_bruto        | number    | sim    | `1000,00`              |
| desconto           | number    | não    | `50,00`                |
| juros              | number    | não    | `0,00`                 |
| multa              | number    | não    | `0,00`                 |
| valor_liquido      | number    | sim    | `950,00`               |
| status             | string    | sim    | `PREVISTO`, `REALIZADO`, `CANCELADO` |
| id_extrato_banco   | string    | não    | `EB2025-000045`        |
| origem             | string    | sim    | `MANUAL` ou `IMPORTADO` |
| observacoes        | string    | não    | `Baixa parcial`        |

#### 4.4.2. `TB_EXTRATOS` – Extratos bancários

| Coluna           | Tipo   | Obrig. | Exemplo                 |
|------------------|--------|--------|-------------------------|
| id               | string | sim    | `EB2025-000045`        |
| data_movimento   | date   | sim    | `09/10/2025`           |
| conta_bancaria   | string | sim    | `BB_AG123_CC456`       |
| historico        | string | sim    | `CRÉDITO CARTÃO`       |
| documento        | string | não    | `NSU123456`            |
| valor            | number | sim    | `950,00`               |
| saldo_apos       | number | não    | `10500,00`             |
| conciliado       | bool   | não    | `TRUE`                 |
| id_lancamento    | string | não    | `L2025-000001`         |

---

### 4.5. Abas de Saída – DRE, DFC, KPIs, Relatórios

#### 4.5.1. `TB_DRE_MENSAL`

| Coluna      | Tipo   | Obrig. | Exemplo          |
|-------------|--------|--------|------------------|
| ano         | number | sim    | `2025`           |
| mes         | number | sim    | `10`             |
| filial      | string | não    | `BOSQUE`         |
| grupo_dre   | string | sim    | `RECEITA_BRUTA` |
| subgrupo    | string | não    | `SERVICOS`       |
| valor       | number | sim    | `100000,00`      |

#### 4.5.2. `TB_DRE_RESUMO`

| Coluna      | Tipo   | Obrig. | Exemplo         |
|-------------|--------|--------|-----------------|
| ano         | number | sim    | `2025`          |
| mes         | number | sim    | `10`            |
| indicador   | string | sim    | `EBITDA_PCT`    |
| valor       | number | sim    | `18,5`          |

#### 4.5.3. `TB_DFC_REAL` – Fluxo de caixa realizado

| Coluna       | Tipo   | Obrig. | Exemplo          |
|--------------|--------|--------|------------------|
| data         | date   | sim    | `09/10/2025`     |
| categoria    | string | sim    | `OPERACIONAL`    |
| tipo         | string | sim    | `ENTRADA`        |
| descricao    | string | sim    | `Recebimento`    |
| valor        | number | sim    | `950,00`         |
| conta_banco  | string | não    | `BB_AG123_CC456` |

#### 4.5.4. `TB_DFC_PROJ` – Fluxo de caixa projetado

| Coluna       | Tipo   | Obrig. | Exemplo         |
|--------------|--------|--------|-----------------|
| ano          | number | sim    | `2025`          |
| mes          | number | sim    | `11`            |
| categoria    | string | sim    | `OPERACIONAL`   |
| tipo         | string | sim    | `ENTRADA`       |
| valor        | number | sim    | `120000,00`     |

#### 4.5.5. `TB_KPI_RESUMO`

| Coluna       | Tipo   | Obrig. | Exemplo               |
|--------------|--------|--------|-----------------------|
| ano          | number | sim    | `2025`                |
| mes          | number | sim    | `10`                  |
| filial       | string | não    | `BOSQUE`              |
| canal        | string | não    | `LOJA_FISICA`         |
| metric       | string | sim    | `DESCONTO_MEDIO`      |
| valor        | number | sim    | `4,5`                 |
| faixa        | string | não    | `BOM`                 |

#### 4.5.6. `TB_KPI_DETALHE`

Estrutura semelhante, com campos adicionais `produto`, `familia`, etc., conforme necessidade.

#### 4.5.7. Abas de relatório formatado (`RPT_COMITE_*`)

Usadas apenas pelo `reporting-service` para gerar visões para comitê:
- `RPT_COMITE_FATURAMENTO`
- `RPT_COMITE_DRE`
- `RPT_COMITE_DFC`
- `RPT_COMITE_KPIS`

Os dados devem vir sempre das tabelas `TB_*` acima, nunca de cálculos “soltos” nas abas de relatório.

---

## 5. Convenções de validação e regras gerais

Claude deve considerar as seguintes regras **mínimas** de validação:

1. **Datas**
   - `data_competencia`: obrigatória, não pode ser vazia.
   - `data_pagamento`: não pode ser futura em relação à data atual.
   - Períodos fechados não podem ser alterados (controle via configuração).

2. **Valores**
   - `valor_bruto` > 0.
   - `valor_liquido` = `valor_bruto` − `desconto` + `juros` + `multa`.
   - Proteger contra células vazias (tratar como 0).

3. **Referências cruzadas**
   - `filial`, `canal`, `centro_custo`, `conta_gerencial` devem existir nas respectivas abas `REF_*`.
   - `status` só pode assumir valores válidos (`PREVISTO`, `REALIZADO`, `CANCELADO`).

4. **Integridade contábil básica**
   - Lançamentos com `status = REALIZADO` devem ter `data_pagamento` preenchida.
   - `reconciliation-service` deve garantir que a soma de lançamentos conciliados por extrato seja coerente (tolerância configurável via `CFG_VALIDATION`).

---

## 6. Paleta de cores e CSS (Neoformula)

Claude deve criar o arquivo `src/frontend/styles/theme-neoformula.css` com as variáveis:

```css
:root {
  --neo-navy:        #06345b;
  --neo-teal:        #009f9a;
  --neo-green:       #00a86b;
  --neo-light-teal:  #4ec7b0;

  --neo-purple:      #3b2e83;
  --neo-orange:      #f7941d;

  --neo-white:       #ffffff;
  --neo-gray-100:    #f5f7fa;
  --neo-gray-300:    #d2d7e0;
  --neo-gray-600:    #5c667a;
  --neo-gray-900:    #212631;

  --kpi-sensacional: #00b894;
  --kpi-excelente:   #27ae60;
  --kpi-bom:         #f1c40f;
  --kpi-ruim:        #e67e22;
  --kpi-pessimo:     #e74c3c;

  --bg-app:          var(--neo-gray-100);
  --bg-card:         var(--neo-white);
  --text-primary:    var(--neo-gray-900);
  --text-secondary:  var(--neo-gray-600);
  --border-soft:     var(--neo-gray-300);
}
```

Diretrizes:
- Cabeçalho/topbar: `--neo-navy` com texto branco.
- Botão primário: `--neo-teal`.
- Estados de KPI: cores `--kpi-*` para faixas de benchmark.

---

## 7. Instruções específicas para Claude Code

1. **Organização**
   - Respeitar a estrutura de pastas e nomes de arquivos desta especificação.
   - Implementar um módulo `sheet-mapping.ts` que centralize todos os nomes de abas.

2. **Acesso ao Google Sheets**
   - Todo acesso a planilhas deve passar por `sheets-client.ts` com funções genéricas:
     - `getSheetValues(sheetName: string): any[][]`
     - `setSheetValues(sheetName: string, range: string, values: any[][])`
     - `appendRows(sheetName: string, rows: any[][])`

3. **Independência dos serviços**
   - Cada serviço deve expor uma interface clara (por exemplo, funções `ConfigService`, `LedgerService`, etc.).
   - Serviços só devem se conhecer via chamadas de função, não compartilhando estado global desnecessário.

4. **Limitações do Apps Script**
   - Usar operações em lote em Sheets (evitar chamadas por célula).
   - Preparar código para ser executado dentro de ~6 minutos por função.
   - Prever uso de `CacheService` para dados de referência (plano de contas, benchmarks, etc.).

5. **Sem lógica excessivamente complexa nesta fase**
   - Implementar stubs para cálculos mais pesados (DRE, DFC, KPIs) com estrutura de função, comentários `TODO` e tipos corretos.
   - Detalhes de cálculo podem ser feitos em iterações futuras, mantendo a arquitetura estável.
