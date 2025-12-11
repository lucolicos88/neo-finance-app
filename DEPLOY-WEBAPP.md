# 🚀 Como Fazer Deploy da Web App

## Passo a Passo para Publicar a Aplicação

### 1️⃣ Preparar a Planilha

Na sua planilha Google Sheets:

1. **Abra a planilha:** https://docs.google.com/spreadsheets/d/1nwEtOMb7uGm0ZXEM_xcQLAJQSOAjhgMSsve_7KXycjI/edit

2. **Execute o Setup (primeira vez):**
   - Menu: `Neoformula Finance` → `Administração` → `⚙️ Setup da Planilha`
   - Aguarde a criação das 26 abas

3. **Crie Dados de Exemplo (para testar):**
   - Menu: `Neoformula Finance` → `Administração` → `📝 Criar Dados de Exemplo`
   - Isso criará 11 lançamentos e 3 extratos

### 2️⃣ Acessar o Editor do Apps Script

**Opção A - Via Menu da Planilha:**
- Menu: `Extensões` → `Apps Script`

**Opção B - Via Comando (mais rápido):**
```bash
npm run open
```

### 3️⃣ Fazer Deploy da Web App

No Editor do Apps Script:

1. **Clique em "Fazer Deploy" (canto superior direito)**
   - Ou: `Deploy` → `New deployment`

2. **Configurar o Deploy:**
   - **Tipo:** Selecione ⚙️ "Web app"
   - **Descrição:** `Versão 1 - DEV`
   - **Executar como:** `Eu (seu email)`
   - **Quem tem acesso:**
     - Para desenvolvimento: `Somente eu`
     - Para produção: `Qualquer pessoa`

3. **Clique em "Fazer Deploy"**

4. **Autorizar o App:**
   - Clique em "Autorizar acesso"
   - Selecione sua conta Google
   - Clique em "Avançado"
   - Clique em "Ir para Neoformula Finance (não seguro)"
   - Clique em "Permitir"

5. **Copiar a URL:**
   - Após o deploy, copie a **URL da Web App**
   - Formato: `https://script.google.com/macros/s/AKfycby.../exec`

### 4️⃣ Acessar a Aplicação

**Opção A - Via URL Direta:**
- Cole a URL copiada no navegador

**Opção B - Via Menu da Planilha:**
- Menu: `Neoformula Finance` → `Administração` → `🌐 Abrir Web App`
- Clique no botão "Abrir Web App"

### 5️⃣ Testar a Aplicação

1. **Dashboard:**
   - Veja os KPIs atualizados
   - Verifique alertas
   - Confira últimos lançamentos

2. **Contas a Pagar:**
   - Visualize contas vencidas e pendentes
   - Teste os filtros
   - Experimente pagar uma conta

3. **Contas a Receber:**
   - Veja recebimentos pendentes
   - Filtre por cliente
   - Teste receber uma conta

4. **Conciliação:**
   - Veja extratos e lançamentos lado a lado
   - Clique para conciliar manualmente
   - Teste a conciliação automática

## 🔄 Atualizar o Deploy

Quando fizer alterações no código:

1. **Build e Push:**
   ```bash
   npm run deploy
   ```

2. **No Apps Script Editor:**
   - Menu: `Deploy` → `Manage deployments`
   - Clique no ✏️ ao lado da versão ativa
   - Mude a **Versão** para "New version"
   - Adicione descrição: `Versão 2 - [sua descrição]`
   - Clique em "Deploy"

3. **Recarregue a Web App:**
   - A URL permanece a mesma
   - Apenas recarregue a página (F5)

## 🐛 Solução de Problemas

### Erro: "Script function not found"
- **Causa:** Build não foi feito ou push falhou
- **Solução:** Execute `npm run deploy` novamente

### Erro: "Authorization required"
- **Causa:** Permissões não foram concedidas
- **Solução:** Refaça o processo de autorização (passo 3.4)

### Aplicação não carrega
- **Causa:** Deploy não está ativo
- **Solução:** Verifique em `Manage deployments` se há um deploy ativo

### Dados não aparecem
- **Causa:** Planilha não tem dados
- **Solução:** Execute o setup de dados de exemplo (passo 1.3)

## 📝 Notas Importantes

- ✅ A URL da Web App é **permanente** - salve-a!
- ✅ Alterações no código requerem novo deploy
- ✅ Alterações na planilha aparecem automaticamente
- ✅ Para produção, mude "Quem tem acesso" para "Qualquer pessoa"
- ⚠️ Cada deploy gera uma nova versão (histórico mantido)
- ⚠️ Limite de 20 versões ativas simultaneamente

## 🎯 Próximos Passos

Após o deploy bem-sucedido:

1. Compartilhe a URL com usuários
2. Configure permissões de acesso
3. Monitore logs: `npm run logs`
4. Implemente funcionalidades adicionais
5. Configure triggers automáticos (Menu: Instalar Triggers)

## 📞 URLs Úteis

- **Planilha:** https://docs.google.com/spreadsheets/d/1nwEtOMb7uGm0ZXEM_xcQLAJQSOAjhgMSsve_7KXycjI/edit
- **Apps Script Editor:** `npm run open` ou via menu Extensões
- **Web App:** Copie após o deploy
- **Repositório Git:** https://github.com/lucolicos88/appPlanNeo

---

🎉 **Pronto!** Sua aplicação web está online e acessível via URL!
