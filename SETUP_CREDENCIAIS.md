# 🔐 Configuração das Credenciais Google Sheets

## 📋 Pré-requisitos

Para que os scripts `infoOpsClock.js` e `inofOutboundDiario.js` funcionem, você precisa criar credenciais de uma Service Account do Google Cloud.

---

## 🚀 Passo a Passo

### 1️⃣ Criar Projeto no Google Cloud

1. Acesse: https://console.cloud.google.com/
2. Clique em **"Criar Projeto"**
3. Dê um nome (ex: "Shopee Manager")
4. Clique em **"Criar"**

---

### 2️⃣ Ativar Google Sheets API

1. No menu lateral, vá em: **APIs e Serviços** → **Biblioteca**
2. Busque por: **"Google Sheets API"**
3. Clique em **"Ativar"**

---

### 3️⃣ Criar Service Account

1. No menu lateral: **APIs e Serviços** → **Credenciais**
2. Clique em **"Criar Credenciais"** → **"Conta de serviço"**
3. Preencha:
   - **Nome**: `shopee-bot`
   - **ID**: `shopee-bot` (gerado automaticamente)
4. Clique em **"Criar e continuar"**
5. Pule as permissões opcionais (clique em **"Continuar"**)
6. Clique em **"Concluir"**

---

### 4️⃣ Baixar Chave JSON

1. Clique na Service Account criada
2. Vá na aba **"Chaves"**
3. Clique em **"Adicionar chave"** → **"Criar nova chave"**
4. Escolha formato: **JSON**
5. Clique em **"Criar"**
6. O arquivo será baixado automaticamente

---

### 5️⃣ Configurar no Projeto

1. Renomeie o arquivo baixado para: **`credenciais.json`**
2. Coloque na raiz do projeto (mesma pasta do `package.json`)
3. ⚠️ **IMPORTANTE**: Este arquivo já está no `.gitignore` e nunca será commitado

---

### 6️⃣ Compartilhar Planilhas com a Service Account

1. Abra o arquivo `credenciais.json`
2. Copie o valor de `client_email` (ex: `shopee-bot@seu-projeto.iam.gserviceaccount.com`)
3. Vá até sua planilha do Google Sheets
4. Clique em **"Compartilhar"**
5. Cole o email da Service Account
6. Defina permissão como **"Leitor"** (ou "Editor" se for escrever)
7. Desmarque **"Notificar pessoas"**
8. Clique em **"Compartilhar"**

---

## ✅ Testar Configuração

Execute os scripts para verificar se está funcionando:

```bash
# Testar infoOpsClock
node infoOpsClock.js

# Testar infoOutboundDiario
node inofOutboundDiario.js
```

Se aparecer: `✅ Arquivo atualizado com X linhas` → **Sucesso!** 🎉

---

## 🛠️ Configurar IDs das Planilhas

Nos arquivos `.js`, você pode alterar os IDs das planilhas:

### `infoOpsClock.js`
```javascript
const ID_PLANILHA = '1Czv3s6ZTKB0t1doydke58JJbgSa1uv3533GGrvxk8Aw';
const INTERVALO = "'Página1'!A4:Y";
```

### `inofOutboundDiario.js`
```javascript
const ID_PLANILHA = '1iJ70tTT_hlUqcWQacHuhP-3CYI8rYNkOdKnBAHXI_eg';
const INTERVALO = "'Resume Out. Capacity'!B5:CE";
```

**Como pegar o ID da planilha:**
- URL da planilha: `https://docs.google.com/spreadsheets/d/SEU_ID_AQUI/edit`
- Copie apenas a parte entre `/d/` e `/edit`

---

## ❌ Troubleshooting

### Erro: "Error: ENOENT: no such file or directory"
- ✅ Verifique se `credenciais.json` está na raiz do projeto

### Erro: "The caller does not have permission"
- ✅ Compartilhe a planilha com o email da Service Account

### Erro: "Unable to parse range"
- ✅ Verifique se o nome da aba está correto
- ✅ Use aspas simples: `'Nome da Aba'!A1:Z`

---

## 🔒 Segurança

- ✅ `credenciais.json` está no `.gitignore`
- ✅ Nunca commite este arquivo
- ✅ Nunca compartilhe a private key
- ❌ Não use em produção sem proteção adicional

---

## 📚 Links Úteis

- [Google Cloud Console](https://console.cloud.google.com/)
- [Google Sheets API Docs](https://developers.google.com/sheets/api)
- [googleapis npm package](https://www.npmjs.com/package/googleapis)

---

**Pronto! Agora você está configurado para usar os scripts de integração com Google Sheets** 🚀
