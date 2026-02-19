# 📦 Guia de Empacotamento e Distribuição

## Shopee - Planejamento Integrado

Este guia explica como gerar os executáveis para distribuição aos analistas.

---

## 🛠️ Pré-requisitos

- **Node.js** instalado (versão 18 ou superior)
- **Windows** (para gerar executáveis `.exe`)
- **Git** (para clonar o repositório)
- **Modo de Desenvolvedor** ativado no Windows (para symbolic links)

---

## 📥 1. Preparar o Ambiente

```powershell
# Clonar o repositório (se ainda não tiver)
git clone https://github.com/Ediiiiipo/Planejamento-Integrado.git
cd Planejamento-Integrado

# Instalar dependências
npm install

# ⚠️ IMPORTANTE: Instalar navegador Chromium do Playwright
npx playwright install chromium
```

---

## 🏗️ 2. Gerar Executáveis

### Opção 1: Gerar TUDO (Instalador + Portable)

```powershell
npm run build
```

**Resultado:**
- `dist/Shopee - Planejamento Integrado-1.0.0-x64.exe` (~500MB) - **Instalador**
- `dist/Shopee - Planejamento Integrado-1.0.0-Portable.exe` (~500MB) - **Portable**

⏱️ **Tempo estimado:** 10-15 minutos (primeira vez)

---

### Opção 2: Gerar Apenas Instalador

```powershell
npm run build:win
```

**Resultado:**
- `dist/Shopee - Planejamento Integrado-1.0.0-x64.exe` - **Instalador NSIS**

---

### Opção 3: Gerar Apenas Portable

```powershell
npm run build:portable
```

**Resultado:**
- `dist/Shopee - Planejamento Integrado-1.0.0-Portable.exe` - **Executável Portátil**

---

## 📦 3. Distribuir aos Analistas

### 🎯 **Recomendação: Versão Portable**

**Por quê?**
- ✅ **Não precisa de permissão de administrador**
- ✅ **Não precisa instalar nada**
- ✅ **Duplo clique e roda**
- ✅ **Ideal para ambientes corporativos restritos**

**Como distribuir:**
1. Copie o arquivo `Shopee - Planejamento Integrado-1.0.0-Portable.exe` (~500MB)
2. Compartilhe via:
   - **Google Drive** / **OneDrive** / **SharePoint**
   - **Rede interna da empresa**
   - **Pendrive** (para instalação offline)

**Instruções para os analistas:**
```
1. Baixe o arquivo "Shopee - Planejamento Integrado-1.0.0-Portable.exe"
2. Salve em uma pasta de sua preferência (ex: C:\Shopee\)
3. Duplo clique no arquivo para abrir
4. Pronto! O aplicativo vai abrir automaticamente
```

---

### 🔧 **Alternativa: Instalador NSIS**

**Quando usar:**
- Analistas têm permissão de administrador
- Querem instalar como aplicativo permanente
- Preferem atalho no Menu Iniciar e Desktop

**Como distribuir:**
1. Copie o arquivo `Shopee - Planejamento Integrado-1.0.0-x64.exe` (~500MB)
2. Compartilhe da mesma forma

**Instruções para os analistas:**
```
1. Baixe o arquivo "Shopee - Planejamento Integrado-1.0.0-x64.exe"
2. Duplo clique para iniciar a instalação
3. Siga as instruções do instalador
4. Escolha a pasta de instalação (ou deixe padrão)
5. Aguarde a instalação concluir
6. Use o atalho criado no Desktop ou Menu Iniciar
```

---

## 📋 4. Informações Técnicas

### O que está incluído no executável?

✅ **Tudo que o aplicativo precisa:**
- Node.js runtime (embutido)
- Chromium (navegador interno do Electron)
- **Playwright + Chromium** (para downloads via navegador)
- Todas as dependências npm
- Código da aplicação
- Ícone da Shopee

✅ **NÃO precisa instalar:**
- Node.js
- Navegador Chrome
- Playwright
- Dependências npm
- Nada!

---

### Tamanho dos arquivos

- **Instalador:** ~500MB
- **Portable:** ~500MB
- **Após instalação:** ~1GB (com cache e dados)

---

### Requisitos do sistema

- **OS:** Windows 10/11 (64-bit)
- **RAM:** 4GB mínimo (8GB recomendado)
- **Disco:** 1.5GB livres
- **Internet:** Necessária para acessar Google Sheets e APIs

---

## 🔄 5. Atualizar Versão

Quando houver uma nova versão:

1. **Atualizar o código:**
   ```powershell
   git pull origin main
   ```

2. **Reinstalar dependências:**
   ```powershell
   npm install
   npx playwright install chromium
   ```

3. **Atualizar a versão no `package.json`:**
   ```json
   {
     "version": "1.1.0"  // Incrementar versão
   }
   ```

4. **Gerar novos executáveis:**
   ```powershell
   npm run build
   ```

5. **Distribuir a nova versão** com as mesmas instruções

---

## ❓ Troubleshooting

### Erro: "Package electron is only allowed in devDependencies"

✅ **Já corrigido!** O `electron` está em `devDependencies` no `package.json`.

---

### Erro: "Cannot create symbolic link"

✅ **Solução:**
1. Ative o **Modo de Desenvolvedor** no Windows:
   - Configurações > Privacidade e segurança > Para desenvolvedores
   - Ative "Modo de Desenvolvedor"
2. OU execute o PowerShell **como Administrador**

---

### Erro: "Cannot find module 'universalify'" ou similar

✅ **Solução:**
- Certifique-se de que executou `npm install` antes do build
- Limpe a pasta `dist/` e tente novamente

---

### Erro: "ENOTDIR, not a directory" ao fazer download

✅ **Solução:**
- Certifique-se de que executou `npx playwright install chromium` antes do build
- Os binários do Playwright devem estar em `node_modules/playwright-core/.local-browsers/`

---

### Build muito lento

✅ **Normal!** O primeiro build pode demorar 10-15 minutos.
- Electron precisa baixar o Chromium (~100MB)
- Playwright adiciona mais ~300MB
- Compactar tudo em um executável

---

### Erro de permissão ao executar

✅ **Solução:**
- Use a versão **Portable** (não precisa de admin)
- OU execute o instalador como administrador (botão direito > "Executar como administrador")

---

## 📞 Suporte

Se tiver problemas:
1. Verifique se o Node.js está instalado: `node --version`
2. Verifique se as dependências foram instaladas: `npm install`
3. Verifique se o Playwright foi instalado: `npx playwright install chromium`
4. Limpe o cache e tente novamente: `Remove-Item -Recurse -Force dist\` e `npm run build`

---

## 🎉 Pronto!

Agora você pode distribuir o aplicativo para os **150+ analistas** sem complicações! 🚀

**Recomendação final:** Use a versão **Portable** para facilitar a vida de todos! 😊

---

## 📝 Notas Importantes

- ⚠️ **Primeira execução pode demorar:** O Playwright precisa configurar o navegador na primeira vez
- ⚠️ **Tamanho do executável:** ~500MB devido ao Chromium embutido
- ✅ **Totalmente offline:** Após o download, funciona sem internet (exceto para acessar Google Sheets)
