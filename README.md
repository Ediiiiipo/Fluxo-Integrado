# 🚀 Shopee - Gerenciador de Pedidos

<div align="center">

![Version](https://img.shields.io/badge/version-1.0.0-orange)
![License](https://img.shields.io/badge/license-MIT-blue)
![Node](https://img.shields.io/badge/node-%3E%3D16.0.0-green)
![Electron](https://img.shields.io/badge/electron-28.0.0-blue)

**Aplicação desktop para automação e gerenciamento de pedidos da Shopee Express**

[Características](#-características) •
[Instalação](#-instalação) •
[Uso](#-uso) •
[Documentação](#-documentação) •
[Contribuir](#-contribuir)

</div>

---

## 📋 Índice

- [Sobre](#-sobre)
- [Características](#-características)
- [Capturas de Tela](#-capturas-de-tela)
- [Pré-requisitos](#-pré-requisitos)
- [Instalação](#-instalação)
- [Uso](#-uso)
  - [Interface Gráfica](#interface-gráfica-electron)
  - [Linha de Comando](#linha-de-comando-terminal)
  - [Múltiplas Stations](#baixar-múltiplas-stations)
- [Configuração](#-configuração)
- [Estrutura do Projeto](#-estrutura-do-projeto)
- [Tecnologias](#-tecnologias)
- [Troubleshooting](#-troubleshooting)
- [Roadmap](#-roadmap)
- [Contribuir](#-contribuir)
- [Licença](#-licença)
- [Contato](#-contato)

---

## 📖 Sobre

O **Shopee Gerenciador de Pedidos** é uma aplicação desktop desenvolvida com Electron que automatiza o processo de download e gerenciamento de pedidos do sistema SPX da Shopee Express.

### 🎯 Problema que resolve:

- ❌ Download manual de relatórios é demorado
- ❌ Troca manual de stations é repetitiva
- ❌ Análise de LH Trips em Excel é trabalhosa
- ❌ Gestão de múltiplas stations é ineficiente

### ✅ Solução:

- ✅ **Automação completa** do download de relatórios
- ✅ **Troca automática** entre stations
- ✅ **Interface tipo Excel** para visualização rápida
- ✅ **Filtros inteligentes** por LH Trip
- ✅ **Processamento em lote** de múltiplas stations

---

## ✨ Características

### 🤖 Automação

- **Download Automático**: Bot que acessa o sistema e baixa relatórios
- **Login Persistente**: Salva sessão para evitar login repetido
- **Seleção de Station**: Troca automaticamente entre stations
- **Processamento em Lote**: Baixa de múltiplas stations em sequência
- **Detecção Inteligente**: Aguarda processamento e detecta quando arquivo está pronto

### 🖥️ Interface Gráfica

- **Dashboard Moderno**: Interface limpa e intuitiva
- **Visualização Tipo Excel**: Tabela com todos os campos do relatório
- **Sidebar com LH Trips**: Lista todas as LH Trips com contadores
- **Filtros Instantâneos**: Clique para filtrar por LH Trip específica
- **Indicadores Visuais**: Loading, progresso e notificações

### 📊 Análise de Dados

- **Agrupamento por LH Trip**: Contagem automática de pedidos por LH
- **Identificação de "Sem LH"**: Destaca pedidos sem LH Trip
- **Suporte a ZIP**: Descompacta e unifica múltiplos arquivos
- **Processamento Excel**: Lê e processa arquivos .xlsx e .xls

### 🔒 Segurança

- **Perfil Isolado**: Navegador com perfil próprio
- **Cookies Seguros**: Sessão salva localmente
- **Sem Senhas no Código**: Credenciais apenas via login manual
- **Screenshots de Erro**: Debug facilitado em caso de falha

---

## 📸 Capturas de Tela

### Interface Principal
```
┌─────────────────────────────────────────────────────────┐
│  📦 Shopee - Gerenciador de Pedidos    🔽 Baixar  📂 Carregar │
├──────────────┬──────────────────────────────────────────┤
│ LH Trips     │  Todos os Pedidos        1.234 pedidos   │
│              │                                           │
│ 📊 TODOS     │  ┌─────────────────────────────────────┐ │
│ 1.234        │  │ ID │ LH Trip │ CEP │ Status │ ...   │ │
│              │  ├─────────────────────────────────────┤ │
│ 🚚 LH001     │  │ BR123... │ LH001 │ 12345-678 │ ...  │ │
│ 450          │  │ BR124... │ LH001 │ 54321-987 │ ...  │ │
│              │  │ ...                                   │ │
│ 🚚 LH002     │  └─────────────────────────────────────┘ │
│ 389          │                                           │
│              │                                           │
│ ⚠️ SEM LH    │                                           │
│ 395          │                                           │
└──────────────┴──────────────────────────────────────────┘
```

---

## 🔧 Pré-requisitos

- **Node.js** >= 16.0.0
- **npm** >= 8.0.0
- **Windows** 10/11 (testado)
- **Git** (para clonar o repositório)

---

## 📦 Instalação

### 1. Clonar o Repositório

```bash
git clone https://github.com/Ediiiiipo/Shopee---Gerenciador-de-Pedidos.git
cd Shopee---Gerenciador-de-Pedidos
```

### 2. Instalar Dependências

```bash
npm install
```

### 3. Instalar Navegador Chromium

```bash
npx playwright install chromium
```

### 4. Verificar Instalação

```bash
npm start
```

Se a janela do aplicativo abrir, está tudo certo! ✅

---

## 🚀 Uso

### Interface Gráfica (Electron)

Iniciar a aplicação desktop:

```bash
npm start
```

**Funcionalidades:**

1. **Baixar da Shopee**: Clique para iniciar automação
2. **Carregar Relatório**: Selecione arquivo Excel local
3. **Filtrar por LH**: Clique na LH Trip desejada na sidebar
4. **Visualizar Pedidos**: Veja todos os dados na tabela

---

### Linha de Comando (Terminal)

Para executar o bot sem interface gráfica:

```bash
node shopee-downloader.js
```

**O que acontece:**

1. ✅ Abre navegador automaticamente
2. ⏳ Aguarda você fazer login (primeira vez)
3. 🔄 Navega para página de exportação
4. 📥 Baixa relatório automaticamente
5. 📊 Processa e salva arquivo unificado
6. 📄 Gera relatório HTML

---

### Baixar Station Específica

Edite o final do arquivo `shopee-downloader.js`:

```javascript
// Trocar para station específica
main('LM Hub_MG_Belo Horizonte_02').catch(console.error);
```

Execute:

```bash
node shopee-downloader.js
```

---

### Baixar Múltiplas Stations

Edite o final do arquivo `shopee-downloader.js`:

```javascript
baixarMultiplasStations([
  'LM Hub_MG_Belo Horizonte_01',
  'LM Hub_MG_Belo Horizonte_02',
  'LM Hub_SP_São Paulo_01',
  'LM Hub_SP_Guarulhos',
  'LM Hub_RJ_Rio de Janeiro_01'
]).catch(console.error);
```

Execute:

```bash
node shopee-downloader.js
```

O bot vai:
1. Fazer login uma vez
2. Baixar de cada station automaticamente
3. Gerar relatório consolidado no final

---

## ⚙️ Configuração

### Adicionar Suas Stations

No arquivo `shopee-downloader.js`, edite a lista:

```javascript
STATIONS: [
  "LM Hub_MG_Belo Horizonte_01",
  "LM Hub_MG_Belo Horizonte_02",
  "LM Hub_SP_São Paulo_01",
  // Adicione suas stations aqui
]
```

### Alterar Pasta de Downloads

```javascript
DOWNLOADS_DIR: path.join(os.homedir(), 'Desktop', 'Shopee_Downloads'),
```

### Ajustar Timeouts

```javascript
TIMEOUT_DEFAULT: 30000,  // 30 segundos
TIMEOUT_LOGIN: 600000,   // 10 minutos
```

---

## 📁 Estrutura do Projeto

```
shopee-bot/
│
├── main.js                    # Processo principal do Electron
├── index.html                 # Interface gráfica (UI)
├── renderer.js                # Lógica da interface
├── shopee-downloader.js       # Bot de automação (core)
│
├── package.json               # Dependências e scripts
├── package-lock.json          # Versões exatas
├── .gitignore                 # Arquivos ignorados pelo Git
│
├── .shopee-bot/              # Perfil do navegador (criado automaticamente)
│   └── profile/              # Sessão e cookies salvos
│
└── Shopee_Downloads/         # Pasta de saída (criada automaticamente)
    └── [Station Name]/
        ├── Relatorio_YYYY-MM-DD.xlsx
        └── Relatorio.html
```

---

## 🛠️ Tecnologias

### Core

- **[Electron](https://www.electronjs.org/)** `^28.0.0` - Framework desktop
- **[Playwright](https://playwright.dev/)** `^1.40.0` - Automação web
- **[Node.js](https://nodejs.org/)** `>=16.0.0` - Runtime JavaScript

### Processamento

- **[XLSX](https://www.npmjs.com/package/xlsx)** `^0.18.5` - Manipulação de Excel
- **[AdmZip](https://www.npmjs.com/package/adm-zip)** `^0.5.10` - Descompactação de arquivos
- **[fs-extra](https://www.npmjs.com/package/fs-extra)** `^11.2.0` - Sistema de arquivos

---

## 🐛 Troubleshooting

### Problema: Bot não clica em "Baixar"

**Solução:**

1. Verifique se o relatório está com status "Pronto"
2. Aguarde mais tempo antes do clique
3. Veja screenshot de erro em: `C:\Users\[USER]\AppData\Local\Temp\shopee_temp\erro.png`

---

### Problema: Não consegue trocar de station

**Solução:**

1. Verifique se o nome da station está correto
2. Confira se tem permissão para acessar a station
3. Veja screenshot: `erro-station.png`

---

### Problema: Login não é detectado

**Solução:**

1. Faça login mais rápido (timeout de 10 min)
2. Verifique se a URL mudou após login
3. Limpe o perfil do navegador: delete `.shopee-bot/`

---

### Problema: Arquivo não aparece na interface

**Solução:**

1. Verifique se o arquivo tem coluna "LH Trip" (coluna I)
2. Confira se o arquivo é .xlsx válido
3. Veja console do Electron (Ctrl+Shift+I)

---

### Problema: "node_modules" muito grande

**Solução:**

```bash
# Limpar node_modules
rm -rf node_modules
npm install --production
```

---

## 🗺️ Roadmap

### ✅ Versão 1.0 (Atual)

- [x] Download automático de relatórios
- [x] Interface Electron
- [x] Filtro por LH Trip
- [x] Suporte a múltiplas stations

### 🚧 Versão 1.1 (Em desenvolvimento)

- [ ] Exportar pedidos filtrados para Excel
- [ ] Gráficos de distribuição por LH
- [ ] Busca por texto na tabela
- [ ] Ordenação de colunas

### 🔮 Versão 2.0 (Futuro)

- [ ] Upload automático de pedidos tratados
- [ ] Dashboard com métricas
- [ ] Integração com Telegram/WhatsApp
- [ ] Agendamento de downloads
- [ ] Tema claro/escuro

---

## 🤝 Contribuir

Contribuições são bem-vindas! Siga os passos:

1. Fork o projeto
2. Crie uma branch para sua feature (`git checkout -b feature/NovaFuncionalidade`)
3. Commit suas mudanças (`git commit -m 'feat: adiciona nova funcionalidade'`)
4. Push para a branch (`git push origin feature/NovaFuncionalidade`)
5. Abra um Pull Request

### Padrão de Commits

- `feat:` Nova funcionalidade
- `fix:` Correção de bug
- `refactor:` Refatoração de código
- `docs:` Documentação
- `style:` Formatação
- `test:` Testes
- `chore:` Manutenção

---

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.

---

## 📞 Contato

**Desenvolvedor:** Ediiiiipo

**GitHub:** [@Ediiiiipo](https://github.com/Ediiiiipo)

**Repositório:** [Shopee---Gerenciador-de-Pedidos](https://github.com/Ediiiiipo/Shopee---Gerenciador-de-Pedidos)

---

## 🙏 Agradecimentos

- Shopee Express pela plataforma
- Comunidade Electron
- Playwright Team
- Todos os contribuidores

---

<div align="center">

**⭐ Se este projeto te ajudou, deixe uma estrela!**

Made with ❤️ for Shopee Express logistics team

</div>