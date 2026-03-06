// ============================================
// SHOPEE DOWNLOAD AUTOMÁTICO v8.0
// - USA NAVEGADOR DO SISTEMA (Chrome/Edge)
// - Validação de popup "Reset" (sessão duplicada)
// - Validação: Aguarda botão "Baixar" no popup
// - Seleção automática de Station
// - Gerenciamento de Stations via JSON
// - Navegação sempre visível para debug
// ============================================

const { chromium } = require('playwright-core');
const fs = require('fs-extra');
const path = require('path');
const os = require('os');
const XLSX = require('xlsx');
const AdmZip = require('adm-zip');

// ⭐ NOVO: Importar troca de station via API (método rápido)
const { trocarStationCompleto } = require('./station-switcher-api');

// ======================= DETECÇÃO DE NAVEGADOR DO SISTEMA =======================
let browserChannel = null;

/**
 * Detecta qual navegador do sistema está disponível
 * @returns {Promise<string|null>} 'chrome', 'msedge' ou null
 */
async function detectSystemBrowser() {
  console.log('🔍 Detectando navegador do sistema...');
  
  // Caminhos comuns do Chrome no Windows
  const chromePaths = [
    'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
    path.join(process.env.LOCALAPPDATA || '', 'Google\\Chrome\\Application\\chrome.exe'),
    path.join(process.env.PROGRAMFILES || '', 'Google\\Chrome\\Application\\chrome.exe'),
    path.join(process.env['PROGRAMFILES(X86)'] || '', 'Google\\Chrome\\Application\\chrome.exe')
  ];
  
  // Caminhos comuns do Edge no Windows
  const edgePaths = [
    'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
    'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
    path.join(process.env.PROGRAMFILES || '', 'Microsoft\\Edge\\Application\\msedge.exe'),
    path.join(process.env['PROGRAMFILES(X86)'] || '', 'Microsoft\\Edge\\Application\\msedge.exe')
  ];
  
  try {
    // Tentar Chrome primeiro
    console.log('   🔍 Procurando Google Chrome...');
    for (const chromePath of chromePaths) {
      if (await fs.pathExists(chromePath)) {
        console.log(`   ✅ Chrome encontrado em: ${chromePath}`);
        try {
          const browser = await chromium.launch({ channel: 'chrome', headless: true });
          await browser.close();
          console.log('✅ Google Chrome verificado e funcional!');
          return 'chrome';
        } catch (e) {
          console.log(`   ⚠️ Chrome encontrado mas não funcional: ${e.message}`);
        }
      }
    }
    
    // Tentar Edge
    console.log('   🔍 Procurando Microsoft Edge...');
    for (const edgePath of edgePaths) {
      if (await fs.pathExists(edgePath)) {
        console.log(`   ✅ Edge encontrado em: ${edgePath}`);
        try {
          const browser = await chromium.launch({ channel: 'msedge', headless: true });
          await browser.close();
          console.log('✅ Microsoft Edge verificado e funcional!');
          return 'msedge';
        } catch (e) {
          console.log(`   ⚠️ Edge encontrado mas não funcional: ${e.message}`);
        }
      }
    }
    
    console.error('❌ Nenhum navegador encontrado!');
    console.error('Por favor, instale Google Chrome ou Microsoft Edge:');
    console.error('Chrome: https://www.google.com/chrome/');
    console.error('Edge: Já vem instalado no Windows 10+');
    return null;
  } catch (error) {
    console.error(`❌ Erro ao detectar navegador: ${error.message}`);
    return null;
  }
}

// ======================= CONFIGURAÇÕES =======================
const CONFIG = {
  // URLs
  URL_LOGIN: 'https://fms.business.accounts.shopee.com.br/authenticate/login/?client_id=25&next=https%3A%2F%2Fspx.shopee.com.br%2Fapi%2Fadmin%2Fbasicserver%2Fops_tob_login%3Frefer%3Dhttps%3A%2F%2Fspx.shopee.com.br%2Faccount%2Flogin%23%2Findex&google_login_redirect=https%3A%2F%2Fspx.shopee.com.br%2Fapi%2Fadmin%2Fbasicserver%2Flogin%2F%3Fredirect%3Dhttps%3A%2F%2Fspx.shopee.com.br%2Faccount%2Flogin%23%2Findex',
  URL_HOME: 'https://spx.shopee.com.br/#/lmRouteCollectionPool',

  // Pastas
  SESSION_FILE: path.join(process.env.APPDATA || os.homedir(), 'shopee-manager', 'shopee_session.json'),
  DESKTOP_DIR: path.join(os.homedir(), 'Desktop'),
  TEMP_DIR: path.join(os.tmpdir(), 'shopee_temp'),

  // Timeouts
  TIMEOUT_DEFAULT: 30000,
  TIMEOUT_LOGIN: 600000,

  // Seletores
  SELECTORS: {
    btnExport: 'button:has-text("Export")',
    btnConfirmarExport: 'div.ssc-dialog-footer button:has-text("Export")',
    btnBaixar: 'button:has-text("Baixar")',
    statusWrapper: '.status-wrapper',
    statusText: '.status-wrapper',
    stationName: 'div[class*="header"] span span',
    popupReset: 'text=selecting reset will log you out',
    btnResetPopup: 'button:has-text("Reset")'
  }
};

// ======================= CLASSE PRINCIPAL =======================
class ShopeeDownloader {
  constructor() {
    this.browser = null;
    this.context = null;
    this.page = null;
    this.isLoggedIn = false;
    this.isHeadless = false;
    this.onProgress = null;
  }

  enviarProgresso(etapa, mensagem) {
    if (this.onProgress) {
      this.onProgress(etapa, mensagem);
    }
  }

  log(message, type = 'info') {
    const colors = {
      info: '\x1b[36m',
      success: '\x1b[32m',
      warning: '\x1b[33m',
      error: '\x1b[31m',
      reset: '\x1b[0m'
    };

    const color = colors[type] || colors.info;
    console.log(`${color}${message}${colors.reset}`);
  }

  // =================== NOVA FUNÇÃO: VERIFICAR E FECHAR POPUP DE RESET ===================
  async verificarEFecharPopupReset() {
    this.log('🔍 Verificando popup de Reset (sessão duplicada)...', 'info');
    
    try {
      await this.page.waitForTimeout(2000);

      const popupTexto = await this.page.locator('text=selecting reset will log you out').isVisible({ timeout: 3000 }).catch(() => false);
      
      if (popupTexto) {
        this.log('⚠️ Popup de sessão duplicada detectado!', 'warning');
        this.log('🔄 Clicando em "Reset" para encerrar sessão anterior...', 'info');
        
        try {
          await this.page.click('button:has-text("Reset")', { timeout: 5000 });
          this.log('✅ Popup de Reset fechado com sucesso!', 'success');
          
          await this.page.waitForTimeout(3000);
          await this.page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
          
          return true;
        } catch (e1) {
          this.log('⚠️ Tentativa 1 falhou, usando estratégia alternativa...', 'warning');
          
          const clicouViaJS = await this.page.evaluate(() => {
            const botoes = Array.from(document.querySelectorAll('button'));
            const btnReset = botoes.find(btn => 
              btn.innerText.trim().toLowerCase() === 'reset'
            );
            
            if (btnReset) {
              btnReset.click();
              return true;
            }
            return false;
          });
          
          if (clicouViaJS) {
            this.log('✅ Popup fechado via JavaScript!', 'success');
            await this.page.waitForTimeout(3000);
            await this.page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
            return true;
          }
          
          this.log('❌ Não foi possível fechar o popup automaticamente', 'error');
          return false;
        }
      } else {
        this.log('✅ Nenhum popup de Reset detectado', 'success');
        return true;
      }
      
    } catch (error) {
      this.log(`⚠️ Erro ao verificar popup de Reset: ${error.message}`, 'warning');
      return true;
    }
  }

  async verificarSeEstaLogado(headless = true, tentativa = 1, maxTentativas = 3) {
    this.log(`🕵️ Verificando status do login... (Tentativa ${tentativa}/${maxTentativas})`, 'info');

    let browserTest = null;
    let contextTest = null;

    try {
      // Verificar se arquivo de sessão existe
      const sessionExists = await fs.pathExists(CONFIG.SESSION_FILE);
      
      if (!sessionExists) {
        this.log('   📁 Primeiro acesso - sessão não existe', 'info');
        return false;
      }

      // Detectar navegador do sistema (apenas uma vez)
      if (!browserChannel) {
        browserChannel = await detectSystemBrowser();
        
        if (!browserChannel) {
          throw new Error('Nenhum navegador compatível encontrado no sistema');
        }
      }
      
      // Lançar navegador
      browserTest = await chromium.launch({
        headless: headless,
        channel: browserChannel,
        args: headless ? ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage'] : ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--start-maximized']
      });

      // Criar contexto com sessão salva
      contextTest = await browserTest.newContext({
        storageState: CONFIG.SESSION_FILE,
        viewport: null
      });

      const pageTest = await contextTest.newPage();

      // ⚡ OTIMIZAÇÃO: domcontentloaded ao invés de networkidle (muito mais rápido)
      await pageTest.goto(CONFIG.URL_HOME, {
        waitUntil: headless ? 'domcontentloaded' : 'networkidle',
        timeout: 60000 // Aumentado de 30s para 60s
      });

      // ⚡ OTIMIZAÇÃO: Timeout reduzido (3s → 1s quando headless)
      await pageTest.waitForTimeout(headless ? 1000 : 3000);

      const url = await pageTest.url();
      this.log(`   🔍 URL: ${url.substring(0, 60)}...`, 'info');

      if (!url.includes('login') && !url.includes('authenticate')) {
        this.log('✅ Login detectado automaticamente!', 'success');

        return {
          logado: true,
          browser: browserTest,
          context: contextTest,
          page: pageTest
        };
      } else {
        this.log('⚠️ Sessão expirada ou não logado.', 'warning');
        await contextTest.close();
        await browserTest.close();
        return {
          logado: false,
          browser: null,
          context: null,
          page: null
        };
      }

    } catch (error) {
      this.log(`   ⚠️ Erro ao verificar login: ${error.message}`, 'warning');
      
      // Limpar recursos
      if (contextTest) {
        try {
          await contextTest.close();
        } catch (e) {}
      }
      if (browserTest) {
        try {
          await browserTest.close();
        } catch (e) {}
      }
      
      // Detectar tipo de erro de conexão
      const errorMsg = error.message;
      const isDisconnected = errorMsg.includes('ERR_INTERNET_DISCONNECTED') || errorMsg.includes('net::ERR_NETWORK_CHANGED');
      const isTimeout = errorMsg.includes('Timeout');
      const isDNS = errorMsg.includes('ERR_NAME_NOT_RESOLVED');
      const isRefused = errorMsg.includes('ERR_CONNECTION_REFUSED');
      
      // Retry automático (exceto se sem internet)
      if (!isDisconnected && tentativa < maxTentativas) {
        this.log(`🔄 Tentando novamente... (${tentativa + 1}/${maxTentativas})`, 'info');
        await new Promise(resolve => setTimeout(resolve, 3000)); // Aguardar 3s antes de tentar novamente
        return this.verificarSeEstaLogado(headless, tentativa + 1, maxTentativas);
      }
      
      // Mensagens específicas por tipo de erro
      this.log('', 'error');
      this.log('═'.repeat(70), 'error');
      
      if (isDisconnected) {
        this.log('❌ ERRO: SEM CONEXÃO COM A INTERNET', 'error');
        this.log('', 'error');
        this.log('📋 O que aconteceu:', 'info');
        this.log('   • Sua conexão com a internet foi perdida', 'info');
        this.log('   • VPN pode ter desconectado durante o processo', 'info');
        this.log('   • Rede Wi-Fi/Ethernet instável', 'info');
        this.log('', 'error');
        this.log('💡 Solução:', 'warning');
        this.log('   1. Verifique sua conexão com a internet', 'warning');
        this.log('   2. Reconecte à VPN (se usar)', 'warning');
        this.log('   3. Tente novamente após estabilizar a conexão', 'warning');
      } else if (isDNS) {
        this.log('❌ ERRO: NÃO FOI POSSÍVEL ENCONTRAR O SERVIDOR', 'error');
        this.log('', 'error');
        this.log('📋 O que aconteceu:', 'info');
        this.log('   • Problema de DNS (servidor de nomes)', 'info');
        this.log('   • Firewall bloqueando acesso ao SPX', 'info');
        this.log('', 'error');
        this.log('💡 Solução:', 'warning');
        this.log('   1. Verifique se consegue acessar spx.shopee.com.br no navegador', 'warning');
        this.log('   2. Desative temporariamente antivírus/firewall', 'warning');
        this.log('   3. Tente usar outra rede (4G/5G)', 'warning');
      } else if (isRefused) {
        this.log('❌ ERRO: CONEXÃO RECUSADA PELO SERVIDOR', 'error');
        this.log('', 'error');
        this.log('📋 O que aconteceu:', 'info');
        this.log('   • O servidor SPX recusou a conexão', 'info');
        this.log('   • Pode estar em manutenção', 'info');
        this.log('', 'error');
        this.log('💡 Solução:', 'warning');
        this.log('   1. Aguarde alguns minutos e tente novamente', 'warning');
        this.log('   2. Verifique se o SPX está acessível no navegador', 'warning');
      } else if (isTimeout) {
        this.log('❌ ERRO: TEMPO LIMITE EXCEDIDO (TIMEOUT)', 'error');
        this.log('', 'error');
        this.log('📋 O que aconteceu:', 'info');
        this.log('   • Conexão muito lenta', 'info');
        this.log('   • Servidor SPX não respondeu a tempo', 'info');
        this.log('', 'error');
        this.log('💡 Solução:', 'warning');
        this.log('   1. Verifique a velocidade da sua internet', 'warning');
        this.log('   2. Feche outros programas que usam internet', 'warning');
        this.log('   3. Tente em outro horário (menos congestionado)', 'warning');
      } else {
        this.log('❌ ERRO: FALHA AO ACESSAR A API DA SHOPEE', 'error');
        this.log('', 'error');
        this.log('📋 Possíveis causas:', 'info');
        this.log('   1. Você não está logado na Shopee', 'info');
        this.log('   2. Sessão expirou', 'info');
        this.log('   3. Problema temporário no servidor', 'info');
        this.log('', 'error');
        this.log('💡 Solução:', 'warning');
        this.log('   1. Feche e abra a ferramenta novamente', 'warning');
        this.log('   2. Faça login manual (aba Download)', 'warning');
      }
      
      this.log('═'.repeat(70), 'error');
      this.log('', 'error');
      
      return {
        logado: false,
        browser: null,
        context: null,
        page: null
      };
    }
  }

  async realizarLoginManual() {
    this.log('\n🛑 LOGIN NECESSÁRIO! Abrindo navegador para você logar... 🛑', 'warning');

    // Detectar navegador do sistema (apenas uma vez)
    if (!browserChannel) {
      browserChannel = await detectSystemBrowser();
      
      if (!browserChannel) {
        throw new Error('Nenhum navegador compatível encontrado no sistema');
      }
    }
    
    // Lançar navegador
    this.browser = await chromium.launch({
      headless: false,
      channel: browserChannel,
      args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--start-maximized']
    });

    // Criar contexto novo (sem sessão)
    this.context = await this.browser.newContext({
      viewport: null
    });

    this.page = await this.context.newPage();
    this.isHeadless = false;

    await this.page.goto(CONFIG.URL_HOME, {
      waitUntil: 'networkidle',
      timeout: 60000
    });

    await this.verificarEFecharPopupReset();

    console.log('');
    console.log('═'.repeat(70));
    this.log('🚨 BOT PAUSADO - AGUARDANDO LOGIN MANUAL 🚨', 'warning');
    console.log('═'.repeat(70));
    console.log('');
    console.log('👤 FAÇA LOGIN NO NAVEGADOR:');
    console.log('   1. Clique em "Login with Google"');
    console.log('   2. Selecione sua conta Google');
    console.log('   3. Complete a validação no celular (se necessário)');
    console.log('   4. Aguarde carregar a página do SPX');
    console.log('');
    console.log('⏰ Tempo máximo: 10 minutos');
    console.log('🤖 O bot continuará automaticamente após login');
    console.log('');
    console.log('═'.repeat(70));
    console.log('');

    const seletoresLogado = [
      '.ssc-layout-sider',
      '.sidebar',
      '#fms-container',
      '.menu-container',
      '.ssc-menu',
      '[class*="layout-sider"]',
      '[class*="sidebar"]'
    ];

    let segundos = 0;
    const maxSegundos = 600;

    while (segundos < maxSegundos) {
      try {
        const minutos = Math.floor(segundos / 60);
        const segs = segundos % 60;
        process.stdout.write(
          `\r⏳ Aguardando login... ${String(minutos).padStart(2, '0')}:${String(segs).padStart(2, '0')} / 10:00`
        );

        for (const seletor of seletoresLogado) {
          try {
            const elemento = await this.page.locator(seletor).first();
            const visivel = await elemento.isVisible({
              timeout: 500
            }).catch(() => false);

            if (visivel) {
              process.stdout.write('\r' + ' '.repeat(80) + '\r');
              console.log('');
              this.log('✅ Login manual detectado!', 'success');

              this.log('🛡️ Verificando se há popup de aviso...', 'info');
              
              await this.page.waitForTimeout(2000);

              try {
                const seletorPopup = '.ssc-dialog-close-icon-wrapper';
                
                await this.page.click(seletorPopup, { timeout: 3000 });
                
                this.log('🧹 Popup fechado com sucesso!', 'success');
                await this.page.waitForTimeout(1000);
              } catch (e) {
                this.log('👌 Nenhum popup obstrutivo detectado.', 'info');
              }

              this.log('🎉 LOGIN CONCLUÍDO COM SUCESSO!', 'success');
              
              // Salvar sessão
              await fs.ensureDir(path.dirname(CONFIG.SESSION_FILE));
              await this.context.storageState({ path: CONFIG.SESSION_FILE });
              this.log('✅ Sessão salva em: shopee_session.json', 'success');
              this.log('💡 Próximas execuções usarão sessão salva', 'info');
              console.log('');

              this.isLoggedIn = true;

              await this.page.waitForTimeout(2000);

              return true;
            }
          } catch (e) {}
        }

        const urlAtual = this.page.url();
        if (urlAtual.includes('spx.shopee.com.br') &&
          !urlAtual.includes('login') &&
          !urlAtual.includes('authenticate')) {

          const temConteudo = await this.page.evaluate(() => {
            return document.body.innerText.length > 500;
          }).catch(() => false);

          if (temConteudo) {
            process.stdout.write('\r' + ' '.repeat(80) + '\r');
            console.log('');
            this.log('✅ Login detectado pela URL!', 'success');
            this.log('🎉 LOGIN CONCLUÍDO COM SUCESSO!', 'success');
            
            // Salvar sessão
            await fs.ensureDir(path.dirname(CONFIG.SESSION_FILE));
            await this.context.storageState({ path: CONFIG.SESSION_FILE });
            this.log('✅ Sessão salva em: shopee_session.json', 'success');
            console.log('');

            this.isLoggedIn = true;
            await this.page.waitForTimeout(2000);
            return true;
          }
        }

        await this.page.waitForTimeout(1500);
        segundos += 1.5;

      } catch (error) {
        if (error.message.includes('Target closed') || error.message.includes('destroyed')) {
          console.log('');
          this.log('❌ Navegador fechado pelo usuário.', 'error');
          return false;
        }

        await this.page.waitForTimeout(1500);
        segundos += 1.5;
      }
    }

    console.log('');
    this.log('❌ TIMEOUT: Login não concluído em 10 minutos', 'error');
    return false;
  }

  async initialize(headless = false) {
    this.log('🚀 Iniciando navegador...', 'info');

    try {
      await fs.ensureDir(CONFIG.TEMP_DIR);
    } catch (error) {
      console.error('❌ ERRO ao criar diretório temporário:', error);
      throw error;
    }

    // 🔍 DETECTAR SE TEM SESSÃO SALVA
    const temSessao = await fs.pathExists(CONFIG.SESSION_FILE);
    
    // ⚡ MODO INTELIGENTE: Respeita preferência do usuário (CTRL+U)
    // Só ativa headless automaticamente se usuário não definiu preferência
    let modoHeadless = headless;
    
    if (!temSessao) {
      // Primeira vez: sempre visível para login
      this.log('🔓 Primeira vez - modo visível para login', 'info');
      modoHeadless = false;
    } else {
      // Tem sessão: usa a preferência passada do main.js (que vem do localStorage)
      // O main.js já lê a preferência do CTRL+U, então só usar o valor que veio
      if (modoHeadless) {
        this.log('⚡ Modo rápido (headless) ativado', 'info');
      } else {
        this.log('👁️ Modo visual ativado', 'info');
      }
    }

    // SEMPRE tentar carregar sessão primeiro
    let resultado;
    try {
      resultado = await this.verificarSeEstaLogado(modoHeadless);
    } catch (error) {
      console.error('⚠️ Erro ao verificar sessão:', error.message);
      resultado = { logado: false };
    }

    if (resultado && resultado.logado && resultado.browser && resultado.context && resultado.page) {
      this.log('🔒 Sessão restaurada com sucesso!', 'success');
      this.browser = resultado.browser;
      this.context = resultado.context;
      this.page = resultado.page;
      this.isHeadless = modoHeadless;
      this.isLoggedIn = true;
      
      // ⚡ OTIMIZAÇÃO: Bloquear recursos pesados se headless
      if (modoHeadless) {
        await this.otimizarCarregamento();
      }
      
      await this.verificarEFecharPopupReset();
      
      this.log('✅ Navegador pronto (sessão reutilizada)', 'success');
      return this.page;
    } else {
      this.log('⚠️ Sessão não encontrada ou expirada - fazendo login...', 'warning');
    }

    let loginOk;
    try {
      loginOk = await this.realizarLoginManual();
    } catch (error) {
      console.error('❌ ERRO em realizarLoginManual():', error);
      console.error('Stack:', error.stack);
      throw error;
    }

    if (!loginOk) {
      throw new Error('Login não concluído');
    }

    return this.page;
  }

  // ⚡ NOVO MÉTODO: Otimizar carregamento (bloquear recursos pesados)
  async otimizarCarregamento() {
    if (!this.page) return;
    
    try {
      await this.page.route('**/*', (route) => {
        const resourceType = route.request().resourceType();
        
        // Bloquear: imagens, fontes, mídia
        if (['image', 'font', 'media', 'stylesheet'].includes(resourceType)) {
          route.abort();
        } else {
          route.continue();
        }
      });
      
      this.log('⚡ Otimização ativada: bloqueando recursos pesados', 'info');
    } catch (error) {
      // Ignorar erro silenciosamente se já tiver rotas configuradas
    }
  }

  async checkAndWaitForLogin() {
    if (this.isLoggedIn) {
      this.log('✅ Já está logado!', 'success');
      return true;
    }

    this.log('❌ Não está logado!', 'error');
    return false;
  }

  // ⭐ NOVO MÉTODO: Troca de station via API (75% mais rápido!)
  async selecionarStation(nomeStation) {
    if (!nomeStation) {
      throw new Error('Nome da station não fornecido');
    }
    
    this.log(`🏬 Selecionando station: ${nomeStation}`, 'info');
    
    try {
      // ⭐ MÉTODO 1: Tentar via API primeiro (RÁPIDO - ~2 segundos)
      this.log('🚀 Tentando via API (método rápido)...', 'info');
      
      const sucessoAPI = await trocarStationCompleto(this.page, nomeStation);
      
      if (sucessoAPI) {
        this.log('✅ Station trocada via API com sucesso!', 'success');
        return true;
      }
      
      // ⚠️ MÉTODO 2: Se API falhar, usar DOM (LENTO - ~8 segundos)
      this.log('⚠️ API falhou, tentando método DOM (fallback)...', 'warning');
      return await this.selecionarStationDOM(nomeStation);
      
    } catch (error) {
      this.log(`❌ Erro ao selecionar station: ${error.message}`, 'error');
      return false;
    }
  }

  // Método antigo renomeado (usado como fallback)
  async selecionarStationDOM(nomeStation) {
    if (!nomeStation) {
      throw new Error('Nome da station não fornecido');
    }
    this.log(`🏬 Selecionando station: ${nomeStation}`, 'info');

    try {
      await this.verificarEFecharPopupReset();

      this.log('   1️⃣ Abrindo dropdown de station...', 'info');

      const clicouDropdown = await this.page.evaluate(() => {
        const elementos = Array.from(document.querySelectorAll('div, span'));

        for (const el of elementos) {
          const texto = (el.innerText || '').trim();
          const rect = el.getBoundingClientRect();

          const isStationName = texto.includes('LM Hub') || texto.includes('SoC_');
          const isInSidebar = rect.left < 250;
          const isInTopArea = rect.top > 40 && rect.top < 150;
          const hasReasonableSize = rect.width > 100 && rect.width < 250 && rect.height > 20 && rect.height < 50;
          const isNotSearchBox = !texto.toLowerCase().includes('busca') && !el.querySelector('input');

          if (isStationName && isInSidebar && isInTopArea && hasReasonableSize && isNotSearchBox) {
            el.click();
            return {
              sucesso: true,
              texto: (texto || '').replace(/\n/g, ' ').substring(0, 50)
            };
          }
        }

        for (const el of elementos) {
          const texto = (el.innerText || '').trim();
          const rect = el.getBoundingClientRect();

          if ((texto.includes('LM Hub') || texto.includes('SoC_')) &&
            rect.left < 250 && rect.top < 200 && rect.top > 30 &&
            texto.length < 40 && !texto.toLowerCase().includes('busca')) {
            el.click();
            return {
              sucesso: true,
              texto: (texto || '').replace(/\n/g, ' ').substring(0, 50)
            };
          }
        }

        return {
          sucesso: false
        };
      });

      if (clicouDropdown.sucesso) {
        this.log(`   ✔ Dropdown aberto (atual: "${clicouDropdown.texto}")`, 'success');
      } else {
        throw new Error('Não encontrou o seletor de station');
      }

      await this.page.waitForTimeout(1500);

      this.log('   2️⃣ Filtrando por nome...', 'info');

      try {
        const inputProcurar = this.page.locator('input[placeholder="Procurar por"]');
        await inputProcurar.waitFor({
          state: 'visible',
          timeout: 3000
        });
        await inputProcurar.click();
        await inputProcurar.fill(nomeStation);
        this.log(`   ✔ Filtro aplicado: "${nomeStation}"`, 'success');
      } catch (e) {
        await this.page.evaluate((stationName) => {
          const inputs = document.querySelectorAll('input[placeholder="Procurar por"], input[placeholder*="procurar"]');
          if (inputs.length > 0) {
            inputs[0].value = stationName;
            inputs[0].dispatchEvent(new Event('input', {
              bubbles: true
            }));
          }
        }, nomeStation);
        this.log('   ✔ Filtro aplicado via JS', 'success');
      }

      await this.page.waitForTimeout(1500);

      this.log('   3️⃣ Selecionando station na lista...', 'info');

      let stationClicada = false;

      try {
        const seletor = `li.ssc-option[title="${nomeStation}"]`;
        const elemento = this.page.locator(seletor);
        await elemento.waitFor({
          state: 'visible',
          timeout: 3000
        });
        await elemento.click();
        stationClicada = true;
        this.log(`   ✔ Clicado via seletor: ${seletor}`, 'success');
      } catch (e) {
        this.log('   ⚠️ Seletor exato não encontrado', 'warning');
      }

      if (!stationClicada) {
        try {
          const seletor = `li.ssc-option[title*="${nomeStation.split('_').slice(-1)[0]}"]`;
          const elemento = this.page.locator(seletor).first();
          await elemento.waitFor({
            state: 'visible',
            timeout: 2000
          });
          await elemento.click();
          stationClicada = true;
          this.log(`   ✔ Clicado via seletor parcial`, 'success');
        } catch (e) {
          this.log('   ⚠️ Seletor parcial não encontrado', 'warning');
        }
      }

      if (!stationClicada) {
        this.log('   ⚠️ Tentando via JavaScript...', 'warning');

        const resultado = await this.page.evaluate((stationName) => {
          const normalizar = (texto) => {
            if (!texto) return '';
            return texto
              .normalize('NFD')
              .replace(/[\u0300-\u036f]/g, '')
              .toLowerCase()
              .trim();
          };

          const stationNameNorm = normalizar(stationName);
          const listaOpcoes = document.querySelectorAll('li.ssc-option');

          for (const li of listaOpcoes) {
            const title = li.getAttribute('title') || '';
            if (title === stationName) {
              li.click();
              return {
                sucesso: true,
                metodo: 'title exato',
                title
              };
            }
          }

          for (const li of listaOpcoes) {
            const title = li.getAttribute('title') || '';
            if (normalizar(title) === stationNameNorm) {
              li.click();
              return {
                sucesso: true,
                metodo: 'normalizado',
                title
              };
            }
          }

          for (const li of listaOpcoes) {
            const title = li.getAttribute('title') || '';
            const titleNorm = normalizar(title);
            if (titleNorm.includes(stationNameNorm) || stationNameNorm.includes(titleNorm)) {
              li.click();
              return {
                sucesso: true,
                metodo: 'parcial',
                title
              };
            }
          }

          const partes = stationName.split('_');
          const cidade = partes[2] || '';
          const cidadeNorm = normalizar(cidade);

          for (const li of listaOpcoes) {
            const title = li.getAttribute('title') || '';
            if (cidadeNorm && normalizar(title).includes(cidadeNorm)) {
              li.click();
              return {
                sucesso: true,
                metodo: 'cidade',
                title
              };
            }
          }

          const opcoes = Array.from(listaOpcoes).map(li => li.getAttribute('title')).slice(0, 5);
          return {
            sucesso: false,
            totalOpcoes: listaOpcoes.length,
            primeiras: opcoes
          };
        }, nomeStation);

        if (resultado.sucesso) {
          stationClicada = true;
          this.log(`   ✔ Clicado via JS (${resultado.metodo}): "${resultado.title}"`, 'success');
        } else {
          this.log(`   ⚠️ Station não encontrada (${resultado.totalOpcoes} opções)`, 'warning');
          if (resultado.primeiras) {
            this.log(`   📋 Primeiras: ${resultado.primeiras.join(', ')}`, 'info');
          }
        }
      }

      if (!stationClicada) {
        throw new Error(`Station "${nomeStation}" não encontrada`);
      }

      this.log('   4️⃣ Aguardando página atualizar...', 'info');
      await this.page.waitForTimeout(2000);

      this.log('   🔄 Navegando para página de pedidos...', 'info');
      await this.page.goto(CONFIG.URL_HOME, {
        waitUntil: 'networkidle',
        timeout: 30000
      });
      await this.page.waitForTimeout(2000);

      this.log('   5️⃣ Validando troca...', 'info');

      const stationAtual = await this.page.evaluate(() => {
        const elementos = document.querySelectorAll('div, span');

        for (const el of elementos) {
          const texto = (el.innerText || '').trim();
          const rect = el.getBoundingClientRect();

          if ((texto.includes('LM Hub') || texto.includes('SoC_')) &&
            rect.left < 250 && rect.top > 40 && rect.top < 150 &&
            texto.length < 50 && !texto.includes('\n')) {
            return texto;
          }
        }
        return null;
      });

      if (stationAtual) {
        this.log(`   ✅ STATION CONFIRMADA: ${stationAtual}`, 'success');
      } else {
        this.log(`   ✅ Clique executado, continuando...`, 'success');
      }

      return true;

    } catch (error) {
      this.log(`   ❌ Erro ao selecionar station: ${error.message}`, 'error');

      try {
        const errorPath = path.join(CONFIG.TEMP_DIR, 'erro_station.png');
        await this.page.screenshot({
          path: errorPath,
          fullPage: true
        });
        this.log(`   📸 Screenshot: ${errorPath}`, 'info');
      } catch (e) {}

      return false;
    }
  }

  async getStationName() {
    try {
      const element = await this.page.locator(CONFIG.SELECTORS.stationName).first();
      await element.waitFor({
        timeout: 5000
      });
      const name = await element.innerText();
      this.log(`✅ Station: ${name}`, 'success');
      return name.trim();
    } catch (error) {
      this.log('⚠️ Não foi possível obter nome da station', 'warning');
      return 'Station_Desconhecida';
    }
  }

  async downloadReport() {
    this.log('\n📥 INICIANDO DOWNLOAD VIA API', 'info');
    console.log('═'.repeat(70));

    await this.verificarEFecharPopupReset();

    this.log('1️⃣ Disparando exportação via API...', 'info');
    this.enviarProgresso(3, 'Disparando exportação via API...');

    // Executar apenas disparo e polling dentro do page.evaluate
    const fileUrl = await this.page.evaluate(async () => {
      const getCookie = (name) => {
        const match = document.cookie.match(new RegExp('(?:^|;\\s*)' + name + '=([^;]*)'));
        return match ? match[1] : '';
      };
      const headers = {
        'accept': 'application/json, text/plain, */*',
        'app': 'FMS Portal',
        'x-csrftoken': getCookie('csrftoken'),
        'device-id': getCookie('spx-admin-device-id'),
      };

      // Step 1: Disparar exportação
      console.log('🚀 Disparando exportação...');
      const exportRes = await fetch(
        '/spx_delivery/admin/delivery/route/collection_pool/list/export?min_length=-1&max_length=-1&min_width=-1&max_width=-1&min_height=-1&max_height=-1&min_weight=-1&max_weight=-1&min_any_dimension=-1&max_any_dimension=-1',
        { headers, credentials: 'include' }
      );
      const exportData = await exportRes.json();
      console.log('📦 Resposta export:', exportData);

      // Step 2: Polling
      const startTs = Math.floor(Date.now() / 1000) - 60;
      console.log('⏳ Aguardando processamento...');

      let fileUrl = null;
      for (let i = 1; i <= 120; i++) {
        await new Promise(r => setTimeout(r, i <= 10 ? 2000 : 5000));

        const res = await fetch(
          `/spxdata/api/export_platform/export_task/list_for_portal?start_time=${startTs}&count=5&pageno=1`,
          { headers, credentials: 'include' }
        );
        const data = await res.json();
        const tasks = data?.data?.task_list || [];

        console.log(`🔄 Polling #${i} — ${tasks.length} tarefa(s)`);

        const ready = tasks.find(t => t.export_status === 2);
        if (ready) {
          console.log(`✅ Pronto! Task ID: ${ready.task_id}`);
          fileUrl = `/${ready.file_name}`;
          break;
        }
      }

      if (!fileUrl) {
        throw new Error('Timeout: tarefa não ficou pronta');
      }

      return fileUrl;
    });

    this.log('2️⃣ Baixando arquivo via Playwright...', 'info');
    this.enviarProgresso(4, 'Baixando arquivo...');

    // Fazer download do arquivo usando Playwright
    const [download] = await Promise.all([
      this.page.waitForEvent('download', { timeout: 60000 }),
      this.page.evaluate((url) => {
        const a = document.createElement('a');
        a.href = url;
        a.download = '';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
      }, fileUrl)
    ]);

    const fileName = download.suggestedFilename() || fileUrl.split('/').pop();
    const tempPath = path.join(CONFIG.TEMP_DIR, fileName);
    await download.saveAs(tempPath);

    this.log(`   ✅ Arquivo baixado: ${fileName}`, 'success');

    // Se for ZIP, processar localmente (Node.js)
    if (fileName.endsWith('.zip')) {
      this.log('3️⃣ Processando arquivo ZIP...', 'info');
      this.enviarProgresso(4, 'Unificando arquivos...');

      const zip = new AdmZip(tempPath);
      const zipEntries = zip.getEntries().filter(e => e.entryName.endsWith('.xlsx'));

      this.log(`   📂 ${zipEntries.length} arquivo(s) XLSX encontrados`, 'info');

      let allRows = [];
      let headerRow = null;

      for (let idx = 0; idx < zipEntries.length; idx++) {
        const entry = zipEntries[idx];
        this.log(`   📄 Processando: ${entry.entryName}`, 'info');

        const xlsxData = entry.getData();
        const workbook = XLSX.read(xlsxData, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        if (idx === 0) {
          headerRow = rows[0];
          allRows = allRows.concat(rows);
        } else {
          const dataOnly = rows.slice(1);
          allRows = allRows.concat(dataOnly);
        }
      }

      this.log(`   📊 Total unificado: ${allRows.length} linhas`, 'success');

      // Gerar XLSX unificado
      const newSheet = XLSX.utils.aoa_to_sheet(allRows);
      const newWorkbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Collection Pool');

      const outputName = fileName.replace('.zip', '_UNIFICADO.xlsx');
      const outputPath = path.join(CONFIG.TEMP_DIR, outputName);
      XLSX.writeFile(newWorkbook, outputPath);

      this.log(`   ✅ Arquivo unificado salvo: ${outputName}`, 'success');
      return outputPath;
    }

    return tempPath;
  }

  async processFile(filePath, stationName) {
    this.log('\n📊 PROCESSANDO ARQUIVO', 'info');
    console.log('═'.repeat(70));

    const outputDir = path.join(CONFIG.DESKTOP_DIR, stationName);
    await fs.ensureDir(outputDir);
    this.log(`📁 Pasta criada: ${outputDir}`, 'info');

    if (filePath.toLowerCase().endsWith('.zip')) {
      this.log('📦 Descompactando e unificando...', 'info');

      const extractionPath = path.join(CONFIG.TEMP_DIR, 'extraidos');
      await fs.ensureDir(extractionPath);
      await fs.emptyDir(extractionPath);

      const zip = new AdmZip(filePath);
      zip.extractAllTo(extractionPath, true);

      const arquivos = (await fs.readdir(extractionPath))
        .filter(f => f.endsWith('.xlsx') || f.endsWith('.xls'))
        .map(f => path.join(extractionPath, f));

      let dadosUnificados = [];

      for (const arquivo of arquivos) {
        const workbook = XLSX.readFile(arquivo);
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        const dataFiltrada = data.filter(row => (row['Destination Hub'] || '').toString().startsWith('LM Hub'));
        dadosUnificados = dadosUnificados.concat(dataFiltrada);
      }

      this.log(`✔ ${arquivos.length} arquivos unificados`, 'success');
      this.log(`✔ ${dadosUnificados.length} registros totais`, 'success');

      const finalPath = path.join(outputDir, `Relatorio_${new Date().toISOString().split('T')[0]}.xlsx`);

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(dadosUnificados);
      XLSX.utils.book_append_sheet(wb, ws, 'Unificado');
      XLSX.writeFile(wb, finalPath);

      this.log(`✅ Arquivo salvo: ${finalPath}`, 'success');

      return {
        filePath: finalPath,
        totalRecords: dadosUnificados.length,
        outputDir
      };
    }

    const workbook = XLSX.readFile(filePath);
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
    const dataFiltrada = data.filter(row => (row['Destination Hub'] || '').toString().startsWith('LM Hub'));

    const finalPath = path.join(outputDir, `Relatorio_${new Date().toISOString().split('T')[0]}.xlsx`);

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(dataFiltrada);
    XLSX.utils.book_append_sheet(wb, ws, 'Filtrado');
    XLSX.writeFile(wb, finalPath);

    this.log(`✅ Arquivo salvo: ${finalPath}`, 'success');

    return {
      filePath: finalPath,
      totalRecords: dataFiltrada.length,
      outputDir
    };
  }

  async run(headless = true, stationParaTrocar = null) {
    try {
      this.enviarProgresso(1, 'Iniciando navegador...');
      await this.initialize(headless);

      console.log('');
      console.log('═'.repeat(70));
      this.log('ETAPA 1: LOGIN', 'info');
      console.log('═'.repeat(70));

      this.enviarProgresso(1, 'Verificando login...');
      const loginOk = await this.checkAndWaitForLogin();
      if (!loginOk) {
        throw new Error('Login não concluído');
      }

      this.enviarProgresso(2, 'Navegando para página de pedidos...');
      await this.page.waitForTimeout(3000);

      console.log('');
      console.log('═'.repeat(70));
      this.log('ETAPA 2: NAVEGAÇÃO', 'info');
      console.log('═'.repeat(70));

      await this.page.goto(CONFIG.URL_HOME, {
        waitUntil: 'networkidle'
      });

      await this.verificarEFecharPopupReset();

      await this.page.waitForTimeout(3000);

      try {
        await this.page.keyboard.press('Escape');
        await this.page.waitForTimeout(500);
        await this.page.keyboard.press('Escape');
      } catch (e) {}

      if (stationParaTrocar) {
        console.log('');
        console.log('═'.repeat(70));
        this.log('ETAPA 2.1: SELEÇÃO DE STATION', 'info');
        console.log('═'.repeat(70));

        this.enviarProgresso(2, `Selecionando station: ${stationParaTrocar}`);
        this.log(`🏢 Trocando para station: ${stationParaTrocar}`, 'info');

        // Retry automático (3 tentativas)
        let trocouStation = false;
        const maxTentativas = 3;
        
        for (let tentativa = 1; tentativa <= maxTentativas; tentativa++) {
          if (tentativa > 1) {
            this.log(`🔄 Tentativa ${tentativa}/${maxTentativas} de trocar station...`, 'info');
            await this.page.waitForTimeout(2000); // Aguardar 2s entre tentativas
          }
          
          trocouStation = await this.selecionarStation(stationParaTrocar);
          
          if (trocouStation) {
            break; // Sucesso!
          }
        }

        if (!trocouStation) {
          this.log('❌ ERRO: Não foi possível trocar para a station após 3 tentativas!', 'error');
          this.log('', 'error');
          this.log('📋 Possíveis causas:', 'error');
          this.log('   1. A station não existe ou foi renomeada no SPX', 'error');
          this.log('   2. Você não tem permissão para acessar essa station', 'error');
          this.log('   3. O seletor da station mudou (atualize a ferramenta)', 'error');
          this.log('   4. Conexão lenta ou instável', 'error');
          this.log('', 'error');
          this.log('💡 Solução:', 'warning');
          this.log('   - Verifique se a station existe no SPX', 'warning');
          this.log('   - Tente trocar manualmente para a station no SPX', 'warning');
          this.log('   - Feche e abra a ferramenta novamente', 'warning');
          this.log('⚠️ Download cancelado.', 'warning');
          throw new Error(`Falha ao trocar para station: ${stationParaTrocar}`);
        }

        this.log('✅ Station trocada com sucesso!', 'success');
        await this.page.waitForTimeout(2000);
      }

      const stationName = stationParaTrocar || await this.getStationName();
      this.log(`📁 Nome da pasta: ${stationName}`, 'info');

      console.log('');
      console.log('═'.repeat(70));
      this.log('ETAPA 3: DOWNLOAD', 'info');
      console.log('═'.repeat(70));

      this.enviarProgresso(3, 'Abrindo menu de exportação...');
      const downloadedFile = await this.downloadReport();

      console.log('');
      console.log('═'.repeat(70));
      this.log('ETAPA 4: PROCESSAMENTO', 'info');
      console.log('═'.repeat(70));

      this.enviarProgresso(5, 'Processando arquivo...');
      const result = await this.processFile(downloadedFile, stationName);

      this.enviarProgresso(6, 'Concluído!');
      console.log('');
      console.log('═'.repeat(70));
      this.log('✅ PROCESSO CONCLUÍDO COM SUCESSO!', 'success');
      console.log('═'.repeat(70));
      console.log('');
      this.log(`📁 Arquivos em: ${result.outputDir}`, 'info');
      this.log(`📊 Total de pedidos: ${result.totalRecords}`, 'info');
      console.log('');

      await this.close();

      return {
        success: true,
        stationName,
        totalRecords: result.totalRecords,
        excelPath: result.filePath,
        filePath: result.filePath,
        outputDir: result.outputDir
      };

    } catch (error) {
      console.log('');
      this.log(`❌ ERRO: ${error.message}`, 'error');

      try {
        const errorPath = path.join(CONFIG.TEMP_DIR, 'erro.png');
        await this.page.screenshot({
          path: errorPath
        });
        this.log(`📸 Screenshot: ${errorPath}`, 'info');
      } catch (e) {}

      await this.close();

      return {
        success: false,
        error: error.message
      };
    }
  }

  async close() {
    if (this.context) {
      await this.context.close();
    }
    if (this.browser) {
      await this.browser.close();
      this.log('🔒 Navegador fechado', 'info');
    }
  }
}

async function main(stationName = null) {
  const downloader = new ShopeeDownloader();
  const result = await downloader.run(false, stationName);

  if (result.success) {
    console.log('\n🎉 SUCESSO TOTAL!');
    console.log(`Station: ${result.stationName}`);
    console.log(`Pedidos: ${result.totalRecords}`);
    console.log(`Excel: ${result.excelPath}`);
  } else {
    console.log(`\n❌ FALHA: ${result.error}`);
  }

  return result;
}

async function baixarMultiplasStations(listaStations) {
  console.log(`\n🏢 Iniciando download de ${listaStations.length} stations...`);

  const downloader = new ShopeeDownloader();
  const resultados = [];

  for (let i = 0; i < listaStations.length; i++) {
    const station = listaStations[i];

    console.log(`\n┌${'─'.repeat(70)}┐`);
    console.log(`│ 📍 Station ${i + 1}/${listaStations.length}: ${station}`);
    console.log(`└${'─'.repeat(70)}┘`);

    try {
      const result = await downloader.run(false, station);
      resultados.push({
        station,
        success: result.success,
        totalRecords: result.totalRecords || 0,
        error: result.error
      });
    } catch (error) {
      resultados.push({
        station,
        success: false,
        totalRecords: 0,
        error: error.message
      });
    }

    if (i < listaStations.length - 1) {
      console.log('\n⏳ Aguardando 5 segundos...');
      await new Promise(resolve => setTimeout(resolve, 5000));
    }
  }

  await downloader.close();

  console.log('\n');
  console.log('═'.repeat(70));
  console.log('📊 RESUMO FINAL');
  console.log('═'.repeat(70));

  const sucesso = resultados.filter(r => r.success).length;
  const falhas = resultados.filter(r => !r.success).length;
  const totalPedidos = resultados.reduce((sum, r) => sum + r.totalRecords, 0);

  console.log(`✅ Sucesso: ${sucesso}`);
  console.log(`❌ Falhas: ${falhas}`);
  console.log(`📦 Total: ${totalPedidos}`);

  return resultados;
}

if (require.main === module) {
  main('LM Hub_MG_Belo Horizonte_02').catch(console.error);
}

// Exportação principal (formato simples para compatibilidade com main.js)
module.exports = ShopeeDownloader;

// Exportações adicionais (se necessário)
module.exports.CONFIG = CONFIG;
module.exports.main = main;
module.exports.baixarMultiplasStations = baixarMultiplasStations;