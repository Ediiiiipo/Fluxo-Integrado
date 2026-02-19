const { chromium } = require('playwright-core');
const { trocarStationCompleto } = require('./station-switcher-api');

async function testarAPI() {
  console.log('🧪 TESTE COM CAPTURA DE LOGS DO NAVEGADOR\n');
  
  let browser, context, page;
  
  try {
    const fs = require('fs-extra');
    const path = require('path');
    const os = require('os');
    
    // Detectar Chrome
    browser = await chromium.launch({
      headless: false,
      channel: 'chrome',
      args: ['--start-maximized']
    });
    
    // Carregar sessão
    const sessionFile = path.join(process.env.APPDATA || os.homedir(), 'shopee-manager', 'shopee_session.json');
    context = await browser.newContext({
      storageState: sessionFile,
      viewport: null
    });
    
    page = await context.newPage();
    
    // ⭐ CAPTURAR TODOS OS LOGS DO CONSOLE DO NAVEGADOR
    page.on('console', msg => {
      console.log(`[BROWSER] ${msg.text()}`);
    });
    
    console.log('🌐 Navegando...\n');
    await page.goto('https://spx.shopee.com.br/#/lmRouteCollectionPool', { waitUntil: 'networkidle' });
    await page.waitForTimeout(3000);
    
    console.log('🔄 Testando troca...\n');
    const sucesso = await trocarStationCompleto(page, 'LM Hub_GO_Goiânia_ St. Empr_02');
    
    console.log('\n' + (sucesso ? '✅ SUCESSO!' : '❌ FALHOU!'));
    
    await page.waitForTimeout(10000);
    
  } catch (error) {
    console.error('❌ Erro:', error.message);
  } finally {
    if (context) await context.close();
    if (browser) await browser.close();
  }
}

testarAPI().catch(console.error);
