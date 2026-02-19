// ============================================
// TESTE RÁPIDO - Station Switcher API
// ============================================

const { chromium } = require('playwright-core');
const { trocarStationCompleto } = require('./station-switcher-api');

async function testarAPI() {
  console.log('');
  console.log('═'.repeat(70));
  console.log('🧪 TESTE DA API DE TROCA DE STATION');
  console.log('═'.repeat(70));
  console.log('');
  
  let browser, context, page;
  
  try {
    // Detectar navegador
    console.log('🔍 Detectando navegador...');
    let channel = null;
    
    try {
      browser = await chromium.launch({ channel: 'chrome', headless: true });
      await browser.close();
      channel = 'chrome';
      console.log('✅ Chrome encontrado!');
    } catch (e) {
      try {
        browser = await chromium.launch({ channel: 'msedge', headless: true });
        await browser.close();
        channel = 'msedge';
        console.log('✅ Edge encontrado!');
      } catch (e2) {
        throw new Error('Nenhum navegador encontrado');
      }
    }
    
    // Abrir navegador
    console.log('🚀 Abrindo navegador...');
    browser = await chromium.launch({
      headless: false,
      channel: channel,
      args: ['--start-maximized']
    });
    
    // Carregar sessão salva
    const fs = require('fs-extra');
    const path = require('path');
    const os = require('os');
    const sessionFile = path.join(process.env.APPDATA || os.homedir(), 'shopee-manager', 'shopee_session.json');
    
    if (!await fs.pathExists(sessionFile)) {
      throw new Error('❌ Sessão não encontrada! Faça login primeiro rodando o app principal.');
    }
    
    console.log('✅ Sessão encontrada!');
    
    context = await browser.newContext({
      storageState: sessionFile,
      viewport: null
    });
    
    page = await context.newPage();
    
    // Ir para o SPX
    console.log('🌐 Navegando para SPX...');
    await page.goto('https://spx.shopee.com.br/#/lmRouteCollectionPool', {
      waitUntil: 'networkidle'
    });
    
    console.log('✅ Página carregada!');
    await page.waitForTimeout(3000);
    
    // TESTAR TROCA DE STATION
    console.log('');
    console.log('═'.repeat(70));
    console.log('🔄 TESTANDO TROCA VIA API');
    console.log('═'.repeat(70));
    
    const stationTeste = 'LM Hub_SP_Sao Paulo'; // ← Coloque o nome de uma station que você tem acesso
    
    console.log(`📍 Station de teste: ${stationTeste}`);
    console.log('');
    
    const sucesso = await trocarStationCompleto(page, stationTeste);
    
    console.log('');
    console.log('═'.repeat(70));
    if (sucesso) {
      console.log('✅ TESTE PASSOU! A API está funcionando!');
    } else {
      console.log('❌ TESTE FALHOU! Houve um erro na API.');
    }
    console.log('═'.repeat(70));
    console.log('');
    console.log('🔍 Aguarde 10 segundos para ver a página...');
    
    await page.waitForTimeout(10000);
    
  } catch (error) {
    console.error('');
    console.error('═'.repeat(70));
    console.error('❌ ERRO NO TESTE:', error.message);
    console.error('═'.repeat(70));
  } finally {
    if (context) await context.close();
    if (browser) await browser.close();
    console.log('');
    console.log('✅ Teste finalizado!');
  }
}

// Executar teste
testarAPI().catch(console.error);
