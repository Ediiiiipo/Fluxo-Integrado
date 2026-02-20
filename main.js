// ============================================
// MAIN.JS - Electron Principal v2.0
// Gerencia janelas, comunicação IPC e Stations
// ============================================

const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const os = require('os');
const XLSX = require('xlsx');
const fs = require('fs-extra');
const packageJson = require('./package.json');

// ============================================
// SISTEMA DE LICENCIAMENTO
// ============================================
const LicenseManager = require('./license-manager');
const licenseManager = new LicenseManager();

// ============================================
// CONTROLE DE MODO HEADLESS (GLOBAL)
// ============================================
const HEADLESS_CONFIG_FILE = path.join(os.homedir(), '.shopee-manager', 'headless-mode.json');

// Função para carregar modo headless do arquivo
async function loadHeadlessMode() {
  try {
    if (await fs.pathExists(HEADLESS_CONFIG_FILE)) {
      const config = await fs.readJson(HEADLESS_CONFIG_FILE);
      return config.headless === true;
    }
  } catch (error) {
    console.error('Erro ao ler modo headless:', error);
  }
  return true; // Padrão: modo rápido
}

// Função para salvar modo headless em arquivo
async function saveHeadlessMode(headless) {
  try {
    await fs.ensureDir(path.dirname(HEADLESS_CONFIG_FILE));
    await fs.writeJson(HEADLESS_CONFIG_FILE, { headless, updatedAt: new Date().toISOString() });
  } catch (error) {
    console.error('Erro ao salvar modo headless:', error);
  }
}

let globalHeadlessMode = true; // Será carregado na inicialização

let mainWindow;

// Caminho do arquivo de stations
const STATIONS_FILE = path.join(__dirname, 'stations.json');

// =================== CRIAR JANELA ===================
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    title: `Shopee - Planejamento Fluxo Integrado v${packageJson.version}`,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true
    },
    icon: path.join(__dirname, 'assets', 'icon.png')
  });

  mainWindow.loadFile('index.html');

  // Abrir DevTools em desenvolvimento
  if (process.env.NODE_ENV === 'development') {
    mainWindow.webContents.openDevTools();
  }

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

// =================== INICIALIZAÇÃO ===================
app.whenReady().then(async () => {
  // ============================================
  // CARREGAR MODO HEADLESS (ANTES DE TUDO)
  // ============================================
  globalHeadlessMode = await loadHeadlessMode();
  console.log(`🎮 Modo headless: ${globalHeadlessMode ? '⚡ RÁPIDO (invisível)' : '👁️ VISUAL (visível)'}`);
  
  // ============================================
  // INICIALIZAR SISTEMA DE LICENCIAMENTO
  // ============================================
  try {
    await licenseManager.initialize();
    console.log('✅ Sistema de licenciamento inicializado');
    
    // Verificar licença
    const licenseCheck = await licenseManager.checkLicense();
    
    if (!licenseCheck.valid) {
      console.log('⚠️ Licença expirada ou inválida');
    } else if (licenseCheck.warning) {
      console.log(`⚠️ Licença expira em ${licenseCheck.daysRemaining} dias`);
    } else {
      console.log(`✅ Licença ativa até ${licenseCheck.expiryDate}`);
    }
  } catch (error) {
    console.error('❌ Erro ao inicializar licença:', error);
  }
  
  // ============================================
  // Atualizar dados do Google Sheets automaticamente ao iniciar
  // ============================================
  console.log('🔄 Atualizando dados do Google Sheets...');
  try {
    const { buscarDadosNoSheets } = require('./infoOpsClock.js');
    await buscarDadosNoSheets();
    console.log('✅ Dados atualizados com sucesso!');
  } catch (error) {
    console.log('⚠️ Erro ao atualizar dados:', error.message);
  }
  
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

// =================== IPC HANDLERS - STATIONS ===================

// Handler: Carregar stations do arquivo JSON
ipcMain.handle('carregar-stations', async () => {
  try {
    // Verificar se arquivo existe
    if (await fs.pathExists(STATIONS_FILE)) {
      const data = await fs.readJson(STATIONS_FILE);
      return {
        success: true,
        stations: data.stations || []
      };
    } else {
      // Retornar lista vazia se arquivo não existe
      return {
        success: true,
        stations: []
      };
    }
  } catch (error) {
    console.error('Erro ao carregar stations:', error);
    return {
      success: false,
      error: error.message,
      stations: []
    };
  }
});

// Handler: Salvar stations no arquivo JSON
ipcMain.handle('salvar-stations', async (event, stations) => {
  try {
    const data = {
      versao: '1.0',
      atualizado_em: new Date().toISOString().split('T')[0],
      stations: stations
    };
    
    await fs.writeJson(STATIONS_FILE, data, { spaces: 2 });
    
    console.log(`✅ ${stations.length} stations salvas em ${STATIONS_FILE}`);
    
    return {
      success: true,
      message: `${stations.length} stations salvas com sucesso`
    };
  } catch (error) {
    console.error('Erro ao salvar stations:', error);
    return {
      success: false,
      error: error.message
    };
  }
});

// =================== IPC HANDLERS - DOWNLOAD ===================

// Handler: Executar download da Shopee (com station opcional e config de navegador)
ipcMain.handle('executar-download', async (event, options = {}) => {
  try {
    const ShopeeDownloader = require('./shopee-downloader.js');
    const downloader = new ShopeeDownloader();
    
    // Função para enviar progresso para o renderer
    const enviarProgresso = (etapa, mensagem) => {
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send('download-progresso', { etapa, mensagem });
      }
    };
    
    // Suportar formato antigo (string) e novo (objeto)
    let stationNome = null;
    let headless = globalHeadlessMode; // ⚡ Usa variável global (muda na hora!)
    
    console.log(`🔍 DEBUG: globalHeadlessMode = ${globalHeadlessMode}`);
    console.log(`🔍 DEBUG: headless inicial = ${headless}`);
    
    if (typeof options === 'string') {
      // Formato antigo: apenas nome da station
      stationNome = options;
    } else if (options && typeof options === 'object') {
      stationNome = options.stationNome || null;
      // Se headless está definido no objeto, usar esse valor
      if (options.headless !== undefined) {
        headless = options.headless === true;
        console.log(`🔍 DEBUG: headless sobrescrito por options = ${headless}`);
      }
    }
    
    console.log(`🔍 DEBUG: headless final = ${headless}`);
    
    console.log('');
    console.log('═'.repeat(70));
    console.log(`🚀 Iniciando download...`);
    if (stationNome) {
      console.log(`🏢 Station solicitada: ${stationNome}`);
    } else {
      console.log(`🏢 Usando station atual do sistema`);
    }
    console.log(`🌐 Modo: ${headless ? 'Headless (invisível) ⚡' : 'Visível 👁️'}`);
    console.log('═'.repeat(70));
    
    // Passar callback de progresso para o downloader
    downloader.onProgress = enviarProgresso;
    
    // Executar download
    const result = await downloader.run(headless, stationNome);
    
    return result;
    
  } catch (error) {
    console.error('❌ Erro no download:', error);
    return {
      success: false,
      error: error.message
    };
  }
});

// Handler: Selecionar arquivo Excel
ipcMain.handle('selecionar-arquivo', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Selecionar Relatório Shopee',
    filters: [
      { name: 'Excel Files', extensions: ['xlsx', 'xls'] }
    ],
    properties: ['openFile']
  });

  if (result.canceled) {
    return { canceled: true };
  }

  return { canceled: false, filePath: result.filePaths[0] };
});

// Handler: Ler arquivo Excel e processar LH Trips
ipcMain.handle('processar-arquivo', async (event, filePath) => {
  try {
    console.log('Processando arquivo:', filePath);
    
    // Ler arquivo Excel
    const workbook = XLSX.readFile(filePath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    
    // Converter para JSON
    const data = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
    
    console.log('═'.repeat(70));
    console.log('📊 DEBUG: PROCESSAMENTO DO EXCEL');
    console.log('═'.repeat(70));
    console.log(`📄 Total de linhas no Excel: ${data.length}`);
    
    // Processar LH Trips (coluna I)
    const lhTripsCount = {};
    const lhTripsDetalhes = {}; // Para debug
    let semLH = 0;
    
    data.forEach((row, index) => {
      // A coluna I no Excel pode ter vários nomes possíveis
      const lhTrip = row['LH Trip'] || row['LH Trip ID'] || row['LH Task'] || row['LH Task ID'] || '';
      
      if (lhTrip && lhTrip.toString().trim() !== '') {
        const lhTripStr = lhTrip.toString().trim();
        lhTripsCount[lhTripStr] = (lhTripsCount[lhTripStr] || 0) + 1;
        
        // Debug: guardar índices
        if (!lhTripsDetalhes[lhTripStr]) {
          lhTripsDetalhes[lhTripStr] = [];
        }
        lhTripsDetalhes[lhTripStr].push(index + 2); // +2 porque Excel começa em 1 e tem header
      } else {
        semLH++;
      }
    });
    
    // Log específico para LT0Q2H01Z0Y11
    const lhEspecifica = 'LT0Q2H01Z0Y11';
    if (lhTripsCount[lhEspecifica]) {
      console.log(`\n🔍 DEBUG: ${lhEspecifica}`);
      console.log(`   Quantidade contada: ${lhTripsCount[lhEspecifica]}`);
      console.log(`   Linhas no Excel: ${lhTripsDetalhes[lhEspecifica].slice(0, 10).join(', ')}${lhTripsDetalhes[lhEspecifica].length > 10 ? '...' : ''}`);
    }
    
    console.log(`\n📊 Total de LH Trips únicas: ${Object.keys(lhTripsCount).length}`);
    console.log(`❌ Pedidos sem LH: ${semLH}`);
    console.log(`✅ Total processado: ${data.length}`);
    console.log('═'.repeat(70));
    
    console.log('LH Trips encontradas:', Object.keys(lhTripsCount).length);
    console.log('Pedidos sem LH:', semLH);
    
    // Preparar dados para retorno
    const lhTripsArray = Object.entries(lhTripsCount).map(([lh, count]) => ({
      lhTrip: lh,
      quantidade: count
    }));
    
    // Ordenar por quantidade (maior para menor)
    lhTripsArray.sort((a, b) => b.quantidade - a.quantidade);
    
    // Adicionar "Sem LH" se houver
    if (semLH > 0) {
      lhTripsArray.push({
        lhTrip: 'SEM LH',
        quantidade: semLH
      });
    }
    
    return {
      success: true,
      totalPedidos: data.length,
      lhTrips: lhTripsArray,
      dadosCompletos: data
    };
    
  } catch (error) {
    console.error('Erro ao processar arquivo:', error);
    return {
      success: false,
      error: error.message
    };
  }
});

// Handler: Filtrar pedidos por LH Trip
ipcMain.handle('filtrar-por-lh', async (event, { dadosCompletos, lhTripSelecionada }) => {
  try {
    let pedidosFiltrados;
    
    if (lhTripSelecionada === 'SEM LH') {
      // Filtrar pedidos sem LH
      pedidosFiltrados = dadosCompletos.filter(row => {
        const lhTrip = row['LH Trip'] || row['LH Trip ID'] || row['LH Task'] || row['LH Task ID'] || '';
        return !lhTrip || lhTrip.toString().trim() === '';
      });
    } else if (lhTripSelecionada === 'TODOS') {
      // Retornar todos
      pedidosFiltrados = dadosCompletos;
    } else {
      // Filtrar por LH específica
      pedidosFiltrados = dadosCompletos.filter(row => {
        const lhTrip = row['LH Trip'] || row['LH Trip ID'] || row['LH Task'] || row['LH Task ID'] || '';
        return lhTrip.toString().trim() === lhTripSelecionada;
      });
    }
    
    return {
      success: true,
      pedidos: pedidosFiltrados
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
});

// Handler: Carregar arquivo Excel (para aba Resumo LH's)
ipcMain.handle('carregar-arquivo', async (event, filePath) => {
  try {
    console.log('Carregando arquivo:', filePath);
    
    const workbook = XLSX.readFile(filePath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const dados = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
    
    console.log('═'.repeat(70));
    console.log('📊 DEBUG: CARREGAR ARQUIVO (DADOS COMPLETOS)');
    console.log('═'.repeat(70));
    console.log(`📄 Total de linhas lidas do Excel: ${dados.length}`);
    
    // Contar LH Trip específica
    const lhEspecifica = 'LT0Q2H01Z0Y11';
    const contagemEspecifica = dados.filter(row => {
      const lhTrip = row['LH Trip'] || row['LH Trip ID'] || row['LH Task'] || row['LH Task ID'] || '';
      return lhTrip.toString().trim() === lhEspecifica;
    }).length;
    
    if (contagemEspecifica > 0) {
      console.log(`\n🔍 DEBUG: ${lhEspecifica}`);
      console.log(`   Pedidos no Excel: ${contagemEspecifica}`);
    }
    
    // Contar todas LH Trips
    const lhTripsCount = {};
    dados.forEach(row => {
      const lhTrip = row['LH Trip'] || row['LH Trip ID'] || row['LH Task'] || row['LH Task ID'] || '';
      if (lhTrip && lhTrip.toString().trim() !== '') {
        const lhTripStr = lhTrip.toString().trim();
        lhTripsCount[lhTripStr] = (lhTripsCount[lhTripStr] || 0) + 1;
      }
    });
    
    console.log(`\n📊 Total de LH Trips únicas: ${Object.keys(lhTripsCount).length}`);
    console.log(`✅ Total enviado para o front: ${dados.length}`);
    console.log('═'.repeat(70));
    
    console.log(`✅ ${dados.length} registros carregados`);
    
    return {
      success: true,
      dados: dados
    };
    
  } catch (error) {
    console.error('Erro ao carregar arquivo:', error);
    return {
      success: false,
      error: error.message
    };
  }
});

// Handler: Exportar LHs selecionadas
ipcMain.handle('exportar-lhs', async (event, { pedidos, nomeArquivo, lhsSelecionadas, pastaDestino }) => {
  try {
    let filePath;
    
    // Se tem pasta da station, salvar direto nela
    if (pastaDestino && await fs.pathExists(pastaDestino)) {
      filePath = path.join(pastaDestino, nomeArquivo);
      console.log(`📁 Salvando na pasta da station: ${pastaDestino}`);
    } else {
      // Se não tem pasta definida, abrir dialog para escolher
      const result = await dialog.showSaveDialog(mainWindow, {
        title: 'Salvar arquivo exportado',
        defaultPath: path.join(require('os').homedir(), 'Desktop', nomeArquivo),
        filters: [
          { name: 'Excel Files', extensions: ['xlsx'] }
        ]
      });
      
      if (result.canceled) {
        return { success: false, error: 'Exportação cancelada' };
      }
      
      filePath = result.filePath;
    }
    
    // Criar workbook
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(pedidos);
    
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Pedidos');
    
    // Salvar arquivo
    XLSX.writeFile(workbook, filePath);
    
    console.log(`✅ Arquivo exportado: ${filePath}`);
    console.log(`   LHs: ${lhsSelecionadas.length}`);
    console.log(`   Pedidos: ${pedidos.length}`);
    
    return {
      success: true,
      filePath: filePath
    };
    
  } catch (error) {
    console.error('Erro ao exportar:', error);
    return {
      success: false,
      error: error.message
    };
  }
});

// =================== IPC HANDLERS - SESSÃO E LIMPEZA ===================

// Handler: Verificar status da sessão
ipcMain.handle('verificar-sessao', async () => {
  try {
    // PERFIL DENTRO DO PROJETO
    const perfilPath = path.join(__dirname, 'chrome_profile_for_bot');
    const temSessao = await fs.pathExists(perfilPath);
    
    // Verificar se tem cookies salvos
    let temCookies = false;
    if (temSessao) {
      const defaultPath = path.join(perfilPath, 'Default');
      temCookies = await fs.pathExists(defaultPath);
    }
    
    return {
      success: true,
      temSessao: temSessao && temCookies,
      perfilPath: perfilPath
    };
  } catch (error) {
    return {
      success: false,
      temSessao: false,
      error: error.message
    };
  }
});

// Handler: Limpar sessão de login (apaga só a pasta do perfil)
ipcMain.handle('limpar-sessao', async () => {
  try {
    // PERFIL DENTRO DO PROJETO
    const perfilPath = path.join(__dirname, 'chrome_profile_for_bot');
    
    if (await fs.pathExists(perfilPath)) {
      await fs.remove(perfilPath);
      console.log('✅ Sessão de login limpa:', perfilPath);
    }
    
    return { success: true };
  } catch (error) {
    console.error('Erro ao limpar sessão:', error);
    return {
      success: false,
      error: error.message
    };
  }
});

// Handler: Limpar todos os dados
ipcMain.handle('limpar-tudo', async () => {
  try {
    // Limpar pasta do perfil (dentro do projeto)
    const perfilPath = path.join(__dirname, 'chrome_profile_for_bot');
    if (await fs.pathExists(perfilPath)) {
      await fs.remove(perfilPath);
      console.log('✅ Pasta do perfil removida');
    }
    
    // Limpar pasta temporária
    const tempPath = path.join(require('os').tmpdir(), 'shopee_temp');
    if (await fs.pathExists(tempPath)) {
      await fs.remove(tempPath);
      console.log('✅ Pasta temporária removida');
    }
    
    // Limpar stations.json
    if (await fs.pathExists(STATIONS_FILE)) {
      await fs.remove(STATIONS_FILE);
      console.log('✅ Arquivo stations.json removido');
    }
    
    return { success: true };
  } catch (error) {
    console.error('Erro ao limpar dados:', error);
    return {
      success: false,
      error: error.message
    };
  }
});

// =================== IPC HANDLERS - PLANILHA GOOGLE SHEETS ===================

const PLANILHA_FILE = path.join(__dirname, 'lhs_planilha.json');
const CREDENCIAIS_FILE = path.join(__dirname, 'credenciais.json');
const ID_PLANILHA = '18P9ZPVmNpFOKhWIq9XzZRuSnhtgYR4aJagC8mc3bNa0';
const INTERVALO_PLANILHA = 'Página1!A:U';

// ===== NOVAS PLANILHAS =====
// InfoOpsClock - Horários dos Ciclos
const OPSCLOCK_FILE = path.join(__dirname, 'dados_infoOpsClock.json');
const ID_PLANILHA_OPSCLOCK = '1BOddilA48UQPD9QsnxAFmVcIs0OYRMYxwhTJ70rrIQ8';
const INTERVALO_OPSCLOCK = "'Página1'!A4:Y";

// InfoOutboundCapacity - Capacidade por Ciclo/Data
const OUTBOUND_FILE = path.join(__dirname, 'dados_outbound_capacity.json');
const ID_PLANILHA_OUTBOUND = '1iJ70tTT_hlUqcWQacHuhP-3CYI8rYNkOdKnBAHXI_eg';
const INTERVALO_OUTBOUND = "'Resume Out. Capacity'!B5:CE";

// Handler: Atualizar planilha do Google Sheets (via API)
ipcMain.handle('atualizar-planilha-google', async (event) => {
  try {
    console.log('📊 Conectando ao Google Sheets via API...');
    
    // Verificar se existe o arquivo de credenciais
    if (!await fs.pathExists(CREDENCIAIS_FILE)) {
      throw new Error('Arquivo credenciais.json não encontrado! Coloque o arquivo na pasta do projeto.');
    }
    
    // Importar googleapis dinamicamente
    const { google } = require('googleapis');
    
    // Autenticar com Service Account
    const auth = new google.auth.GoogleAuth({
      keyFile: CREDENCIAIS_FILE,
      scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
    });
    
    const client = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: client });
    
    console.log('   ✅ Autenticado com sucesso!');
    
    // ========== 1. PLANILHA PRINCIPAL (LHs) ==========
    console.log('');
    console.log('📋 [1/3] Atualizando planilha de LHs...');
    
    const resposta = await sheets.spreadsheets.values.get({
      spreadsheetId: ID_PLANILHA,
      range: INTERVALO_PLANILHA,
    });
    
    const linhas = resposta.data.values;
    
    if (!linhas || linhas.length === 0) {
      throw new Error('A planilha de LHs está vazia.');
    }
    
    console.log(`   ✅ ${linhas.length} linhas encontradas`);
    
    // Converter para JSON (usando trip_number como chave)
    // IMPORTANTE: Se houver múltiplas linhas com o mesmo trip_number,
    // apenas a última linha será mantida (sobrescreve as anteriores)
    const cabecalhos = linhas[0];
    const dados = {};
    
    // Identificar índice da coluna 'destination' para filtrar depois
    const idxDestination = cabecalhos.findIndex(col => 
      col && col.toLowerCase() === 'destination'
    );
    
    // Identificar índice da coluna 'update_datetime' para usar como última atualização
    const idxUpdateDatetime = cabecalhos.findIndex(col => 
      col && col.toLowerCase() === 'update_datetime'
    );
    
    console.log(`   🔍 Índice da coluna 'destination': ${idxDestination}`);
    console.log(`   🔍 Índice da coluna 'update_datetime': ${idxUpdateDatetime}`);
    
    for (let i = 1; i < linhas.length; i++) {
      const linha = linhas[i];
      const tripNumber = linha[0]; // Coluna A = trip_number
      
      if (tripNumber) {
        const registro = {};
        cabecalhos.forEach((col, idx) => {
          registro[col] = linha[idx] || '';
        });
        
        // Armazenar com chave composta: trip_number + destination
        // Isso permite múltiplas entradas da mesma LH para diferentes destinations
        const destination = linha[idxDestination] || '';
        const chaveComposta = `${tripNumber}|${destination}`;
        
        dados[chaveComposta] = registro;
      }
    }
    
    console.log(`   ✅ Convertido: ${Object.keys(dados).length} registros (trip_number + destination)`);
    
    // Capturar update_datetime da primeira linha de dados (todas têm o mesmo valor)
    let updateDatetimeValue = null;
    if (linhas.length > 1 && idxUpdateDatetime >= 0) {
      updateDatetimeValue = linhas[1][idxUpdateDatetime];
      console.log(`   📅 update_datetime capturado: ${updateDatetimeValue}`);
    }
    
    // Salvar JSON localmente
    const dadosParaSalvar = {
      ultimaAtualizacao: updateDatetimeValue || new Date().toISOString(),  // Usar update_datetime do Sheets
      totalRegistros: Object.keys(dados).length,
      dados: dados,
      usaChaveComposta: true  // Flag para indicar que usa trip_number|destination
    };
    
    await fs.writeJson(PLANILHA_FILE, dadosParaSalvar, { spaces: 2 });
    console.log(`   ✅ Salvo em: ${PLANILHA_FILE}`);
    
    // ========== 2. PLANILHA OPSCLOCK (Horários dos Ciclos) ==========
    console.log('');
    console.log('⏰ [2/3] Atualizando planilha OpsClock (horários)...');
    
    let dadosOpsClock = [];
    try {
      const respostaOpsClock = await sheets.spreadsheets.values.get({
        spreadsheetId: ID_PLANILHA_OPSCLOCK,
        range: INTERVALO_OPSCLOCK,
      });
      
      const linhasOpsClock = respostaOpsClock.data.values;
      
      if (linhasOpsClock && linhasOpsClock.length > 0) {
        const cabecalhosOps = linhasOpsClock[0];
        dadosOpsClock = linhasOpsClock.slice(1).map(linha => {
          let obj = {};
          cabecalhosOps.forEach((col, idx) => {
            obj[col] = linha[idx] || '';
          });
          return obj;
        });
        
        await fs.writeJson(OPSCLOCK_FILE, {
          ultimaAtualizacao: new Date().toISOString(),
          totalRegistros: dadosOpsClock.length,
          dados: dadosOpsClock
        }, { spaces: 2 });
        
        console.log(`   ✅ ${dadosOpsClock.length} registros de ciclos salvos`);
      } else {
        console.log('   ⚠️ Planilha OpsClock vazia');
      }
    } catch (errOps) {
      console.log(`   ⚠️ Erro ao carregar OpsClock: ${errOps.message}`);
    }
    
    // ========== 3. PLANILHA OUTBOUND (Capacidade por Ciclo/Data) ==========
    console.log('');
    console.log('📊 [3/3] Atualizando planilha Outbound Capacity...');
    
    let dadosOutbound = [];
    try {
      const respostaOutbound = await sheets.spreadsheets.values.get({
        spreadsheetId: ID_PLANILHA_OUTBOUND,
        range: INTERVALO_OUTBOUND,
      });
      
      const linhasOutbound = respostaOutbound.data.values;
      
      if (linhasOutbound && linhasOutbound.length > 0) {
        const cabecalhosOut = linhasOutbound[0];
        dadosOutbound = linhasOutbound.slice(1).map(linha => {
          let obj = {};
          cabecalhosOut.forEach((col, idx) => {
            obj[col] = linha[idx] || '';
          });
          return obj;
        });
        
        await fs.writeJson(OUTBOUND_FILE, {
          ultimaAtualizacao: new Date().toISOString(),
          totalRegistros: dadosOutbound.length,
          dados: dadosOutbound
        }, { spaces: 2 });
        
        console.log(`   ✅ ${dadosOutbound.length} registros de capacidade salvos`);
      } else {
        console.log('   ⚠️ Planilha Outbound vazia');
      }
    } catch (errOut) {
      console.log(`   ⚠️ Erro ao carregar Outbound: ${errOut.message}`);
    }
    
    console.log('');
    console.log('═'.repeat(50));
    console.log('✅ TODAS AS PLANILHAS ATUALIZADAS!');
    console.log('═'.repeat(50));
    
    return {
      success: true,
      dados: dados,
      ultimaAtualizacao: dadosParaSalvar.ultimaAtualizacao,
      totalRegistros: dadosParaSalvar.totalRegistros,
      opsClock: dadosOpsClock.length,
      outbound: dadosOutbound.length
    };
    
  } catch (error) {
    console.error('❌ Erro ao atualizar planilha:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
});

// Handler: Carregar planilha local (JSON salvo)
ipcMain.handle('carregar-planilha-local', async () => {
  try {
    if (await fs.pathExists(PLANILHA_FILE)) {
      const dados = await fs.readJson(PLANILHA_FILE);
      console.log(`📊 Planilha local carregada: ${dados.totalRegistros} LHs`);
      return {
        success: true,
        dados: dados
      };
    } else {
      return {
        success: false,
        error: 'Arquivo não encontrado'
      };
    }
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
});

// Handler: Carregar dados OpsClock local (horários dos ciclos)
ipcMain.handle('carregar-opsclock-local', async () => {
  try {
    if (await fs.pathExists(OPSCLOCK_FILE)) {
      const dados = await fs.readJson(OPSCLOCK_FILE);
      console.log(`⏰ OpsClock local carregado: ${dados.totalRegistros} registros`);
      return {
        success: true,
        dados: dados.dados || [],
        ultimaAtualizacao: dados.ultimaAtualizacao
      };
    } else {
      return {
        success: false,
        error: 'Arquivo não encontrado'
      };
    }
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
});

// Handler: Carregar dados Outbound Capacity local (capacidade por ciclo)
ipcMain.handle('carregar-outbound-local', async () => {
  try {
    if (await fs.pathExists(OUTBOUND_FILE)) {
      const dados = await fs.readJson(OUTBOUND_FILE);
      console.log(`📊 Outbound local carregado: ${dados.totalRegistros} registros`);
      return {
        success: true,
        dados: dados.dados || [],
        ultimaAtualizacao: dados.ultimaAtualizacao
      };
    } else {
      return {
        success: false,
        error: 'Arquivo não encontrado'
      };
    }
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
});

// =================== IPC HANDLERS - PLANEJAMENTO ===================

// Handler: Gerar arquivo de planejamento
ipcMain.handle('gerar-planejamento', async (event, dados) => {
  try {
    const { pedidos, nomeArquivo, lhsSelecionadas, qtdBacklog, pastaStation, tosComplemento } = dados;
    
    if (!pedidos || pedidos.length === 0) {
      return { success: false, error: 'Nenhum pedido para exportar' };
    }
    
    // Criar workbook Excel
    const wb = XLSX.utils.book_new();
    
    // Sheet principal com todos os pedidos
    const ws = XLSX.utils.json_to_sheet(pedidos);
    XLSX.utils.book_append_sheet(wb, ws, 'Planejamento');
    
    
    // Sheet de TOs de Complemento (se houver)
    if (tosComplemento && tosComplemento.length > 0) {
      const wsTOs = XLSX.utils.json_to_sheet(tosComplemento);
      XLSX.utils.book_append_sheet(wb, wsTOs, 'TOs Complemento');
      console.log(`✅ Aba 'TOs Complemento' criada com ${tosComplemento.length} registros`);
    }
    
    // Definir pasta de saída - usar pasta da station se disponível
    let outputDir;
    
    if (pastaStation && await fs.pathExists(pastaStation)) {
      // Usar pasta da station atual
      outputDir = pastaStation;
    } else {
      // Fallback para Shopee_Downloads
      outputDir = path.join(app.getPath('desktop'), 'Shopee_Downloads');
      await fs.ensureDir(outputDir);
    }
    
    // Salvar arquivo diretamente na pasta da station
    const filePath = path.join(outputDir, nomeArquivo);
    XLSX.writeFile(wb, filePath);
    
    console.log(`✅ Planejamento gerado: ${filePath}`);
    
    return {
      success: true,
      filePath: filePath,
      qtdPedidos: pedidos.length
    };
    
  } catch (error) {
    console.error('❌ Erro ao gerar planejamento:', error);
    return {
      success: false,
      error: error.message
    };
  }
});

// =================== IPC HANDLER - EXPORTAR LHs ===================

const ExportadorLHs = require('./exportar-lhs.js');

ipcMain.handle('exportar-lhs-spx', async (event, config) => {
  try {
    const { headless, diasPendentes, diasFinalizados } = config;
    
    console.log('');
    console.log('═'.repeat(70));
    console.log(`📦 Iniciando exportação de LHs...`);
    console.log(`🌐 Modo: ${headless ? 'Headless (invisível)' : 'Visível'}`);
    console.log(`📅 Período Pendentes/Expedidos: ${diasPendentes} dias`);
    console.log(`📅 Período Finalizados: ${diasFinalizados} dias`);
    console.log('═'.repeat(70));
    
    // Callback para enviar progresso à UI
    const enviarProgresso = (etapa, pagina, total) => {
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send('exportar-lhs-progresso', {
          etapa,
          pagina,
          total
        });
      }
    };
    
    // Criar instância do exportador
    const exportador = new ExportadorLHs(enviarProgresso);
    
    // Executar exportação
    const result = await exportador.run(headless, diasPendentes, diasFinalizados);
    
    return result;
    
  } catch (error) {
    console.error('❌ Erro na exportação de LHs:', error);
    return {
      success: false,
      error: error.message
    };
  }
});

// ==================== IPC HANDLER: VALIDAR LHs COM SPX ====================
const ValidadorLHs = require('./validar-lhs');

ipcMain.handle('validar-lhs-spx', async (event, config) => {
  try {
    const { stationFiltro, sequenceFilter } = config || {};
    
    console.log('');
    console.log('═'.repeat(70));
    console.log(`🔍 Iniciando validação cruzada de LHs...`);
    if (stationFiltro) {
      console.log(`📍 Station: ${stationFiltro}`);
    }
    if (sequenceFilter) {
      console.log(`🎯 Filtro: ${sequenceFilter}`);
    }
    console.log('═'.repeat(70));
    
    // Criar instância do validador
    const validador = new ValidadorLHs();
    
    // Executar validação com filtro de station e sequence_number
    const result = await validador.validar(stationFiltro, sequenceFilter || 'todos');
    
    return result;
    
  } catch (error) {
    console.error('❌ Erro na validação de LHs:', error);
    return {
      success: false,
      error: error.message
    };
  }
});

// ==================== IPC HANDLER: SALVAR LOG DE PLANEJAMENTO ====================

ipcMain.handle('salvar-log-planejamento', async (event, logData) => {
  try {
    console.log('📝 Salvando log de planejamento no Google Sheets...');
    
    const { google } = require('googleapis');
    const credenciaisPath = path.join(__dirname, 'credenciais.json');
    
    // Verificar se credenciais existem
    if (!await fs.pathExists(credenciaisPath)) {
      throw new Error('Arquivo credenciais.json não encontrado');
    }
    
    // Carregar credenciais
    const credenciais = await fs.readJson(credenciaisPath);
    const auth = new google.auth.GoogleAuth({
      credentials: credenciais,
      scopes: ['https://www.googleapis.com/auth/spreadsheets']
    });
    
    const sheets = google.sheets({ version: 'v4', auth });
    
    // ID da planilha de controle
    const SPREADSHEET_ID = '1oKhwpY3yWpcb0w6CYqvAoTT2Ss689bwfPw7U6l2jKNo';
    const SHEET_NAME = 'Sheet2';
    
    // Preparar dados para inserir
    const dataHoraBR = new Date(logData.dataHora).toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
    
    // Criar array com dados em colunas separadas
    const valores = [
      [
        logData.email,                    // Coluna A: E-MAIL
        dataHoraBR,                       // Coluna B: DATA/HORA
        logData.estacao,                  // Coluna C: ESTAÇÃO
        logData.ciclo,                    // Coluna D: CICLO
        logData.dataExpedicao,            // Coluna E: DATA EXPEDIÇÃO
        logData.capAutomatico,            // Coluna F: CAP AUTOMÁTICO
        logData.pedidosTotais,            // Coluna G: PEDIDOS TOTAIS
        logData.pedidosPlanejados,        // Coluna H: PEDIDOS PLANEJADOS
        logData.quantidadeLHs,            // Coluna I: QTD LHs
        logData.backlog                   // Coluna J: BACKLOG
      ]
    ];
    
    // Adicionar linha no Google Sheets
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A:J`,
      valueInputOption: 'RAW',
      resource: {
        values: valores
      }
    });
    
    console.log('✅ Log salvo com sucesso no Google Sheets!');
    
    return { success: true };
    
  } catch (error) {
    console.error('❌ Erro ao salvar log:', error);
    return {
      success: false,
      error: error.message
    };
  }
});

// ============================================
// IPC HANDLERS - SISTEMA DE LICENCIAMENTO
// ============================================

// Handler: Verificar licença
ipcMain.handle('license-check', async () => {
  try {
    return await licenseManager.checkLicense();
  } catch (error) {
    console.error('❌ Erro ao verificar licença:', error);
    return { valid: false, reason: 'error' };
  }
});

// Handler: Solicitar renovação
ipcMain.handle('license-request-renewal', async (event, { nome, email }) => {
  try {
    return await licenseManager.requestRenewal(nome, email);
  } catch (error) {
    console.error('❌ Erro ao solicitar renovação:', error);
    return { success: false, error: 'Erro interno' };
  }
});

// Handler: Ativar licença
ipcMain.handle('license-activate', async (event, password) => {
  try {
    return await licenseManager.activateLicense(password);
  } catch (error) {
    console.error('❌ Erro ao ativar licença:', error);
    return { success: false, error: 'Erro interno' };
  }
});

// Handler: Buscar solicitação (admin)
ipcMain.handle('license-get-request', async (event, code) => {
  try {
    const request = await licenseManager.getRequest(code);
    if (!request) {
      return { success: false, error: 'Solicitação não encontrada' };
    }
    return { success: true, request };
  } catch (error) {
    console.error('❌ Erro ao buscar solicitação:', error);
    return { success: false, error: 'Erro interno' };
  }
});

// Handler: Aprovar solicitação (admin)
ipcMain.handle('license-approve-request', async (event, { code, approvedBy }) => {
  try {
    return await licenseManager.approveRequest(code, approvedBy);
  } catch (error) {
    console.error('❌ Erro ao aprovar solicitação:', error);
    return { success: false, error: 'Erro interno' };
  }
});

// Handler: Obter histórico (admin)
ipcMain.handle('license-get-history', async () => {
  try {
    return await licenseManager.getRecentRequests(10);
  } catch (error) {
    console.error('❌ Erro ao obter histórico:', error);
    return [];
  }
});

// Handler: Estender licença manualmente (admin)
ipcMain.handle('license-extend', async () => {
  try {
    return await licenseManager.extendLicenseManually(6);
  } catch (error) {
    console.error('❌ Erro ao estender licença:', error);
    return { success: false, error: 'Erro interno' };
  }
});

// ============================================
// IPC HANDLER - TOGGLE HEADLESS (MUDA NA HORA)
// ============================================

// Handler: Alternar modo headless (CTRL+U)
ipcMain.handle('toggle-headless-mode', async (event, newMode) => {
  try {
    // newMode: true = headless (rápido), false = visível
    globalHeadlessMode = newMode;
    
    // Salvar em arquivo para persistir
    await saveHeadlessMode(newMode);
    
    console.log(`🎮 Modo headless alterado: ${newMode ? '⚡ RÁPIDO (invisível)' : '👁️ VISUAL (visível)'}`);
    
    return {
      success: true,
      currentMode: globalHeadlessMode
    };
  } catch (error) {
    console.error('❌ Erro ao alternar modo:', error);
    return {
      success: false,
      error: error.message
    };
  }
});

// Handler: Obter modo headless atual
ipcMain.handle('get-headless-mode', async () => {
  return {
    success: true,
    headless: globalHeadlessMode
  };
});

// ============================================
// SINCRONIZAÇÃO SPX - BUSCAR LHS E GERAR CSV
// ============================================
const { chromium } = require('playwright-core');
const { trocarStationCompleto, buscarStationIdPorNome } = require('./station-switcher-api');

const SPX_CONFIG = {
    SESSION_FILE: path.join(process.env.APPDATA || os.homedir(), 'shopee-manager', 'shopee_session.json'),
    FILE_SPECIFIC_JSON: path.join(__dirname, 'LH_busca_especifica.json'),
    FILE_FINAL_CSV: path.join(__dirname, 'Relatorio_LHs_Pronto.csv'),
    API_SEARCH: "https://spx.shopee.com.br/api/admin/transportation/trip/list",
    API_HISTORY: "https://spx.shopee.com.br/api/admin/transportation/trip/history/list"
};

// Mapeamento de Status
const statusMap = { 
    10: "Criado", 
    20: "Aguardando Motorista", 
    30: "Embarcando", 
    40: "Em Trânsito", 
    50: "Chegou no Destino", 
    60: "Desembarcando", 
    80: "Finalizado", 
    90: "Finalizado", 
    100: "Cancelado",
    200: "Cancelado"
};

// Formatar Timestamps Unix para Data/Hora BR
const formatData = (ts) => {
    if (!ts || ts === 0) return null;
    return new Date(ts * 1000).toLocaleString('pt-BR');
};

async function detectSystemBrowser() {
    const paths = [
        'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
        'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'
    ];
    for (const p of paths) { 
        if (await fs.pathExists(p)) return p; 
    }
    return null;
}

/**
 * Converte dados JSON para CSV
 */
async function exportToCSV(dataArray, outputFolder) {
    console.log(`📊 Formatando ${dataArray.length} registro(s) para Excel...`);
    
    const rows = dataArray.map(item => {
        const d = item.dados || item;
        const stations = d.trip_station || [];
        
        const origem = stations[0] || {};
        const destino = stations[stations.length - 1] || {};

        // Lógica Smart: Se Real (atd/ata) for 0, usa Estimada (etd/eta)
        const saida = formatData(origem.atd) || (origem.etd ? `Est: ${formatData(origem.etd)}` : "Pendente");
        const chegada = formatData(destino.ata) || (destino.eta ? `Est: ${formatData(destino.eta)}` : "Em trânsito");
        const planejado = formatData(d.trip_date) || "N/A";
        const custo = (d.total_cost && d.total_cost !== 0) ? d.total_cost : (d.agency_cost || "0");

        return [
            d.trip_number,
            statusMap[d.trip_status] || d.trip_status,
            planejado,
            d.driver_name || "N/A",
            d.vehicle_number || "N/A",
            d.vehicle_type_name || "N/A",
            origem.station_name || "N/A",
            saida,
            destino.station_name || "N/A",
            chegada,
            custo
        ].join(';');
    });

    const header = "LH;Status;Data Planejada;Motorista;Placa;Tipo Veiculo;Origem;Saida Real;Destino;Chegada Real;Custo";
    const csvContent = "\ufeff" + [header, ...rows].join('\n');
    
    // Salvar na pasta da station com nome simples
    const csvPath = path.join(outputFolder, 'LHs_SPX.csv');
    await fs.writeFile(csvPath, csvContent, 'utf8');
    console.log(`✨ Relatório gerado: ${csvPath}`);
    
    return csvPath;
}

/**
 * Busca múltiplas LHs no SPX (modo Admin) e gera CSV
 */
async function buscarLHsNoSPXComCSV(lhIds, stationFolder, currentStationName) {
    console.log(`\n🚀 INICIANDO BUSCA DE ${lhIds.length} LHs (MODO ADMIN)...`);
    
    if (!await fs.pathExists(SPX_CONFIG.SESSION_FILE)) {
        throw new Error("❌ Sessão SPX não encontrada. Faça login manual primeiro.");
    }

    const browserPath = await detectSystemBrowser();
    if (!browserPath) {
        throw new Error("❌ Navegador não encontrado.");
    }

    const browser = await chromium.launch({ 
        executablePath: browserPath, 
        headless: globalHeadlessMode 
    });
    const context = await browser.newContext({ storageState: SPX_CONFIG.SESSION_FILE });
    const page = await context.newPage();
    
    let resultadosGerais = [];
    let encontradas = 0;
    let stationOriginal = currentStationName;

    try {
        // Navegar para SPX
        await page.goto('https://spx.shopee.com.br/#/hubLinehaulTrips/trip', { 
            waitUntil: 'networkidle',
            timeout: 30000 
        });
        await page.waitForTimeout(3000);

        // 1. TROCAR PARA STATION ADMIN
        console.log('\n🔄 Trocando para station admin...');
        const trocouAdmin = await trocarStationCompleto(page, 'admin');
        
        if (!trocouAdmin) {
            console.warn('⚠️ Não foi possível trocar para admin, continuando na station atual...');
        } else {
            console.log('✅ Agora operando na station admin');
        }
        
        await page.waitForTimeout(2000);

        // 2. BUSCAR LHs
        for (const lhId of lhIds) {
            console.log(`   🔡 Buscando ${lhId}...`);
            let achou = false;

            try {
                // Tenta nas duas APIs
                for (const api of [SPX_CONFIG.API_SEARCH, SPX_CONFIG.API_HISTORY]) {
                    const res = await page.evaluate(async ({url, id}) => {
                        try {
                            const r = await fetch(`${url}?trip_number=${id}&pageno=1&count=10`);
                            return r.ok ? await r.json() : null;
                        } catch (e) { return null; }
                    }, { url: api, id: lhId });

                    if (res?.data?.list?.length > 0) {
                        const match = res.data.list.find(i => i.trip_number.trim() === lhId.trim()) || res.data.list[0];
                        resultadosGerais.push({ lh_id: lhId, dados: match });
                        encontradas++;
                        console.log(`      ✅ ${lhId} encontrada!`);
                        achou = true;
                        break;
                    }
                }
            } catch (err) {
                console.error(`      ⚠️ Erro: ${err.message}`);
            }

            if (!achou) {
                console.log(`      ❌ ${lhId} não encontrada`);
            }
            
            await page.waitForTimeout(800);
        }
        
        // 3. VOLTAR PARA STATION ORIGINAL
        if (stationOriginal && stationOriginal !== 'admin') {
            console.log(`\n🔄 Voltando para station original: ${stationOriginal}`);
            const voltou = await trocarStationCompleto(page, stationOriginal);
            
            if (voltou) {
                console.log(`✅ Restaurado para: ${stationOriginal}`);
            } else {
                console.warn(`⚠️ Não foi possível voltar para ${stationOriginal}`);
            }
        }
        
    } finally {
        await browser.close();
    }

    console.log(`\n✅ BUSCA CONCLUÍDA: ${encontradas}/${lhIds.length}`);

    let csvPath = null;
    
    if (resultadosGerais.length > 0) {
        // Salvar JSON
        await fs.writeJson(SPX_CONFIG.FILE_SPECIFIC_JSON, resultadosGerais, { spaces: 2 });
        
        // Gerar CSV na pasta da station
        csvPath = await exportToCSV(resultadosGerais, stationFolder);
    }

    return {
        total: lhIds.length,
        encontradas: encontradas,
        erros: lhIds.length - encontradas,
        resultados: resultadosGerais,
        csvPath: csvPath,
        jsonPath: resultadosGerais.length > 0 ? SPX_CONFIG.FILE_SPECIFIC_JSON : null
    };
}

ipcMain.handle('sincronizar-lhs-spx', async (event, { lhIds, stationFolder, currentStationName }) => {
    try {
        console.log('\n🔄 INICIANDO SINCRONIZAÇÃO SPX...');
        console.log(`📍 Station atual: ${currentStationName}`);
        const resultado = await buscarLHsNoSPXComCSV(lhIds, stationFolder, currentStationName);
        return { success: true, data: resultado };
    } catch (error) {
        console.error('❌ Erro na sincronização SPX:', error);
        return { success: false, error: error.message };
    }
});

// Handler: Abrir arquivo no sistema
ipcMain.handle('abrir-arquivo', async (event, filePath) => {
    try {
        const { shell } = require('electron');
        await shell.openPath(filePath);
        return { success: true };
    } catch (error) {
        console.error('❌ Erro ao abrir arquivo:', error);
        return { success: false, error: error.message };
    }
});