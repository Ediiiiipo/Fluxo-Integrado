// ============================================
// RENDERER.JS - Interface do Gerenciador v2.0
// Sistema de Planejamento Hub e Backlog
// ============================================

const {
    ipcRenderer
} = require('electron');

// ======================= ESTADO GLOBAL =======================
let dadosAtuais = []; // Todos os pedidos carregados
let lhTrips = {}; // Agrupamento por LH Trip
let lhTripAtual = null; // LH Trip selecionada na aba principal
let stationsCadastradas = []; // Lista de stations
let todasColunas = []; // Todas as colunas do arquivo
let colunasVisiveis = []; // Colunas selecionadas para exibir
let pastaStationAtual = null; // Pasta da station do último download
let stationAtualNome = null; // Nome da station atual (extraído da pasta)
let emailUsuario = null; // E-mail do usuário logado

// Estado para Planejamento Hub - Seleção de LHs
let lhsSelecionadasPlan = new Set(); // LHs selecionadas para o planejamento

// Estado para Backlog - Pedidos sem LH
let pedidosBacklogSelecionados = new Set(); // IDs dos pedidos selecionados
let backlogConfirmado = false; // Se o backlog foi confirmado

// Estado para separação Backlog vs Planejável (baseado no STATUS)
let pedidosBacklogPorStatus = []; // Pedidos com STATUS de backlog (LMHub_Received, Return_LMHub_Received)
let pedidosPlanejáveis = []; // Pedidos com outros status (planejáveis)
let lhTripsBacklog = {}; // LH Trips que são backlog (agrupado)
let lhTripsPlanejáveis = {}; // LH Trips planejáveis (agrupado)
let lhsLixoSistemico = []; // LHs filtradas automaticamente (sem origin/destination/previsão)

// Status que identificam Backlog (lowercase, sem espaços)
const STATUS_BACKLOG = [
    'lmhub_received', 
    'return_lmhub_received', 
    'hub_received', 
    'return_hub_received',
    'sinalizar_inventário',  // 🆕 Adicionado para mover para backlog
    'sinalizar_inventario'    // 🆕 Versão sem acento
];

// Função para verificar se status é de backlog
function isStatusBacklog(status) {
    if (!status) return false;
    // Normalizar: lowercase, remover espaços, remover underscores extras
    const statusNorm = String(status).toLowerCase().trim().replace(/\s+/g, '_');
    
    // Verificar se contém algum dos status de backlog
    return STATUS_BACKLOG.some(sb => statusNorm.includes(sb) || sb.includes(statusNorm));
}

/**
 * 🆕 Reclassifica pedidos com "Sinalizar Inventário" para Backlog
 * Esta função roda APÓS o carregamento inicial, quando os status já foram calculados
 */
function reclassificarSinalizarInventarioParaBacklog() {
    console.log('\n🔄 [RECLASSIFICAÇÃO] Movendo "Sinalizar Inventário" para Backlog...');
    
    let pedidosMovidos = 0;
    let lhsMovidas = [];
    
    // Calcular estatísticas de volume se ainda não existir
    const volumes = Object.keys(lhTripsPlanejáveis).map(lhTrip => lhTripsPlanejáveis[lhTrip].length);
    const estatisticas = {
        media: volumes.length > 0 ? volumes.reduce((a, b) => a + b, 0) / volumes.length : 0,
        percentil10: volumes.length > 0 ? volumes.sort((a, b) => a - b)[Math.floor(volumes.length * 0.1)] : 0
    };
    
    // Encontrar nomes das colunas
    const primeiroRegistro = dadosAtuais[0] || {};
    const todasColunasDisponiveis = Object.keys(primeiroRegistro);
    
    const colunaLH = todasColunasDisponiveis.find(col =>
        col.toLowerCase().includes('lh trip') ||
        col.toLowerCase().includes('lh_trip') ||
        col.toLowerCase().includes('lhtask')
    ) || 'LH Trip';
    
    const colunaStatus = todasColunasDisponiveis.find(col =>
        col.toLowerCase() === 'status' ||
        col.toLowerCase().includes('status')
    ) || 'Status';
    
    // Percorrer todas as LHs planejáveis
    const lhsParaRemover = [];
    
    for (const lhTrip in lhTripsPlanejáveis) {
        const pedidos = lhTripsPlanejáveis[lhTrip];
        const qtdPedidos = pedidos.length;
        
        // Buscar dados da planilha para esta LH
        const dadosPlanilhaLH = buscarDadosPlanilhaPorStation(lhTrip);
        
        // Verificar se é baixo volume
        const isBaixoVolume = verificarLHBaixoVolume(qtdPedidos, estatisticas);
        
        // Caso 1: Tem dados na planilha E é baixo volume → verificar se previsão passou
        if (isBaixoVolume && dadosPlanilhaLH) {
            // Verificar se previsão já passou
            let previsaoPassou = false;
            try {
                const previsaoFinalCandidatos = [
                    dadosPlanilhaLH.previsao_final,
                    dadosPlanilhaLH['Previsão Final'],
                    dadosPlanilhaLH['previsão_final'],
                    dadosPlanilhaLH.PREVISAO_FINAL
                ].filter(p => p && String(p).trim() !== '');
                
                if (previsaoFinalCandidatos.length > 0) {
                    const previsaoFinal = String(previsaoFinalCandidatos[0]).trim();
                    const dataPrevisao = new Date(previsaoFinal);
                    const dataHoje = new Date();
                    dataHoje.setHours(23, 59, 59, 999); // Fim do dia de hoje
                    
                    if (!isNaN(dataPrevisao.getTime())) {
                        // Considera que passou se for hoje ou antes
                        previsaoPassou = dataPrevisao <= dataHoje;
                    }
                }
            } catch (e) {
                // Se der erro, considerar que passou
                previsaoPassou = true;
            }
            
            // Se é baixo volume E previsão já passou → Sinalizar Inventário → Backlog
            if (previsaoPassou) {
                console.log(`   📦 Movendo LH ${lhTrip} (${pedidos.length} pedidos, baixo volume + previsão passada) para Backlog`);
                
                // Mover pedidos para backlog
                pedidos.forEach(pedido => {
                    // Marcar a LH original
                    pedido._lhOriginal = lhTrip;
                    // Renomear para Backlog
                    pedido[colunaLH] = 'Backlog';
                    // Marcar status
                    pedido[colunaStatus] = 'Sinalizar Inventário';
                    
                    // Adicionar ao array de backlog
                    pedidosBacklogPorStatus.push(pedido);
                    pedidosMovidos++;
                });
                
                // Adicionar ao objeto de backlog agrupado
                if (!lhTripsBacklog[lhTrip]) {
                    lhTripsBacklog[lhTrip] = [];
                }
                lhTripsBacklog[lhTrip].push(...pedidos);
                
                // Marcar para remoção
                lhsParaRemover.push(lhTrip);
                lhsMovidas.push(lhTrip);
            }
        }
        // Caso 2: NÃO tem dados na planilha E é baixo volume → mover para Backlog
        else if (isBaixoVolume && !dadosPlanilhaLH) {
            console.log(`   📦 Movendo LH ${lhTrip} (${pedidos.length} pedidos, baixo volume + sem dados) para Backlog`);
            
            // Mover pedidos para backlog
            pedidos.forEach(pedido => {
                // Marcar a LH original
                pedido._lhOriginal = lhTrip;
                // Renomear para Backlog
                pedido[colunaLH] = 'Backlog';
                // Marcar status
                pedido[colunaStatus] = 'Sinalizar Inventário';
                
                // Adicionar ao array de backlog
                pedidosBacklogPorStatus.push(pedido);
                pedidosMovidos++;
            });
            
            // Adicionar ao objeto de backlog agrupado
            if (!lhTripsBacklog[lhTrip]) {
                lhTripsBacklog[lhTrip] = [];
            }
            lhTripsBacklog[lhTrip].push(...pedidos);
            
            // Marcar para remoção
            lhsParaRemover.push(lhTrip);
            lhsMovidas.push(lhTrip);
        }
        // Caso 3: Alto volume sem dados → NÃO mover (mantém planejável)
        else if (!isBaixoVolume && !dadosPlanilhaLH) {
            console.log(`   ✅ LH ${lhTrip} (${pedidos.length} pedidos, ALTO volume sem dados) → Mantém planejável`);
        }
    }
    
    // Remover LHs das planejáveis
    lhsParaRemover.forEach(lhTrip => {
        delete lhTripsPlanejáveis[lhTrip];
    });
    
    if (pedidosMovidos > 0) {
        console.log(`✅ [RECLASSIFICAÇÃO] ${pedidosMovidos} pedidos movidos para Backlog`);
        console.log(`📊 [RECLASSIFICAÇÃO] ${lhsMovidas.length} LHs reclassificadas: ${lhsMovidas.join(', ')}`);
        console.log(`📊 [RECLASSIFICAÇÃO] Novo total Backlog: ${pedidosBacklogPorStatus.length} pedidos\n`);
    } else {
        console.log(`ℹ️ [RECLASSIFICAÇÃO] Nenhum pedido "Sinalizar Inventário" encontrado\n`);
    }
}

// Estado para Ciclos (OpsClock e Outbound)
let dadosOpsClock = []; // Horários dos ciclos por station
let dadosOutbound = []; // Capacidade por ciclo/data
let cicloSelecionado = 'Todos'; // Filtro de ciclo atual

// Cache de validações SPX (persiste entre trocas de ciclo)
let cacheSPX = new Map(); // Map<lhId, {status, statusCodigo, chegadaReal, timestamp}>
let dataCicloSelecionada = null; // Data selecionada para o planejamento do ciclo

// Estado para CAP Manual
let capsManual = {};

// Estado para medição de tempo de execução
let tempoInicioExecucao = null;
let tempoFimExecucao = null;

// Função para extrair nome da station da pasta
function extrairNomeStation(caminhoOuNome) {
    if (!caminhoOuNome) return null;
    
    // Se for um caminho, pegar só o nome da pasta
    let nome = caminhoOuNome;
    if (caminhoOuNome.includes('\\') || caminhoOuNome.includes('/')) {
        const partes = caminhoOuNome.split(/[\\\/]/);
        nome = partes[partes.length - 1];
    }
    
    // Garantir que nome seja string válida
    if (!nome || typeof nome !== 'string') {
        return null;
    }
    
    // Remover apenas sufixos específicos que NÃO são parte do nome da station
    // Exemplo: "LM Hub_GO_Aparecida de Goiânia_ St. Empr_02" -> remove "_ St. Empr_02"
    // MAS manter: "LM Hub_MG_Belo Horizonte_02" -> mantém "_02" pois é outra station
    nome = nome.replace(/[_\s]*St\.?\s*Empr[_\s]*\d*/gi, '');
    
    // NÃO remover mais o _\d+ do final, pois pode ser parte do nome da station
    // nome = nome.replace(/_\d+$/, '');  // REMOVIDO
    
    nome = nome.trim();
    
    return nome;
}

// ======================= INICIALIZAÇÃO =======================
document.addEventListener('DOMContentLoaded', () => {
    console.log('🚀 Interface carregada');
    
    // ✅ DEFINIR TÍTULO E VERSÃO
    const packageJson = require('./package.json');
    document.title = `Shopee - Planejamento Fluxo Integrado v${packageJson.version}`;
    
    // Atualizar versão no cabeçalho interno
    const appVersionElement = document.getElementById('appVersion');
    if (appVersionElement) {
        appVersionElement.textContent = `v${packageJson.version}`;
        console.log(`🏷️ Versão definida: v${packageJson.version}`);
    }

    // ✅ VERIFICAR SE USUÁRIO JÁ FEZ LOGIN
    verificarLoginUsuario();

    // Carregar stations
    carregarStations();
    carregarCapsManual();

    // Inicializar autocomplete de stations
    initStationAutocomplete();

    // Carregar configuração de colunas salvas
    carregarConfigColunas();

    // Carregar configurações do navegador
    carregarConfigNavegador();

    // Event Listeners - Header
    document.getElementById('btnDownload').addEventListener('click', iniciarDownload);
    document.getElementById('btnSelectFile').addEventListener('click', selecionarArquivo);
    document.getElementById('btnConfigStations').addEventListener('click', abrirModalStations);

    // Event Listeners - Tabs
    document.querySelectorAll('.tab').forEach(tab => {
        tab.addEventListener('click', () => trocarAba(tab.dataset.tab));
    });

    // Event Listeners - Modal Stations
    document.getElementById('closeModalStations').addEventListener('click', fecharModalStations);
    document.getElementById('btnFecharModalStations').addEventListener('click', fecharModalStations);
    document.getElementById('btnAddStation').addEventListener('click', adicionarStation);

    // Event Listeners - Aba Configurações (Navegador)
    document.getElementById('configHeadless') ?.addEventListener('change', salvarConfigNavegador);
    document.getElementById('btnLimparSessao') ?.addEventListener('click', limparSessaoLogin);

    // Event Listeners - Aba Planejamento Hub
    document.getElementById('btnAtualizarPlanilha')?.addEventListener('click', atualizarPlanilhaGoogle);
    document.getElementById('filtroPlanejamentoBusca')?.addEventListener('input', renderizarTabelaPlanejamento);
    
    // Event Listeners - Seleção de LHs no Planejamento
    document.getElementById('btnSugerirPlanejamento')?.addEventListener('click', sugerirPlanejamentoAutomatico);
    document.getElementById('btnGerarPlanejamento')?.addEventListener('click', iniciarGeracaoPlanejamento);
    
    // Event Listener - Toggle painel colapsável
    document.getElementById('btnTogglePainel')?.addEventListener('click', togglePainelColapsavel);
    
    // Event Listeners - Seletor de Data do Ciclo
    document.getElementById('dataCicloSelecionada')?.addEventListener('change', onDataCicloChange);
    document.getElementById('btnDataHoje')?.addEventListener('click', setDataCicloHoje);
    
    // Inicializar data do ciclo com hoje
    inicializarDataCiclo();
    
    // Event Listeners - Aba Backlog
    document.getElementById('btnSelecionarTodosBacklog')?.addEventListener('click', selecionarTodosBacklog);
    document.getElementById('btnLimparSelecaoBacklog')?.addEventListener('click', limparSelecaoBacklog);
    document.getElementById('btnConfirmarBacklog')?.addEventListener('click', confirmarBacklog);
    document.getElementById('checkTodosBacklog')?.addEventListener('change', toggleTodosBacklog);
    


    // Fechar modal ao clicar fora
    document.getElementById('modalStations').addEventListener('click', (e) => {
        if (e.target.id === 'modalStations') fecharModalStations();
    });

    // Verificar status da sessão
    verificarStatusSessao();

    // Carregar dados locais da planilha
    carregarDadosPlanilhaLocal();
    
    // ===== ATALHOS DE TECLADO =====
    // CTRL+U: Usar fase de captura para garantir que seja pego primeiro!
    document.addEventListener('keydown', (e) => {
        if (e.ctrlKey && e.key === 'u') {
            console.log('🔍 DEBUG: CTRL+U CAPTURADO NA FASE DE CAPTURA!');
            e.preventDefault();
            e.stopPropagation();
            e.stopImmediatePropagation();
            
            try {
                console.log('🔍 DEBUG: Chamando toggleModoHeadless()...');
                toggleModoHeadless();
                console.log('🔍 DEBUG: toggleModoHeadless() chamado com sucesso!');
            } catch (error) {
                console.error('❌ ERRO ao chamar toggleModoHeadless:', error);
            }
            
            return false;
        }
    }, true); // true = capture phase (pega primeiro!)
    
    // Outros atalhos
    document.addEventListener('keydown', (e) => {
        
        // Ctrl+/: Abrir Easter Egg (Sobre o Projeto)
        if (e.ctrlKey && e.key === '/') {
            e.preventDefault();
            abrirModalEasterEgg();
        }
        
        // Esc: Fechar aba Configurações se estiver aberta
        if (e.key === 'Escape') {
            const tabConfig = document.getElementById('tab-config');
            if (tabConfig && tabConfig.classList.contains('active')) {
                fecharAbaConfiguracoes();
            }
        }
    });
    
    // Event Listener - Botão fechar configurações
    document.getElementById('btnFecharConfig')?.addEventListener('click', fecharAbaConfiguracoes);
});

// ======================= TABS =======================
function trocarAba(tabId) {
    // Atualizar botões das tabs
    document.querySelectorAll('.tab').forEach(tab => {
        tab.classList.toggle('active', tab.dataset.tab === tabId);
    });

    // Atualizar conteúdo das tabs
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.toggle('active', content.id === `tab-${tabId}`);
    });

    // Se mudou para Configurações, atualizar lista de colunas
    if (tabId === 'config') {
        atualizarListaColunas();
    }

    // Se mudou para Planejamento Hub, atualizar tabela
    if (tabId === 'planejamento') {
        renderizarTabelaPlanejamento();
    }
    
    // Se mudou para Backlog, atualizar tabela
    if (tabId === 'backlog') {
        renderizarBacklog();
    }
}

// ======================= STATIONS =======================
async function carregarStations() {
    try {
        const resultado = await ipcRenderer.invoke('carregar-stations');
        if (resultado.success) {
            stationsCadastradas = resultado.stations;
            atualizarSelectStations();
            console.log(`✅ ${stationsCadastradas.length} stations carregadas`);
        }
    } catch (error) {
        console.error('Erro ao carregar stations:', error);
    }
}

function atualizarSelectStations() {
    const select = document.getElementById('stationSelect');
    const searchInput = document.getElementById('stationSearchInput');

    if (select) {
        select.innerHTML = '<option value="">-- Station atual do sistema --</option>';

        // Agrupar por UF
        const porUF = {};
        stationsCadastradas.forEach(s => {
            const uf = s.uf || 'Outros';
            if (!porUF[uf]) porUF[uf] = [];
            porUF[uf].push(s);
        });

        // Criar optgroups
        Object.keys(porUF).sort().forEach(uf => {
            const group = document.createElement('optgroup');
            group.label = `📍 ${uf}`;

            porUF[uf].forEach(station => {
                const option = document.createElement('option');
                option.value = station.nome;
                option.textContent = station.nome;
                group.appendChild(option);
            });

            select.appendChild(group);
        });
    }

    // Atualizar placeholder do input de busca
    if (searchInput) {
        const stationAtual = document.getElementById('stationSelecionada') ?.value;
        if (stationAtual) {
            searchInput.value = stationAtual;
        } else {
            searchInput.placeholder = stationsCadastradas.length > 0 ?
                `Buscar entre ${stationsCadastradas.length} stations...` :
                'Nenhuma station cadastrada';
        }
    }
}

// Inicializar autocomplete de stations
function initStationAutocomplete() {
    const searchInput = document.getElementById('stationSearchInput');
    const suggestions = document.getElementById('stationSuggestions');
    const hiddenInput = document.getElementById('stationSelecionada');

    if (!searchInput || !suggestions) return;

    // Ao digitar
    searchInput.addEventListener('input', () => {
        const termo = searchInput.value.toLowerCase().trim();

        if (termo.length === 0) {
            suggestions.classList.remove('active');
            hiddenInput.value = '';
            return;
        }

        // Filtrar stations
        const resultados = stationsCadastradas.filter(s => {
            const nome = s.nome.toLowerCase();
            const codigo = (s.codigo || '').toLowerCase();
            const uf = (s.uf || '').toLowerCase();
            return nome.includes(termo) || codigo.includes(termo) || uf.includes(termo);
        });

        if (resultados.length === 0) {
            suggestions.innerHTML = '<div class="station-no-results">Nenhuma station encontrada</div>';
        } else {
            suggestions.innerHTML = resultados.map(s => `
                <div class="station-suggestion-item" data-nome="${s.nome}">
                    <div class="station-suggestion-name">${s.nome}</div>
                    <div class="station-suggestion-info">
                        ${s.codigo ? `Código: ${s.codigo}` : ''}
                        ${s.codigo && s.uf ? ' | ' : ''}
                        ${s.uf ? `UF: ${s.uf}` : ''}
                    </div>
                </div>
            `).join('');

            // Event listeners nos itens
            suggestions.querySelectorAll('.station-suggestion-item').forEach(item => {
                item.addEventListener('click', () => {
                    const nome = item.dataset.nome;
                    searchInput.value = nome;
                    hiddenInput.value = nome;
                    suggestions.classList.remove('active');
                    // Atualizar ciclos quando station for selecionada
                    atualizarInfoCiclos();
                });
            });
        }

        suggestions.classList.add('active');
    });

    // Ao focar
    searchInput.addEventListener('focus', () => {
        if (searchInput.value.length > 0 || stationsCadastradas.length > 0) {
            // Mostrar todas se não tem busca
            if (searchInput.value.length === 0 && stationsCadastradas.length > 0) {
                suggestions.innerHTML = stationsCadastradas.slice(0, 10).map(s => `
                    <div class="station-suggestion-item" data-nome="${s.nome}">
                        <div class="station-suggestion-name">${s.nome}</div>
                        <div class="station-suggestion-info">
                            ${s.codigo ? `Código: ${s.codigo}` : ''}
                            ${s.codigo && s.uf ? ' | ' : ''}
                            ${s.uf ? `UF: ${s.uf}` : ''}
                        </div>
                    </div>
                `).join('');

                if (stationsCadastradas.length > 10) {
                    suggestions.innerHTML += `<div class="station-no-results">Digite para buscar mais ${stationsCadastradas.length - 10} stations...</div>`;
                }

                suggestions.querySelectorAll('.station-suggestion-item').forEach(item => {
                    item.addEventListener('click', () => {
                        const nome = item.dataset.nome;
                        searchInput.value = nome;
                        hiddenInput.value = nome;
                        suggestions.classList.remove('active');
                        // Atualizar ciclos quando station for selecionada
                        atualizarInfoCiclos();
                    });
                });

                suggestions.classList.add('active');
            }
        }
    });

    // Fechar ao clicar fora
    document.addEventListener('click', (e) => {
        if (!searchInput.contains(e.target) && !suggestions.contains(e.target)) {
            suggestions.classList.remove('active');
        }
    });

    // Limpar ao pressionar Escape ou ao apagar tudo
    searchInput.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            suggestions.classList.remove('active');
            searchInput.blur();
        }
    });
}

function abrirModalStations() {
    document.getElementById('modalStations').classList.add('active');
    renderizarListaStations();
}

function fecharModalStations() {
    document.getElementById('modalStations').classList.remove('active');
}

function renderizarListaStations() {
    const container = document.getElementById('stationsList');

    if (stationsCadastradas.length === 0) {
        container.innerHTML = '<p style="text-align:center;color:#999;padding:20px;">Nenhuma station cadastrada</p>';
        return;
    }

    container.innerHTML = stationsCadastradas.map((station, index) => `
        <div class="station-item">
            <div class="station-info">
                <span class="station-nome">${station.nome}</span>
                <span class="station-meta">Código: ${station.codigo || '-'} | UF: ${station.uf || '-'}</span>
            </div>
            <button class="btn btn-delete" onclick="removerStation(${index})">🗑️ Remover</button>
        </div>
    `).join('');
}

async function adicionarStation() {
    const nome = document.getElementById('inputStationNome').value.trim();
    const codigo = document.getElementById('inputStationCodigo') ?.value.trim() || '';
    const uf = document.getElementById('inputStationUF') ?.value.trim().toUpperCase() || '';

    if (!nome) {
        alert('Digite o nome da station');
        return;
    }

    stationsCadastradas.push({
        nome,
        codigo,
        uf
    });

    await salvarStations();
    atualizarSelectStations();
    renderizarListaStations();

    // Limpar campos
    document.getElementById('inputStationNome').value = '';
    if (document.getElementById('inputStationCodigo')) {
        document.getElementById('inputStationCodigo').value = '';
    }
    if (document.getElementById('inputStationUF')) {
        document.getElementById('inputStationUF').value = '';
    }
}

async function removerStation(index) {
    if (confirm(`Remover station "${stationsCadastradas[index].nome}"?`)) {
        stationsCadastradas.splice(index, 1);
        await salvarStations();
        atualizarSelectStations();
        renderizarListaStations();
    }
}

async function salvarStations() {
    try {
        await ipcRenderer.invoke('salvar-stations', stationsCadastradas);
    } catch (error) {
        console.error('Erro ao salvar stations:', error);
    }
}

// ======================= DOWNLOAD =======================
async function iniciarDownload() {
    // Tentar pegar do autocomplete primeiro, depois do select antigo
    let stationSelecionada = document.getElementById('stationSelecionada')?.value ||
        document.getElementById('stationSearchInput')?.value ||
        document.getElementById('stationSelect')?.value ||
        null;
    
    // Garantir que seja uma string válida
    if (!stationSelecionada || stationSelecionada === 'null' || stationSelecionada === 'undefined') {
        stationSelecionada = null; // Deixar null para o backend usar station padrão
    }
    const configNav = getConfigNavegador();

    mostrarLoading('Baixando dados...', stationSelecionada ?
        `Station: ${stationSelecionada}` :
        'Usando station atual do sistema', true); // true = mostrar progresso

    atualizarProgresso(1, 'Verificando login...');

    try {
        // Passar apenas o nome da station
        // O headless é controlado pelo CTRL+U (variável global no main.js)
        const resultado = await ipcRenderer.invoke('executar-download', {
            stationNome: stationSelecionada
        });

        esconderLoading();

        if (resultado.success) {
            // GUARDAR A PASTA DA STATION para usar nos exports
            if (resultado.outputDir) {
                pastaStationAtual = resultado.outputDir;
                console.log('📁 Pasta da station:', pastaStationAtual);
                
                // Extrair nome da station
                stationAtualNome = extrairNomeStation(pastaStationAtual);
                console.log('📍 Station atual:', stationAtualNome);
            }

            alert('✅ Download concluído com sucesso!');

            // CARREGAR AUTOMATICAMENTE o arquivo baixado
            if (resultado.filePath) {
                mostrarLoading('Carregando dados...', resultado.filePath);
                await carregarArquivo(resultado.filePath);
                esconderLoading();
            }
        } else {
            alert(`❌ Erro no download: ${resultado.error}`);
        }
    } catch (error) {
        esconderLoading();
        alert(`❌ Erro: ${error.message}`);
    }
}

// Listener para atualizações de progresso do main process
if (typeof ipcRenderer !== 'undefined') {
    ipcRenderer.on('download-progresso', (event, dados) => {
        if (dados.etapa && dados.mensagem) {
            atualizarProgresso(dados.etapa, dados.mensagem);
        }
    });
}

async function selecionarArquivo() {
    try {
        const resultado = await ipcRenderer.invoke('selecionar-arquivo');

        if (!resultado.canceled && resultado.filePath) {
            mostrarLoading('Carregando arquivo...', resultado.filePath);
            await carregarArquivo(resultado.filePath);
            esconderLoading();
        }
    } catch (error) {
        esconderLoading();
        alert(`❌ Erro ao selecionar arquivo: ${error.message}`);
    }
}

async function carregarArquivo(filePath) {
    try {
        const resultado = await ipcRenderer.invoke('carregar-arquivo', filePath);

        if (resultado.success) {
            dadosAtuais = resultado.dados;
            
            // 🔍 DEBUG: Contar LH específica
            const lhEspecifica = 'LT0Q2H01Z0Y11';
            const pedidosLH = dadosAtuais.filter(row => {
                const lhTrip = row['LH Trip'] || row['LH Trip ID'] || row['LH Task'] || row['LH Task ID'] || '';
                return lhTrip.toString().trim() === lhEspecifica;
            });
            
            if (pedidosLH.length > 0) {
                console.log(`\n🔍 DEBUG FRONT: ${lhEspecifica}`);
                console.log(`   Pedidos recebidos do main.js: ${pedidosLH.length}`);
                console.log(`   Total de dados: ${dadosAtuais.length}`);
            }
            
            // EXTRAIR PASTA DA STATION do caminho do arquivo
            if (filePath) {
                const path = require('path');
                pastaStationAtual = path.dirname(filePath);
                console.log('📁 Pasta da station (do arquivo):', pastaStationAtual);
                
                // Extrair nome da station da pasta
                stationAtualNome = extrairNomeStation(pastaStationAtual);
                console.log('📍 Station atual:', stationAtualNome);
                
                // Atualizar input da station
                const stationInput = document.getElementById('stationSearchInput');
                if (stationInput && stationAtualNome) {
                    stationInput.value = stationAtualNome;
                }
                
                // Atualizar ciclos para a station
                atualizarInfoCiclos();
            }
            
            processarDados();
            console.log(`✅ ${dadosAtuais.length} registros carregados`);
            
            // 🔄 SINCRONIZAÇÃO AUTOMÁTICA após download
            console.log('🔄 Sincronizando planilhas automaticamente...');
            setTimeout(() => {
                atualizarPlanilhaGoogle();
            }, 1000); // Aguarda 1 segundo para garantir que dados foram processados
        } else {
            alert(`❌ Erro ao carregar: ${resultado.error}`);
        }
    } catch (error) {
        alert(`❌ Erro: ${error.message}`);
    }
}

// ======================= PROCESSAMENTO DE DADOS =======================
function processarDados() {
    // Capturar TODAS as colunas do arquivo
    if (dadosAtuais.length > 0) {
        todasColunas = Object.keys(dadosAtuais[0]);

        // Se não tem colunas configuradas, usar todas
        if (colunasVisiveis.length === 0) {
            colunasVisiveis = [...todasColunas];
        } else {
            // Filtrar apenas colunas que existem no arquivo atual
            colunasVisiveis = colunasVisiveis.filter(col => todasColunas.includes(col));

            // Se ficou vazia, usar todas
            if (colunasVisiveis.length === 0) {
                colunasVisiveis = [...todasColunas];
            }
        }
    }

    // Encontrar coluna de LH Trip
    const colunaLH = todasColunas.find(col =>
        col.toLowerCase().includes('lh trip') ||
        col.toLowerCase().includes('lh_trip') ||
        col.toLowerCase().includes('lhtask') ||
        col.toLowerCase().includes('lh task')
    ) || 'LH Trip';
    
    // Encontrar coluna de STATUS
    const colunaStatus = todasColunas.find(col =>
        col.toLowerCase() === 'status' ||
        col.toLowerCase().includes('status')
    ) || 'STATUS';
    
    console.log(`📋 Coluna LH: "${colunaLH}", Coluna Status: "${colunaStatus}"`);

    // Resetar agrupamentos
    lhTrips = {};
    pedidosBacklogPorStatus = [];
    pedidosPlanejáveis = [];
    lhTripsBacklog = {};
    lhTripsPlanejáveis = {};
    
    // Debug: coletar todos os status únicos para verificação
    const statusUnicos = new Set();

    // Separar pedidos por STATUS
    // 🔍 DEBUG: Rastrear LH específica ANTES do processamento
    const lhEspecificaDebug = 'LT0Q2H01Z0Y11';
    const pedidosLHAntes = dadosAtuais.filter(row => {
        const lhTrip = row['LH Trip'] || row['LH Trip ID'] || row['LH Task'] || row['LH Task ID'] || '';
        return lhTrip.toString().trim() === lhEspecificaDebug;
    });
    console.log(`\n🔍 DEBUG ANTES PROCESSAMENTO: ${lhEspecificaDebug}`);
    console.log(`   Total de pedidos: ${pedidosLHAntes.length}`);
    
    dadosAtuais.forEach(row => {
        let lh = row[colunaLH] || '(vazio)';
        const status = row[colunaStatus] || '';
        
        // Coletar status únicos para debug
        statusUnicos.add(status);
        
        // Verificar se é backlog pelo status (função mais robusta)
        const isBacklog = isStatusBacklog(status);
        
        // Se é backlog, renomear a LH para "Backlog" (igual ao VBA)
        if (isBacklog) {
            // Manter a LH original numa propriedade auxiliar para referência
            row._lhOriginal = lh;
            // Renomear para Backlog
            row[colunaLH] = 'Backlog';
            lh = 'Backlog';
        }
        
        // Agrupar todos por LH (para a aba LH Trips - visualização geral)
        if (!lhTrips[lh]) {
            lhTrips[lh] = [];
        }
        lhTrips[lh].push(row);
        
        if (isBacklog) {
            // É backlog - vai para aba "Tratar Backlog"
            pedidosBacklogPorStatus.push(row);
            
            // Agrupar por LH original também (para referência)
            const lhOriginal = row._lhOriginal || '(vazio)';
            if (!lhTripsBacklog[lhOriginal]) {
                lhTripsBacklog[lhOriginal] = [];
            }
            lhTripsBacklog[lhOriginal].push(row);
        } else {
            // É planejável - vai para aba "Planejamento Hub"
            pedidosPlanejáveis.push(row);
            
            // Agrupar por LH também (só se tiver LH válido)
            if (lh && lh !== '(vazio)' && lh !== 'Backlog') {
                if (!lhTripsPlanejáveis[lh]) {
                    lhTripsPlanejáveis[lh] = [];
                }
                lhTripsPlanejáveis[lh].push(row);
            }
        }
    });
    
    // 🔍 DEBUG: Contar pedidos da LH específica após loop
    const pedidosLHBacklog = lhTripsBacklog[lhEspecificaDebug] || [];
    const pedidosLHPlanejavel = lhTripsPlanejáveis[lhEspecificaDebug] || [];
    const totalProcessado = pedidosLHBacklog.length + pedidosLHPlanejavel.length;
    
    console.log(`\n🔍 DEBUG APÓS LOOP: ${lhEspecificaDebug}`);
    console.log(`   Backlog: ${pedidosLHBacklog.length}`);
    console.log(`   Planejável: ${pedidosLHPlanejavel.length}`);
    console.log(`   Total: ${totalProcessado}`);
    console.log(`   Diferença: ${pedidosLHAntes.length - totalProcessado} pedidos`);
    
    if (pedidosLHAntes.length !== totalProcessado) {
        console.log(`\n❌ PERDEU ${pedidosLHAntes.length - totalProcessado} PEDIDOS!`);
        console.log(`   Investigando...`);
        
        // Encontrar os pedidos perdidos
        const pedidosPerdidos = pedidosLHAntes.filter(pedido => {
            const estaNoPlanejavel = pedidosLHPlanejavel.includes(pedido);
            const estaNoBacklog = pedidosLHBacklog.includes(pedido);
            return !estaNoPlanejavel && !estaNoBacklog;
        });
        
        console.log(`   Pedidos perdidos: ${pedidosPerdidos.length}`);
        pedidosPerdidos.forEach((pedido, i) => {
            const lhCol = pedido[colunaLH];
            const status = pedido[colunaStatus];
            const isBack = isStatusBacklog(status);
            console.log(`   [${i+1}] LH: "${lhCol}", Status: "${status}", isBacklog: ${isBack}`);
        });
    }
    
    // Log de status únicos encontrados
    console.log(`📋 Status únicos encontrados:`, [...statusUnicos]);
    
    // Log de separação
    const totalBacklog = pedidosBacklogPorStatus.length;
    const totalPlanejavel = pedidosPlanejáveis.length;
    const lhsBacklog = Object.keys(lhTripsBacklog).filter(lh => lh !== '(vazio)').length;
    const lhsPlanejáveis = Object.keys(lhTripsPlanejáveis).length;
    
    console.log(`📊 SEPARAÇÃO POR STATUS:`);
    console.log(`   🔴 Backlog: ${totalBacklog} pedidos em ${lhsBacklog} LHs originais`);
    console.log(`   🟢 Planejável: ${totalPlanejavel} pedidos em ${lhsPlanejáveis} LHs`);
    
    // 🆕 PÓS-PROCESSAMENTO: Mover pedidos "Sinalizar Inventário" para Backlog
    reclassificarSinalizarInventarioParaBacklog();
    
    // Atualizar interface
    renderizarListaLHs();

    // Resetar seleções
    lhTripAtual = null;

    console.log(`📊 ${todasColunas.length} colunas encontradas`);
    console.log(`👁️ ${colunasVisiveis.length} colunas visíveis`);
}

// ======================= ABA LH TRIPS =======================
function renderizarListaLHs() {
    const container = document.getElementById('lhList');
    const countEl = document.getElementById('lhCount');

    if (!container) return;

    // 🗑️ Filtrar LHs lixo sistêmico da lista
    const lhsLixoSet = new Set(lhsLixoSistemico.map(row => row.lh_trip));
    
    const lhKeys = Object.keys(lhTrips)
        .filter(lh => !lhsLixoSet.has(lh)) // Remover LHs lixo
        .sort((a, b) => {
            // "TODOS" primeiro, depois "(vazio)", depois ordem alfabética
            if (a === 'TODOS') return -1;
            if (b === 'TODOS') return 1;
            if (a === '(vazio)') return 1;
            if (b === '(vazio)') return -1;
            return lhTrips[b].length - lhTrips[a].length; // Ordenar por quantidade
        });

    if (countEl) countEl.textContent = `${lhKeys.length} LH Trips encontradas`;

    // Adicionar "TODOS" no início
    let html = `
        <div class="lh-item ${lhTripAtual === 'TODOS' ? 'active' : ''}" data-lh="TODOS">
            <div class="lh-name">📊 TODOS</div>
            <div class="lh-count">${dadosAtuais.length} pedidos</div>
        </div>
    `;

    lhKeys.forEach(lh => {
        const count = lhTrips[lh].length;
        const isVazio = lh === '(vazio)';
        html += `
            <div class="lh-item ${lhTripAtual === lh ? 'active' : ''}" data-lh="${lh}">
                <div class="lh-name">${isVazio ? '⚠️ ' : '🚚 '}${lh}</div>
                <div class="lh-count">${count} pedidos</div>
            </div>
        `;
    });

    container.innerHTML = html;

    // Event listeners para os itens
    container.querySelectorAll('.lh-item').forEach(item => {
        item.addEventListener('click', () => selecionarLH(item.dataset.lh));
    });
}

function selecionarLH(lh) {
    console.log(`🔍 Selecionando LH: "${lh}"`);
    console.log(`   Existe em lhTrips:`, !!lhTrips[lh]);
    console.log(`   Quantidade:`, lhTrips[lh]?.length || 0);
    
    lhTripAtual = lh;

    // Atualizar visual
    document.querySelectorAll('.lh-item').forEach(item => {
        item.classList.toggle('active', item.dataset.lh === lh);
    });

    // Atualizar título
    document.getElementById('selectedLhTitle').textContent = `LH Trip: ${lh}`;

    // Renderizar tabela - usar dadosAtuais para TODOS, senão filtrar por LH
    let pedidos;
    if (lh === 'TODOS') {
        pedidos = dadosAtuais;
    } else {
        pedidos = lhTrips[lh] || [];
    }
    
    console.log(`   Pedidos a exibir: ${pedidos.length}`);
    
    document.getElementById('pedidosCount').textContent = `${pedidos.length} pedidos`;

    renderizarTabela(pedidos);
}

// ======================= FILTRO ESTILO EXCEL =======================
// Estado dos filtros (compartilhado)
let filtrosAtivos = {}; // { coluna: { valores: Set, ordenacao: 'asc'|'desc'|null } }
let filtrosAtivosPlan = {}; // Para Planejamento Hub

// Criar popup de filtro Excel
function criarPopupFiltroExcel(coluna, valores, container, tipoTabela = 'lhtrips') {
    // Remover popup existente
    const popupExistente = document.querySelector('.excel-filter-popup');
    if (popupExistente) popupExistente.remove();
    
    // Valores únicos ordenados
    const valoresUnicos = [...new Set(valores.map(v => String(v ?? '-')))].sort((a, b) => {
        // Tentar ordenar numericamente se possível
        const numA = parseFloat(a);
        const numB = parseFloat(b);
        if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
        return a.localeCompare(b, 'pt-BR');
    });
    
    // Selecionar referência de filtros correta
    let filtrosRef;
    if (tipoTabela === 'planejamento') {
        filtrosRef = filtrosAtivosPlan;
    } else if (tipoTabela === 'backlog') {
        filtrosRef = filtrosAtivosBacklog;
    } else {
        filtrosRef = filtrosAtivos;
    }
    
    const filtroAtual = filtrosRef[coluna] || { valores: new Set(valoresUnicos), ordenacao: null };
    
    const popup = document.createElement('div');
    popup.className = 'excel-filter-popup';
    popup.innerHTML = `
        <div class="excel-filter-header">
            <span>Filtrar: ${coluna}</span>
            <button class="excel-filter-close">&times;</button>
        </div>
        <div class="excel-filter-search">
            <input type="text" placeholder="🔍 Buscar..." class="excel-filter-search-input">
        </div>
        <div class="excel-filter-sort">
            <button class="excel-sort-btn" data-sort="asc">↑ A-Z / Menor</button>
            <button class="excel-sort-btn" data-sort="desc">↓ Z-A / Maior</button>
        </div>
        <div class="excel-filter-select-all">
            <label>
                <input type="checkbox" class="excel-select-all" checked>
                <span>(Selecionar Todos)</span>
            </label>
        </div>
        <div class="excel-filter-list">
            ${valoresUnicos.map(valor => `
                <label class="excel-filter-item" data-valor="${valor}">
                    <input type="checkbox" value="${valor}" ${filtroAtual.valores.has(valor) ? 'checked' : ''}>
                    <span>${valor}</span>
                </label>
            `).join('')}
        </div>
        <div class="excel-filter-actions">
            <button class="btn-excel-aplicar">✓ Aplicar</button>
            <button class="btn-excel-limpar">✕ Limpar</button>
        </div>
    `;
    
    document.body.appendChild(popup);
    
    // Posicionar popup
    const rect = container.getBoundingClientRect();
    popup.style.top = `${rect.bottom + 5}px`;
    popup.style.left = `${Math.min(rect.left, window.innerWidth - 280)}px`;
    
    // Event Listeners
    const searchInput = popup.querySelector('.excel-filter-search-input');
    const selectAll = popup.querySelector('.excel-select-all');
    const checkboxes = popup.querySelectorAll('.excel-filter-item input');
    const lista = popup.querySelector('.excel-filter-list');
    
    // Busca
    searchInput.addEventListener('input', (e) => {
        const termo = e.target.value.toLowerCase();
        popup.querySelectorAll('.excel-filter-item').forEach(item => {
            const valor = item.dataset.valor.toLowerCase();
            item.style.display = valor.includes(termo) ? 'flex' : 'none';
        });
    });
    
    // Selecionar todos
    selectAll.addEventListener('change', (e) => {
        const visiveisCheckboxes = [...checkboxes].filter(cb => cb.closest('.excel-filter-item').style.display !== 'none');
        visiveisCheckboxes.forEach(cb => cb.checked = e.target.checked);
    });
    
    // Atualizar "selecionar todos" quando checkboxes mudam
    checkboxes.forEach(cb => {
        cb.addEventListener('change', () => {
            const visiveisCheckboxes = [...checkboxes].filter(cb => cb.closest('.excel-filter-item').style.display !== 'none');
            const todosMarcados = visiveisCheckboxes.every(cb => cb.checked);
            selectAll.checked = todosMarcados;
        });
    });
    
    // Função para re-renderizar após filtro
    const reRenderizar = () => {
        if (tipoTabela === 'planejamento') {
            renderizarTabelaPlanejamento();
        } else if (tipoTabela === 'backlog') {
            renderizarBacklog();
        } else {
            const pedidosAtuais = lhTripAtual === 'TODOS' ? dadosAtuais : (lhTrips[lhTripAtual] || []);
            renderizarTabela(pedidosAtuais);
        }
    };
    
    // Ordenação
    popup.querySelectorAll('.excel-sort-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const sort = btn.dataset.sort;
            filtrosRef[coluna] = filtrosRef[coluna] || { valores: new Set(valoresUnicos), ordenacao: null };
            filtrosRef[coluna].ordenacao = sort;
            
            popup.remove();
            reRenderizar();
        });
    });
    
    // Aplicar filtro
    popup.querySelector('.btn-excel-aplicar').addEventListener('click', () => {
        const selecionados = new Set();
        checkboxes.forEach(cb => {
            if (cb.checked) selecionados.add(cb.value);
        });
        
        filtrosRef[coluna] = filtrosRef[coluna] || { valores: selecionados, ordenacao: null };
        filtrosRef[coluna].valores = selecionados;
        
        popup.remove();
        reRenderizar();
    });
    
    // Limpar filtro
    popup.querySelector('.btn-excel-limpar').addEventListener('click', () => {
        delete filtrosRef[coluna];
        popup.remove();
        reRenderizar();
    });
    
    // Fechar popup
    popup.querySelector('.excel-filter-close').addEventListener('click', () => popup.remove());
    
    // Fechar ao clicar fora
    setTimeout(() => {
        document.addEventListener('click', function fecharPopup(e) {
            if (!popup.contains(e.target) && !container.contains(e.target)) {
                popup.remove();
                document.removeEventListener('click', fecharPopup);
            }
        });
    }, 100);
    
    // Foco no campo de busca
    searchInput.focus();
}

// Aplicar filtros aos dados
function aplicarFiltrosExcel(dados, colunas, filtrosRef) {
    let resultado = [...dados];
    
    // Aplicar filtros de valores
    Object.keys(filtrosRef).forEach(coluna => {
        const filtro = filtrosRef[coluna];
        if (filtro && filtro.valores && filtro.valores.size > 0) {
            resultado = resultado.filter(row => {
                let valor = row[coluna];
                
                // Tratar colunas especiais que são objetos
                if (coluna === 'status_lh' && valor && typeof valor === 'object') {
                    valor = valor.texto || '-';
                } else if (coluna === 'tempo_corte' && valor && typeof valor === 'object') {
                    valor = valor.texto || '-';
                } else {
                    valor = String(valor ?? '-');
                }
                
                return filtro.valores.has(valor);
            });
        }
    });
    
    // Aplicar ordenação (última coluna com ordenação)
    const colunaOrdenada = Object.keys(filtrosRef).find(col => filtrosRef[col]?.ordenacao);
    if (colunaOrdenada) {
        const ordenacao = filtrosRef[colunaOrdenada].ordenacao;
        resultado.sort((a, b) => {
            let valA = a[colunaOrdenada];
            let valB = b[colunaOrdenada];
            
            // Tratar colunas especiais que são objetos
            if (colunaOrdenada === 'status_lh') {
                valA = valA?.texto ?? '';
                valB = valB?.texto ?? '';
            } else if (colunaOrdenada === 'tempo_corte') {
                // Ordenar por minutos para tempo_corte
                valA = valA?.minutos ?? 9999;
                valB = valB?.minutos ?? 9999;
                return ordenacao === 'asc' ? valA - valB : valB - valA;
            } else {
                valA = valA ?? '';
                valB = valB ?? '';
            }
            
            // Tentar ordenar numericamente
            const numA = parseFloat(valA);
            const numB = parseFloat(valB);
            
            if (!isNaN(numA) && !isNaN(numB)) {
                return ordenacao === 'asc' ? numA - numB : numB - numA;
            }
            
            // Ordenar como string
            const strA = String(valA);
            const strB = String(valB);
            return ordenacao === 'asc' 
                ? strA.localeCompare(strB, 'pt-BR') 
                : strB.localeCompare(strA, 'pt-BR');
        });
    }
    
    return resultado;
}

// Verificar se coluna tem filtro ativo
function temFiltroAtivo(coluna, filtrosRef) {
    const filtro = filtrosRef[coluna];
    if (!filtro) return false;
    if (filtro.ordenacao) return true;
    // Verificar se algum valor foi desmarcado (não todos selecionados)
    return true; // Simplificado - sempre mostra ícone se tem filtro
}

function renderizarTabela(pedidos) {
    const container = document.getElementById('tableContainer');

    if (pedidos.length === 0) {
        container.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">📋</div>
                <h3>Nenhum pedido encontrado</h3>
            </div>
        `;
        return;
    }

    // Usar colunas configuradas, ou todas se não houver configuração
    const colunas = colunasVisiveis.length > 0 ? colunasVisiveis : Object.keys(pedidos[0]);

    // Aplicar filtros Excel
    let pedidosFiltrados = aplicarFiltrosExcel(pedidos, colunas, filtrosAtivos);

    // Construir tabela
    let html = '<table><thead><tr>';
    colunas.forEach(col => {
        const temFiltro = filtrosAtivos[col];
        const icone = temFiltro ? '🔽' : '▼';
        const classeAtivo = temFiltro ? 'filtro-ativo' : '';
        html += `<th class="${classeAtivo}">
            <div class="th-content">
                <span class="th-titulo">${col}</span>
                <button class="btn-filtro-excel" data-coluna="${col}">${icone}</button>
            </div>
        </th>`;
    });
    html += '</tr></thead><tbody>';

    // Mostrar pedidos filtrados (limitar para performance)
    const maxLinhas = 1000;
    const pedidosExibir = pedidosFiltrados.slice(0, maxLinhas);

    pedidosExibir.forEach(row => {
        html += '<tr>';
        colunas.forEach(col => {
            const valor = row[col];
            const valorExibir = valor !== null && valor !== undefined ? valor : '-';
            html += `<td>${valorExibir}</td>`;
        });
        html += '</tr>';
    });

    html += '</tbody></table>';

    // Info de registros
    const totalOriginal = pedidos.length;
    const totalFiltrado = pedidosFiltrados.length;
    
    html += `<div class="tabela-info">`;
    if (totalFiltrado !== totalOriginal) {
        html += `<span>🔍 Filtrado: ${totalFiltrado.toLocaleString('pt-BR')} de ${totalOriginal.toLocaleString('pt-BR')}</span>`;
    } else {
        html += `<span>📊 Total: ${totalOriginal.toLocaleString('pt-BR')} registros</span>`;
    }
    if (totalFiltrado > maxLinhas) {
        html += `<span>⚠️ Mostrando ${maxLinhas.toLocaleString('pt-BR')} linhas</span>`;
    }
    if (Object.keys(filtrosAtivos).length > 0) {
        html += `<button class="btn-limpar-todos-filtros" onclick="limparTodosFiltros()">🗑️ Limpar Filtros</button>`;
    }
    html += `</div>`;

    container.innerHTML = html;
    
    // Adicionar event listeners nos botões de filtro
    container.querySelectorAll('.btn-filtro-excel').forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            const coluna = btn.dataset.coluna;
            const valoresColuna = pedidos.map(row => row[coluna]);
            criarPopupFiltroExcel(coluna, valoresColuna, btn, 'lhtrips');
        });
    });
}

// Limpar todos os filtros (LH Trips)
function limparTodosFiltros() {
    filtrosAtivos = {};
    const pedidosAtuais = lhTripAtual === 'TODOS' ? dadosAtuais : (lhTrips[lhTripAtual] || []);
    renderizarTabela(pedidosAtuais);
}

// Limpar todos os filtros (Planejamento)
function limparTodosFiltrosPlan() {
    filtrosAtivosPlan = {};
    renderizarTabelaPlanejamento();
}

// ======================= SUGESTÃO AUTOMÁTICA DE PLANEJAMENTO =======================
function sugerirPlanejamentoAutomatico() {
    // Marcar início da execução
    tempoInicioExecucao = Date.now();
    
    console.log("🎯 Função sugerirPlanejamentoAutomatico chamada");
    console.log("📍 cicloSelecionado atual:", cicloSelecionado);
    
    // DETECÇÃO AUTOMÁTICA: Se tem CAP Manual definido, usa ele
    let cicloParaUsar = cicloSelecionado;
    
    if (!cicloParaUsar || cicloParaUsar === 'Todos') {
        // Procurar se tem algum CAP Manual definido
        const ciclosComCapManual = [];
        ['AM', 'PM1', 'PM2'].forEach(c => {
            if (obtemCapManual(c) !== null) {
                ciclosComCapManual.push(c);
            }
        });
        
        if (ciclosComCapManual.length === 1) {
            // Tem exatamente 1 CAP Manual definido - usar ele automaticamente!
            cicloParaUsar = ciclosComCapManual[0];
            cicloSelecionado = cicloParaUsar; // Atualizar variável global
            console.log("✅ CAP Manual detectado automaticamente:", cicloParaUsar);
        } else if (ciclosComCapManual.length > 1) {
            // Tem mais de 1 CAP Manual - precisa escolher
            alert('⚠️ Múltiplos CAPs Manuais definidos!\n\nSelecione o ciclo desejado clicando em AM, PM1 ou PM2.');
            return;
        } else {
            // Não tem CAP Manual - precisa selecionar ciclo manualmente
            alert('⚠️ Selecione um ciclo primeiro!\n\nClique em AM, PM1 ou PM2 nos cards de ciclos para definir o target.');
            return;
        }
    }
    
    // Verificar se está usando CAP Manual
    const temCapManual = obtemCapManual(cicloParaUsar) !== null;
    const tipoCAP = temCapManual ? 'CAP MANUAL' : 'CAP Automático';
    
    // Pegar CAP do ciclo
    const capCiclo = obterCapacidadeCiclo(cicloParaUsar);
    
    if (!capCiclo || capCiclo === 0) {
        alert(`⚠️ Não foi possível obter a capacidade do ciclo ${cicloParaUsar}.\n\nVerifique se a planilha de capacidade está atualizada.`);
        return;
    }
    
    console.log('🎯 Iniciando sugestão automática para ciclo ' + cicloParaUsar);
    console.log('📊 Usando: ' + tipoCAP);
    console.log('📊 CAP do ciclo: ' + capCiclo.toLocaleString('pt-BR') + ' pedidos');
    
    // Mostrar confirmação quando usar CAP Manual
    if (temCapManual) {
        const confirmar = confirm(
            '🎯 Sugestão de Planejamento\n\n' +
            'Ciclo: ' + cicloParaUsar + '\n' +
            'Usando: CAP MANUAL\n' +
            'Capacidade: ' + capCiclo.toLocaleString('pt-BR') + ' pedidos\n\n' +
            'Deseja continuar?'
        );
        if (!confirmar) return;
    }
    
    // Limpar seleções anteriores
    lhsSelecionadasPlan.clear();
    pedidosBacklogSelecionados.clear();
    
    // Variáveis para sugestão de complemento
    let lhsNoPisoComEstouro = []; // LHs P0 que estouram CAP (para sugestão de TOs)
    
    let totalSelecionado = 0;
    const lhsSugeridas = [];
    const backlogSugerido = [];
    let lhsBloqueadas = 0;
    let lhsForaDoCorte = 0;
    
    // ===== PASSO 1: BACKLOG (pedidos com status LMHub_Received ou Return_LMHub_Received) =====
    // Isso inclui pedidos COM ou SEM LH que tenham esse status
    console.log(`📦 Backlog total disponível: ${pedidosBacklogPorStatus.length} pedidos`);
    
    // Adicionar backlog (prioridade FIFO)
    pedidosBacklogPorStatus.forEach((pedido, index) => {
        if (totalSelecionado < capCiclo) {
            const id = getShipmentIdFromPedido(pedido, index);
            pedidosBacklogSelecionados.add(id);
            backlogSugerido.push(id);
            totalSelecionado++;
        }
    });
    
    console.log(`✅ Backlog selecionado: ${backlogSugerido.length} pedidos`);
    console.log(`📊 Total após backlog: ${totalSelecionado.toLocaleString('pt-BR')}`);
    
    // ===== PASSO 2: LHs PLANEJÁVEIS - PRIORIDADE FIFO (mais antigas primeiro) =====
    if (totalSelecionado < capCiclo) {
        // Pegar apenas LHs planejáveis (não backlog por status)
        const lhsDoSPX = Object.keys(lhTripsPlanejáveis).filter(lh => lh && lh !== '(vazio)' && lh.trim() !== '');
        
        // Montar array com info de cada LH incluindo status e data
        const lhsComInfo = lhsDoSPX.map(lhTrip => {
            const qtdPedidos = lhTripsPlanejáveis[lhTrip]?.length || 0;
            // ✅ USAR FUNÇÃO DE FILTRO POR STATION
            const dadosPlanilhaLH = buscarDadosPlanilhaPorStation(lhTrip);
            
            // ✅ PASSAR CICLO COMO PARÂMETRO para calcular tempo de corte correto
            const tempoCorte = calcularTempoCorte(dadosPlanilhaLH, cicloParaUsar);
            const statusLH = calcularStatusLH(dadosPlanilhaLH, null, null, cicloParaUsar);
            
            // Extrair data da previsão para ordenação FIFO
            const previsaoFinal = dadosPlanilhaLH?.previsao_final || '';
            const dataPrevisao = extrairDataParaOrdenacao(previsaoFinal);
            
            return {
                lhTrip,
                qtdPedidos,
                tempoCorte,
                minutosCorte: tempoCorte.minutos,
                dentroLimite: tempoCorte.dentroLimite,
                statusLH,
                isBacklogPiso: statusLH.codigo === 'P0B',
                dataPrevisao
            };
        });
        
        // Ordenar por: Data → Hora → Status (mais antigo primeiro - FIFO)
        console.log('📊 Ordenando LHs...');
        lhsComInfo.forEach(lh => {
            console.log(`  ${lh.lhTrip}: ${lh.dataPrevisao ? lh.dataPrevisao.toISOString() : 'SEM DATA'} (${lh.statusLH?.codigo}) - Corte: ${lh.minutosCorte !== null ? lh.minutosCorte + 'min' : 'N/A'}`);
        });
        
        lhsComInfo.sort((a, b) => {
            // Prioridade 1: DATA de previsão (mais cedo primeiro - CRESCENTE)
            if (a.dataPrevisao && b.dataPrevisao) {
                const diffData = a.dataPrevisao - b.dataPrevisao;
                if (diffData !== 0) return diffData; // Datas diferentes - MAIS ANTIGA PRIMEIRO
            }
            // Se só uma tem data, priorizar a que tem
            if (a.dataPrevisao && !b.dataPrevisao) return -1;
            if (!a.dataPrevisao && b.dataPrevisao) return 1;
            
            // Prioridade 2: HORA de previsão (já considerada na dataPrevisao acima)
            // Se chegaram aqui, têm a mesma data+hora
            
            // Prioridade 3: STATUS - Full antes de Atrasada
            const statusPrioridade = {
                'F0': 1,  // Full - máxima prioridade
                'P0': 2,  // No Prazo
                'P1': 3,  // Aguardando Descarga
                'A1': 4,  // Atrasada
                'A2': 5,  // Muito Atrasada
                'P0B': 6  // Backlog Piso (já processado antes, mas por segurança)
            };
            
            const prioA = statusPrioridade[a.statusLH?.codigo] || 99;
            const prioB = statusPrioridade[b.statusLH?.codigo] || 99;
            
            if (prioA !== prioB) return prioA - prioB;
            
            // Prioridade 4: Quantidade de pedidos (maior primeiro)
            return b.qtdPedidos - a.qtdPedidos;
        });
        
        console.log('📊 LHs após ordenação:');
        lhsComInfo.forEach(lh => {
            console.log(`  ${lh.lhTrip}: ${lh.dataPrevisao ? lh.dataPrevisao.toISOString() : 'SEM DATA'} (${lh.statusLH?.codigo}) - Corte: ${lh.minutosCorte !== null ? lh.minutosCorte + 'min' : 'N/A'}`);
        });
        
        // Selecionar LHs até atingir o CAP
        // ⚠️ VALIDAÇÃO: Só incluir LHs que chegam ANTES do horário de corte (minutosCorte >= 0)
        let backlogsPisoSelecionados = 0;
        
        for (const lhInfo of lhsComInfo) {
            // 🔒 VALIDAÇÃO DE STATUS: Ignorar LHs bloqueadas (status P3 - fora do prazo)
            if (lhInfo.statusLH?.codigo === 'P3') {
                lhsBloqueadas++;
                console.log(`🔒 LH ${lhInfo.lhTrip} BLOQUEADA: status P3 (em trânsito - fora do prazo)`);
                continue; // Pular esta LH
            }
            
            // ✅ VALIDAÇÃO DE CORTE: Só incluir se chegar A TEMPO
            // minutosCorte >= 0 significa que chegará antes do corte
            // minutosCorte < 0 significa que já passou do corte (atrasada)
            if (lhInfo.minutosCorte !== null && lhInfo.minutosCorte < 0) {
                lhsForaDoCorte++;
                console.log(`⛔ LH ${lhInfo.lhTrip} EXCLUÍDA: não chegará a tempo (${lhInfo.minutosCorte} min)`);
                continue; // Pular esta LH
            }
            
            // Verificar se cabe no CAP
            if (totalSelecionado + lhInfo.qtdPedidos <= capCiclo) {
                // LH cabe perfeitamente no CAP
                lhsSelecionadasPlan.add(lhInfo.lhTrip);
                lhsSugeridas.push(lhInfo);
                totalSelecionado += lhInfo.qtdPedidos;
                
                if (lhInfo.isBacklogPiso) backlogsPisoSelecionados++;
                
                console.log(`✅ LH ${lhInfo.lhTrip} INCLUÍDA: ${lhInfo.qtdPedidos} pedidos (corte em ${lhInfo.minutosCorte || '?'} min)`);
            } else if (totalSelecionado < capCiclo) {
                // 🎯 FIFO: Próxima LH que não cabe no CAP
                // Não incluir LH completa, mas marcar para sugestão de TOs
                const faltam = capCiclo - totalSelecionado;
                
                console.log(`💡 LH ${lhInfo.lhTrip} NÃO CABE (${lhInfo.qtdPedidos} pedidos, faltam ${faltam}). Sugerir TOs para completar CAP.`);
                
                // Marcar como candidata para TOs (não adicionar à seleção ainda)
                lhInfo.candidataParaTOs = true;
                lhInfo.qtdNecessaria = faltam;
                
                // Armazenar para sugestão posterior
                window.lhCandidataParaTOs = lhInfo;
                
                // 💚 MOSTRAR BANNER automaticamente após renderização
                setTimeout(() => {
                    console.log(`💚 Banner: Chamando mostrarBannerLHCandidata para ${lhInfo.lhTrip}`);
                    mostrarBannerLHCandidata(lhInfo, faltam, capCiclo);
                }, 800);
                
                break; // Parar após primeira LH que não cabe (prioridade FIFO)
            } else if (totalSelecionado >= capCiclo) {
                // CAP já atingido
                break;
            }
        }
        
        // 💬 Armazenar LHs com estouro para uso posterior (modal de TOs)
        if (lhsNoPisoComEstouro.length > 0) {
            window.lhsComEstouroPiso = lhsNoPisoComEstouro;
            console.log(`🟡 ${lhsNoPisoComEstouro.length} LH(s) No Piso com estouro tolerado - sugestão de TOs disponível`);
        }
        
        console.log(`✅ LHs selecionadas: ${lhsSugeridas.length}`);
        console.log(`📦 Backlogs do piso: ${backlogsPisoSelecionados}`);
        console.log(`🔒 LHs bloqueadas (P3): ${lhsBloqueadas}`);
        console.log(`⛔ LHs excluídas (fora do corte): ${lhsForaDoCorte}`);
        console.log(`📊 Total final: ${totalSelecionado.toLocaleString('pt-BR')}`);
    }
    
    // ===== MOSTRAR RESUMO =====
    const percentualCap = ((totalSelecionado / capCiclo) * 100).toFixed(1);
    const lhsDentroLimite = lhsSugeridas.filter(lh => lh.dentroLimite).length;
    const backlogsPiso = lhsSugeridas.filter(lh => lh.isBacklogPiso).length;
    
    // Mostrar info de sugestão
    mostrarInfoSugestao({
        ciclo: cicloParaUsar,
        cap: capCiclo,
        totalSelecionado,
        percentualCap,
        backlog: backlogSugerido.length,
        lhs: lhsSugeridas.length,
        lhsDentroLimite,
        backlogsPiso,
        lhsBloqueadas: lhsBloqueadas || 0
    });
    
    // Atualizar interface
    renderizarTabelaPlanejamento();
    renderizarBacklog();
    // atualizarInfoSelecao() - Função removida
    
    // Ir para aba de backlog se tiver backlog selecionado
    if (backlogSugerido.length > 0) {
        // Marcar backlog como confirmado
        backlogConfirmado = true;
    }
    
    // 🟡 ABRIR MODAL DE TOs AUTOMATICAMENTE se houver LH No Piso com estouro
    if (lhsNoPisoComEstouro.length > 0) {
        const lhComEstouro = lhsNoPisoComEstouro[0];
        console.log(`🟡 Abrindo modal de TOs automaticamente para LH ${lhComEstouro.lhTrip} (estouro: +${lhComEstouro.estouroQtd} pedidos)`);
        
        // Aguardar renderização da tabela antes de abrir modal
        setTimeout(() => {
            abrirModalTOs(lhComEstouro.lhTrip);
            
            // Mostrar mensagem informativa
            setTimeout(() => {
                alert('🟡 LH No Piso Detectada!\n\n' +
                      `LH: ${lhComEstouro.lhTrip}\n` +
                      `Pedidos: ${lhComEstouro.qtdPedidos}\n` +
                      `Estouro: +${lhComEstouro.estouroQtd} pedidos acima do CAP\n\n` +
                      '💡 Sugestão: Ajuste por TOs para otimizar o CAP!\n' +
                      'As TOs já foram pré-selecionadas via FIFO.');
            }, 500);
        }, 300);
    }
    
    // 💡 SUGESTÃO DE COMPLEMENTO: Identificar próxima LH FIFO se não completou o CAP
    const faltamParaCompletar = capCiclo - totalSelecionado;
    let lhComplementoSugerida = null;
    
    console.log(`🔍 DEBUG - Verificação de complemento:`);
    console.log(`   faltamParaCompletar: ${faltamParaCompletar}`);
    console.log(`   totalSelecionado: ${totalSelecionado}`);
    console.log(`   Condição (faltam > 0 && total > 0): ${faltamParaCompletar > 0 && totalSelecionado > 0}`);
    
    if (faltamParaCompletar > 0 && totalSelecionado > 0) {
        console.log(`🔍 DEBUG - Procurando LH candidata...`);
        console.log(`   Total de LHs disponíveis: ${lhsComInfo.length}`);
        console.log(`   LHs já selecionadas: ${lhsSelecionadasPlan.size}`);
        
        // Buscar próxima LH FIFO que não foi selecionada
        let lhsAnalisadas = 0;
        let lhsIgnoradasSelecionadas = 0;
        let lhsIgnoradasBloqueadas = 0;
        let lhsIgnoradasForaCorte = 0;
        
        for (const lhInfo of lhsComInfo) {
            lhsAnalisadas++;
            // Ignorar LHs já selecionadas
            if (lhsSelecionadasPlan.has(lhInfo.lhTrip)) {
                lhsIgnoradasSelecionadas++;
                continue;
            }
            
            // Ignorar LHs bloqueadas (P3)
            if (lhInfo.statusLH?.codigo === 'P3') {
                lhsIgnoradasBloqueadas++;
                continue;
            }
            
            // Ignorar LHs fora do corte
            if (lhInfo.minutosCorte !== null && lhInfo.minutosCorte < 0) {
                lhsIgnoradasForaCorte++;
                continue;
            }
            
            // Esta é a próxima LH FIFO disponível!
            lhComplementoSugerida = lhInfo;
            console.log(`💡 LH de complemento sugerida: ${lhInfo.lhTrip} (${lhInfo.qtdPedidos} pedidos, faltam ${faltamParaCompletar})`);
            break;
        }
        
        console.log(`🔍 DEBUG - Resultado da busca:`);
        console.log(`   LHs analisadas: ${lhsAnalisadas}`);
        console.log(`   LHs ignoradas (já selecionadas): ${lhsIgnoradasSelecionadas}`);
        console.log(`   LHs ignoradas (bloqueadas P3): ${lhsIgnoradasBloqueadas}`);
        console.log(`   LHs ignoradas (fora do corte): ${lhsIgnoradasForaCorte}`);
        console.log(`   LH candidata encontrada: ${lhComplementoSugerida ? lhComplementoSugerida.lhTrip : 'NENHUMA'}`);
        
        // Armazenar para uso posterior (destaque visual + modal)
        if (lhComplementoSugerida) {
            window.lhComplementoSugerida = {
                lhTrip: lhComplementoSugerida.lhTrip,
                qtdPedidos: lhComplementoSugerida.qtdPedidos,
                faltam: faltamParaCompletar,
                statusLH: lhComplementoSugerida.statusLH
            };
        }
    }
    
    console.log('═'.repeat(50));
    console.log('🎯 SUGESTÃO CONCLUÍDA');
    console.log(`   Ciclo: ${cicloParaUsar}`);
    console.log(`   CAP: ${capCiclo.toLocaleString('pt-BR')}`);
    console.log(`   Total: ${totalSelecionado.toLocaleString('pt-BR')} (${percentualCap}%)`);
    console.log(`   Backlog (sem LH): ${backlogSugerido.length}`);
    console.log(`   Backlog (piso): ${backlogsPiso}`);
    console.log(`   LHs: ${lhsSugeridas.length} (${lhsDentroLimite} dentro do limite)`);
    if (lhComplementoSugerida) {
        console.log(`   💡 Complemento sugerido: ${lhComplementoSugerida.lhTrip} (${lhComplementoSugerida.qtdPedidos} pedidos)`);
    }
    console.log('═'.repeat(50));
    
    // 💬 MOSTRAR BANNER VERDE se houver LH candidata para complemento
    if (window.lhComplementoSugerida) {
        const lhCandidata = window.lhComplementoSugerida;
        const faltam = lhCandidata.faltam;
        
        console.log(`💚 LH Candidata detectada: ${lhCandidata.lhTrip} (${lhCandidata.qtdPedidos} pedidos)`);
        console.log(`💚 Faltam: ${faltam} pedidos para completar CAP`);
        
        // Aguardar renderização da tabela antes de mostrar banner
        setTimeout(() => {
            mostrarBannerLHCandidata(lhCandidata, faltam, capCiclo);
        }, 500);
    }
}

// Extrair data para ordenação FIFO
function extrairDataParaOrdenacao(previsaoFinal) {
    if (!previsaoFinal) return null;
    
    try {
        const str = String(previsaoFinal).trim();
        
        if (str.includes('/')) {
            // DD/MM/YYYY HH:MM:SS ou DD/MM/YYYY HH:MM
            const partesPrincipais = str.split(' ');
            const data = partesPrincipais[0];
            const hora = partesPrincipais[1] || '00:00:00';
            
            const partesData = data.split('/');
            if (partesData.length === 3) {
                const dia = parseInt(partesData[0]);
                const mes = parseInt(partesData[1]) - 1;
                const ano = parseInt(partesData[2]);
                
                // Extrair hora, minuto, segundo
                const partesHora = hora.split(':');
                const hh = parseInt(partesHora[0]) || 0;
                const mm = parseInt(partesHora[1]) || 0;
                const ss = parseInt(partesHora[2]) || 0;
                
                return new Date(ano, mes, dia, hh, mm, ss);
            }
        } else if (str.includes('-')) {
            // YYYY-MM-DD HH:MM:SS ou YYYY-MM-DD HH:MM
            const partesPrincipais = str.split(' ');
            const data = partesPrincipais[0];
            const hora = partesPrincipais[1] || '00:00:00';
            
            const partesData = data.split('-');
            if (partesData.length === 3) {
                const ano = parseInt(partesData[0]);
                const mes = parseInt(partesData[1]) - 1;
                const dia = parseInt(partesData[2]);
                
                // Extrair hora, minuto, segundo
                const partesHora = hora.split(':');
                const hh = parseInt(partesHora[0]) || 0;
                const mm = parseInt(partesHora[1]) || 0;
                const ss = parseInt(partesHora[2]) || 0;
                
                return new Date(ano, mes, dia, hh, mm, ss);
            }
        }
    } catch (e) {}
    
    return null;
}

// Função auxiliar para obter ID do pedido (mesmo padrão do backlog)
function getShipmentIdFromPedido(pedido, fallbackIndex) {
    // Verificar várias variações de nome de coluna
    const possiveisNomes = ['SHIPMENT ID', 'Shipment ID', 'shipment_id', 'SHIPMENT_ID', 'shipmentid', 'ShipmentId', 'Shipment Id', 'ID', 'id'];
    
    // Debug: verificar se é pedido de LH lixo
    const lhTripDebug = pedido['LH Trip'] || pedido['LH_TRIP'] || pedido['lh_trip'] || '';
    const isLixo = lhTripDebug && (lhTripDebug === 'LT0Q2F01YEIJ1' || lhTripDebug === 'LT1Q2I01ZC4C1' || lhTripDebug === 'LT0Q2B01YKHZ1' || lhTripDebug === 'LT1Q2901YTE01');
    
    if (isLixo) {
        console.log(`🔍 [DEBUG LIXO] Buscando Shipment ID para LH ${lhTripDebug}`);
        console.log(`   Colunas disponíveis:`, Object.keys(pedido).slice(0, 10));
    }
    
    for (const nome of possiveisNomes) {
        if (pedido[nome]) {
            const valor = String(pedido[nome]).trim();
            if (valor) {
                if (isLixo) {
                    console.log(`   ✅ Encontrado em coluna '${nome}': ${valor}`);
                }
                return valor;
            }
        }
    }
    
    // ✅ FALLBACK MELHORADO: usar chave única baseada em outras colunas
    // Para pedidos sem Shipment ID (como os das LHs lixo), criar ID baseado em:
    // Zipcode + City + LH Trip + Destination Address
    const zipcode = pedido['Zipcode'] || pedido['ZIPCODE'] || pedido['zipcode'] || '';
    const city = pedido['City'] || pedido['CITY'] || pedido['city'] || '';
    const lhTrip = pedido['LH Trip'] || pedido['LH_TRIP'] || pedido['lh_trip'] || '';
    const destAddress = pedido['Destination Address'] || pedido['DESTINATION ADDRESS'] || pedido['destination_address'] || '';
    
    // Se temos informações suficientes, criar chave única
    if (zipcode || city || lhTrip) {
        const chave = `${zipcode}_${city}_${lhTrip}_${destAddress}`.replace(/[\s]+/g, '_');
        const id = `pedido_${chave}`;
        // Log apenas para LHs lixo sistêmico
        if (lhTrip && (lhTrip === 'LT0Q2F01YEIJ1' || lhTrip === 'LT1Q2I01ZC4C1' || lhTrip === 'LT0Q2B01YKHZ1' || lhTrip === 'LT1Q2901YTE01')) {
            console.log(`🔑 [DEBUG LIXO] ID gerado para pedido sem Shipment ID: ${id} (LH: ${lhTrip})`);
        }
        return id;
    }
    
    // Último fallback: usar índice fixo (sem Date.now para ser consistente)
    return `backlog_${fallbackIndex}`;
}

// Obter capacidade do ciclo selecionado
function obterCapacidadeCiclo(ciclo) {
    console.log("🔍 obterCapacidadeCiclo chamada para:", ciclo);
    console.log("📦 capsManual atual:", capsManual);
    const capManual = obtemCapManual(ciclo);
    console.log("🔎 obtemCapManual(" + ciclo + ") retornou:", capManual);
    if (capManual !== null) {
        console.log("Usando CAP Manual:", ciclo, capManual);
        return capManual;
    }
    console.log("⚙️ CAP Manual não encontrado, buscando CAP automático...");
    if (!ciclo || ciclo === 'Todos') return 0;
    
    const stationSelecionada = stationAtualNome || '';
    const stationNormalizada = stationSelecionada.toLowerCase().replace(/lm\s*hub[_\s]*/gi, '').replace(/[_\s]+/g, '');
    
    // Buscar capacidade para esta station e ciclo
    const capacidadeStation = dadosOutbound.filter(item => {
        const sortCodeName = item['Sort Code Name'] || item['sort_code_name'] || '';
        const itemNorm = sortCodeName.toLowerCase().replace(/lm\s*hub[_\s]*/gi, '').replace(/[_\s]+/g, '');
        return itemNorm.includes(stationNormalizada) || stationNormalizada.includes(itemNorm);
    });
    
    // Encontrar registro do ciclo
    const registroCiclo = capacidadeStation.find(cap => {
        const tipoCap = cap['Type Outbound'] || cap['type_outbound'] || '';
        return tipoCap.toUpperCase() === ciclo.toUpperCase();
    });
    
    if (!registroCiclo) return 0;
    
    // USAR DATA DO CICLO SELECIONADA (não hoje!)
    const dataCiclo = getDataCicloSelecionada();
    const diaHoje = String(dataCiclo.getDate()).padStart(2, '0');
    const diaSemZero = String(dataCiclo.getDate());
    const mesHoje = String(dataCiclo.getMonth() + 1).padStart(2, '0');
    const mesSemZero = String(dataCiclo.getMonth() + 1);
    const anoHoje = dataCiclo.getFullYear();
    const anoCurto = String(anoHoje).slice(2);
    
    const formatosData = [
        `${diaHoje}/${mesHoje}/${anoHoje}`,       // 11/01/2026
        `${diaSemZero}/${mesSemZero}/${anoHoje}`, // 11/1/2026
        `${diaHoje}/${mesHoje}/${anoCurto}`,      // 11/01/26
        `${anoHoje}-${mesHoje}-${diaHoje}`,       // 2026-01-11
    ];
    
    console.log(`📊 obterCapacidadeCiclo: buscando CAP para ${ciclo} em ${formatosData[0]}`);
    
    for (const formato of formatosData) {
        if (registroCiclo[formato] !== undefined && registroCiclo[formato] !== '') {
            let valor = registroCiclo[formato];
            if (typeof valor === 'string') {
                valor = valor.replace(/\./g, '').replace(',', '.');
            }
            const cap = parseFloat(valor) || 0;
            console.log(`✅ CAP encontrado para ${ciclo} em ${formato}: ${cap}`);
            return cap;
        }
    }
    
    console.log(`⚠️ CAP não encontrado para ${ciclo} na data ${formatosData[0]}`);
    return 0;
}

// Mostrar info da sugestão na interface
function mostrarInfoSugestao(info) {
    // Remover info anterior se existir
    const infoAnterior = document.querySelector('.sugestao-info');
    if (infoAnterior) infoAnterior.remove();
    
    // Verificar se está usando CAP manual
    const usandoCapManual = obtemCapManual(info.ciclo) !== null;
    const badgeCapManual = usandoCapManual ? '<span class="badge-cap-manual-sugestao">CAP MANUAL</span>' : '';
    
    // Montar detalhe com backlogs do piso se houver
    let detalhe = info.backlog + ' do backlog + ' + info.lhs + ' LHs';
    if (info.backlogsPiso && info.backlogsPiso > 0) {
        detalhe += ' (' + info.backlogsPiso + ' backlogs piso, ' + info.lhsDentroLimite + ' dentro do limite)';
    } else {
        detalhe += ' (' + info.lhsDentroLimite + ' dentro do limite de 45 min)';
    }
    
    // Adicionar informação de LHs bloqueadas se houver
    if (info.lhsBloqueadas && info.lhsBloqueadas > 0) {
        detalhe += ' | 🔒 ' + info.lhsBloqueadas + ' LH' + (info.lhsBloqueadas > 1 ? 's' : '') + ' bloqueada' + (info.lhsBloqueadas > 1 ? 's' : '') + ' (fora do prazo)';
    }
    
    // Adicionar informação de LHs No Piso com estouro se houver
    if (window.lhsComEstouroPiso && window.lhsComEstouroPiso.length > 0) {
        const qtdEstouro = window.lhsComEstouroPiso.reduce((sum, lh) => sum + lh.estouroQtd, 0);
        detalhe += ' | 🟡 ' + window.lhsComEstouroPiso.length + ' LH' + (window.lhsComEstouroPiso.length > 1 ? 's' : '') + ' no piso (+' + qtdEstouro + ' estouro)';
    }
    
    // Criar novo elemento de info
    const infoDiv = document.createElement('div');
    infoDiv.className = 'sugestao-info';
    infoDiv.innerHTML = '<div class="sugestao-info-texto">' +
            '<span class="sugestao-info-titulo">🎯 Sugestão para ' + info.ciclo + ' ' + badgeCapManual + '</span>' +
            '<span class="sugestao-info-detalhe">' + detalhe + '</span>' +
        '</div>' +
        '<div style="display: flex; align-items: center; gap: 15px;">' +
            '<div style="text-align: right;">' +
                '<div class="sugestao-info-cap">' + info.totalSelecionado.toLocaleString('pt-BR') + ' / ' + info.cap.toLocaleString('pt-BR') + '</div>' +
                '<div style="font-size: 13px; color: #155724;">' + info.percentualCap + '% do CAP</div>' +
            '</div>' +
            '<button class="btn btn-confirmar-sugestao" onclick="confirmarSugestaoEGerar()">' +
                '✅ Confirmar e Gerar' +
            '</button>' +
        '</div>';
    
    // Inserir antes da barra de seleção
    const selecaoBar = document.querySelector('.planejamento-selecao-bar');
    if (selecaoBar) {
        selecaoBar.parentNode.insertBefore(infoDiv, selecaoBar);
    }
}

// Mostrar banner de LH candidata para complemento
function mostrarBannerLHCandidata(lhCandidata, faltam, capTotal) {
    // Remover banner anterior se existir
    const bannerAnterior = document.querySelector('.banner-lh-candidata');
    if (bannerAnterior) bannerAnterior.remove();
    
    // Criar banner
    const banner = document.createElement('div');
    banner.className = 'banner-lh-candidata';
    banner.innerHTML = `
        <div class="banner-lh-candidata-header">
            <div class="banner-lh-candidata-icon">💚</div>
            <h3 class="banner-lh-candidata-title">CAP pode ser completado!</h3>
        </div>
        <div class="banner-lh-candidata-body">
            <p style="margin: 0 0 8px 0;">Sobram <strong>${faltam.toLocaleString('pt-BR')} pedidos</strong> para completar o CAP.</p>
            <p style="margin: 0 0 8px 0;">Próxima LH FIFO disponível:</p>
            <div class="banner-lh-candidata-lh">${lhCandidata.lhTrip} (${lhCandidata.qtdPedidos.toLocaleString('pt-BR')} pedidos)</div>
            <p style="margin: 8px 0 0 0; font-size: 14px; opacity: 0.9;">💡 Use TOs para completar exatamente o CAP!</p>
        </div>
        <div class="banner-lh-candidata-footer">
            <button class="banner-lh-candidata-btn banner-lh-candidata-btn-secondary" onclick="fecharBannerLHCandidata()">Ignorar</button>
            <button class="banner-lh-candidata-btn banner-lh-candidata-btn-primary" onclick="abrirTOsLHCandidata('${lhCandidata.lhTrip}')">📦 Abrir TOs</button>
        </div>
    `;
    
    // Adicionar ao body
    document.body.appendChild(banner);
    
    console.log('💚 Banner de LH candidata exibido');
}

// Fechar banner de LH candidata
function fecharBannerLHCandidata() {
    const banner = document.querySelector('.banner-lh-candidata');
    if (banner) {
        banner.classList.add('fade-out');
        setTimeout(() => banner.remove(), 500);
    }
}

// Abrir modal de TOs da LH candidata
function abrirTOsLHCandidata(lhTrip) {
    fecharBannerLHCandidata();
    abrirModalTOs(lhTrip);
}

// Atualizar card de sugestão com totais atualizados (usado quando TOs são selecionadas)
function atualizarCardSugestao() {
    const infoDiv = document.querySelector('.sugestao-info');
    if (!infoDiv) return; // Não tem card para atualizar
    
    // Calcular novo total
    const totalSelecionado = calcularTotalSelecionado();
    const capCiclo = obterCapacidadeCicloAtual();
    const percentualCap = capCiclo > 0 ? ((totalSelecionado / capCiclo) * 100).toFixed(1) : 0;
    
    // Atualizar apenas os números do card
    const capDiv = infoDiv.querySelector('.sugestao-info-cap');
    const percentDiv = infoDiv.querySelector('div[style*="font-size: 13px"]');
    
    if (capDiv) {
        capDiv.textContent = totalSelecionado.toLocaleString('pt-BR') + ' / ' + capCiclo.toLocaleString('pt-BR');
    }
    
    if (percentDiv) {
        percentDiv.textContent = percentualCap + '% do CAP';
    }
}

// Confirmar sugestão e gerar planejamento
async function confirmarSugestaoEGerar() {
    // Verificar se tem seleção
    const totalLHs = lhsSelecionadasPlan.size;
    const totalBacklog = pedidosBacklogSelecionados.size;
    
    if (totalLHs === 0 && totalBacklog === 0) {
        alert('⚠️ Nenhuma LH ou pedido do backlog selecionado!');
        return;
    }
    
    // ✅ USAR calcularTotalSelecionado() que já considera TOs parciais
    console.log('🔍 DEBUG - Calculando total para modal de confirmação:');
    console.log('  - LHs selecionadas:', totalLHs);
    console.log('  - Backlog:', totalBacklog);
    console.log('  - TOs selecionadas:', Object.keys(tosSelecionadasPorLH));
    
    const totalPedidos = calcularTotalSelecionado();
    console.log('  - Total calculado:', totalPedidos);
    
    // Confirmar com o usuário
    const ciclo = cicloSelecionado || 'Geral';
    const confirmacao = confirm(
        `📋 CONFIRMAR PLANEJAMENTO - ${ciclo}\n\n` +
        `📦 ${totalBacklog} pedidos do backlog\n` +
        `🚚 ${totalLHs} LHs selecionadas\n` +
        `📊 Total: ${totalPedidos.toLocaleString('pt-BR')} pedidos\n\n` +
        `Deseja gerar o arquivo de planejamento?`
    );
    
    if (!confirmacao) return;
    
    // Marcar backlog como confirmado
    backlogConfirmado = true;
    
    // Gerar o planejamento
    await gerarArquivoPlanejamento();
    
    // Remover info de sugestão após gerar
    const infoSugestao = document.querySelector('.sugestao-info');
    if (infoSugestao) infoSugestao.remove();
}

// Expor função globalmente
window.confirmarSugestaoEGerar = confirmarSugestaoEGerar;

// ======================= LOADING COM PROGRESSO =======================
function mostrarLoading(titulo, mensagem, comProgresso = false) {
    document.getElementById('loadingTitle').textContent = titulo;
    document.getElementById('loadingMessage').textContent = mensagem;
    document.getElementById('loadingOverlay').classList.add('active');

    const progressContainer = document.getElementById('progressContainer');
    if (progressContainer) {
        if (comProgresso) {
            progressContainer.classList.add('active');
            resetarProgresso();
        } else {
            progressContainer.classList.remove('active');
        }
    }
}

function esconderLoading() {
    document.getElementById('loadingOverlay').classList.remove('active');
    const progressContainer = document.getElementById('progressContainer');
    if (progressContainer) {
        progressContainer.classList.remove('active');
    }
}

function atualizarProgresso(etapa, mensagem) {
    // etapa: 1 a 5
    const porcentagens = {
        1: 10,
        2: 30,
        3: 50,
        4: 70,
        5: 90,
        6: 100
    };
    const porcentagem = porcentagens[etapa] || 0;

    document.getElementById('progressBar').style.width = `${porcentagem}%`;
    document.getElementById('progressText').textContent = mensagem || `Etapa ${etapa} de 5`;
    document.getElementById('loadingMessage').textContent = mensagem;

    // Atualizar steps visuais
    for (let i = 1; i <= 5; i++) {
        const stepEl = document.getElementById(`step${i}`);
        if (stepEl) {
            stepEl.classList.remove('active', 'completed');
            if (i < etapa) {
                stepEl.classList.add('completed');
            } else if (i === etapa) {
                stepEl.classList.add('active');
            }
        }
    }
}

function resetarProgresso() {
    document.getElementById('progressBar').style.width = '0%';
    document.getElementById('progressText').textContent = 'Iniciando...';
    for (let i = 1; i <= 5; i++) {
        const stepEl = document.getElementById(`step${i}`);
        if (stepEl) {
            stepEl.classList.remove('active', 'completed');
        }
    }
}

// ======================= TORNAR FUNÇÕES GLOBAIS =======================
window.removerStation = removerStation;
window.toggleColuna = toggleColuna;

// ======================= CONFIGURAÇÃO DO NAVEGADOR =======================
function carregarConfigNavegador() {
    try {
        const config = localStorage.getItem('shopee_config_navegador');
        if (config) {
            const parsed = JSON.parse(config);
            document.getElementById('configHeadless').checked = parsed.headless !== false; // default true
            document.getElementById('configAutoDetect').checked = parsed.autoDetect !== false; // default true
        }
    } catch (e) {
        console.log('Usando configurações padrão do navegador');
    }
}

function salvarConfigNavegador() {
    try {
        const config = {
            headless: document.getElementById('configHeadless').checked,
            autoDetect: document.getElementById('configAutoDetect').checked
        };
        localStorage.setItem('shopee_config_navegador', JSON.stringify(config));
        console.log('✅ Configurações do navegador salvas:', config);
    } catch (e) {
        console.error('Erro ao salvar configurações:', e);
    }
}

function getConfigNavegador() {
    try {
        const config = localStorage.getItem('shopee_config_navegador');
        if (config) {
            return JSON.parse(config);
        }
    } catch (e) {}
    return {
        headless: true,
        autoDetect: true
    };
}

async function limparSessaoLogin() {
    if (!confirm('Isso irá limpar a sessão de login.\n\nVocê precisará fazer login novamente na próxima execução.\n\nContinuar?')) {
        return;
    }

    mostrarLoading('Limpando sessão...', 'Removendo dados de login');

    try {
        const resultado = await ipcRenderer.invoke('limpar-sessao');

        esconderLoading();

        if (resultado.success) {
            alert('✅ Sessão limpa com sucesso!\n\nNa próxima execução, você precisará fazer login novamente.');
            verificarStatusSessao();
        } else {
            alert(`❌ Erro ao limpar sessão: ${resultado.error}`);
        }
    } catch (error) {
        esconderLoading();
        alert(`❌ Erro: ${error.message}`);
    }
}

async function limparTodosOsDados() {
    if (!confirm('⚠️ ATENÇÃO!\n\nIsso irá limpar:\n- Sessão de login\n- Cookies\n- Todas as configurações\n- Stations cadastradas\n\nEssa ação não pode ser desfeita!\n\nContinuar?')) {
        return;
    }

    // Confirmar novamente
    if (!confirm('Tem certeza ABSOLUTA? Todos os dados serão perdidos!')) {
        return;
    }

    mostrarLoading('Limpando tudo...', 'Removendo todos os dados');

    try {
        // Limpar localStorage
        localStorage.clear();

        // Limpar sessão no backend
        const resultado = await ipcRenderer.invoke('limpar-tudo');

        esconderLoading();

        if (resultado.success) {
            alert('✅ Todos os dados foram limpos!\n\nO aplicativo será reiniciado.');
            // Recarregar a página para aplicar as mudanças
            window.location.reload();
        } else {
            alert(`❌ Erro ao limpar dados: ${resultado.error}`);
        }
    } catch (error) {
        esconderLoading();
        alert(`❌ Erro: ${error.message}`);
    }
}

async function verificarStatusSessao() {
    try {
        const resultado = await ipcRenderer.invoke('verificar-sessao');

        const statusEl = document.getElementById('sessionStatus');
        const statusTextEl = document.getElementById('sessionStatusText');
        const indicatorEl = statusEl ?.querySelector('.session-indicator');

        if (!statusEl || !statusTextEl) return;

        if (resultado.temSessao) {
            statusEl.classList.remove('inactive');
            statusEl.classList.add('active');
            if (indicatorEl) {
                indicatorEl.classList.remove('inactive');
                indicatorEl.classList.add('active');
            }
            statusTextEl.textContent = '✅ Sessão ativa';
        } else {
            statusEl.classList.remove('active');
            statusEl.classList.add('inactive');
            if (indicatorEl) {
                indicatorEl.classList.remove('active');
                indicatorEl.classList.add('inactive');
            }
            statusTextEl.textContent = '⚠️ Não autenticado';
        }
    } catch (error) {
        console.log('Erro ao verificar sessão:', error);
    }
}

// ======================= CONFIGURAÇÃO DE COLUNAS =======================
function carregarConfigColunas() {
    try {
        const config = localStorage.getItem('shopee_colunas_visiveis');
        if (config) {
            colunasVisiveis = JSON.parse(config);
            console.log(`✅ Configuração de colunas carregada: ${colunasVisiveis.length} colunas`);
        }
    } catch (e) {
        console.log('Nenhuma configuração de colunas salva');
    }
}

function salvarConfigColunas() {
    try {
        // Capturar colunas selecionadas
        const container = document.getElementById('colunasContainer') || document.getElementById('colunasRelatorioGrid');
        if (!container) {
            alert('❌ Carregue um arquivo primeiro');
            return;
        }

        const checkboxes = container.querySelectorAll('input[type="checkbox"]:checked');
        colunasVisiveis = Array.from(checkboxes).map(cb => cb.dataset.coluna).filter(Boolean);

        // Salvar no localStorage
        localStorage.setItem('shopee_colunas_visiveis', JSON.stringify(colunasVisiveis));

        alert(`✅ Configuração salva!\n${colunasVisiveis.length} colunas selecionadas`);

        // Atualizar tabela se tiver LH selecionada
        if (lhTripAtual) {
            const pedidos = lhTripAtual === 'TODOS' ? dadosAtuais : (lhTrips[lhTripAtual] || []);
            renderizarTabela(pedidos);
        }

    } catch (e) {
        alert('❌ Erro ao salvar configuração');
        console.error(e);
    }
}

function resetarConfigColunas() {
    if (confirm('Resetar configuração e mostrar todas as colunas?')) {
        colunasVisiveis = [...todasColunas];
        localStorage.removeItem('shopee_colunas_visiveis');
        atualizarListaColunas();

        // Atualizar tabela
        if (lhTripAtual) {
            const pedidos = lhTripAtual === 'TODOS' ? dadosAtuais : (lhTrips[lhTripAtual] || []);
            renderizarTabela(pedidos);
        }

        alert('✅ Configuração resetada! Todas as colunas serão exibidas.');
    }
}

function atualizarListaColunas() {
    const container = document.getElementById('colunasContainer') || document.getElementById('colunasRelatorioGrid');
    const countEl = document.getElementById('totalColunasDisponiveis');

    if (!container) return;

    if (todasColunas.length === 0) {
        container.innerHTML = `
            <p style="color:#999;text-align:center;grid-column:1/-1;padding:20px;">
                Carregue um arquivo para ver as colunas disponíveis
            </p>
        `;
        if (countEl) countEl.textContent = '0';
        return;
    }

    if (countEl) countEl.textContent = todasColunas.length;

    container.innerHTML = todasColunas.map((col, index) => {
        const isChecked = colunasVisiveis.includes(col);
        return `
            <div class="coluna-item">
                <input type="checkbox" id="col_${index}" data-coluna="${col}" ${isChecked ? 'checked' : ''} onchange="toggleColuna('${col}')">
                <label for="col_${index}">${col}</label>
            </div>
        `;
    }).join('');
}

function toggleColuna(coluna) {
    const item = document.querySelector(`input[data-coluna="${coluna}"]`);
    if (item) {
        item.checked = !item.checked;
    }
    atualizarContadorColunas();
    atualizarPreviewConfig();
}

function selecionarTodasColunas() {
    const container = document.getElementById('colunasContainer') || document.getElementById('colunasRelatorioGrid');
    if (!container) return;

    container.querySelectorAll('input[type="checkbox"]').forEach(cb => {
        cb.checked = true;
    });
    atualizarContadorColunas();
    atualizarPreviewConfig();
}

function limparTodasColunas() {
    const container = document.getElementById('colunasContainer') || document.getElementById('colunasRelatorioGrid');
    if (!container) return;

    container.querySelectorAll('input[type="checkbox"]').forEach(cb => {
        cb.checked = false;
    });
    atualizarContadorColunas();
    atualizarPreviewConfig();
}

function atualizarContadorColunas() {
    const container = document.getElementById('colunasContainer') || document.getElementById('colunasRelatorioGrid');
    const countEl = document.getElementById('colunasSelecionadasCount');

    if (!container || !countEl) return;

    const total = container.querySelectorAll('input[type="checkbox"]:checked').length;
    countEl.textContent = total;
}

function atualizarPreviewConfig() {
    const container = document.getElementById('previewColunasContainer');
    if (!container) return;

    const colunasContainer = document.getElementById('colunasContainer') || document.getElementById('colunasRelatorioGrid');
    if (!colunasContainer) return;

    const checkboxes = colunasContainer.querySelectorAll('input[type="checkbox"]:checked');
    const colunasSelecionadas = Array.from(checkboxes).map(cb => cb.dataset.coluna);

    if (colunasSelecionadas.length === 0) {
        container.innerHTML = `
            <div class="preview-empty" style="padding:40px;">
                <div class="preview-empty-icon">👁️</div>
                <h3>Nenhuma coluna selecionada</h3>
                <p>Selecione as colunas que deseja exibir</p>
            </div>
        `;
        return;
    }

    // Mostrar preview da tabela com as colunas selecionadas
    let html = '<table style="width:100%; font-size:12px;"><thead><tr>';
    colunasSelecionadas.forEach(col => {
        html += `<th style="padding:8px; background:#f8f9fa;">${col}</th>`;
    });
    html += '</tr></thead><tbody>';

    // Mostrar 3 linhas de exemplo
    const exemplos = dadosAtuais.slice(0, 3);
    exemplos.forEach(row => {
        html += '<tr>';
        colunasSelecionadas.forEach(col => {
            const valor = row[col];
            const valorExibir = valor !== null && valor !== undefined ? valor : '-';
            // Truncar valores muito longos
            const valorTruncado = String(valorExibir).length > 30 ?
                String(valorExibir).substring(0, 30) + '...' :
                valorExibir;
            html += `<td style="padding:8px; border-bottom:1px solid #eee;">${valorTruncado}</td>`;
        });
        html += '</tr>';
    });

    html += '</tbody></table>';

    if (dadosAtuais.length > 3) {
        html += `<p style="text-align:center; color:#999; padding:10px; font-size:12px;">
            ... e mais ${(dadosAtuais.length - 3).toLocaleString('pt-BR')} registros
        </p>`;
    }

    container.innerHTML = html;
}

// ======================= ABA PLANEJAMENTO HUB =======================

// Dados da planilha (JSON convertido)
let dadosPlanilha = {};
let ultimaAtualizacaoPlanilha = null;

// Inicializar event listeners da aba Planejamento
function initPlanejamentoHub() {
    // Botão atualizar planilha
    const btnAtualizar = document.getElementById('btnAtualizarPlanilha');
    if (btnAtualizar) {
        btnAtualizar.addEventListener('click', atualizarPlanilhaGoogle);
    }

    // Filtros
    const filtroStatus = document.getElementById('filtroPlanejamentoStatus');
    const filtroBusca = document.getElementById('filtroPlanejamentoBusca');

    // Listener para o filtro de Ciclo (novo)
    const filtroCiclo = document.getElementById('filtroCiclo');
    if (filtroCiclo) {
        filtroCiclo.addEventListener('change', (e) => {
            cicloSelecionado = e.target.value || 'Todos';
            renderizarTabelaPlanejamento();
            // Atualizar visual dos cards
            document.querySelectorAll('.ciclo-card').forEach(card => {
                card.classList.remove('ativo');
                if (card.dataset.ciclo === cicloSelecionado) {
                    card.classList.add('ativo');
                }
            });
        });
    }

    if (filtroStatus) {
        filtroStatus.addEventListener('change', renderizarTabelaPlanejamento);
    }
    if (filtroBusca) {
        filtroBusca.addEventListener('input', renderizarTabelaPlanejamento);
    }

    // Carregar dados salvos localmente
    carregarDadosPlanilhaLocal();
    carregarDadosOpsClockLocal();
    carregarDadosOutboundLocal();
}

// Função auxiliar para buscar dados da planilha filtrando por station
function buscarDadosPlanilhaPorStation(lhTrip) {
    // Se não tem station selecionada, buscar pela chave antiga (sem filtro)
    const stationSelecionada = stationAtualNome || 
        document.getElementById('stationSearchInput')?.value || '';
    
    if (!stationSelecionada) {
        // Tentar buscar pela chave antiga (sem destination)
        return dadosPlanilha[lhTrip] || null;
    }
    
    // Normalizar nome da station para comparação
    const stationNormalizada = stationSelecionada.trim();
    
    // Tentar buscar com chave composta: trip_number|destination
    const chaveComposta = `${lhTrip}|${stationNormalizada}`;
    
    if (dadosPlanilha[chaveComposta]) {
        console.log(`✅ [FILTRO] LH ${lhTrip} encontrada para station ${stationNormalizada}`);
        return dadosPlanilha[chaveComposta];
    }
    
    // Se não encontrou com chave composta, buscar manualmente
    // (para compatibilidade com dados antigos ou variações de nome)
    for (const chave in dadosPlanilha) {
        if (chave.startsWith(lhTrip + '|')) {
            const registro = dadosPlanilha[chave];
            const destination = registro.destination || registro.Destination || registro.DESTINATION || '';
            
            // Comparar destination com station selecionada (case-insensitive e normalizado)
            if (destination.trim().toLowerCase() === stationNormalizada.toLowerCase()) {
                console.log(`✅ [FILTRO] LH ${lhTrip} encontrada para station ${stationNormalizada} (busca manual)`);
                return registro;
            }
        }
    }
    
    // Se não encontrou com filtro, tentar buscar pela chave antiga (sem destination)
    const dadosSemFiltro = dadosPlanilha[lhTrip];
    if (dadosSemFiltro) {
        console.log(`⚠️ [FILTRO] LH ${lhTrip} encontrada SEM filtro de station (dados antigos)`);
        return dadosSemFiltro;
    }
    
    console.log(`❌ [FILTRO] LH ${lhTrip} NÃO encontrada para station ${stationNormalizada}`);
    return null;
}

// Carregar dados da planilha salvos localmente
async function carregarDadosPlanilhaLocal() {
    try {
        const resultado = await ipcRenderer.invoke('carregar-planilha-local');
        if (resultado.success && resultado.dados) {
            dadosPlanilha = resultado.dados.dados || {};
            ultimaAtualizacaoPlanilha = resultado.dados.ultimaAtualizacao;

            atualizarInfoPlanilha();
            renderizarTabelaPlanejamento();
        }
    } catch (error) {
        console.log('Nenhum dado local da planilha encontrado');
    }
}

// Carregar dados OpsClock (horários dos ciclos)
async function carregarDadosOpsClockLocal() {
    try {
        const resultado = await ipcRenderer.invoke('carregar-opsclock-local');
        if (resultado.success && resultado.dados) {
            dadosOpsClock = resultado.dados;
            console.log(`⏰ OpsClock carregado: ${dadosOpsClock.length} registros`);
            atualizarInfoCiclos();
        }
    } catch (error) {
        console.log('Nenhum dado OpsClock encontrado');
    }
}

// Carregar dados Outbound (capacidade por ciclo)
async function carregarDadosOutboundLocal() {
    try {
        const resultado = await ipcRenderer.invoke('carregar-outbound-local');
        if (resultado.success && resultado.dados) {
            dadosOutbound = resultado.dados;
            console.log(`📊 Outbound carregado: ${dadosOutbound.length} registros`);
            atualizarInfoCiclos();
        }
    } catch (error) {
        console.log('Nenhum dado Outbound encontrado');
    }
}

// Atualizar dados da planilha Google Sheets (via API)
async function atualizarPlanilhaGoogle() {
    mostrarLoading('Atualizando planilhas...', 'Conectando ao Google Sheets via API...');

    try {
        const resultado = await ipcRenderer.invoke('atualizar-planilha-google');

        esconderLoading();

        if (resultado.success) {
            dadosPlanilha = resultado.dados;
            ultimaAtualizacaoPlanilha = resultado.ultimaAtualizacao;

            // Recarregar dados das novas planilhas
            await carregarDadosOpsClockLocal();
            await carregarDadosOutboundLocal();

            atualizarInfoPlanilha();
            atualizarInfoCiclos();
            renderizarTabelaPlanejamento();

            const msg = `✅ Planilhas atualizadas!\n\n` +
                `📋 ${Object.keys(dadosPlanilha).length} LHs\n` +
                `⏰ ${resultado.opsClock || 0} registros de ciclos\n` +
                `📊 ${resultado.outbound || 0} registros de capacidade`;
            
            alert(msg);
        } else {
            alert(`❌ Erro ao atualizar: ${resultado.error}`);
        }
    } catch (error) {
        esconderLoading();
        alert(`❌ Erro: ${error.message}`);
    }
}

// Atualizar informações da planilha na interface
function atualizarInfoPlanilha() {
    // Última atualização
    const ultimaAtualizacaoEl = document.getElementById('ultimaAtualizacaoPlanilha');
    if (ultimaAtualizacaoEl) {
        if (ultimaAtualizacaoPlanilha) {
            const data = new Date(ultimaAtualizacaoPlanilha);
            ultimaAtualizacaoEl.textContent = data.toLocaleString('pt-BR');
        } else {
            ultimaAtualizacaoEl.textContent = 'Nunca';
        }
    }

    // Total na planilha
    const totalPlanilhaEl = document.getElementById('statTotalPlanilha');
    if (totalPlanilhaEl) {
        totalPlanilhaEl.textContent = Object.keys(dadosPlanilha).length.toLocaleString('pt-BR');
    }

    // Iniciar contador de próxima atualização
    iniciarContadorProximaAtualizacao();
}

// Atualizar informações de ciclos na interface
function atualizarInfoCiclos() {
    // Usar a station do arquivo carregado
    let stationSelecionada = stationAtualNome || 
        document.getElementById('stationSearchInput')?.value || '';
    
    // Garantir que não seja undefined ou string 'undefined'
    if (!stationSelecionada || stationSelecionada === 'undefined' || typeof stationSelecionada !== 'string') {
        stationSelecionada = '';
    }
    
    if (!stationSelecionada) {
        console.log('⚠️ Nenhuma station definida');
        const containerCiclos = document.getElementById('containerCiclos');
        if (containerCiclos) {
            containerCiclos.innerHTML = '<span class="sem-ciclos">Carregue um arquivo para ver os ciclos</span>';
        }
        return;
    }
    
    console.log(`🔍 Buscando ciclos para: "${stationSelecionada}"`);
    console.log(`📊 Total OpsClock: ${dadosOpsClock.length}`);
    console.log(`📊 Total Outbound: ${dadosOutbound.length}`);
    
    // Debug: mostrar primeiros registros para entender estrutura
    if (dadosOpsClock.length > 0) {
        console.log('🔍 Exemplo OpsClock:', dadosOpsClock[0]);
    }
    if (dadosOutbound.length > 0) {
        console.log('🔍 Exemplo Outbound:', dadosOutbound[0]);
    }
    
    // Normalizar nome da station para comparação (versão completa e reduzida)
    const stationCompleta = stationSelecionada.toLowerCase().replace(/[_\s]+/g, '');
    const stationBase = stationSelecionada
        .toLowerCase()
        .replace(/lm\s*hub[_\s]*/gi, '')  // Remover "LM Hub_"
        .replace(/[_\s]+/g, '')            // Remover underscores e espaços
        .replace(/st\.?\s*empr/gi, '');    // Remover "St. Empr" (mas NÃO remover números!)
    
    console.log(`🔍 Station completa: "${stationCompleta}"`);
    console.log(`🔍 Station base: "${stationBase}"`);
    
    // Função para normalizar nome de station da planilha
    const normalizarStation = (nome) => {
        return nome
            .toLowerCase()
            .replace(/lm\s*hub[_\s]*/gi, '')
            .replace(/[_\s]+/g, '')
            .replace(/st\.?\s*empr/gi, '');
    };
    
    // Filtrar dados de ciclos para a station selecionada
    const ciclosStation = dadosOpsClock.filter(item => {
        const stationName = item['Station name'] || item['Station Name'] || item['station_name'] || 
                           item['Sort Code'] || item['sort_code'] || '';
        const itemNorm = normalizarStation(stationName);
        const status = item['Status'] || '';
        
        // ✅ FILTRAR APENAS REGISTROS ATIVOS
        const isActive = status.includes('Active');
        const matchStation = itemNorm === stationBase || 
                           itemNorm.includes(stationBase) || 
                           stationBase.includes(itemNorm);
        
        return isActive && matchStation;
    });
    
    // Filtrar capacidade para a station selecionada - PRIORIZAR EXATA
    let capacidadeStation = dadosOutbound.filter(item => {
        const sortCodeName = item['Sort Code Name'] || item['sort_code_name'] || 
                            item['Station'] || item['station'] || '';
        const itemNorm = normalizarStation(sortCodeName);
        
        // Match EXATO primeiro
        return itemNorm === stationBase;
    });
    
    // Se não encontrou exata, tentar parcial
    if (capacidadeStation.length === 0) {
        capacidadeStation = dadosOutbound.filter(item => {
            const sortCodeName = item['Sort Code Name'] || item['sort_code_name'] || 
                                item['Station'] || item['station'] || '';
            const itemNorm = normalizarStation(sortCodeName);
            
            return itemNorm.includes(stationBase) || stationBase.includes(itemNorm);
        });
    }
    
    console.log(`✅ Ciclos encontrados: ${ciclosStation.length}`);
    console.log(`✅ Capacidade encontrada: ${capacidadeStation.length}`);
    
    if (ciclosStation.length > 0) {
        console.log('📋 Ciclos filtrados:', ciclosStation);
    }
    if (capacidadeStation.length > 0) {
        console.log('📋 Capacidade filtrada:', capacidadeStation);
        // Mostrar detalhes de cada registro de capacidade
        capacidadeStation.forEach((cap, idx) => {
            const tipo = cap['Type Outbound'] || cap['type_outbound'] || '';
            const sortCode = cap['Sort Code Name'] || cap['sort_code_name'] || '';
            console.log(`   [${idx}] ${sortCode} - ${tipo}`);
        });
    }
    
    // Atualizar dropdown de ciclos
    atualizarDropdownCiclos(ciclosStation);
    
    // Atualizar cards de ciclos
    atualizarCardsCiclos(ciclosStation, capacidadeStation);
}

// Atualizar dropdown de filtro de ciclos
function atualizarDropdownCiclos(ciclosStation) {
    const selectCiclo = document.getElementById('filtroCiclo');
    if (!selectCiclo) return;
    
    // Limpar opções anteriores (manter "Todos")
    selectCiclo.innerHTML = '<option value="">Todos</option>';
    
    // Pegar ciclos únicos
    const ciclosUnicos = [...new Set(ciclosStation.map(item => {
        return item['Dispatch Window'] || item['dispatch_window'] || item['Ciclo'] || '';
    }))].filter(c => c && c !== 'Total');
    
    // Adicionar opções
    ciclosUnicos.forEach(ciclo => {
        const option = document.createElement('option');
        option.value = ciclo;
        option.textContent = ciclo;
        selectCiclo.appendChild(option);
    });
}

// Atualizar cards de ciclos no painel
function atualizarCardsCiclos(ciclosStation, capacidadeStation) {
    const containerCiclos = document.getElementById('containerCiclos');
    if (!containerCiclos) return;
    
    // Usar data do ciclo selecionada (ou hoje)
    const dataCiclo = getDataCicloSelecionada();
    const diaHoje = String(dataCiclo.getDate()).padStart(2, '0');
    const diaSemZero = String(dataCiclo.getDate()); // sem zero à esquerda
    const mesHoje = String(dataCiclo.getMonth() + 1).padStart(2, '0');
    const mesSemZero = String(dataCiclo.getMonth() + 1); // sem zero à esquerda
    const anoHoje = dataCiclo.getFullYear();
    const anoCurto = String(anoHoje).slice(2);
    
    // Formatos possíveis da coluna de data (muitos formatos para garantir compatibilidade)
    const formatosData = [
        `${diaHoje}/${mesHoje}/${anoHoje}`,       // 10/01/2026
        `${diaSemZero}/${mesSemZero}/${anoHoje}`, // 10/1/2026
        `${diaHoje}/${mesHoje}/${anoCurto}`,      // 10/01/26
        `${anoHoje}-${mesHoje}-${diaHoje}`,       // 2026-01-10
        `${mesHoje}/${diaHoje}/${anoHoje}`,       // 01/10/2026 (formato americano)
        `${mesSemZero}/${diaSemZero}/${anoHoje}`, // 1/10/2026 (formato americano sem zero)
    ];
    
    console.log('📅 Data do ciclo selecionada:', formatosData[0]);
    console.log('📅 Formatos testados:', formatosData);
    
    // Debug: mostrar colunas disponíveis na capacidade
    if (capacidadeStation.length > 0) {
        const colunas = Object.keys(capacidadeStation[0]);
        // Encontrar colunas de data (qualquer coisa com números e barras ou traços)
        const colunasData = colunas.filter(col => 
            col.match(/\d+[\/\-]\d+[\/\-]\d+/) 
        );
        console.log('📅 Colunas de data na planilha:', colunasData.slice(0, 10));
        
        // Mostrar se algum formato bate
        formatosData.forEach(fmt => {
            if (colunasData.includes(fmt)) {
                console.log(`✅ Formato "${fmt}" encontrado na planilha!`);
            }
        });
    }
    
    // Montar HTML dos cards
    let html = '';
    
    ciclosStation.forEach(ciclo => {
        const nomeCiclo = ciclo['Dispatch Window'] || ciclo['dispatch_window'] || ciclo['Ciclo'] || '';
        // Usar colunas de Unloading (X e Y) em vez de Routing (V e W)
        const startTime = ciclo['Start time2'] || ciclo['start_time2'] || ciclo['Start time'] || ciclo['start_time'] || '';
        const endTime = ciclo['End time2'] || ciclo['end_time2'] || ciclo['End time'] || ciclo['end_time'] || '';
        
        if (!nomeCiclo || nomeCiclo === 'Total') return;
        
        // Buscar capacidade para este ciclo (comparação case-insensitive)
        const capacidade = capacidadeStation.find(cap => {
            const tipoCap = cap['Type Outbound'] || cap['type_outbound'] || '';
            return tipoCap.toUpperCase() === nomeCiclo.toUpperCase();
        });
        
        console.log(`🔍 Buscando capacidade para ciclo "${nomeCiclo}":`, capacidade ? 'ENCONTRADO' : 'NÃO ENCONTRADO');
        
        // Pegar capacidade do dia de hoje
        let capHoje = 0;
        if (capacidade) {
            // Procurar coluna com data de hoje em vários formatos
            for (const formato of formatosData) {
                if (capacidade[formato] !== undefined && capacidade[formato] !== '') {
                    // Converter para número (remover pontos de milhar, trocar vírgula por ponto)
                    let valor = capacidade[formato];
                    console.log(`   📅 Valor bruto em "${formato}":`, valor, `(tipo: ${typeof valor})`);
                    if (typeof valor === 'string') {
                        valor = valor.replace(/\./g, '').replace(',', '.');
                    }
                    capHoje = parseFloat(valor) || 0;
                    console.log(`✅ Cap ${nomeCiclo} em "${formato}": ${capHoje}`);
                    break;
                }
            }
            
            // Se não encontrou com formato exato, tentar busca parcial nas chaves
            if (capHoje === 0) {
                const todasChaves = Object.keys(capacidade);
                for (const key of todasChaves) {
                    // Verificar se a chave contém o dia e mês de hoje
                    const keyNorm = key.replace(/\s/g, '');
                    if (keyNorm.includes(`${diaHoje}/${mesHoje}`) || 
                        keyNorm.includes(`${diaSemZero}/${mesSemZero}`) ||
                        keyNorm.includes(`${mesHoje}/${diaHoje}`) ||
                        keyNorm.includes(`${anoHoje}-${mesHoje}-${diaHoje}`)) {
                        let valor = capacidade[key];
                        if (typeof valor === 'string') {
                            valor = valor.replace(/\./g, '').replace(',', '.');
                        }
                        capHoje = parseFloat(valor) || 0;
                        console.log(`✅ Cap ${nomeCiclo} (parcial) em "${key}": ${capHoje}`);
                        break;
                    }
                }
            }
            
            // Debug se não encontrou
            if (capHoje === 0) {
                console.log(`⚠️ Não encontrou capacidade para ${nomeCiclo}. Chaves disponíveis:`, Object.keys(capacidade).slice(0, 15));
            }
        }
        
        // Formatar horários (remover segundos se tiver)
        const startFormatado = startTime ? startTime.split(':').slice(0, 2).join(':') : '-';
        const endFormatado = endTime ? endTime.split(':').slice(0, 2).join(':') : '-';
        
        const isAtivo = cicloSelecionado === nomeCiclo ? 'ativo' : '';
        
        // Formatar capacidade com separador de milhar
        const capFormatada = capHoje.toLocaleString('pt-BR');
        
        // Usar mesmo estilo dos cards superiores
        html += `
            <div class="ciclo-stat ${isAtivo}" data-ciclo="${nomeCiclo}">
                <span class="ciclo-nome">${nomeCiclo}</span>
                <span class="ciclo-horario">${startFormatado} - ${endFormatado}</span>
                <span class="ciclo-cap">${capFormatada}</span>
            </div>
        `;
    });
    
    containerCiclos.innerHTML = html || '<span class="sem-ciclos">Nenhum ciclo encontrado</span>';
    
    // Adicionar event listeners nos cards
    containerCiclos.querySelectorAll('.ciclo-stat').forEach(card => {
        card.addEventListener('click', () => {
            const cicloClicado = card.dataset.ciclo;
            console.log('🖱️ Card clicado:', cicloClicado);
            
            // Toggle - se clicar no mesmo, desseleciona
            if (cicloSelecionado === cicloClicado) {
                cicloSelecionado = 'Todos';
                console.log('↩️ Ciclo desselecionado. Ciclo atual:', cicloSelecionado);
            } else {
                cicloSelecionado = cicloClicado;
                console.log('✅ Ciclo selecionado:', cicloSelecionado);
            }
            
            // Atualizar visual dos cards
            containerCiclos.querySelectorAll('.ciclo-stat').forEach(c => {
                c.classList.remove('ativo');
                if (c.dataset.ciclo === cicloSelecionado) {
                    c.classList.add('ativo');
                }
            });
            
            // Re-renderizar tabela
            renderizarTabelaPlanejamento();
        });
    });
}

// Função para toggle do painel colapsável
function togglePainelColapsavel() {
    const painel = document.getElementById('painelColapsavel');
    const btn = document.getElementById('btnTogglePainel');
    
    if (painel && btn) {
        painel.classList.toggle('recolhido');
        btn.classList.toggle('recolhido');
        btn.textContent = painel.classList.contains('recolhido') ? '🔽' : '🔼';
    }
}

// ===== FUNÇÕES PARA GERENCIAR DATA DO CICLO =====

// Inicializar data do ciclo com hoje
function inicializarDataCiclo() {
    const inputData = document.getElementById('dataCicloSelecionada');
    if (inputData) {
        const hoje = new Date();
        const dataFormatada = hoje.toISOString().split('T')[0]; // YYYY-MM-DD
        inputData.value = dataFormatada;
        dataCicloSelecionada = hoje;
    }
}

// Quando a data do ciclo é alterada
function onDataCicloChange(e) {
    const dataStr = e.target.value;
    console.log('📅 [onChange] Evento change disparado!');
    console.log('📅 [onChange] Valor recebido:', dataStr);
    
    if (dataStr) {
        dataCicloSelecionada = new Date(dataStr + 'T00:00:00');
        console.log('📅 [onChange] Data do ciclo alterada para:', dataCicloSelecionada.toLocaleDateString('pt-BR'));
        
        // Atualizar capacidade para a nova data
        atualizarInfoCiclos();
        
        // Re-renderizar tabela com novos cálculos
        renderizarTabelaPlanejamento();
    }
}

// Definir data do ciclo como hoje
function setDataCicloHoje() {
    const inputData = document.getElementById('dataCicloSelecionada');
    if (inputData) {
        const hoje = new Date();
        const dataFormatada = hoje.toISOString().split('T')[0];
        
        console.log('📅 [HOJE] Definindo data como:', hoje.toLocaleDateString('pt-BR'), '(', dataFormatada, ')');
        
        // Atualizar valor do input
        inputData.value = dataFormatada;
        dataCicloSelecionada = hoje;
        
        // ✅ AGUARDAR DOM ATUALIZAR antes de disparar evento
        setTimeout(() => {
            console.log('📅 [HOJE] Disparando evento change...');
            console.log('📅 [HOJE] Valor atual do input:', inputData.value);
            
            const event = new Event('change', { bubbles: true });
            inputData.dispatchEvent(event);
        }, 0);
    }
}

// Obter data do ciclo selecionada (ou hoje se não houver)
function getDataCicloSelecionada() {
    if (dataCicloSelecionada) {
        return new Date(dataCicloSelecionada);
    }
    return new Date();
}

// Contador para próxima atualização
let contadorInterval = null;
let proximaAtualizacaoTime = null;

function iniciarContadorProximaAtualizacao() {
    // Limpar intervalo anterior
    if (contadorInterval) {
        clearInterval(contadorInterval);
    }

    // A planilha atualiza a cada 30 minutos
    const intervaloMinutos = 30;

    if (ultimaAtualizacaoPlanilha) {
        const ultima = new Date(ultimaAtualizacaoPlanilha);
        proximaAtualizacaoTime = new Date(ultima.getTime() + intervaloMinutos * 60 * 1000);
    } else {
        proximaAtualizacaoTime = new Date();
    }

    // Atualizar a cada segundo
    contadorInterval = setInterval(() => {
        atualizarContador();
    }, 1000);

    // Atualizar imediatamente
    atualizarContador();
}

function atualizarContador() {
    const proximaEl = document.getElementById('proximaAtualizacao');
    if (!proximaEl || !proximaAtualizacaoTime) return;

    const agora = new Date();
    const diff = proximaAtualizacaoTime - agora;

    if (diff <= 0) {
        proximaEl.textContent = '⏳ Atualização disponível!';
        proximaEl.classList.add('urgente');
    } else {
        const minutos = Math.floor(diff / 60000);
        const segundos = Math.floor((diff % 60000) / 1000);
        proximaEl.textContent = `⏳ Próxima em: ${minutos.toString().padStart(2, '0')}:${segundos.toString().padStart(2, '0')}`;
        proximaEl.classList.toggle('urgente', minutos < 5);
    }
}

// Estado dos filtros por coluna (Planejamento Hub) - Agora usando filtrosAtivosPlan definido acima

// Função para calcular tempo até o horário de corte do ciclo
function calcularTempoCorte(dadosPlanilhaLH, cicloParam = null) {
    // Se não tiver dados da planilha, retornar sem cálculo
    if (!dadosPlanilhaLH) {
        return { texto: '-', minutos: null, dentroLimite: false, status: 'sem-dados' };
    }
    
    // Pegar previsão final
    const previsaoFinal = dadosPlanilhaLH.previsao_final || dadosPlanilhaLH['PREVISAO FINAL'] || '';
    if (!previsaoFinal || previsaoFinal === '-') {
        return { texto: '-', minutos: null, dentroLimite: false, status: 'sem-dados' };
    }
    
    // Pegar horário de corte do ciclo - USAR PARÂMETRO SE FORNECIDO
    const cicloAtual = cicloParam || cicloSelecionado || 'Todos';
    let horarioInicio = null;
    let horarioCorte = null;
    
    console.log(`🔍 [DEBUG] Buscando horários para ciclo "${cicloAtual}"`);
    console.log(`🔍 [DEBUG] Total registros OpsClock: ${dadosOpsClock.length}`);
    
    // Buscar Start Time e End Time do ciclo selecionado
    if (cicloAtual !== 'Todos' && dadosOpsClock.length > 0) {
        const stationSelecionada = stationAtualNome || '';
        const stationBase = stationSelecionada.toLowerCase().replace(/lm\s*hub[_\s]*/gi, '').replace(/[_\s]+/g, '');
        
        console.log(`🔍 [DEBUG] Station base para busca: "${stationBase}"`);
        
        const cicloInfo = dadosOpsClock.find(item => {
            const stationName = item['Station name'] || item['Station Name'] || '';
            const itemNorm = stationName.toLowerCase().replace(/lm\s*hub[_\s]*/gi, '').replace(/[_\s]+/g, '');
            const dispatchWindow = item['Dispatch Window'] || '';
            const status = item['Status'] || '';
            
            // ✅ FILTRAR APENAS REGISTROS ATIVOS
            const isActive = status.includes('Active');
            const matchStation = (itemNorm === stationBase || itemNorm.includes(stationBase) || stationBase.includes(itemNorm));
            const matchCiclo = dispatchWindow.toUpperCase() === cicloAtual.toUpperCase();
            
            const match = isActive && matchStation && matchCiclo;
            
            if (match) {
                console.log(`✅ [DEBUG] MATCH encontrado!`);
                console.log(`   Station: "${stationName}" (norm: "${itemNorm}")`);
                console.log(`   Dispatch Window: "${dispatchWindow}"`);
                console.log(`   Status: "${status}"`);
                console.log(`   🔑 TODAS AS CHAVES DO REGISTRO:`, Object.keys(item));
                console.log(`   📋 Registro completo:`, item);
            }
            
            return match;
        });
        
        if (cicloInfo) {
            // ✅ PRIORIZAR COLUNAS DE UNLOADING (time2 = X e Y)
            // Routing (V e W) são menos precisas que Unloading (X e Y)
            horarioInicio = cicloInfo['Start time2'] || cicloInfo['start_time2'] || 
                          cicloInfo['Start time'] || cicloInfo['start_time'] || '';
            horarioCorte = cicloInfo['End time2'] || cicloInfo['end_time2'] || 
                         cicloInfo['End time'] || cicloInfo['end_time'] || '';
            
            console.log(`  📋 Dados OpsClock para ${cicloAtual}:`);
            console.log(`     - Start time: ${cicloInfo['Start time'] || 'N/A'}`);
            console.log(`     - Start time2: ${cicloInfo['Start time2'] || 'N/A'}`);
            console.log(`     - End time: ${cicloInfo['End time'] || 'N/A'}`);
            console.log(`     - End time2: ${cicloInfo['End time2'] || 'N/A'}`);
            console.log(`     - ✅ USADO: ${horarioInicio} - ${horarioCorte}`);
        }
    }
    
    // Se não tiver horário de corte definido
    if (!horarioCorte) {
        return { texto: '⏳', minutos: null, dentroLimite: false, status: 'sem-ciclo' };
    }
    
    try {
        // AGORA - hora atual real
        const agora = new Date();
        
        // DATA DO CICLO SELECIONADA (ou hoje)
        const dataCiclo = getDataCicloSelecionada();
        dataCiclo.setHours(0, 0, 0, 0);
        
        console.log(`🔍 calcularTempoCorte - LH: ${dadosPlanilhaLH.lh_trip || 'N/A'}`);
        console.log(`  📅 Data do ciclo: ${dataCiclo.toLocaleDateString('pt-BR')}`);
        console.log(`  📅 Previsão final: ${previsaoFinal}`);
        
        // Parse da previsão final
        const previsaoDate = parsearDataHora(previsaoFinal);
        if (!previsaoDate) {
            return { texto: '-', minutos: null, dentroLimite: false, status: 'erro' };
        }
        
        console.log(`  ⏰ Previsão parseada: ${previsaoDate.toLocaleString('pt-BR')}`);
        
        // Parse dos horários do ciclo
        const [horaInicio, minutoInicio] = (horarioInicio || '00:00').split(':').map(Number);
        const [horaCorte, minutoCorte] = horarioCorte.split(':').map(Number);
        
        console.log(`  🕐 Horário ciclo: ${horaInicio}:${String(minutoInicio || 0).padStart(2, '0')} - ${horaCorte}:${String(minutoCorte || 0).padStart(2, '0')}`);
        
        // Criar datas de início e fim do ciclo de UNLOADING
        // Detectar se janela atravessa meia-noite comparando Start time2 com End time2
        const atravessaMeiaNoite = horaInicio > horaCorte;
        
        let dataInicioCiclo = new Date(dataCiclo);
        let dataFimCiclo = new Date(dataCiclo);
        
        if (atravessaMeiaNoite) {
            // Janela atravessa meia-noite (ex: AM 20:00 - 01:00)
            // Start time2 no DIA ANTERIOR a data de expedicao
            dataInicioCiclo.setDate(dataInicioCiclo.getDate() - 1);
            dataInicioCiclo.setHours(horaInicio, minutoInicio || 0, 0, 0);
            
            // End time2 no DIA DA EXPEDICAO
            dataFimCiclo.setHours(horaCorte, minutoCorte || 0, 0, 0);
            
            console.log(`  🌙 Janela atravessa meia-noite: Start ${dataInicioCiclo.toLocaleString('pt-BR')} - End ${dataFimCiclo.toLocaleString('pt-BR')}`);
        } else {
            // Janela no mesmo dia (ex: PM1 09:00 - 11:00)
            // Ambos Start time2 e End time2 no DIA DA EXPEDICAO
            dataInicioCiclo.setHours(horaInicio, minutoInicio || 0, 0, 0);
            dataFimCiclo.setHours(horaCorte, minutoCorte || 0, 0, 0);
            
            console.log(`  ☀ Janela no mesmo dia: ${dataInicioCiclo.toLocaleString('pt-BR')} - ${dataFimCiclo.toLocaleString('pt-BR')}`);
        }
        
        // Corte com margem de 45 min
        const corteComMargem = new Date(dataFimCiclo.getTime() - 45 * 60 * 1000);
        
        console.log(`  ⏰ Corte (com 45min): ${corteComMargem.toLocaleString('pt-BR')}`);
        
        // Verificar se está no piso
        const etaRealized = dadosPlanilhaLH.eta_destination_realized || 
                           dadosPlanilhaLH['ETA_DESTINATION_REALIZED'] || '';
        const unloaded = dadosPlanilhaLH.unloaded_destination_datetime || 
                        dadosPlanilhaLH['UNLOADED_DESTINATION_DATETIME'] || '';
        const estaNoPiso = etaRealized !== '' && unloaded !== '';
        
        // ===== LÓGICA DE CÁLCULO =====
        let texto = '';
        let dentroLimite = false;
        let minutosCalculado = null;
        let status = '';
        
        // Verificar se o ciclo já passou (comparando com AGORA)
        const cicloJaPassou = agora > dataFimCiclo;
        
        if (cicloJaPassou) {
            // Ciclo já encerrou
            return { 
                texto: '⛔ Ciclo encerrado', 
                minutos: null, 
                dentroLimite: false, 
                status: 'ciclo-encerrado',
                tooltip: 'Este ciclo já encerrou. Selecione outra data ou ciclo para planejar.'
            };
        }
        
        // ✅ LÓGICA CORRIGIDA: Sempre comparar PREVISÃO DA LH vs CORTE COM MARGEM
        // Tempo restante = Corte (com 45min de margem) - Previsão da LH
        minutosCalculado = Math.floor((corteComMargem.getTime() - previsaoDate.getTime()) / (1000 * 60));
        
        console.log(`  Minutos calculados: ${minutosCalculado} (${minutosCalculado >= 0 ? 'DENTRO DO PRAZO' : 'Em transito - fora do prazo'})`);
        
        // ===== FORMATAR RESULTADO =====
        if (minutosCalculado >= 60) {
            // Mais de 1 hora - POSITIVO (verde)
            const h = Math.floor(minutosCalculado / 60);
            const m = minutosCalculado % 60;
            texto = `✅ ${h}h${m > 0 ? m + 'm' : ''}`;
            dentroLimite = true;
            status = 'ok';
        } else if (minutosCalculado >= 45) {
            // Entre 45-60 min - POSITIVO (verde)
            texto = `✅ ${minutosCalculado} min`;
            dentroLimite = true;
            status = 'ok';
        } else if (minutosCalculado >= 15) {
            // Entre 15-45 min - ATENÇÃO (amarelo)
            texto = `⚠️ ${minutosCalculado} min`;
            dentroLimite = true;
            status = 'atencao';
        } else if (minutosCalculado >= 0) {
            // Menos de 15 min - URGENTE (laranja)
            texto = `🔶 ${minutosCalculado} min`;
            dentroLimite = true;
            status = 'urgente';
        } else {
            // NEGATIVO - passou do corte (vermelho)
            const atraso = Math.abs(minutosCalculado);
            if (atraso >= 60) {
                const h = Math.floor(atraso / 60);
                const m = atraso % 60;
                texto = `❌ -${h}h${m > 0 ? m + 'm' : ''}`;
            } else {
                texto = `❌ -${atraso} min`;
            }
            dentroLimite = false;
            status = 'atrasado';
        }
        
        return { texto, minutos: minutosCalculado, dentroLimite, status };
        
    } catch (error) {
        console.error('Erro ao calcular tempo de corte:', error);
        return { texto: '-', minutos: null, dentroLimite: false, status: 'erro' };
    }
}

// Função auxiliar para parsear data/hora em vários formatos
function parsearDataHora(str) {
    if (!str) return null;
    
    try {
        const partes = String(str).trim().split(' ');
        if (partes.length < 2) return null;
        
        const [dataStr, horaStr] = partes;
        let dia, mes, ano;
        
        if (dataStr.includes('/')) {
            // Formato DD/MM/YYYY ou D/M/YYYY
            const dataParts = dataStr.split('/');
            dia = parseInt(dataParts[0]);
            mes = parseInt(dataParts[1]) - 1; // Mês é 0-indexed
            ano = parseInt(dataParts[2]);
        } else if (dataStr.includes('-')) {
            // Formato YYYY-MM-DD
            const dataParts = dataStr.split('-');
            ano = parseInt(dataParts[0]);
            mes = parseInt(dataParts[1]) - 1;
            dia = parseInt(dataParts[2]);
        } else {
            return null;
        }
        
        const [hora, minuto, segundo] = horaStr.split(':').map(Number);
        
        return new Date(ano, mes, dia, hora, minuto || 0, segundo || 0);
    } catch (e) {
        return null;
    }
}

// Função para calcular o status da LH baseado nas colunas da planilha
function calcularStatusLH(dadosPlanilhaLH, qtdPedidos = null, estatisticas = null, cicloSelecionado = null) {
    // Extrair LH Trip com múltiplos fallbacks (PRIORIDADE: trip_number é o campo correto!)
    const lhTrip = dadosPlanilhaLH?.trip_number ||  // ← CORRETO!
                   dadosPlanilhaLH?.lh_trip || 
                   dadosPlanilhaLH?.['LH Trip'] || 
                   dadosPlanilhaLH?.['LH_TRIP'] ||
                   dadosPlanilhaLH?.['lh trip'] ||
                   dadosPlanilhaLH?.lhTrip ||
                   'N/A';
    
    // PRIORIDADE 1: Verificar cache SPX primeiro (sobrescreve qualquer cálculo)
    console.log(`   🔍 [CACHE CHECK] Verificando cache SPX para ${lhTrip}... (cache tem ${cacheSPX.size} entradas)`);
    
    // Debug: mostrar chaves do cache
    if (cacheSPX.size > 0 && lhTrip !== 'N/A') {
        const cacheKeys = Array.from(cacheSPX.keys());
        console.log(`   📋 [CACHE KEYS] Chaves no cache:`, cacheKeys);
        console.log(`   🔍 [CACHE SEARCH] Procurando por: "${lhTrip}"`);
    }
    
    if (cacheSPX.has(lhTrip)) {
        const validacaoSPX = cacheSPX.get(lhTrip);
        console.log(`   💾 [CACHE SPX HIT] ✅✅✅ Usando validação SPX PERMANENTE: ${lhTrip} → ${validacaoSPX.statusCodigo}`);
        return {
            codigo: validacaoSPX.statusCodigo,
            texto: validacaoSPX.status,
            classe: `status-${validacaoSPX.statusCodigo.toLowerCase()}`,
            icone: validacaoSPX.statusCodigo === 'P0' ? '✅' : '🚚',
            isBloqueada: false, // Nunca bloqueia se veio do SPX
            _spxValidado: true
        };
    } else {
        if (lhTrip === 'N/A') {
            const props = Object.keys(dadosPlanilhaLH || {});
            console.log(`   ⚠️ [CACHE WARNING] LH Trip = N/A! Propriedades do objeto (${props.length}):`, props);
        } else {
            console.log(`   ❌ [CACHE SPX MISS] ${lhTrip} não encontrado no cache`);
        }
    }
    
    if (!dadosPlanilhaLH) {
        // LH sem dados na planilha
        console.log(`🔍 [SEM DADOS] LH sem dados na planilha - ${qtdPedidos} pedidos`);
        
        // Verificar se é BAIXO volume antes de sinalizar inventário
        const isBaixoVolume = qtdPedidos !== null && estatisticas && verificarLHBaixoVolume(qtdPedidos, estatisticas);
        
        if (isBaixoVolume) {
            // Baixo volume + sem dados → Sinalizar Inventário
            console.log(`🔍 LH sem dados + baixo volume (${qtdPedidos} pedidos) → Sinalizar Inventário`);
            return { codigo: 'P0I', texto: 'Sinalizar Inventário', classe: 'status-p0i', icone: '🔍' };
        } else {
            // Alto volume + sem dados → Status genérico P3 (não bloqueia)
            console.log(`⚠️ LH sem dados mas ALTO volume (${qtdPedidos} pedidos) → P3 genérico`);
            return { 
                codigo: 'P3', 
                texto: 'Em transito - fora do prazo', 
                classe: 'status-p3', 
                icone: '⛔',
                isBloqueada: true 
            };
        }
    }
    
    // VERIFICAR BAIXO VOLUME (mas só sinalizar inventário se já passou do prazo)
    const isBaixoVolume = qtdPedidos !== null && estatisticas && verificarLHBaixoVolume(qtdPedidos, estatisticas);
    
    // Se tem baixo volume, verificar se a previsão já passou
    if (isBaixoVolume) {
        const lhTrip = dadosPlanilhaLH.lh_trip || dadosPlanilhaLH['LH Trip'] || dadosPlanilhaLH['LH_TRIP'] || 'N/A';
        console.log(`🔍 LH com baixo volume detectada: ${lhTrip} - ${qtdPedidos} pedidos`);
        
        // Buscar previsão final para verificar se já passou do prazo
        const buscarValor = (obj, ...nomes) => {
            for (const nome of nomes) {
                if (obj[nome] !== undefined && obj[nome] !== null && String(obj[nome]).trim() !== '') {
                    return String(obj[nome]).trim();
                }
                const nomeLower = nome.toLowerCase();
                if (obj[nomeLower] !== undefined && obj[nomeLower] !== null && String(obj[nomeLower]).trim() !== '') {
                    return String(obj[nomeLower]).trim();
                }
            }
            return '';
        };
        
        const previsaoFinal = buscarValor(dadosPlanilhaLH, 'previsao_final', 'PREVISAO FINAL', 'Previsao Final');
        console.log(`   📅 Previsão Final encontrada: "${previsaoFinal}"`);
        console.log(`   📊 Dados disponíveis:`, Object.keys(dadosPlanilhaLH));
        
        // Verificar se a previsão é futura (não passou do prazo)
        let previsaoFutura = false;
        if (previsaoFinal) {
            try {
                // Tentar converter a data completa (com hora) primeiro
                let dataPrevisao = new Date(previsaoFinal);
                
                // Se não conseguiu, tentar formato brasileiro
                if (isNaN(dataPrevisao.getTime())) {
                    const apenasData = previsaoFinal.split(' ')[0];
                    console.log(`   📅 Data extraída: "${apenasData}"`);
                    
                    const partesData = apenasData.split('/');
                    if (partesData.length === 3) {
                        const dia = parseInt(partesData[0]);
                        const mes = parseInt(partesData[1]) - 1;
                        const ano = parseInt(partesData[2]);
                        
                        // Se tem hora, extrair também
                        const partesHora = previsaoFinal.split(' ')[1]?.split(':');
                        if (partesHora && partesHora.length >= 2) {
                            const hora = parseInt(partesHora[0]);
                            const minuto = parseInt(partesHora[1]);
                            dataPrevisao = new Date(ano, mes, dia, hora, minuto);
                        } else {
                            dataPrevisao = new Date(ano, mes, dia);
                        }
                    }
                }
                
                if (!isNaN(dataPrevisao.getTime())) {
                    const agora = new Date();
                    
                    console.log(`   📅 Data previsão: ${dataPrevisao.toLocaleString('pt-BR')}`);
                    console.log(`   📅 Agora: ${agora.toLocaleString('pt-BR')}`);
                    
                    // Se previsão é no futuro, não sinalizar inventário
                    if (dataPrevisao > agora) {
                        previsaoFutura = true;
                        console.log(`   ✅ Previsão futura - NÃO sinalizar inventário`);
                    } else {
                        console.log(`   ⚠️ Previsão já passou - PODE sinalizar inventário`);
                    }
                }
            } catch (e) {
                console.log(`   ⚠️ Erro ao processar data: ${e.message}`);
            }
        }
        
        // Só sinalizar inventário se NÃO for previsão futura
        if (!previsaoFutura) {
            console.log(`   🔍 Baixo volume + previsão passada → Sinalizar Inventário`);
            return { 
                codigo: 'P0I', 
                texto: 'Sinalizar Inventário', 
                classe: 'status-p0i',
                icone: '🔍',
                isBaixoVolume: true
            };
        } else {
            console.log(`   ➡️ Baixo volume MAS previsão futura → Seguir fluxo normal`);
            // Continuar para determinar status normal (P2/P3)
        }
    }
    
    // Função auxiliar para buscar valor em múltiplas variações de nome
    const buscarValor = (obj, ...nomes) => {
        for (const nome of nomes) {
            // Tentar nome exato
            if (obj[nome] !== undefined && obj[nome] !== null && String(obj[nome]).trim() !== '') {
                return String(obj[nome]).trim();
            }
            // Tentar lowercase
            const nomeLower = nome.toLowerCase();
            if (obj[nomeLower] !== undefined && obj[nomeLower] !== null && String(obj[nomeLower]).trim() !== '') {
                return String(obj[nomeLower]).trim();
            }
        }
        return '';
    };
    
    // COLUNAS CORRETAS:
    // Coluna M - eta_destination_realized (LH CHEGOU no hub)
    const etaDestinationRealized = buscarValor(dadosPlanilhaLH, 
        'eta_destination_realized', 
        'ETA_DESTINATION_REALIZED'
    );
    
    // Coluna O - unloaded_destination_datetime (LH FOI DESCARREGADA)
    const unloadedDatetime = buscarValor(dadosPlanilhaLH, 
        'unloaded_destination_datetime', 
        'UNLOADED_DESTINATION_DATETIME'
    );
    
    // Verificar se os campos estão preenchidos
    const chegouNoHub = etaDestinationRealized !== '';
    const foiDescarregada = unloadedDatetime !== '';
    
    // P0 = LH no piso (chegou E descarregou)
    if (chegouNoHub && foiDescarregada) {
        // Verificar se é backlog (previsão < hoje - 1 dia)
        const previsaoFinal = dadosPlanilhaLH.previsao_final || dadosPlanilhaLH['PREVISAO FINAL'] || '';
        const isBacklogPiso = verificarSeBacklogPiso(previsaoFinal);
        
        if (isBacklogPiso) {
            return { 
                codigo: 'P0B', 
                texto: 'Backlog', 
                classe: 'status-p0b',
                icone: '📦',
                isBacklog: true
            };
        }
        
        return { 
            codigo: 'P0', 
            texto: 'No Piso', 
            classe: 'status-p0',
            icone: '✅'
        };
    }
    
    // P1 = LH chegou no Hub (chegou mas não descarregou ainda)
    if (chegouNoHub && !foiDescarregada) {
        return { 
            codigo: 'P1', 
            texto: 'Aguard. Descarregamento', 
            classe: 'status-p1',
            icone: '🚚'
        };
    }
    
    // P2 ou P3 = LH em trânsito (não chegou ainda)
    // Verificar pelo tempo até o corte
    // ✅ IMPORTANTE: Usar o ciclo selecionado para calcular o tempo de corte corretamente
    const tempoCorte = calcularTempoCorte(dadosPlanilhaLH, cicloSelecionado);
    
    console.log(`🔍 calcularStatusLH - LH: ${lhTrip}`);
    console.log(`   📅 Ciclo recebido: ${cicloSelecionado || 'NULL'}`);
    console.log(`   ⏱️ Tempo corte: ${tempoCorte.minutos} min (${tempoCorte.dentroLimite ? 'DENTRO' : 'FORA'})`);
    
    if (tempoCorte.dentroLimite) {
        // P2 = Vai chegar no prazo
        console.log(`   ✅ Status: P2 (Em Trânsito)`);
        return { 
            codigo: 'P2', 
            texto: 'Em Trânsito', 
            classe: 'status-p2',
            icone: '🔄'
        };
    } else {
        // P3 = Não vai chegar no prazo
        console.log(`   ⛔ Status: P3 (Fora do prazo)`);
        return { 
            codigo: 'P3', 
            texto: 'Em transito - fora do prazo', 
            classe: 'status-p3',
            icone: '⚠️'
        };
    }
}

// Verificar se a LH no piso é backlog (previsão < hoje - 1 dia)
function verificarSeBacklogPiso(previsaoFinal) {
    if (!previsaoFinal) return false;
    
    try {
        // Extrair data da previsão
        let dataPrevisao;
        const str = String(previsaoFinal).trim();
        
        // Formato: "09/01/2026 19:48:40" ou "2026-01-09 19:48:40"
        if (str.includes('/')) {
            // DD/MM/YYYY
            const partes = str.split(' ')[0].split('/');
            if (partes.length === 3) {
                const dia = parseInt(partes[0]);
                const mes = parseInt(partes[1]) - 1; // Mês é 0-indexed
                const ano = parseInt(partes[2]);
                dataPrevisao = new Date(ano, mes, dia);
            }
        } else if (str.includes('-')) {
            // YYYY-MM-DD
            const partes = str.split(' ')[0].split('-');
            if (partes.length === 3) {
                const ano = parseInt(partes[0]);
                const mes = parseInt(partes[1]) - 1;
                const dia = parseInt(partes[2]);
                dataPrevisao = new Date(ano, mes, dia);
            }
        }
        
        if (!dataPrevisao || isNaN(dataPrevisao.getTime())) return false;
        
        // Data limite: hoje - 1 dia (margem de 1 dia)
        const hoje = new Date();
        hoje.setHours(0, 0, 0, 0);
        const dataLimite = new Date(hoje);
        dataLimite.setDate(dataLimite.getDate() - 1);
        
        // Se previsão < dataLimite, é backlog
        return dataPrevisao < dataLimite;
        
    } catch (error) {
        console.error('Erro ao verificar backlog piso:', error);
        return false;
    }
}

// Função para calcular estatísticas de volume de pedidos e identificar outliers
function calcularEstatisticasVolume(lhsData) {
    if (!lhsData || lhsData.length === 0) {
        return { media: 0, mediana: 0, percentil10: 0, percentil25: 0, desvio: 0 };
    }
    
    // Extrair quantidade de pedidos de cada LH
    const volumes = lhsData.map(lh => lh.pedidos || 0).filter(v => v > 0);
    
    if (volumes.length === 0) {
        return { media: 0, mediana: 0, percentil10: 0, percentil25: 0, desvio: 0 };
    }
    
    // Ordenar para calcular percentis e mediana
    const volumesOrdenados = [...volumes].sort((a, b) => a - b);
    
    // Calcular média
    const soma = volumes.reduce((acc, v) => acc + v, 0);
    const media = soma / volumes.length;
    
    // Calcular mediana
    const meio = Math.floor(volumesOrdenados.length / 2);
    const mediana = volumesOrdenados.length % 2 === 0
        ? (volumesOrdenados[meio - 1] + volumesOrdenados[meio]) / 2
        : volumesOrdenados[meio];
    
    // Calcular percentis
    const percentil10 = volumesOrdenados[Math.floor(volumesOrdenados.length * 0.10)];
    const percentil25 = volumesOrdenados[Math.floor(volumesOrdenados.length * 0.25)];
    
    // Calcular desvio padrão
    const variancia = volumes.reduce((acc, v) => acc + Math.pow(v - media, 2), 0) / volumes.length;
    const desvio = Math.sqrt(variancia);
    
    return {
        media: Math.round(media),
        mediana: Math.round(mediana),
        percentil10: Math.round(percentil10),
        percentil25: Math.round(percentil25),
        desvio: Math.round(desvio)
    };
}

// Função para verificar se uma LH tem volume anormalmente baixo (outlier)
function verificarLHBaixoVolume(qtdPedidos, estatisticas) {
    if (!estatisticas || qtdPedidos === 0) return false;
    
    // Limite absoluto máximo: nunca considerar baixo volume se tiver 100+ pedidos
    // LHs fechadas normalmente têm pelo menos 100 pedidos
    const LIMITE_ABSOLUTO_MAX = 100;
    
    // Usar 30% da média como critério, mas com teto de 100 pedidos
    const limite30Porcento = Math.round(estatisticas.media * 0.30);
    const limiteMedia = Math.min(limite30Porcento, LIMITE_ABSOLUTO_MAX);
    
    // É baixo volume se estiver abaixo do limite
    const isBaixo = qtdPedidos < limiteMedia;
    
    // Debug detalhado para volumes próximos aos limites
    if (qtdPedidos > 50 && qtdPedidos < 200) {
        console.log(`🔍 [DEBUG VOLUME] ${qtdPedidos} pedidos:`);
        console.log(`   📊 Média: ${Math.round(estatisticas.media)}`);
        console.log(`   📊 Limite 30% média: ${limite30Porcento}`);
        console.log(`   📊 Limite máximo absoluto: ${LIMITE_ABSOLUTO_MAX}`);
        console.log(`   📊 Limite USADO: ${limiteMedia}`);
        console.log(`   📊 É baixo? ${isBaixo ? 'SIM' : 'NÃO'} (${qtdPedidos} < ${limiteMedia})`);
    }
    
    return isBaixo;
}

// Função auxiliar para converter previsão (data + hora) em timestamp
function parsePrevisaoParaTimestamp(previsaoStr) {
    try {
        // Formato esperado: "DD/MM/YYYY HH:MM:SS" ou "DD/MM/YYYY HH:MM"
        const partes = previsaoStr.trim().split(' ');
        if (partes.length < 2) return null;
        
        const [data, hora] = partes;
        const [dia, mes, ano] = data.split('/');
        const [horas, minutos, segundos = '0'] = hora.split(':');
        
        // Criar objeto Date
        const timestamp = new Date(
            parseInt(ano),
            parseInt(mes) - 1, // Mês começa em 0
            parseInt(dia),
            parseInt(horas),
            parseInt(minutos),
            parseInt(segundos)
        );
        
        return timestamp.getTime();
    } catch (error) {
        console.error('❌ Erro ao parsear previsão:', previsaoStr, error);
        return null;
    }
}

// ======================= FILTRO DE LIXO SISTÊMICO =======================
/**
 * Detecta LHs "lixo sistêmico" - sem origin, destination, previsão e com poucos pedidos
 * Essas LHs devem ser movidas automaticamente para Backlog
 */
function isLixoSistemico(rowData) {
    // ✅ NOVA LÓGICA: Identificar LHs não encontradas na planilha SPX
    // Se a LH não foi encontrada (dadosPlanilhaLH === null), considerar lixo sistêmico
    const naoEncontrada = !rowData.dadosPlanilhaLH || rowData.encontrada === false;
    const pedidos = rowData.pedidos || 0;
    const poucosPedidos = pedidos <= 2;
    
    // Se não foi encontrada na planilha E tem poucos pedidos, é lixo
    if (naoEncontrada && poucosPedidos) {
        console.log(`🗑️ Lixo sistêmico detectado (não encontrada): ${rowData.lh_trip}`, {
            encontrada: rowData.encontrada,
            pedidos
        });
        return true;
    }
    
    // Lógica antiga: Buscar valores de origin e destination
    const origin = rowData.origin || rowData.dadosPlanilhaLH?.origin || 
                   rowData.dadosPlanilhaLH?.Origin || rowData.dadosPlanilhaLH?.ORIGIN || '';
    const destination = rowData.destination || rowData.dadosPlanilhaLH?.destination || 
                        rowData.dadosPlanilhaLH?.Destination || rowData.dadosPlanilhaLH?.DESTINATION || '';
    const previsaoData = rowData.previsao_data || '';
    
    // Verificar se todos os campos estão vazios/inválidos
    const semOrigin = !origin || origin === '-' || origin.trim() === '';
    const semDestination = !destination || destination === '-' || destination.trim() === '';
    const semPrevisao = !previsaoData || previsaoData === '-' || previsaoData.trim() === '';
    
    const isLixo = semOrigin && semDestination && semPrevisao && poucosPedidos;
    
    // Log para debug
    if (isLixo) {
        console.log(`🗑️ Lixo sistêmico detectado: ${rowData.lh_trip}`, {
            origin: origin || '(vazio)',
            destination: destination || '(vazio)',
            previsao: previsaoData || '(vazio)',
            pedidos
        });
    }
    
    return isLixo;
}

// Renderizar tabela de planejamento (cruzamento SPX x Planilha)
function renderizarTabelaPlanejamento() {
    const tbody = document.getElementById('tbodyPlanejamento');
    const thead = document.getElementById('theadPlanejamento');
    if (!tbody) return;

    // USAR APENAS LHs PLANEJÁVEIS (não backlog por status)
    // lhTripsPlanejáveis contém apenas pedidos com status diferente de LMHub_Received e Return_LMHub_Received
    const lhsDoSPX = Object.keys(lhTripsPlanejáveis).filter(lh => lh && lh !== '(vazio)' && lh.trim() !== '');
    
    // Pedidos sem LH que são planejáveis (não backlog por status)
    const pedidosSemLHPlanejáveis = pedidosPlanejáveis.filter(p => {
        const colunaLH = Object.keys(p).find(col =>
            col.toLowerCase().includes('lh trip') ||
            col.toLowerCase().includes('lh_trip')
        ) || 'LH Trip';
        const lh = p[colunaLH] || '';
        return !lh || lh.trim() === '';
    }).length;
    
    // Total de backlog (para mostrar info)
    const totalBacklog = pedidosBacklogPorStatus.length;
    const lhsBacklogCount = Object.keys(lhTripsBacklog).filter(lh => lh !== '(vazio)').length;

    if (lhsDoSPX.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="20" style="text-align: center; padding: 40px; color: #999;">
                    <div style="font-size: 48px; margin-bottom: 15px;">📦</div>
                    <h3>Nenhuma LH planejável encontrada</h3>
                    <p>Baixe os dados do SPX primeiro para fazer o cruzamento</p>
                    ${totalBacklog > 0 ? `<p style="margin-top:10px;color:#ff9800;">📋 ${totalBacklog} pedidos identificados como Backlog (${lhsBacklogCount} LHs) → Aba "Tratar Backlog"</p>` : ''}
                    ${pedidosSemLHPlanejáveis > 0 ? `<p style="margin-top:10px;color:#999;">⚠️ ${pedidosSemLHPlanejáveis} pedidos sem LH Trip</p>` : ''}
                </td>
            </tr>
        `;
        atualizarEstatisticasPlanejamento(0, 0, 0, 0);
        return;
    }

    // CAPTURAR FILTRO DE BUSCA
    const filtroBuscaGeral = document.getElementById('filtroPlanejamentoBusca')?.value.toLowerCase() || '';

    // CARREGAR COLUNAS CONFIGURADAS
    let colunasPlanejamento = JSON.parse(localStorage.getItem('colunasPlanejamento') || '[]');
    
    if (colunasPlanejamento.length === 0) {
        // Colunas padrão: origin, destination, status_lh, previsão separada, tempo para corte
        colunasPlanejamento = ['origin', 'destination', 'status_lh', 'previsao_data', 'previsao_hora', 'tempo_corte'];
    }
    
    // Remover colunas antigas se existirem (migração)
    colunasPlanejamento = colunasPlanejamento.filter(col => 
        col !== 'eta_destination_edited' && col !== 'is_full'
    );
    
    // Garantir que status_lh e tempo_corte estão nas colunas
    if (!colunasPlanejamento.includes('status_lh')) {
        // Inserir antes de previsao_data se existir, senão no final
        const idxPrevisao = colunasPlanejamento.indexOf('previsao_data');
        if (idxPrevisao >= 0) {
            colunasPlanejamento.splice(idxPrevisao, 0, 'status_lh');
        } else {
            colunasPlanejamento.push('status_lh');
        }
    }
    if (!colunasPlanejamento.includes('tempo_corte')) {
        colunasPlanejamento.push('tempo_corte');
    }

    // Colunas fixas + configuradas
    const colunasFixas = ['tipo', 'status', 'lh_trip', 'pedidos', 'pedidos_tos'];
    const todasColunasPlan = [...colunasFixas, ...colunasPlanejamento];

    // Preparar dados para filtragem
    let dadosTabela = [];
    
    // Calcular estatísticas de volume para identificar outliers
    const dadosParaEstatisticas = lhsDoSPX.map(lhTrip => ({
        pedidos: lhTripsPlanejáveis[lhTrip]?.length || 0
    }));
    const estatisticasVolume = calcularEstatisticasVolume(dadosParaEstatisticas);
    
    console.log('📊 Estatísticas de Volume de LHs:', estatisticasVolume);
    console.log(`   ➡️ Média: ${estatisticasVolume.media} pedidos`);
    console.log(`   ➡️ Percentil 10: ${estatisticasVolume.percentil10} pedidos`);
    console.log(`   ➡️ Limite 30% da média: ${Math.round(estatisticasVolume.media * 0.3)} pedidos`);
    console.log(`   ➡️ LHs com baixo volume serão: ≤ ${Math.max(estatisticasVolume.percentil10, Math.round(estatisticasVolume.media * 0.3))} pedidos`);
    
    lhsDoSPX.forEach(lhTrip => {
        // Usar lhTripsPlanejáveis para contar pedidos (apenas os planejáveis)
        const qtdPedidos = lhTripsPlanejáveis[lhTrip]?.length || 0;
        // ✅ USAR FUNÇÃO DE FILTRO POR STATION
        const dadosPlanilhaLH = buscarDadosPlanilhaPorStation(lhTrip);
        const encontrada = !!dadosPlanilhaLH;

        // Identificar LHs FBS (Full by Shopee) pela coluna ORIGIN
        // Buscar em várias variações de nome de coluna (case-insensitive)
        let valorOrigin = null;
        if (dadosPlanilhaLH) {
            // Tentar várias variações de nome
            valorOrigin = dadosPlanilhaLH.origin || dadosPlanilhaLH.Origin || dadosPlanilhaLH.ORIGIN || 
                          dadosPlanilhaLH.origem || dadosPlanilhaLH.Origem;
            
            // Se não encontrou, buscar case-insensitive em todas as chaves
            if (!valorOrigin) {
                const chaveOrigin = Object.keys(dadosPlanilhaLH).find(k => k.toLowerCase() === 'origin');
                if (chaveOrigin) valorOrigin = dadosPlanilhaLH[chaveOrigin];
            }
        }
        
        const isFBS = valorOrigin && typeof valorOrigin === 'string' && 
                      valorOrigin.toUpperCase().startsWith('FBS_');
        
        // DEBUG: Log para verificar detecção FBS
        if (isFBS) {
            console.log(`⚡ LH FBS detectada: ${lhTrip}, ORIGIN: ${valorOrigin}`);
        }
        
        // LH é considerada FULL se:
        // 1. Tem FBS_ no ORIGIN (nova lógica prioritária), OU
        // 2. Tem flag is_full/is_full_truck/tipo_carga = Full (lógica antiga)
        const isFull = isFBS || 
                       dadosPlanilhaLH?.is_full === 'Full' || dadosPlanilhaLH?.is_full === 'Sim' ||
                       dadosPlanilhaLH?.is_full_truck === 'Full' || dadosPlanilhaLH?.tipo_carga === 'Full';

        // Aplicar filtro de busca
        if (filtroBuscaGeral && !lhTrip.toLowerCase().includes(filtroBuscaGeral)) return;

        // Verificar se é LH No Piso com estouro (marcada na sugestão)
        const estouroPiso = window.lhsComEstouroPiso && 
                            window.lhsComEstouroPiso.some(lh => lh.lhTrip === lhTrip);
        
        // Verificar se esta LH é a sugerida para complemento
        const complementoSugerido = window.lhComplementoSugerida && 
                                    window.lhComplementoSugerida.lhTrip === lhTrip;
        
        // Montar objeto com dados para filtro
        const rowData = {
            tipo: isFull ? 'FULL' : 'Normal',
            status: encontrada ? 'Encontrada' : 'Não encontrada',
            lh_trip: lhTrip,
            pedidos: qtdPedidos,
            pedidos_tos: calcularTotalTOsSelecionadas(lhTrip),
            isFull,
            encontrada,
            dadosPlanilhaLH,
            estouroPiso,         // Flag para destaque visual (LH no piso com estouro)
            complementoSugerido  // Flag para destaque visual (LH sugerida para complemento)
        };

        // Adicionar colunas da planilha
        colunasPlanejamento.forEach(col => {
            if (col === 'status_lh') {
                // Calcular status da LH baseado nas colunas da planilha
                // ✅ PASSAR CICLO SELECIONADO para calcular status P3 corretamente
                rowData['status_lh'] = calcularStatusLH(dadosPlanilhaLH, qtdPedidos, estatisticasVolume, cicloSelecionado);
            } else if (col === 'previsao_data' || col === 'previsao_hora') {
                // Separar previsao_final em data e hora
                const previsaoFinal = dadosPlanilhaLH?.previsao_final || dadosPlanilhaLH?.['PREVISAO FINAL'] || '';
                if (previsaoFinal && previsaoFinal !== '-') {
                    const { data, hora } = formatarPrevisaoFinal(previsaoFinal);
                    rowData['previsao_data'] = data;
                    rowData['previsao_hora'] = hora;
                } else {
                    rowData['previsao_data'] = '-';
                    rowData['previsao_hora'] = '-';
                }
            } else if (col === 'tempo_corte') {
                // Calcular tempo até o horário de corte do ciclo
                rowData['tempo_corte'] = calcularTempoCorte(dadosPlanilhaLH);
                rowData['tempo_corte_minutos'] = rowData['tempo_corte'].minutos; // Para ordenação
            } else {
                rowData[col] = dadosPlanilhaLH?.[col] || '-';
            }
        });

        dadosTabela.push(rowData);
    });

    // 🗑️ FILTRAR LIXO SISTÊMICO (mover para backlog automaticamente)
    lhsLixoSistemico = dadosTabela.filter(row => isLixoSistemico(row));
    dadosTabela = dadosTabela.filter(row => !isLixoSistemico(row));
    
    // Log de LHs filtradas
    if (lhsLixoSistemico.length > 0) {
        const totalPedidosLixo = lhsLixoSistemico.reduce((sum, row) => sum + (row.pedidos || 0), 0);
        console.log(`🗑️ ${lhsLixoSistemico.length} LHs de lixo sistêmico filtradas automaticamente (${totalPedidosLixo} pedidos):`);
        lhsLixoSistemico.forEach(row => {
            console.log(`   - ${row.lh_trip} (${row.pedidos} pedidos)`);
        });
        
        // ✅ ADICIONAR PEDIDOS AO BACKLOG
        // Buscar os pedidos dessas LHs nos dados originais e marcar como backlog
        const lhsLixoSet = new Set(lhsLixoSistemico.map(row => row.lh_trip));
        const colunaLH = todasColunas.find(c => 
            c.toLowerCase() === 'lh trip' || 
            c.toLowerCase() === 'lh_trip' ||
            c.toLowerCase() === 'lhtrip'
        ) || 'LH Trip';
        
        // Adicionar pedidos dessas LHs ao backlog
        const pedidosLixo = dadosAtuais.filter(pedido => {
            const lhTrip = pedido[colunaLH];
            return lhTrip && lhsLixoSet.has(lhTrip);
        });
        
        console.log(`🔍 [DEBUG LIXO] Pedidos encontrados para LHs lixo:`);
        pedidosLixo.forEach(pedido => {
            const shipmentId = pedido['Shipment ID'] || pedido['SHIPMENT ID'] || pedido['shipment_id'] || '(sem ID)';
            const lhTrip = pedido[colunaLH];
            console.log(`   - ${shipmentId} (LH: ${lhTrip})`);
        });
        
        // Adicionar ao array de backlog (se ainda não estiverem lá)
        let pedidosAdicionados = 0;
        pedidosLixo.forEach(pedido => {
            if (!pedidosBacklogPorStatus.find(p => p === pedido)) {
                pedidosBacklogPorStatus.push(pedido);
                pedidosAdicionados++;
            }
        });
        
        console.log(`✅ ${pedidosAdicionados} pedidos adicionados ao BACKLOG (${pedidosLixo.length - pedidosAdicionados} já estavam lá)`);
        console.log(`📊 [DEBUG LIXO] Total de pedidos no backlog após adição: ${pedidosBacklogPorStatus.length}`);

        
        // 🗑️ REMOVER LHs LIXO DO OBJETO lhTrips (para não aparecerem na aba LH Trips)
        lhsLixoSet.forEach(lhTrip => {
            if (lhTrips[lhTrip]) {
                console.log(`🗑️ Removendo LH ${lhTrip} do objeto lhTrips`);
                delete lhTrips[lhTrip];
            }
            if (lhTripsPlanejáveis[lhTrip]) {
                console.log(`🗑️ Removendo LH ${lhTrip} do objeto lhTripsPlanejáveis`);
                delete lhTripsPlanejáveis[lhTrip];
            }
        });
    }

    const totalAntesFiltro = dadosTabela.length;

    // Aplicar filtros Excel
    dadosTabela = aplicarFiltrosExcel(dadosTabela, todasColunasPlan, filtrosAtivosPlan);

    // Ordenar por FIFO/FEFO se não houver ordenação específica
    const temOrdenacao = Object.keys(filtrosAtivosPlan).some(col => filtrosAtivosPlan[col]?.ordenacao);
    if (!temOrdenacao) {
        dadosTabela.sort((a, b) => {
            // Verificar se está bloqueada (status P3 - fora do prazo)
            const aBloqueada = a.status_lh?.codigo === 'P3';
            const bBloqueada = b.status_lh?.codigo === 'P3';
            
            // LHs bloqueadas (P3) vão para o final, independente de serem FULL
            if (aBloqueada && !bBloqueada) return 1;
            if (!aBloqueada && bBloqueada) return -1;
            
            // PRIORIDADE ABSOLUTA: LHs FULL não bloqueadas ficam no topo (independente de CAP)
            // Isso garante que LHs FBS sempre fiquem no topo (exceto se P3)
            const aFullPrioritaria = a.isFull && !aBloqueada;
            const bFullPrioritaria = b.isFull && !bBloqueada;
            
            if (aFullPrioritaria && !bFullPrioritaria) return -1;
            if (!aFullPrioritaria && bFullPrioritaria) return 1;
            
            // Segundo: Ordenar por data/hora de previsão (FIFO/FEFO - mais antiga primeiro)
            const aPrevisao = a.previsao_data && a.previsao_hora && a.previsao_data !== '-' && a.previsao_hora !== '-'
                ? `${a.previsao_data} ${a.previsao_hora}`
                : null;
            const bPrevisao = b.previsao_data && b.previsao_hora && b.previsao_data !== '-' && b.previsao_hora !== '-'
                ? `${b.previsao_data} ${b.previsao_hora}`
                : null;
            
            if (aPrevisao && bPrevisao) {
                // Converter para timestamp para comparação
                const aTimestamp = parsePrevisaoParaTimestamp(aPrevisao);
                const bTimestamp = parsePrevisaoParaTimestamp(bPrevisao);
                
                if (aTimestamp && bTimestamp) {
                    return aTimestamp - bTimestamp; // Mais antiga primeiro
                }
            }
            
            // Se uma tem previsão e outra não, priorizar a que tem
            if (aPrevisao && !bPrevisao) return -1;
            if (!aPrevisao && bPrevisao) return 1;
            
            // Terceiro: por quantidade de pedidos (maior primeiro)
            return b.pedidos - a.pedidos;
        });
    }

    // Contabilizar
    let encontradas = 0;
    let naoEncontradas = 0;
    let totalPedidos = 0;

    dadosTabela.forEach(row => {
        if (row.encontrada) encontradas++;
        else naoEncontradas++;
        totalPedidos += row.pedidos;
    });

    // GERAR CABEÇALHO DINÂMICO COM FILTROS EXCEL
    if (thead) {
        let headerHtml = '<tr>';
        
        // Coluna de checkbox
        headerHtml += `<th style="width: 50px;">
            <input type="checkbox" id="checkTodasLHsPlan" title="Selecionar todas" 
                   onchange="toggleTodasLHsPlan(this.checked)">
        </th>`;
        
        // Colunas fixas
        const nomesColunas = {
            'tipo': 'Tipo',
            'status': 'Status', 
            'lh_trip': 'LH Trip',
            'pedidos': 'Pedidos',
            'pedidos_tos': 'Pedidos TOs',
        };
        
        colunasFixas.forEach(col => {
            const temFiltro = filtrosAtivosPlan[col];
            const icone = temFiltro ? '🔽' : '▼';
            const classeAtivo = temFiltro ? 'filtro-ativo' : '';
            headerHtml += `<th class="${classeAtivo}">
                <div class="th-content">
                    <span class="th-titulo">${nomesColunas[col]}</span>
                    <button class="btn-filtro-excel-plan" data-coluna="${col}">${icone}</button>
                </div>
            </th>`;
        });
        
        // Colunas da planilha
        const nomesColunasPlanilha = {
            'status_lh': 'Status LH',
            'origin': 'Origin',
            'destination': 'Destination',
            'previsao_data': 'Previsão Data',
            'previsao_hora': 'Previsão Hora',
            'tempo_corte': 'Tempo p/ Corte',
            'is_full': 'Is Full',
            'previsao_final': 'Previsão Final'
        };
        
        colunasPlanejamento.forEach(col => {
            const temFiltro = filtrosAtivosPlan[col];
            const icone = temFiltro ? '🔽' : '▼';
            const classeAtivo = temFiltro ? 'filtro-ativo' : '';
            const nomeExibicao = nomesColunasPlanilha[col] || col.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase());
            headerHtml += `<th class="${classeAtivo}">
                <div class="th-content">
                    <span class="th-titulo">${nomeExibicao}</span>
                    <button class="btn-filtro-excel-plan" data-coluna="${col}">${icone}</button>
                </div>
            </th>`;
        });
        
        headerHtml += '</tr>';
        thead.innerHTML = headerHtml;
    }

    // Gerar HTML da tabela
    let html = '';

    dadosTabela.forEach(row => {
        const selecionada = lhsSelecionadasPlan.has(row.lh_trip);
        // Adicionar classe especial para LHs dentro do limite de corte
        // LHs FULL também recebem o destaque azul (aderentes)
        const dentroLimite = row.tempo_corte?.dentroLimite || row.isFull;
        
        // Verificar se a LH está bloqueada (status P3 - fora do prazo)
        const statusLH = row.status_lh;
        
        // IMPORTANTE: Se foi validado pelo SPX, NUNCA bloqueia
        const lhBloqueada = statusLH && !statusLH._spxValidado && statusLH.isBloqueada;
        const motivoBloqueio = lhBloqueada ? 'LH em trânsito - não chegará a tempo para este ciclo' : '';
        
        // DEBUG: Log do status na renderização
        if (statusLH && (statusLH.codigo === 'P2' || statusLH.codigo === 'P3' || statusLH._spxValidado)) {
            console.log(`🖥️ RENDERIZAÇÃO - LH: ${row.lh_trip}, Status: ${statusLH.codigo} (${statusLH.texto}), SPX Validado: ${!!statusLH._spxValidado}, Bloqueada: ${lhBloqueada}`);
        }
        
        // Verificar se é LH com estouro no piso (para destaque especial)
        const comEstouroPiso = row.estouroPiso || false;
        const tituloEstouro = comEstouroPiso ? '🟡 LH no piso - sugerida para ajuste por TOs' : '';
        
        // Verificar se é LH candidata para complemento de CAP via TOs
        const candidataComplemento = window.lhCandidataParaTOs && 
                                      window.lhCandidataParaTOs.lhTrip === row.lh_trip;
        const tituloComplemento = candidataComplemento ? '💚 Próxima LH FIFO - pode ser usada para completar CAP via TOs' : '';
        
        html += `<tr class="${row.isFull ? 'lh-full' : ''} ${selecionada ? 'row-selecionada' : ''} ${dentroLimite ? 'lh-dentro-limite' : ''} ${lhBloqueada ? 'lh-bloqueada' : ''} ${comEstouroPiso ? 'lh-estouro-piso' : ''} ${candidataComplemento ? 'lh-candidata-complemento' : ''}" ${lhBloqueada ? `title="🔒 ${motivoBloqueio}"` : (candidataComplemento ? `title="${tituloComplemento}"` : (comEstouroPiso ? `title="${tituloEstouro}"` : ''))}>`;
        
        // Coluna de checkbox
        html += `<td>
            ${lhBloqueada ? '<span class="icone-cadeado" title="🔒 LH bloqueada - não pode ser planejada">🔒</span>' : ''}
            <input type="checkbox" class="checkbox-lh-plan" data-lh="${row.lh_trip}" 
                   ${selecionada ? 'checked' : ''} 
                   ${lhBloqueada ? 'disabled' : ''}
                   onchange="toggleSelecaoLH('${row.lh_trip}', this.checked)">
        </td>`;
        
        // Adicionar ícone ⚡ para LHs FBS (Full by Shopee)
        let valorOriginBadge = null;
        if (row.dadosPlanilhaLH) {
            valorOriginBadge = row.dadosPlanilhaLH.origin || row.dadosPlanilhaLH.Origin || row.dadosPlanilhaLH.ORIGIN || 
                               row.dadosPlanilhaLH.origem || row.dadosPlanilhaLH.Origem;
            if (!valorOriginBadge) {
                const chaveOrigin = Object.keys(row.dadosPlanilhaLH).find(k => k.toLowerCase() === 'origin');
                if (chaveOrigin) valorOriginBadge = row.dadosPlanilhaLH[chaveOrigin];
            }
        }
        const isFBS = valorOriginBadge && typeof valorOriginBadge === 'string' && 
                      valorOriginBadge.toUpperCase().startsWith('FBS_');
        const badgeFull = isFBS ? '<span class="badge-full">⚡ FULL</span>' : '<span class="badge-full">⭐ FULL</span>';
        html += `<td>${row.isFull ? badgeFull : '<span class="badge-normal">Normal</span>'}</td>`;
        html += `<td class="${row.encontrada ? 'status-encontrada' : 'status-nao-encontrada'}">${row.encontrada ? '✅' : '❌'}</td>`;
        html += `<td class="lh-trip-cell">${row.lh_trip}</td>`;
        html += `<td>${row.pedidos}</td>`;
        
        // Bloquear seleção de TOs para LHs bloqueadas
        if (lhBloqueada) {
            html += `<td class="pedidos-tos-cell bloqueada" title="🔒 LH bloqueada - não pode selecionar TOs">${row.pedidos_tos > 0 ? '🔹 ' : ''}${row.pedidos_tos || 0}</td>`;
        } else {
            html += `<td class="pedidos-tos-cell ${row.pedidos_tos > 0 ? 'tem-tos' : ''}" onclick="abrirModalTOs('${row.lh_trip}')" title="${row.pedidos_tos > 0 ? 'TOs selecionadas' : 'Clique para selecionar TOs'}">${row.pedidos_tos > 0 ? '🔹 ' : ''}${row.pedidos_tos || 0}</td>`;
        }

        colunasPlanejamento.forEach(col => {
            let valor = row[col];
            
            // Tratamento especial para status_lh
            if (col === 'status_lh') {
                const statusLH = row['status_lh'];
                if (statusLH && typeof statusLH === 'object') {
                    valor = `<span class="badge-status-lh ${statusLH.classe}">${statusLH.icone} ${statusLH.texto}</span>`;
                } else {
                    valor = '-';
                }
            // Tratamento especial para tempo_corte
            } else if (col === 'tempo_corte') {
                const tempoCorte = row['tempo_corte'];
                const statusLH = row['status_lh'];
                
                // ✅ MOSTRAR TEMPO APENAS SE STATUS FOR "EM TRÂNSITO" (P2 ou P3)
                const isEmTransito = statusLH && (statusLH.codigo === 'P2' || statusLH.codigo === 'P3');
                
                if (isEmTransito && tempoCorte && typeof tempoCorte === 'object') {
                    // Definir classe baseado no status
                    let classe = 'tempo-corte-neutro';
                    if (tempoCorte.status === 'ok') {
                        classe = 'tempo-corte-ok';
                    } else if (tempoCorte.status === 'atencao') {
                        classe = 'tempo-corte-atencao';
                    } else if (tempoCorte.status === 'urgente') {
                        classe = 'tempo-corte-urgente';
                    } else if (tempoCorte.status === 'atrasado') {
                        classe = 'tempo-corte-atrasado';
                    } else if (tempoCorte.status === 'ciclo-encerrado') {
                        classe = 'tempo-corte-encerrado';
                    }
                    
                    // Adicionar tooltip se existir
                    const tooltip = tempoCorte.tooltip ? `title="${tempoCorte.tooltip}"` : '';
                    valor = `<span class="badge-tempo-corte ${classe}" ${tooltip}>${tempoCorte.texto}</span>`;
                } else {
                    // Para outros status (No Piso, Sinalizar Inventário, etc.), mostrar "-"
                    valor = '-';
                }
            } else if (col === 'previsao_data' || col === 'previsao_hora') {
                // Já está no formato correto
                valor = valor || '-';
            } else if (col.includes('eta') || col.includes('date') || col.includes('previsao') || col.includes('datetime')) {
                valor = formatarData(valor);
            } else {
                valor = truncarTexto(String(valor), 30);
            }
            
            html += `<td>${valor}</td>`;
        });

        html += '</tr>';
    });

    if (html === '') {
        const totalColunas = todasColunasPlan.length + 1; // +1 para checkbox
        tbody.innerHTML = `
            <tr>
                <td colspan="${totalColunas}" style="text-align: center; padding: 40px; color: #999;">
                    <h3>Nenhum resultado encontrado</h3>
                    <p>Tente ajustar os filtros</p>
                </td>
            </tr>
        `;
    } else {
        tbody.innerHTML = html;
    }

    // Atualizar estatísticas
    atualizarEstatisticasPlanejamento(dadosTabela.length, encontradas, naoEncontradas, totalPedidos);
    
    // Info de filtros
    const tabelaContainer = document.querySelector('.planejamento-table-container');
    let infoExistente = tabelaContainer?.querySelector('.tabela-info-plan');
    if (infoExistente) infoExistente.remove();
    
    if (tabelaContainer && (dadosTabela.length !== totalAntesFiltro || Object.keys(filtrosAtivosPlan).length > 0)) {
        const infoDiv = document.createElement('div');
        infoDiv.className = 'tabela-info tabela-info-plan';
        infoDiv.innerHTML = `
            <span>🔍 Mostrando: ${dadosTabela.length.toLocaleString('pt-BR')} de ${totalAntesFiltro.toLocaleString('pt-BR')}</span>
            ${Object.keys(filtrosAtivosPlan).length > 0 ? '<button class="btn-limpar-todos-filtros" onclick="limparTodosFiltrosPlan()">🗑️ Limpar Filtros</button>' : ''}
        `;
        tabelaContainer.appendChild(infoDiv);
    }
    
    // Adicionar event listeners nos botões de filtro Excel
    document.querySelectorAll('.btn-filtro-excel-plan').forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            const coluna = btn.dataset.coluna;
            
            // Coletar valores da coluna (precisa recalcular dados originais)
            let valoresColuna = [];
            lhsDoSPX.forEach(lhTrip => {
                const qtdPedidos = lhTrips[lhTrip]?.length || 0;
                // ✅ USAR FUNÇÃO DE FILTRO POR STATION
                const dadosPlanilhaLH = buscarDadosPlanilhaPorStation(lhTrip);
                const encontrada = !!dadosPlanilhaLH;
                // Identificar LHs FBS (Full by Shopee) pela coluna ORIGIN (case-insensitive)
                let valorOrigin = null;
                if (dadosPlanilhaLH) {
                    valorOrigin = dadosPlanilhaLH.origin || dadosPlanilhaLH.Origin || dadosPlanilhaLH.ORIGIN || 
                                  dadosPlanilhaLH.origem || dadosPlanilhaLH.Origem;
                    if (!valorOrigin) {
                        const chaveOrigin = Object.keys(dadosPlanilhaLH).find(k => k.toLowerCase() === 'origin');
                        if (chaveOrigin) valorOrigin = dadosPlanilhaLH[chaveOrigin];
                    }
                }
                const isFBS = valorOrigin && typeof valorOrigin === 'string' && 
                              valorOrigin.toUpperCase().startsWith('FBS_');
                const isFull = isFBS || 
                               dadosPlanilhaLH?.is_full === 'Full' || dadosPlanilhaLH?.is_full === 'Sim' ||
                               dadosPlanilhaLH?.is_full_truck === 'Full' || dadosPlanilhaLH?.tipo_carga === 'Full';
                
                if (coluna === 'tipo') {
                    valoresColuna.push(isFull ? 'FULL' : 'Normal');
                } else if (coluna === 'status') {
                    valoresColuna.push(encontrada ? 'Encontrada' : 'Não encontrada');
                } else if (coluna === 'lh_trip') {
                    valoresColuna.push(lhTrip);
                } else if (coluna === 'pedidos') {
                    valoresColuna.push(String(qtdPedidos));
                } else if (coluna === 'status_lh') {
                    // Status calculado dinamicamente
                    // ✅ PASSAR CICLO SELECIONADO para calcular status P3 corretamente
                    const statusObj = calcularStatusLH(dadosPlanilhaLH, null, null, cicloSelecionado);
                    valoresColuna.push(statusObj.texto);
                } else if (coluna === 'tempo_corte') {
                    // Tempo de corte calculado dinamicamente
                    const tempoCorte = calcularTempoCorte(dadosPlanilhaLH, cicloSelecionado);
                    valoresColuna.push(tempoCorte.texto || '-');
                } else if (coluna === 'previsao_data' || coluna === 'previsao_hora') {
                    // Previsão separada
                    const previsaoFinal = dadosPlanilhaLH?.previsao_final || '';
                    if (previsaoFinal) {
                        const { data, hora } = formatarPrevisaoFinal(previsaoFinal);
                        valoresColuna.push(coluna === 'previsao_data' ? data : hora);
                    } else {
                        valoresColuna.push('-');
                    }
                } else {
                    valoresColuna.push(dadosPlanilhaLH?.[coluna] || '-');
                }
            });
            
            criarPopupFiltroExcel(coluna, valoresColuna, btn, 'planejamento');
        });
    });
}

// Limpar filtros do Planejamento (já definida acima como limparTodosFiltrosPlan)

// Atualizar estatísticas do planejamento
function atualizarEstatisticasPlanejamento(total, encontradas, naoEncontradas, totalPedidos) {
    const totalEl = document.getElementById('statTotalLHs');
    const encontradasEl = document.getElementById('statLHsEncontradas');
    const naoEncontradasEl = document.getElementById('statLHsNaoEncontradas');
    const totalPedidosEl = document.getElementById('statTotalPedidos');

    if (totalEl) totalEl.textContent = total.toLocaleString('pt-BR');
    if (encontradasEl) encontradasEl.textContent = encontradas.toLocaleString('pt-BR');
    if (naoEncontradasEl) naoEncontradasEl.textContent = naoEncontradas.toLocaleString('pt-BR');
    if (totalPedidosEl) totalPedidosEl.textContent = (totalPedidos || 0).toLocaleString('pt-BR');
    
    // ✅ ATUALIZAR CARD DE BACKLOG
    const backlogEl = document.getElementById('statBacklogPedidos');
    if (backlogEl) {
        const totalBacklog = pedidosBacklogPorStatus?.length || 0;
        backlogEl.textContent = totalBacklog.toLocaleString('pt-BR');
    }
    
    // Atualizar contador de selecionadas
    atualizarContadorSelecaoLHs();
}

// ======================= SELEÇÃO DE LHs NO PLANEJAMENTO =======================

// Atualizar contador de LHs selecionadas
function atualizarContadorSelecaoLHs() {
    const selecionadasEl = document.getElementById('statLHsSelecionadas');
    const infoEl = document.getElementById('selecaoInfo');
    
    const qtdLHs = lhsSelecionadasPlan.size;
    const qtdBacklog = pedidosBacklogSelecionados.size;
    let qtdPedidosLHs = 0;
    let lhsComTOsParciais = 0;
    
    lhsSelecionadasPlan.forEach(lh => {
        // Verificar se tem TOs parciais selecionadas
        if (tosSelecionadasPorLH[lh] && tosSelecionadasPorLH[lh].size > 0) {
            // Usar apenas pedidos das TOs selecionadas
            qtdPedidosLHs += calcularTotalTOsSelecionadas(lh);
            lhsComTOsParciais++;
        } else {
            // Usar todos os pedidos da LH
            qtdPedidosLHs += lhTrips[lh]?.length || 0;
        }
    });
    
    const totalPedidos = qtdPedidosLHs + qtdBacklog;
    
    // Mostrar total de PEDIDOS selecionados no card (não quantidade de LHs)
    if (selecionadasEl) selecionadasEl.textContent = totalPedidos.toLocaleString('pt-BR');
    
    // Mostrar info completa incluindo backlog e TOs parciais
    if (infoEl) {
        let info = `${qtdLHs} LHs`;
        if (lhsComTOsParciais > 0) {
            info += ` (${lhsComTOsParciais} parcial)`;
        }
        info += ` = ${qtdPedidosLHs.toLocaleString('pt-BR')} pedidos`;
        
        if (qtdBacklog > 0) {
            info += ` + ${qtdBacklog.toLocaleString('pt-BR')} backlog = ${totalPedidos.toLocaleString('pt-BR')} total`;
        }
        infoEl.textContent = info;
    }
}

// Selecionar todas as LHs visíveis
function selecionarTodasLHsPlanejamento() {
    const checkboxes = document.querySelectorAll('.checkbox-lh-plan');
    checkboxes.forEach(cb => {
        cb.checked = true;
        lhsSelecionadasPlan.add(cb.dataset.lh);
    });
    atualizarContadorSelecaoLHs();
}

// Limpar seleção de LHs
function limparSelecaoLHsPlanejamento() {
    lhsSelecionadasPlan.clear();
    pedidosBacklogSelecionados.clear();
    backlogConfirmado = false;
    
    const checkboxes = document.querySelectorAll('.checkbox-lh-plan');
    checkboxes.forEach(cb => cb.checked = false);
    atualizarContadorSelecaoLHs();
    
    // Remover info de sugestão
    const infoSugestao = document.querySelector('.sugestao-info');
    if (infoSugestao) infoSugestao.remove();
    
    // Atualizar backlog
    renderizarBacklog();
}

// Toggle seleção de uma LH
function toggleSelecaoLH(lhTrip, checked) {
    if (checked) {
        // ✅ VALIDAR SE A LH CHEGA A TEMPO ANTES DE PERMITIR SELEÇÃO
        const dadosLH = buscarDadosPlanilhaPorStation(lhTrip);
        if (dadosLH && cicloSelecionado && cicloSelecionado !== 'Todos') {
            const tempoCorte = calcularTempoCorte(dadosLH, cicloSelecionado);
            
            // Se minutosCorte < 0, significa que não chegará a tempo
            if (tempoCorte.minutos !== null && tempoCorte.minutos < 0) {
                alert(`⚠️ LH ${lhTrip} não pode ser selecionada!\n\n` +
                      `Esta LH não chegará a tempo para o ciclo ${cicloSelecionado}.\n` +
                      `Tempo de corte: ${tempoCorte.minutos} minutos (já passou do horário limite).\n\n` +
                      `Selecione apenas LHs que chegarão antes do horário de corte.`);
                
                // Desmarcar checkbox
                const checkbox = document.querySelector(`input[data-lh="${lhTrip}"]`);
                if (checkbox) checkbox.checked = false;
                return;
            }
        }
        
        lhsSelecionadasPlan.add(lhTrip);
    } else {
        lhsSelecionadasPlan.delete(lhTrip);
        // Também remover TOs selecionadas desta LH
        delete tosSelecionadasPorLH[lhTrip];
    }
    atualizarContadorSelecaoLHs();
    
    // Atualizar visual da linha
    const row = document.querySelector(`input[data-lh="${lhTrip}"]`)?.closest('tr');
    if (row) {
        row.classList.toggle('row-selecionada', checked);
    }
}

// Toggle todas as LHs visíveis na tabela
function toggleTodasLHsPlan(checked) {
    const checkboxes = document.querySelectorAll('.checkbox-lh-plan');
    let lhsExcluidas = [];
    
    checkboxes.forEach(cb => {
        const lh = cb.dataset.lh;
        
        if (checked) {
            // ✅ VALIDAR SE A LH CHEGA A TEMPO
            const dadosLH = dadosPlanilha[lh];
            if (dadosLH && cicloSelecionado && cicloSelecionado !== 'Todos') {
                const tempoCorte = calcularTempoCorte(dadosLH, cicloSelecionado);
                
                // Se minutosCorte < 0, não selecionar
                if (tempoCorte.minutos !== null && tempoCorte.minutos < 0) {
                    lhsExcluidas.push(lh);
                    cb.checked = false;
                    return; // Pular esta LH
                }
            }
            
            lhsSelecionadasPlan.add(lh);
            cb.checked = true;
        } else {
            lhsSelecionadasPlan.delete(lh);
            cb.checked = false;
            // Também remover TOs selecionadas
            delete tosSelecionadasPorLH[lh];
        }
        
        cb.closest('tr')?.classList.toggle('row-selecionada', checked);
    });
    
    // Mostrar alerta se houver LHs excluídas
    if (checked && lhsExcluidas.length > 0) {
        alert(`⚠️ ${lhsExcluidas.length} LH(s) não foram selecionadas!\n\n` +
              `Estas LHs não chegarão a tempo para o ciclo ${cicloSelecionado}:\n` +
              lhsExcluidas.join(', ') + '\n\n' +
              `Apenas LHs que chegam antes do horário de corte podem ser selecionadas.`);
    }
    
    atualizarContadorSelecaoLHs();
}

// ======================= ABA BACKLOG (PEDIDOS SEM LH) =======================

// Renderizar aba de Backlog
function renderizarBacklog() {
    const tbody = document.getElementById('tbodyBacklog');
    const thead = document.getElementById('theadBacklog');
    if (!tbody) return;
    
    // USAR TODOS OS PEDIDOS COM STATUS DE BACKLOG (LMHub_Received, Return_LMHub_Received)
    // Inclui pedidos COM LH e SEM LH que tenham esses status
    const pedidosBacklog = pedidosBacklogPorStatus || [];
    
    // Separar backlog por tipo para estatísticas
    const colunaLH = Object.keys(pedidosBacklog[0] || {}).find(col =>
        col.toLowerCase().includes('lh trip') ||
        col.toLowerCase().includes('lh_trip')
    ) || 'LH Trip';
    
    const backlogComLH = pedidosBacklog.filter(p => p[colunaLH] && p[colunaLH].trim() !== '');
    const backlogSemLH = pedidosBacklog.filter(p => !p[colunaLH] || p[colunaLH].trim() === '');
    const lhsNoBacklog = [...new Set(backlogComLH.map(p => p[colunaLH]))].length;
    
    // Atualizar estatísticas
    const statTotal = document.getElementById('statBacklogTotal');
    const statSelecionados = document.getElementById('statBacklogSelecionados');
    const backlogInfo = document.getElementById('backlogInfo');
    
    if (statTotal) statTotal.textContent = pedidosBacklog.length.toLocaleString('pt-BR');
    if (backlogInfo) {
        let info = `${pedidosBacklog.length} pedidos de backlog`;
        if (lhsNoBacklog > 0) info += ` (${lhsNoBacklog} LHs)`;
        if (backlogSemLH.length > 0) info += ` | ${backlogSemLH.length} sem LH`;
        backlogInfo.textContent = info;
    }
    
    if (pedidosBacklog.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="20" style="text-align: center; padding: 40px; color: #999;">
                    <div style="font-size: 48px; margin-bottom: 15px;">✅</div>
                    <h3>Nenhum backlog encontrado</h3>
                    <p>Não há pedidos com status LMHub_Received ou Return_LMHub_Received</p>
                    <p style="font-size: 12px; color: #aaa; margin-top: 10px;">
                        Pedidos com esses status são automaticamente identificados como backlog
                    </p>
                </td>
            </tr>
        `;
        return;
    }
    
    // Pegar TODAS as colunas do primeiro pedido
    const colunas = Object.keys(pedidosBacklog[0]);
    
    // Gerar cabeçalho
    let headerHtml = '<tr>';
    headerHtml += `<th style="width: 50px; position: sticky; left: 0; background: #f8f9fa; z-index: 11;">
        <input type="checkbox" id="checkTodosBacklog" title="Selecionar todos" 
               ${pedidosBacklogSelecionados.size === pedidosBacklog.length ? 'checked' : ''}>
    </th>`;
    
    colunas.forEach(col => {
        const temFiltro = filtrosAtivosBacklog?.[col];
        const icone = temFiltro ? '🔽' : '▼';
        const classeAtivo = temFiltro ? 'filtro-ativo' : '';
        headerHtml += `<th class="${classeAtivo}">
            <div class="th-content">
                <span class="th-titulo">${col}</span>
                <button class="btn-filtro-excel-backlog" data-coluna="${col}">${icone}</button>
            </div>
        </th>`;
    });
    headerHtml += '</tr>';
    thead.innerHTML = headerHtml;
    
    // Aplicar filtros
    let pedidosFiltrados = aplicarFiltrosExcel(pedidosBacklog, colunas, filtrosAtivosBacklog || {});
    
    // Gerar linhas
    let html = '';
    
    // ✅ SEÇÃO DE LHS LIXO SISTÊMICO
    if (lhsLixoSistemico.length > 0) {
        const totalPedidosLixo = lhsLixoSistemico.reduce((sum, row) => sum + (row.pedidos || 0), 0);
        
        html += `
            <tr style="background: #fff3cd; font-weight: bold;">
                <td colspan="${colunas.length + 1}" style="padding: 15px; text-align: center;">
                    🗑️ LHs Lixo Sistêmico (${lhsLixoSistemico.length} LHs, ${totalPedidosLixo} pedidos)
                    <span style="font-size: 0.9em; font-weight: normal; color: #856404; display: block; margin-top: 5px;">
                        Sem origin/destination/previsão - Pedidos já incluídos no backlog
                    </span>
                </td>
            </tr>
        `;
        
        // Buscar os pedidos dessas LHs para exibir detalhes
        const lhsLixoSet = new Set(lhsLixoSistemico.map(row => row.lh_trip));
        const colunaLH = colunas.find(c => 
            c.toLowerCase().includes('lh trip') || 
            c.toLowerCase().includes('lh_trip')
        ) || 'LH Trip';
        
        // Agrupar pedidos por LH
        const pedidosPorLH = {};
        pedidosBacklog.forEach(pedido => {
            const lhTrip = pedido[colunaLH];
            if (lhTrip && lhsLixoSet.has(lhTrip)) {
                if (!pedidosPorLH[lhTrip]) {
                    pedidosPorLH[lhTrip] = [];
                }
                pedidosPorLH[lhTrip].push(pedido);
            }
        });
        
        // Exibir pedidos agrupados por LH
        Object.keys(pedidosPorLH).forEach(lhTrip => {
            const pedidosLH = pedidosPorLH[lhTrip];
            
            // Linha de cabeçalho da LH
            html += `
                <tr style="background: #fffaeb; font-weight: bold;">
                    <td colspan="${colunas.length + 1}" style="padding: 10px; padding-left: 30px;">
                        🗑️ ${lhTrip} (${pedidosLH.length} pedidos)
                    </td>
                </tr>
            `;
            
            // Exibir primeiros 3 pedidos como exemplo
            pedidosLH.slice(0, 3).forEach(pedido => {
                html += `<tr style="background: #fffdf5;">`;
                html += `<td style="position: sticky; left: 0; background: #fffdf5; z-index: 1;">
                    <span style="color: #ccc; font-size: 0.8em;">●</span>
                </td>`;
                
                colunas.forEach(col => {
                    const valor = pedido[col] || '-';
                    html += `<td style="font-size: 0.9em; color: #666;">${valor}</td>`;
                });
                
                html += `</tr>`;
            });
            
            // Se tiver mais de 3, mostrar quantos foram ocultados
            if (pedidosLH.length > 3) {
                html += `
                    <tr style="background: #fffdf5;">
                        <td colspan="${colunas.length + 1}" style="padding: 5px; padding-left: 50px; font-size: 0.85em; color: #999; font-style: italic;">
                            ... e mais ${pedidosLH.length - 3} pedidos
                        </td>
                    </tr>
                `;
            }
        });
    }
    
    pedidosFiltrados.forEach((pedido, index) => {
        // Usar mesma função de ID em todos os lugares
        const id = getShipmentIdFromPedido(pedido, index);
        const selecionado = pedidosBacklogSelecionados.has(id);
        
        html += `<tr class="${selecionado ? 'row-selecionada' : ''}">`;
        html += `<td style="position: sticky; left: 0; background: ${selecionado ? '#e8f5e9' : 'white'}; z-index: 1;">
            <input type="checkbox" class="checkbox-backlog" data-id="${id}" 
                   ${selecionado ? 'checked' : ''} onchange="toggleSelecaoBacklog('${id}', this.checked)">
        </td>`;
        
        colunas.forEach(col => {
            html += `<td>${pedido[col] || '-'}</td>`;
        });
        html += '</tr>';
    });
    
    tbody.innerHTML = html;
    
    // Atualizar contador
    atualizarContadorBacklog();
    
    // Re-adicionar listener do checkbox geral
    document.getElementById('checkTodosBacklog')?.addEventListener('change', toggleTodosBacklog);
    
    // Adicionar listeners dos filtros
    document.querySelectorAll('.btn-filtro-excel-backlog').forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            const coluna = btn.dataset.coluna;
            const valoresColuna = pedidosBacklogPorStatus.map(p => p[coluna]);
            criarPopupFiltroExcel(coluna, valoresColuna, btn, 'backlog');
        });
    });
}

// Estado de filtros do Backlog
let filtrosAtivosBacklog = {};

// Atualizar contador de backlog selecionados
function atualizarContadorBacklog() {
    const statSelecionados = document.getElementById('statBacklogSelecionados');
    const infoEl = document.getElementById('backlogSelecaoInfo');
    
    if (statSelecionados) statSelecionados.textContent = pedidosBacklogSelecionados.size.toLocaleString('pt-BR');
    if (infoEl) infoEl.textContent = `${pedidosBacklogSelecionados.size} pedidos selecionados`;
}

// Toggle seleção de pedido do backlog
function toggleSelecaoBacklog(id, checked) {
    if (checked) {
        pedidosBacklogSelecionados.add(id);
    } else {
        pedidosBacklogSelecionados.delete(id);
    }
    atualizarContadorBacklog();
    
    // Atualizar visual da linha
    const row = document.querySelector(`input[data-id="${id}"]`)?.closest('tr');
    if (row) {
        row.classList.toggle('row-selecionada', checked);
    }
    
    // Atualizar checkbox "selecionar todos" baseado nos visíveis
    atualizarCheckboxTodosBacklog();
}

// Atualizar estado do checkbox "selecionar todos" baseado nos itens visíveis
function atualizarCheckboxTodosBacklog() {
    const checkAll = document.getElementById('checkTodosBacklog');
    if (!checkAll) return;
    
    const checkboxesVisiveis = document.querySelectorAll('.checkbox-backlog');
    if (checkboxesVisiveis.length === 0) {
        checkAll.checked = false;
        return;
    }
    
    const todosMarcados = [...checkboxesVisiveis].every(cb => cb.checked);
    checkAll.checked = todosMarcados;
}

// Toggle todos os pedidos VISÍVEIS do backlog (respeitando filtros)
function toggleTodosBacklog(e) {
    const checked = e?.target?.checked ?? true;
    
    // Selecionar apenas os checkboxes VISÍVEIS na tabela (já filtrados)
    const checkboxesVisiveis = document.querySelectorAll('.checkbox-backlog');
    
    checkboxesVisiveis.forEach(cb => {
        const id = cb.dataset.id;
        cb.checked = checked;
        
        if (checked) {
            pedidosBacklogSelecionados.add(id);
        } else {
            pedidosBacklogSelecionados.delete(id);
        }
        
        cb.closest('tr')?.classList.toggle('row-selecionada', checked);
    });
    
    atualizarContadorBacklog();
}

// Selecionar todos do backlog (visíveis)
function selecionarTodosBacklog() {
    const checkAll = document.getElementById('checkTodosBacklog');
    if (checkAll) checkAll.checked = true;
    toggleTodosBacklog({ target: { checked: true } });
}

// Limpar seleção do backlog (visíveis)
function limparSelecaoBacklog() {
    const checkAll = document.getElementById('checkTodosBacklog');
    if (checkAll) checkAll.checked = false;
    toggleTodosBacklog({ target: { checked: false } });
}

// Confirmar backlog e voltar ao planejamento
function confirmarBacklog() {
    backlogConfirmado = true;
    trocarAba('planejamento');
    
    // Mostrar mensagem
    const qtd = pedidosBacklogSelecionados.size;
    if (qtd > 0) {
        alert(`✅ ${qtd} pedidos do backlog serão incluídos no planejamento.`);
    }
}

// ======================= GERAÇÃO DO ARQUIVO DE PLANEJAMENTO =======================

// Iniciar processo de geração do planejamento
function iniciarGeracaoPlanejamento() {
    // Verificar se tem LHs selecionadas
    if (lhsSelecionadasPlan.size === 0) {
        alert('⚠️ Selecione pelo menos uma LH para gerar o planejamento.');
        return;
    }
    
    // Verificar se tem pedidos de backlog
    const totalBacklog = pedidosBacklogPorStatus.length;
    
    if (totalBacklog > 0 && !backlogConfirmado) {
        const resposta = confirm(
            `📦 Existem ${totalBacklog} pedidos de Backlog (status LMHub_Received ou Return_LMHub_Received).\n\n` +
            `Deseja tratar o Backlog antes de gerar o planejamento?\n\n` +
            `• SIM - Abre a aba "Tratar Backlog" para selecionar pedidos\n` +
            `• NÃO - Gera o planejamento apenas com as LHs selecionadas`
        );
        
        if (resposta) {
            trocarAba('backlog');
            renderizarBacklog();
            return;
        }
    }
    
    // Gerar o arquivo
    gerarArquivoPlanejamento();
}

// Gerar arquivo de planejamento
async function gerarArquivoPlanejamento() {
    // Coletar todos os pedidos das LHs selecionadas (considerando TOs parciais)
    let pedidosPlanejamento = [];
    let lhsCompletas = 0;
    let lhsParciais = 0;
    
    lhsSelecionadasPlan.forEach(lh => {
        // Verificar se tem TOs parciais selecionadas
        if (tosSelecionadasPorLH[lh] && tosSelecionadasPorLH[lh].size > 0) {
            // SELEÇÃO PARCIAL: pegar apenas pedidos das TOs selecionadas
            const pedidosLH = lhTripsPlanejáveis[lh] || lhTrips[lh] || [];
            const tosSelecionadas = tosSelecionadasPorLH[lh];
            
            pedidosLH.forEach(pedido => {
                // Encontrar coluna de TO
                const colunaTO = Object.keys(pedido).find(col =>
                    col.toLowerCase().includes('to id') ||
                    col.toLowerCase().includes('to_id') ||
                    col.toLowerCase().includes('transfer order') ||
                    col.toLowerCase() === 'to'
                );
                
                const toId = pedido[colunaTO] || 'SEM_TO';
                
                // Só incluir se a TO foi selecionada
                if (tosSelecionadas.has(toId)) {
                    pedidosPlanejamento.push(pedido);
                }
            });
            
            lhsParciais++;
            console.log(`📦 LH ${lh}: ${tosSelecionadas.size} TOs selecionadas (parcial)`);
        } else {
            // SELEÇÃO COMPLETA: pegar todos os pedidos da LH
            const pedidosLH = lhTripsPlanejáveis[lh] || [];
            pedidosPlanejamento = pedidosPlanejamento.concat(pedidosLH);
            lhsCompletas++;
        }
    });
    
    console.log(`📋 LHs completas: ${lhsCompletas}, LHs parciais (TOs): ${lhsParciais}`);
    console.log(`📋 Pedidos das LHs: ${pedidosPlanejamento.length}`);
    console.log(`📦 Backlog selecionado: ${pedidosBacklogSelecionados.size} IDs`);
    
    // Adicionar pedidos do backlog selecionados
    if (pedidosBacklogSelecionados.size > 0) {
        let backlogAdicionado = 0;
        
        // Debug: mostrar alguns IDs selecionados
        const idsSelecionados = [...pedidosBacklogSelecionados];
        console.log(`📋 Primeiros 5 IDs selecionados:`, idsSelecionados.slice(0, 5));
        
        // Debug: mostrar primeiros IDs dos pedidos
        const primeirosIds = pedidosBacklogPorStatus.slice(0, 5).map((p, i) => getShipmentIdFromPedido(p, i));
        console.log(`📋 Primeiros 5 IDs dos pedidos:`, primeirosIds);
        
        pedidosBacklogPorStatus.forEach((pedido, index) => {
            // Usar mesma função de ID que foi usada na seleção
            const id = getShipmentIdFromPedido(pedido, index);
            
            if (pedidosBacklogSelecionados.has(id)) {
                // 🔥 SEMPRE criar cópia e renomear para "Backlog"
                // Todo pedido que está em pedidosBacklogPorStatus DEVE ser Backlog
                const pedidoCopia = { ...pedido };
                
                // Substituir LH Trip por "Backlog" em TODAS as variações
                if (pedidoCopia['LH Trip']) pedidoCopia['LH Trip'] = 'Backlog';
                if (pedidoCopia['LH_TRIP']) pedidoCopia['LH_TRIP'] = 'Backlog';
                if (pedidoCopia['lh_trip']) pedidoCopia['lh_trip'] = 'Backlog';
                if (pedidoCopia['LH TRIP']) pedidoCopia['LH TRIP'] = 'Backlog';
                
                pedidosPlanejamento.push(pedidoCopia);
                backlogAdicionado++;
            }
        });
        
        console.log(`✅ Backlog adicionado: ${backlogAdicionado} pedidos`);
    }
    
    console.log(`📊 TOTAL FINAL: ${pedidosPlanejamento.length} pedidos`);
    
    if (pedidosPlanejamento.length === 0) {
        alert('⚠️ Nenhum pedido encontrado para gerar o planejamento.');
        return;
    }
    
    // Gerar nome do arquivo - formato: Planejamento_DD.MM.AA
    const hoje = new Date();
    const dataFormatada = hoje.toLocaleDateString('pt-BR', {
        day: '2-digit',
        month: '2-digit',
        year: '2-digit'
    }).replace(/\//g, '.');
    
    const nomeArquivo = `Planejamento_${dataFormatada}.xlsx`;
    
    // Mostrar loading
    mostrarLoading('Gerando planejamento...', `${pedidosPlanejamento.length} pedidos`);
    
    try {
        // Preparar dados das TOs de complemento para o Excel
        const tosComplementoExcel = [];
        lhsSelecionadasPlan.forEach(lh => {
            if (tosSelecionadasPorLH[lh] && tosSelecionadasPorLH[lh].size > 0) {
                const pedidosLH = lhTrips[lh] || [];
                const tosSelecionadas = tosSelecionadasPorLH[lh];
                
                tosSelecionadas.forEach(toId => {
                    // Buscar pedidos desta TO
                    const pedidosTO = pedidosLH.filter(pedido => {
                        const colunaTO = Object.keys(pedido).find(col =>
                            col.toLowerCase().includes('to id') ||
                            col.toLowerCase().includes('to_id') ||
                            col.toLowerCase().includes('transfer order') ||
                            col.toLowerCase() === 'to'
                        );
                        return pedido[colunaTO] === toId;
                    });
                    
                    // Adicionar uma linha para cada BR (pedido)
                    pedidosTO.forEach(pedido => {
                        const colunaBR = Object.keys(pedido).find(col =>
                            col.toLowerCase().includes('tracking') ||
                            col.toLowerCase().includes('br') ||
                            col.toLowerCase().includes('parcel') ||
                            col.toLowerCase().includes('shipment') ||
                            col.toLowerCase().includes('package')
                        );
                        const brId = pedido[colunaBR] || 'N/A';
                        
                        tosComplementoExcel.push({
                            'LH TRIP': lh,
                            'TO ID': toId,
                            'BR': brId
                        });
                    });
                });
            }
        });
        
        // Chamar main process para gerar o arquivo
        const resultado = await ipcRenderer.invoke('gerar-planejamento', {
            pedidos: pedidosPlanejamento,
            nomeArquivo: nomeArquivo,
            lhsSelecionadas: Array.from(lhsSelecionadasPlan),
            qtdBacklog: pedidosBacklogSelecionados.size,
            pastaStation: pastaStationAtual, // Passar a pasta da station atual
            tosComplemento: tosComplementoExcel // ⭐ Passar TOs de complemento
        });
        
        esconderLoading();
        
        // Marcar fim da execução
        tempoFimExecucao = Date.now();
        
        if (resultado.success) {
            // Mostrar resumo detalhado
            let resumo = `✅ Planejamento gerado com sucesso!\n\n`;
            resumo += `📄 Arquivo: ${nomeArquivo}\n`;
            resumo += `📍 Local: ${resultado.filePath}\n\n`;
            resumo += `📊 Resumo:\n`;
            resumo += `• ${lhsCompletas} LHs completas\n`;
            if (lhsParciais > 0) {
                resumo += `• ${lhsParciais} LHs parciais (TOs selecionadas)\n`;
            }
            resumo += `• ${pedidosPlanejamento.length} pedidos total\n`;
            if (pedidosBacklogSelecionados.size > 0) {
                resumo += `• ${pedidosBacklogSelecionados.size} do backlog`;
            }
            
            alert(resumo);
            
            // ✅ ENVIAR LOG PARA GOOGLE SHEETS
            try {
                const dadosRelatorio = extrairDadosRelatorio();
                await enviarLogPlanejamento(dadosRelatorio);
            } catch (logError) {
                console.error('❌ Erro ao enviar log:', logError);
                // Não bloquear o fluxo se o log falhar
            }
            
            // Resetar estados
            backlogConfirmado = false;
            // ⚠️ NÃO limpar tosSelecionadasPorLH aqui!
            // O relatório HTML precisa desses dados.
            // tosSelecionadasPorLH = {};
        } else {
            alert(`❌ Erro ao gerar planejamento: ${resultado.error}`);
        }
    } catch (error) {
        esconderLoading();
        alert(`❌ Erro: ${error.message}`);
    }
}

// Função auxiliar para truncar texto
function truncarTexto(texto, maxLength) {
    if (!texto || texto === '-') return '-';
    const str = String(texto);
    return str.length > maxLength ? str.substring(0, maxLength) + '...' : str;
}

// Função para padronizar previsão final em data e hora separadas
function formatarPrevisaoFinal(previsaoFinal) {
    if (!previsaoFinal || previsaoFinal === '-') {
        return { data: '-', hora: '-' };
    }
    
    const str = String(previsaoFinal).trim();
    let data = '-';
    let hora = '-';
    
    try {
        // Verificar se é formato ISO (2026-01-08T06:00:00 ou 2026-01-08 06:00:00)
        if (str.includes('T') || (str.includes('-') && str.includes(':'))) {
            // Formato ISO: 2026-01-08T06:00:00.000 ou 2026-01-08 06:00:00.000
            const dateObj = new Date(str.replace(' ', 'T'));
            
            if (!isNaN(dateObj.getTime())) {
                // Data no formato DD/MM/YYYY
                data = dateObj.toLocaleDateString('pt-BR', {
                    day: '2-digit',
                    month: '2-digit',
                    year: 'numeric'
                });
                
                // Hora no formato HH:MM:SS
                hora = dateObj.toLocaleTimeString('pt-BR', {
                    hour: '2-digit',
                    minute: '2-digit',
                    second: '2-digit'
                });
            } else {
                // Fallback: separar manualmente
                const partes = str.replace('T', ' ').split(' ');
                if (partes.length >= 1) {
                    data = formatarDataPadrao(partes[0]);
                }
                if (partes.length >= 2) {
                    hora = formatarHoraPadrao(partes[1]);
                }
            }
        } else {
            // Formato tradicional: "09/01/2026 19:47:51" ou "9/1/2026 07:21:10"
            const partes = str.split(' ');
            
            if (partes.length >= 1) {
                data = formatarDataPadrao(partes[0]);
            }
            if (partes.length >= 2) {
                hora = formatarHoraPadrao(partes[1]);
            }
        }
    } catch (error) {
        console.error('Erro ao formatar previsão:', error, str);
        // Fallback: retornar como veio
        const partes = str.split(' ');
        data = partes[0] || '-';
        hora = partes[1] || '-';
    }
    
    return { data, hora };
}

// Padronizar data para DD/MM/YYYY
function formatarDataPadrao(dataStr) {
    if (!dataStr || dataStr === '-') return '-';
    
    // Se já está no formato DD/MM/YYYY ou D/M/YYYY
    if (dataStr.includes('/')) {
        const partes = dataStr.split('/');
        if (partes.length === 3) {
            const dia = partes[0].padStart(2, '0');
            const mes = partes[1].padStart(2, '0');
            const ano = partes[2].length === 2 ? '20' + partes[2] : partes[2];
            return `${dia}/${mes}/${ano}`;
        }
    }
    
    // Se está no formato YYYY-MM-DD
    if (dataStr.includes('-')) {
        const partes = dataStr.split('-');
        if (partes.length === 3) {
            const ano = partes[0];
            const mes = partes[1].padStart(2, '0');
            const dia = partes[2].padStart(2, '0');
            return `${dia}/${mes}/${ano}`;
        }
    }
    
    return dataStr;
}

// Padronizar hora para HH:MM:SS
function formatarHoraPadrao(horaStr) {
    if (!horaStr || horaStr === '-') return '-';
    
    // Remover milissegundos se tiver (06:00:00.000 -> 06:00:00)
    let hora = horaStr.split('.')[0];
    
    // Separar partes
    const partes = hora.split(':');
    
    if (partes.length >= 2) {
        const hh = partes[0].padStart(2, '0');
        const mm = partes[1].padStart(2, '0');
        const ss = partes[2] ? partes[2].padStart(2, '0') : '00';
        return `${hh}:${mm}:${ss}`;
    }
    
    return hora;
}

// Função auxiliar para formatar data
function formatarData(dataStr) {
    if (!dataStr || dataStr === '-') return '-';

    try {
        // Tentar parse de diferentes formatos
        let data;
        if (dataStr.includes('T')) {
            data = new Date(dataStr);
        } else if (dataStr.includes(' ')) {
            // Formato: 2025-12-21 00:30:00.000
            data = new Date(dataStr.replace(' ', 'T'));
        } else {
            return dataStr;
        }

        if (isNaN(data.getTime())) return dataStr;

        return data.toLocaleString('pt-BR', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
    } catch (e) {
        return dataStr;
    }
}

// ======================= MENU DE CONFIGURAÇÕES =======================
function trocarPainelConfig(painelId) {
    // Atualizar menu
    document.querySelectorAll('.config-menu-item').forEach(item => {
        item.classList.toggle('active', item.dataset.config === painelId);
    });

    // Atualizar painéis
    document.querySelectorAll('.config-panel').forEach(panel => {
        panel.classList.toggle('active', panel.id === `config-${painelId}`);
    });

    // Se for painel de planejamento, carregar colunas disponíveis
    if (painelId === 'planejamento') {
        carregarColunasPlanejamento();
    }
}

// Carregar colunas disponíveis da planilha
function carregarColunasPlanejamento() {
    const grid = document.getElementById('colunasPlanejamentoGrid');
    if (!grid) return;

    // Pegar colunas da primeira LH na planilha
    const primeiraLH = Object.values(dadosPlanilha)[0];

    if (!primeiraLH) {
        grid.innerHTML = '<p style="color:#999;text-align:center;grid-column:1/-1;">Atualize a planilha primeiro para ver as colunas disponíveis</p>';
        return;
    }

    const colunas = Object.keys(primeiraLH);
    const colunasSalvas = JSON.parse(localStorage.getItem('colunasPlanejamento') || '[]');

    let html = '';
    colunas.forEach((col, index) => {
        const checked = colunasSalvas.length === 0 || colunasSalvas.includes(col);
        html += `
            <div class="coluna-item">
                <input type="checkbox" id="colPlan_${index}" data-coluna="${col}" ${checked ? 'checked' : ''}>
                <label for="colPlan_${index}">${col}</label>
            </div>
        `;
    });

    grid.innerHTML = html;
}

// Selecionar todas as colunas do Planejamento
function selecionarTodasColunasPlanejamento() {
    const grid = document.getElementById('colunasPlanejamentoGrid');
    if (!grid) return;
    
    grid.querySelectorAll('input[type="checkbox"]').forEach(cb => {
        cb.checked = true;
    });
}

// Limpar todas as colunas do Planejamento
function limparTodasColunasPlanejamento() {
    const grid = document.getElementById('colunasPlanejamentoGrid');
    if (!grid) return;
    
    grid.querySelectorAll('input[type="checkbox"]').forEach(cb => {
        cb.checked = false;
    });
}

// Salvar configuração de colunas do Planejamento
function salvarConfigColunasPlanejamento() {
    const grid = document.getElementById('colunasPlanejamentoGrid');
    if (!grid) {
        alert('❌ Carregue a planilha primeiro');
        return;
    }
    
    const checkboxes = grid.querySelectorAll('input[type="checkbox"]:checked');
    const colunasSelecionadas = Array.from(checkboxes).map(cb => cb.dataset.coluna).filter(Boolean);
    
    if (colunasSelecionadas.length === 0) {
        alert('⚠️ Selecione pelo menos uma coluna');
        return;
    }
    
    localStorage.setItem('colunasPlanejamento', JSON.stringify(colunasSelecionadas));
    
    alert(`✅ Configuração salva!\n${colunasSelecionadas.length} colunas selecionadas`);
    
    // Atualizar tabela com novas colunas
    renderizarTabelaPlanejamento();
}
// ======================= SISTEMA DE TOs (TRANSFER ORDERS) =======================

// Estado das TOs
let tosSelecionadasPorLH = {}; // { lhTrip: Set([to1, to2]) }
let lhAtualModal = null; // LH sendo editada no modal

// ===== FUNÇÕES DO MODAL DE TOs =====

function abrirModalTOs(lhTrip) {
    lhAtualModal = lhTrip;
    
    // Buscar TOs da LH na planilha
    // ✅ USAR FUNÇÃO DE FILTRO POR STATION
    const dadosPlanilhaLH = buscarDadosPlanilhaPorStation(lhTrip);
    if (!dadosPlanilhaLH) {
        alert('⚠️ Esta LH não foi encontrada na planilha Google Sheets.\n\nNão é possível visualizar as TOs.');
        return;
    }
    
    // Extrair TOs dos pedidos
    const tosArray = extrairTOsDaLH(lhTrip, dadosPlanilhaLH);
    
    if (tosArray.length === 0) {
        alert('⚠️ Esta LH não possui TOs identificadas.\n\nVerifique se os pedidos possuem TO ID.');
        return;
    }
    
    // Detectar status da LH (no piso ou em trânsito)
    const statusLH = calcularStatusLH(dadosPlanilhaLH);
    const estaNoPiso = statusLH.codigo === 'P0' || statusLH.codigo === 'P0B';
    
    // Ordenar TOs por FIFO (data mais antiga primeiro)
    tosArray.sort((a, b) => {
        return (a.dataMaisAntiga || new Date()) - (b.dataMaisAntiga || new Date());
    });
    
    // Pegar total de pedidos da LH
    const totalPedidosLH = lhTrips[lhTrip]?.length || 0;
    
    // Calcular métricas do modal
    const capCiclo = obterCapacidadeCicloAtual();
    const jaSelecionado = calcularTotalSelecionado() - (lhsSelecionadasPlan.has(lhTrip) ? totalPedidosLH : 0);
    const faltam = Math.max(0, capCiclo - jaSelecionado);
    
    // Atualizar header do modal
    const locationIcon = estaNoPiso ? '📍 No Piso' : '🚚 Em Trânsito';
    document.getElementById('modalToLhName').innerHTML = `LH: ${lhTrip} (${totalPedidosLH.toLocaleString('pt-BR')} pedidos) - ${locationIcon}`;
    
    // Atualizar info
    document.getElementById('modalToCapCiclo').textContent = capCiclo.toLocaleString('pt-BR');
    document.getElementById('modalToJaSelecionado').textContent = jaSelecionado.toLocaleString('pt-BR');
    document.getElementById('modalToFaltam').textContent = faltam.toLocaleString('pt-BR');
    
    // Inicializar seleção de TOs se não existir
    if (!tosSelecionadasPorLH[lhTrip]) {
        tosSelecionadasPorLH[lhTrip] = new Set();
    }
    
    // Renderizar tabela de TOs
    renderizarTabelaTOs(tosArray, lhTrip, faltam, estaNoPiso);
    
    // Mostrar modal
    document.getElementById('modalTOs').style.display = 'flex';
    
    // Atualizar info de seleção
    atualizarInfoSelecaoTOs(lhTrip, tosArray);
}

function fecharModalTOs() {
    document.getElementById('modalTOs').style.display = 'none';
    lhAtualModal = null;
}

function extrairTOsDaLH(lhTrip, dadosPlanilhaLH) {
    const tosArray = [];
    
    // Buscar pedidos dessa LH no SPX
    const pedidosLH = lhTrips[lhTrip] || [];
    
    // Agrupar por TO
    const tosPorId = {};
    
    pedidosLH.forEach(pedido => {
        // Encontrar coluna de TO
        const colunaTO = Object.keys(pedido).find(col =>
            col.toLowerCase().includes('to id') ||
            col.toLowerCase().includes('to_id') ||
            col.toLowerCase().includes('transfer order') ||
            col.toLowerCase() === 'to'
        );
        
        const toId = pedido[colunaTO] || 'SEM_TO';
        
        if (!tosPorId[toId]) {
            tosPorId[toId] = {
                toId,
                pedidos: [],
                dataMaisAntiga: null
            };
        }
        
        tosPorId[toId].pedidos.push(pedido);
        
        // Encontrar data mais antiga para FIFO
        const colunaData = Object.keys(pedido).find(col =>
            col.toLowerCase().includes('create') ||
            col.toLowerCase().includes('date') ||
            col.toLowerCase().includes('data')
        );
        
        if (colunaData && pedido[colunaData]) {
            const data = parsearDataHora(pedido[colunaData]);
            if (data) {
                if (!tosPorId[toId].dataMaisAntiga || data < tosPorId[toId].dataMaisAntiga) {
                    tosPorId[toId].dataMaisAntiga = data;
                }
            }
        }
    });
    
    // Converter para array
    Object.values(tosPorId).forEach(to => {
        tosArray.push({
            toId: to.toId,
            qtdPedidos: to.pedidos.length,
            dataMaisAntiga: to.dataMaisAntiga || new Date(),
            pedidos: to.pedidos
        });
    });
    
    return tosArray;
}

function renderizarTabelaTOs(tosArray, lhTrip, faltam, estaNoPiso) {
    const tbody = document.getElementById('tbodyTOs');
    const thead = document.getElementById('theadTOs');
    
    // Atualizar header da coluna de data baseado na localização
    const colunaDataLabel = estaNoPiso ? 'Chegou em (FIFO)' : 'Previsão Chegada';
    if (thead) {
        const headers = thead.querySelectorAll('th');
        if (headers.length >= 5) {
            headers[4].textContent = colunaDataLabel;
        }
    }
    
    let html = '';
    let acumulado = 0;
    const capCiclo = obterCapacidadeCicloAtual();
    
    tosArray.forEach((to, index) => {
        const selecionada = tosSelecionadasPorLH[lhTrip]?.has(to.toId);
        acumulado += to.qtdPedidos;
        
        // Verificar se esta TO estouraria o CAP
        const totalAtualSemLH = calcularTotalSelecionado() - (lhsSelecionadasPlan.has(lhTrip) ? (lhTrips[lhTrip]?.length || 0) : 0);
        const totalComTOsSelecionadas = totalAtualSemLH + calcularTotalTOsSelecionadas(lhTrip);
        const estouraria = (totalComTOsSelecionadas + to.qtdPedidos) > capCiclo && capCiclo > 0;
        
        let statusBadge = '';
        if (selecionada) {
            statusBadge = '<span class="badge-to-selecionada">✅ Selecionada</span>';
        } else if (estouraria) {
            statusBadge = '<span class="badge-to-estoura">⚠️ Estoura CAP</span>';
        } else {
            statusBadge = '<span class="badge-to-ok">✅ Cabe</span>';
        }
        
        const dataFormatada = to.dataMaisAntiga ? 
            to.dataMaisAntiga.toLocaleDateString('pt-BR') + ' ' + 
            to.dataMaisAntiga.toLocaleTimeString('pt-BR', {hour: '2-digit', minute: '2-digit'}) : 
            '-';
        
        html += `
            <tr class="${selecionada ? 'to-selecionada' : ''} ${estouraria && !selecionada ? 'to-estouraria' : ''}">
                <td>
                    <input type="checkbox" class="checkbox-to" data-to="${to.toId}" 
                           ${selecionada ? 'checked' : ''} 
                           onchange="toggleSelecaoTO('${lhTrip}', '${to.toId}', this.checked)">
                </td>
                <td class="to-id">${to.toId}</td>
                <td class="to-pedidos">${to.qtdPedidos.toLocaleString('pt-BR')}</td>
                <td class="to-acumulado">${acumulado.toLocaleString('pt-BR')}</td>
                <td class="to-data">${dataFormatada}</td>
                <td class="to-status">${statusBadge}</td>
            </tr>
        `;
    });
    
    tbody.innerHTML = html;
    
    // Event listener para checkbox de selecionar todas
    const checkTodasTOs = document.getElementById('checkTodasTOs');
    if (checkTodasTOs) {
        checkTodasTOs.checked = false;
        checkTodasTOs.onchange = (e) => toggleTodasTOs(lhTrip, e.target.checked);
    }
}

function toggleSelecaoTO(lhTrip, toId, checked) {
    if (!tosSelecionadasPorLH[lhTrip]) {
        tosSelecionadasPorLH[lhTrip] = new Set();
    }
    
    if (checked) {
        tosSelecionadasPorLH[lhTrip].add(toId);
    } else {
        tosSelecionadasPorLH[lhTrip].delete(toId);
    }
    
    // Recalcular e atualizar modal
    // ✅ USAR FUNÇÃO DE FILTRO POR STATION
    const dadosPlanilhaLH = buscarDadosPlanilhaPorStation(lhTrip);
    const tosArray = extrairTOsDaLH(lhTrip, dadosPlanilhaLH);
    atualizarInfoSelecaoTOs(lhTrip, tosArray);
    
    // Atualizar visual da linha
    const row = document.querySelector(`input[data-to="${toId}"]`)?.closest('tr');
    if (row) {
        row.classList.toggle('to-selecionada', checked);
    }
    
    // 🔄 RE-RENDERIZAR TABELA PRINCIPAL para atualizar coluna PEDIDOS TOS
    renderizarTabelaPlanejamento();
    
    // Atualizar card de sugestão se existir
    atualizarCardSugestao();
}

function toggleTodasTOs(lhTrip, checked) {
    const checkboxes = document.querySelectorAll('.checkbox-to');
    checkboxes.forEach(cb => {
        cb.checked = checked;
        const toId = cb.dataset.to;
        if (checked) {
            if (!tosSelecionadasPorLH[lhTrip]) tosSelecionadasPorLH[lhTrip] = new Set();
            tosSelecionadasPorLH[lhTrip].add(toId);
        } else {
            tosSelecionadasPorLH[lhTrip]?.delete(toId);
        }
        cb.closest('tr')?.classList.toggle('to-selecionada', checked);
    });
    
    // ✅ USAR FUNÇÃO DE FILTRO POR STATION
    const dadosPlanilhaLH = buscarDadosPlanilhaPorStation(lhTrip);
    const tosArray = extrairTOsDaLH(lhTrip, dadosPlanilhaLH);
    atualizarInfoSelecaoTOs(lhTrip, tosArray);
    
    // 🔄 RE-RENDERIZAR TABELA PRINCIPAL para atualizar coluna PEDIDOS TOS
    renderizarTabelaPlanejamento();
    
    // Atualizar card de sugestão se existir
    atualizarCardSugestao();
}

function limparSelecaoTOs() {
    if (!lhAtualModal) return;
    
    tosSelecionadasPorLH[lhAtualModal] = new Set();
    
    const checkboxes = document.querySelectorAll('.checkbox-to');
    checkboxes.forEach(cb => {
        cb.checked = false;
        cb.closest('tr')?.classList.remove('to-selecionada');
    });
    
    const checkTodasTOs = document.getElementById('checkTodasTOs');
    if (checkTodasTOs) checkTodasTOs.checked = false;
    
    const dadosPlanilhaLH = dadosPlanilha[lhAtualModal];
    const tosArray = extrairTOsDaLH(lhAtualModal, dadosPlanilhaLH);
    atualizarInfoSelecaoTOs(lhAtualModal, tosArray);
    
    // 🔄 RE-RENDERIZAR TABELA PRINCIPAL para atualizar coluna PEDIDOS TOS
    renderizarTabelaPlanejamento();
    
    // Atualizar card de sugestão se existir
    atualizarCardSugestao();
}

function sugerirTOsAutomatico() {
    if (!lhAtualModal) return;
    
    const lhTrip = lhAtualModal;
    // ✅ USAR FUNÇÃO DE FILTRO POR STATION
    const dadosPlanilhaLH = buscarDadosPlanilhaPorStation(lhTrip);
    const tosArray = extrairTOsDaLH(lhTrip, dadosPlanilhaLH);
    
    // Detectar status da LH
    const statusLH = calcularStatusLH(dadosPlanilhaLH);
    const estaNoPiso = statusLH.codigo === 'P0' || statusLH.codigo === 'P0B';
    
    // Ordenar por FIFO (mais antiga primeiro)
    tosArray.sort((a, b) => (a.dataMaisAntiga || new Date()) - (b.dataMaisAntiga || new Date()));
    
    // Calcular quanto falta
    const capCiclo = obterCapacidadeCicloAtual();
    const totalAtualSemLH = calcularTotalSelecionado() - (lhsSelecionadasPlan.has(lhTrip) ? (lhTrips[lhTrip]?.length || 0) : 0);
    const faltam = Math.max(0, capCiclo - totalAtualSemLH);
    
    // Limpar seleção anterior
    tosSelecionadasPorLH[lhTrip] = new Set();
    
    // Selecionar TOs até atingir o faltam (FIFO)
    let totalSelecionado = 0;
    
    for (const to of tosArray) {
        if (totalSelecionado + to.qtdPedidos <= faltam) {
            tosSelecionadasPorLH[lhTrip].add(to.toId);
            totalSelecionado += to.qtdPedidos;
        }
    }
    
    // Re-renderizar tabela
    renderizarTabelaTOs(tosArray, lhTrip, faltam, estaNoPiso);
    atualizarInfoSelecaoTOs(lhTrip, tosArray);
    
    // 🔄 RE-RENDERIZAR TABELA PRINCIPAL para atualizar coluna PEDIDOS TOS
    renderizarTabelaPlanejamento();
    
    // Atualizar card de sugestão se existir
    atualizarCardSugestao();
}

function calcularTotalTOsSelecionadas(lhTrip) {
    if (!tosSelecionadasPorLH[lhTrip] || tosSelecionadasPorLH[lhTrip].size === 0) {
        return 0;
    }
    
    const tosArray = extrairTOsDaLH(lhTrip, dadosPlanilha[lhTrip]);
    let total = 0;
    
    tosArray.forEach(to => {
        if (tosSelecionadasPorLH[lhTrip].has(to.toId)) {
            total += to.qtdPedidos;
        }
    });
    
    return total;
}

function atualizarInfoSelecaoTOs(lhTrip, tosArray) {
    const tosSelecionadas = tosSelecionadasPorLH[lhTrip] || new Set();
    
    let qtdTOs = tosSelecionadas.size;
    let pedidosTOs = 0;
    
    tosArray.forEach(to => {
        if (tosSelecionadas.has(to.toId)) {
            pedidosTOs += to.qtdPedidos;
        }
    });
    
    const capCiclo = obterCapacidadeCicloAtual();
    const totalAtualSemLH = calcularTotalSelecionado() - (lhsSelecionadasPlan.has(lhTrip) ? (lhTrips[lhTrip]?.length || 0) : 0);
    const novoTotal = totalAtualSemLH + pedidosTOs;
    
    document.getElementById('modalToQtdTOs').textContent = qtdTOs.toLocaleString('pt-BR');
    document.getElementById('modalToPedidosTOs').textContent = pedidosTOs.toLocaleString('pt-BR');
    document.getElementById('modalToNovoTotal').textContent = novoTotal.toLocaleString('pt-BR');
    
    // Atualizar resumo
    const resumoEl = document.getElementById('modalToResumo');
    if (qtdTOs > 0) {
        const percentual = capCiclo > 0 ? ((novoTotal / capCiclo) * 100).toFixed(1) : 0;
        resumoEl.textContent = `${qtdTOs} TOs selecionadas = ${pedidosTOs.toLocaleString('pt-BR')} pedidos (${percentual}% do CAP)`;
    } else {
        resumoEl.textContent = 'Selecione as TOs para complementar o planejamento';
    }
}

function confirmarSelecaoTOs() {
    if (!lhAtualModal) return;
    
    const lhTrip = lhAtualModal;
    const tosSelecionadas = tosSelecionadasPorLH[lhTrip];
    
    if (!tosSelecionadas || tosSelecionadas.size === 0) {
        // Se não tem TOs selecionadas, remover a LH da seleção
        lhsSelecionadasPlan.delete(lhTrip);
        delete tosSelecionadasPorLH[lhTrip];
    } else {
        // Marcar LH como selecionada (parcialmente via TOs)
        lhsSelecionadasPlan.add(lhTrip);
    }
    
    fecharModalTOs();
    
    // Atualizar tabela e contadores
    renderizarTabelaPlanejamento();
    atualizarContadorSelecaoLHs();
    
    // ✅ ATUALIZAR CARD DE SUGESTÃO com novo total
    atualizarCardSugestao();
}

// Função auxiliar para obter capacidade do ciclo atual
function obterCapacidadeCicloAtual() {
    if (!cicloSelecionado || cicloSelecionado === 'Todos') return 0;
    
    // VERIFICAR CAP MANUAL PRIMEIRO
    const capManual = obtemCapManual(cicloSelecionado);
    if (capManual !== null) {
        console.log('✅ Usando CAP Manual no modal:', cicloSelecionado, '=', capManual);
        return capManual;
    }
    
    // Se não tem CAP Manual, pegar do Google Sheets
    const stationSelecionada = stationAtualNome || '';
    const stationBase = stationSelecionada.toLowerCase().replace(/lm\s*hub[_\s]*/gi, '').replace(/[_\s]+/g, '');
    
    const capacidadeStation = dadosOutbound.filter(item => {
        const sortCodeName = item['Sort Code Name'] || item['sort_code_name'] || '';
        const itemNorm = sortCodeName.toLowerCase().replace(/lm\s*hub[_\s]*/gi, '').replace(/[_\s]+/g, '');
        return itemNorm.includes(stationBase) || stationBase.includes(itemNorm);
    });
    
    const registroCiclo = capacidadeStation.find(cap => {
        const tipoCap = cap['Type Outbound'] || cap['type_outbound'] || '';
        return tipoCap.toUpperCase() === cicloSelecionado.toUpperCase();
    });
    
    if (!registroCiclo) return 0;
    
    const dataCiclo = getDataCicloSelecionada();
    const diaHoje = String(dataCiclo.getDate()).padStart(2, '0');
    const diaSemZero = String(dataCiclo.getDate());
    const mesHoje = String(dataCiclo.getMonth() + 1).padStart(2, '0');
    const mesSemZero = String(dataCiclo.getMonth() + 1);
    const anoHoje = dataCiclo.getFullYear();
    const anoCurto = String(anoHoje).slice(2);
    
    const formatosData = [
        `${diaHoje}/${mesHoje}/${anoHoje}`,
        `${diaSemZero}/${mesSemZero}/${anoHoje}`,
        `${diaHoje}/${mesHoje}/${anoCurto}`,
        `${anoHoje}-${mesHoje}-${diaHoje}`,
    ];
    
    for (const formato of formatosData) {
        if (registroCiclo[formato] !== undefined && registroCiclo[formato] !== '') {
            let valor = registroCiclo[formato];
            if (typeof valor === 'string') {
                valor = valor.replace(/\./g, '').replace(',', '.');
            }
            return parseFloat(valor) || 0;
        }
    }
    
    return 0;
}

// Calcular total selecionado (LHs + Backlog, considerando TOs parciais)
function calcularTotalSelecionado() {
    let total = 0;
    
    // Backlog selecionado
    total += pedidosBacklogSelecionados.size;
    console.log('📊 calcularTotalSelecionado - Backlog:', pedidosBacklogSelecionados.size);
    
    // LHs selecionadas
    lhsSelecionadasPlan.forEach(lh => {
        // Verificar se tem TOs parciais
        if (tosSelecionadasPorLH[lh] && tosSelecionadasPorLH[lh].size > 0) {
            // Usar apenas pedidos das TOs selecionadas
            const totalTOs = calcularTotalTOsSelecionadas(lh);
            console.log(`  LH ${lh}: ${totalTOs} pedidos (TOs parciais)`);
            total += totalTOs;
        } else {
            // Usar todos os pedidos da LH
            const totalLH = lhTrips[lh]?.length || 0;
            console.log(`  LH ${lh}: ${totalLH} pedidos (LH completa)`);
            total += totalLH;
        }
    });
    
    console.log('📊 Total final calculado:', total);
    return total;
}

// Inicializar evento de duplo clique para abrir modal de TOs
document.addEventListener('DOMContentLoaded', () => {
    // Aguardar um pouco para garantir que a tabela foi renderizada
    setTimeout(() => {
        document.addEventListener('dblclick', (e) => {
            const cell = e.target.closest('.lh-trip-cell');
            if (cell) {
                const lhTrip = cell.textContent.trim().split('\n')[0].trim();
                if (lhTrip && lhTrip !== '-') {
                    abrirModalTOs(lhTrip);
                }
            }
        });
    }, 500);
});

// Expor funções globalmente para uso nos onclick do HTML
window.toggleSelecaoTO = toggleSelecaoTO;
window.toggleTodasTOs = toggleTodasTOs;
window.limparSelecaoTOs = limparSelecaoTOs;
window.confirmarSelecaoTOs = confirmarSelecaoTOs;
window.fecharModalTOs = fecharModalTOs;
window.abrirModalTOs = abrirModalTOs;
window.sugerirTOsAutomatico = sugerirTOsAutomatico;
// ======================= FUNÇÕES CAP MANUAL =======================

function carregarCapsManual() {
    const saved = localStorage.getItem('capsManual');
    if (saved) {
        try {
            capsManual = JSON.parse(saved);
            console.log('📊 CAPs Manual carregados:', capsManual);
        } catch (e) {
            console.error('Erro ao carregar CAPs manual:', e);
            capsManual = {};
        }
    }
}

function salvarCapsManual() {
    localStorage.setItem('capsManual', JSON.stringify(capsManual));
    console.log('💾 CAPs Manual salvos:', capsManual);
}

function definirCapManual(ciclo, capacidade) {
    if (!ciclo || ciclo === 'Todos') {
        alert('⚠️ Selecione um ciclo específico (AM, PM1 ou PM2)');
        return false;
    }
    
    const cap = parseInt(capacidade);
    if (isNaN(cap) || cap <= 0) {
        alert('⚠️ Informe uma capacidade válida maior que zero');
        return false;
    }
    
    capsManual[ciclo] = cap;
    salvarCapsManual();
    
    // SELECIONAR CICLO ANTES de recarregar
    cicloSelecionado = ciclo;
    console.log('✅ Ciclo selecionado:', ciclo);
    
    // Recarregar dados para atualizar interface
    carregarDadosOpsClockLocal();
    renderizarTabelaPlanejamento();
    
    // REFORÇAR seleção do ciclo DEPOIS do recarregamento
    setTimeout(() => {
        cicloSelecionado = ciclo;
        
        // Atualizar visual dos cards
        const containerCiclos = document.getElementById('containerCiclos');
        if (containerCiclos) {
            containerCiclos.querySelectorAll('.ciclo-stat').forEach(card => {
                card.classList.remove('ativo');
                if (card.dataset.ciclo === ciclo) {
                    card.classList.add('ativo');
                }
            });
        }
        
        console.log('✅ Ciclo selecionado após reload:', ciclo);
    }, 100);
    
    console.log('✅ CAP Manual definido:', ciclo, '=', cap.toLocaleString('pt-BR'));
    return true;
}

function removerCapManual(ciclo) {
    if (capsManual[ciclo]) {
        delete capsManual[ciclo];
        salvarCapsManual();
        
        // Recarregar dados para atualizar interface
        carregarDadosOpsClockLocal();
        renderizarTabelaPlanejamento();
        
        console.log('🗑️ CAP Manual removido:', ciclo);
        return true;
    }
    return false;
}

function obtemCapManual(ciclo) {
    return capsManual[ciclo] || null;
}

function abrirModalCapManual() {
    console.log('🔓 Abrindo modal CAP Manual...');
    
    const modal = document.getElementById('modalCapManual');
    const inputValor = document.getElementById('capManualValor');
    
    // Preencher valores
    const cicloAtual = cicloSelecionado !== 'Todos' ? cicloSelecionado : 'AM';
    document.getElementById('capManualCiclo').value = cicloAtual;
    inputValor.value = capsManual[cicloAtual] || '';
    
    // Remover todos os bloqueios possíveis
    inputValor.removeAttribute('readonly');
    inputValor.removeAttribute('disabled');
    inputValor.readOnly = false;
    inputValor.disabled = false;
    inputValor.contentEditable = false; // Não usar contenteditable
    
    // Mostrar modal
    modal.style.display = 'flex';
    
    // SOLUÇÃO DEFINITIVA: Capturar teclado globalmente
    const handleGlobalKeyPress = function(e) {
        // Se modal está visível E tecla é alfanumérica
        if (modal.style.display === 'flex') {
            // Focar input se não estiver focado
            if (document.activeElement !== inputValor) {
                e.preventDefault();
                inputValor.focus();
                // Simular a digitação da tecla
                if (e.key.length === 1) {
                    inputValor.value += e.key;
                    atualizarPreviewCapManual();
                }
            }
        }
    };
    
    // Armazenar referência para remover depois
    window._capManualKeyHandler = handleGlobalKeyPress;
    document.addEventListener('keypress', handleGlobalKeyPress, true);
    
    // Atualizar preview
    setTimeout(() => atualizarPreviewCapManual(), 10);
    
    // Tentativa de foco
    setTimeout(() => {
        inputValor.click();
        inputValor.focus();
        console.log('✅ Foco aplicado + listener de teclado ativo');
    }, 100);
}

function fecharModalCapManual() {
    const modal = document.getElementById('modalCapManual');
    modal.style.display = 'none';
    
    // Remover listener global
    if (window._capManualKeyHandler) {
        document.removeEventListener('keypress', window._capManualKeyHandler, true);
        window._capManualKeyHandler = null;
    }
}

function atualizarPreviewCapManual() {
    const ciclo = document.getElementById('capManualCiclo').value;
    const inputValor = document.getElementById('capManualValor');
    
    // Remover tudo que não é número
    const apenasNumeros = inputValor.value.replace(/\D/g, '');
    const valor = parseInt(apenasNumeros) || 0;
    
    // Pegar CAP automático para comparação
    const capAuto = obterCapacidadeCiclo(ciclo);
    const capManualAtual = capsManual[ciclo];
    
    let html = '<div class="cap-manual-preview">';
    
    // CAP Automático
    const capAutoStr = capAuto > 0 ? capAuto.toLocaleString('pt-BR') : 'Não encontrado';
    html += '<div class="preview-item">';
    html += '<span class="preview-label">CAP Automático (Google Sheets):</span>';
    html += '<span class="preview-valor">' + capAutoStr + '</span>';
    html += '</div>';
    
    // CAP Manual Atual
    if (capManualAtual) {
        html += '<div class="preview-item destaque">';
        html += '<span class="preview-label">✅ CAP Manual Atual:</span>';
        html += '<span class="preview-valor">' + capManualAtual.toLocaleString('pt-BR') + '</span>';
        html += '</div>';
    }
    
    // Novo CAP
    if (valor > 0) {
        const diferenca = valor - capAuto;
        const sinal = diferenca > 0 ? '+' : '';
        const cor = diferenca > 0 ? '#28a745' : (diferenca < 0 ? '#dc3545' : '#666');
        
        html += '<div class="preview-item novo">';
        html += '<span class="preview-label">➡️ Novo CAP Manual:</span>';
        html += '<span class="preview-valor">' + valor.toLocaleString('pt-BR') + '</span>';
        html += '</div>';
        
        if (capAuto > 0) {
            const percDif = ((diferenca/capAuto)*100).toFixed(1);
            html += '<div class="preview-item">';
            html += '<span class="preview-label">Diferença:</span>';
            html += '<span class="preview-valor" style="color: ' + cor + '; font-weight: 700;">';
            html += sinal + diferenca.toLocaleString('pt-BR') + ' (' + percDif + '%)';
            html += '</span>';
            html += '</div>';
        }
    }
    
    html += '</div>';
    
    document.getElementById('capManualPreview').innerHTML = html;
}

function confirmarCapManual() {
    const ciclo = document.getElementById('capManualCiclo').value;
    const inputValor = document.getElementById('capManualValor').value;
    
    // Remover tudo que não é número
    const valor = inputValor.replace(/\D/g, '');
    
    if (definirCapManual(ciclo, valor)) {
        fecharModalCapManual();
        
        // SELECIONAR AUTOMATICAMENTE O CICLO
        cicloSelecionado = ciclo;
        console.log('✅ Ciclo selecionado automaticamente:', cicloSelecionado);
        
        // Atualizar visual dos cards
        const containerCiclos = document.getElementById('containerCiclos');
        if (containerCiclos) {
            containerCiclos.querySelectorAll('.ciclo-stat').forEach(card => {
                card.classList.remove('ativo');
                if (card.dataset.ciclo === ciclo) {
                    card.classList.add('ativo');
                }
            });
        }
        
        // Re-renderizar tabela com o filtro do ciclo
        renderizarTabelaPlanejamento();
        
        const msg = '✅ CAP Manual definido e ciclo selecionado!\n\nCiclo: ' + ciclo + '\nCapacidade: ' + parseInt(valor).toLocaleString('pt-BR') + ' pedidos\n\n➡️ Agora é só clicar em "Sugerir Planejamento"';
        alert(msg);
    }
}

function limparCapManual() {
    const ciclo = document.getElementById('capManualCiclo').value;
    
    if (capsManual[ciclo]) {
        const msg = 'Deseja remover o CAP Manual do ciclo ' + ciclo + '?\n\nO sistema voltará a usar o CAP automático do Google Sheets.';
        if (confirm(msg)) {
            removerCapManual(ciclo);
            fecharModalCapManual();
        }
    } else {
        alert('⚠️ Não há CAP Manual definido para o ciclo ' + ciclo);
    }
}

// Expor funções globalmente
window.abrirModalCapManual = abrirModalCapManual;
window.fecharModalCapManual = fecharModalCapManual;
window.confirmarCapManual = confirmarCapManual;
window.limparCapManual = limparCapManual;
window.atualizarPreviewCapManual = atualizarPreviewCapManual;




// ==================== CAP MANUAL INLINE ====================

function aplicarCapManualInline() {
    const cicloSelect = document.getElementById('capManualCicloInline');
    const valorInput = document.getElementById('capManualValorInline');
    
    const ciclo = cicloSelect.value;
    const valor = valorInput.value.replace(/\D/g, ''); // Remove não-números
    
    if (!ciclo) {
        alert('⚠️ Selecione um ciclo (AM, PM1 ou PM2)');
        cicloSelect.focus();
        return;
    }
    
    if (!valor || parseInt(valor) <= 0) {
        alert('⚠️ Digite uma capacidade válida maior que zero');
        valorInput.focus();
        return;
    }
    
    const capacidade = parseInt(valor);
    
    // Definir CAP Manual
    if (definirCapManual(ciclo, capacidade)) {
        // Limpar campos
        valorInput.value = '';
        
        // Mensagem de sucesso
        const msg = '✅ CAP Manual aplicado!\n\n' +
                    'Ciclo: ' + ciclo + '\n' +
                    'Capacidade: ' + capacidade.toLocaleString('pt-BR') + ' pedidos\n\n' +
                    '➡️ Ciclo selecionado automaticamente!\n' +
                    'Clique em "Sugerir Planejamento" para usar.';
        alert(msg);
        
        console.log('✅ CAP Manual aplicado via inline:', ciclo, '=', capacidade);
    }
}

function removerCapManualInline() {
    const cicloSelect = document.getElementById('capManualCicloInline');
    const valorInput = document.getElementById('capManualValorInline');
    
    const ciclo = cicloSelect.value;
    
    if (!ciclo) {
        alert('⚠️ Selecione um ciclo para remover');
        return;
    }
    
    if (!capsManual[ciclo]) {
        alert('ℹ️ Não há CAP Manual definido para ' + ciclo);
        return;
    }
    
    const confirma = confirm('Deseja remover o CAP Manual do ciclo ' + ciclo + '?\n\nO sistema voltará a usar o CAP automático do Google Sheets.');
    
    if (confirma) {
        removerCapManual(ciclo);
        
        // Limpar campos
        cicloSelect.value = '';
        valorInput.value = '';
        
        alert('✅ CAP Manual removido!\n\nSistema voltou ao CAP automático.');
        
        console.log('🗑️ CAP Manual removido via inline:', ciclo);
    }
}

// Atualizar select quando carregar CAPs
function atualizarSelectCapManual() {
    const cicloSelect = document.getElementById('capManualCicloInline');
    if (!cicloSelect) return;
    
    // Adicionar indicador visual nos options que têm CAP Manual
    const options = cicloSelect.querySelectorAll('option');
    options.forEach(opt => {
        const ciclo = opt.value;
        if (ciclo && capsManual[ciclo]) {
            opt.textContent = ciclo + ' ✓ (' + capsManual[ciclo].toLocaleString('pt-BR') + ')';
        } else if (ciclo) {
            opt.textContent = ciclo;
        }
    });
}

// Chamar ao carregar CAPs
const _carregarCapsManualOriginal = carregarCapsManual;
carregarCapsManual = function() {
    _carregarCapsManualOriginal();
    atualizarSelectCapManual();
};

// Preencher valor ao selecionar ciclo com CAP existente
document.addEventListener('DOMContentLoaded', () => {
    const cicloSelect = document.getElementById('capManualCicloInline');
    const valorInput = document.getElementById('capManualValorInline');
    
    if (cicloSelect && valorInput) {
        cicloSelect.addEventListener('change', () => {
            const ciclo = cicloSelect.value;
            if (ciclo && capsManual[ciclo]) {
                valorInput.value = capsManual[ciclo];
                console.log('📝 CAP Manual carregado no input:', ciclo, '=', capsManual[ciclo]);
            } else {
                valorInput.value = '';
            }
        });
    }
});

console.log('✅ Funções CAP Manual Inline carregadas');


// ==================== LOGS DETALHADOS - CAP MANUAL INLINE ====================

// Monitorar eventos do input
document.addEventListener('DOMContentLoaded', () => {
    const valorInput = document.getElementById('capManualValorInline');
    
    if (valorInput) {
        console.log('✅ Input CAP Manual encontrado:', valorInput);
        
        // Log de todos os eventos
        valorInput.addEventListener('focus', () => {
            console.log('🎯 INPUT FOCADO');
            console.log('  - activeElement:', document.activeElement.id);
            console.log('  - readOnly:', valorInput.readOnly);
            console.log('  - disabled:', valorInput.disabled);
            console.log('  - contentEditable:', valorInput.contentEditable);
        });
        
        valorInput.addEventListener('blur', () => {
            console.log('😔 INPUT PERDEU FOCO');
        });
        
        valorInput.addEventListener('keydown', (e) => {
            console.log('⌨️ KEYDOWN:', e.key, 'code:', e.code);
        });
        
        valorInput.addEventListener('keypress', (e) => {
            console.log('⌨️ KEYPRESS:', e.key, 'code:', e.code);
        });
        
        valorInput.addEventListener('input', (e) => {
            console.log('✍️ INPUT EVENT - Valor atual:', e.target.value);
        });
        
        valorInput.addEventListener('change', (e) => {
            console.log('🔄 CHANGE EVENT - Valor:', e.target.value);
        });
        
        // Verificar estado inicial
        setTimeout(() => {
            console.log('📊 ESTADO INICIAL DO INPUT:');
            console.log('  - ID:', valorInput.id);
            console.log('  - Type:', valorInput.type);
            console.log('  - ReadOnly:', valorInput.readOnly);
            console.log('  - Disabled:', valorInput.disabled);
            console.log('  - TabIndex:', valorInput.tabIndex);
            console.log('  - Display:', window.getComputedStyle(valorInput).display);
            console.log('  - Visibility:', window.getComputedStyle(valorInput).visibility);
            console.log('  - PointerEvents:', window.getComputedStyle(valorInput).pointerEvents);
        }, 1000);
        
        // Monitorar mudanças de atributos
        const observer = new MutationObserver((mutations) => {
            mutations.forEach((mutation) => {
                if (mutation.type === 'attributes') {
                    console.log('🔧 ATRIBUTO MUDOU:', mutation.attributeName);
                    console.log('  - Novo valor:', valorInput.getAttribute(mutation.attributeName));
                }
            });
        });
        
        observer.observe(valorInput, {
            attributes: true,
            attributeOldValue: true
        });
        
        console.log('✅ Monitoramento do input ativo');
    } else {
        console.error('❌ Input CAP Manual NÃO encontrado!');
    }
});

// Monitorar clicks no documento
document.addEventListener('click', (e) => {
    const target = e.target;
    if (target.id === 'capManualValorInline') {
        console.log('🖱️ CLICK NO INPUT CAP MANUAL');
        console.log('  - Target:', target);
        console.log('  - ReadOnly:', target.readOnly);
        console.log('  - Disabled:', target.disabled);
    }
});

// Log de aplicação de CAP
const _aplicarCapManualInlineOriginal = aplicarCapManualInline;
aplicarCapManualInline = function() {
    console.log('🚀 APLICAR CAP MANUAL INLINE chamada');
    const cicloSelect = document.getElementById('capManualCicloInline');
    const valorInput = document.getElementById('capManualValorInline');
    console.log('  - Ciclo selecionado:', cicloSelect.value);
    console.log('  - Valor digitado:', valorInput.value);
    console.log('  - Input readOnly:', valorInput.readOnly);
    console.log('  - Input disabled:', valorInput.disabled);
    _aplicarCapManualInlineOriginal();
};

console.log('✅ Sistema de logs detalhados ativado');


// ==================== DEBUG - QUEM ESTÁ ROUBANDO O FOCO ====================

let lastFocusedElement = null;

document.addEventListener('focus', (e) => {
    if (lastFocusedElement && lastFocusedElement.id === 'capManualValorInline' && e.target.id !== 'capManualValorInline') {
        console.log('🚨 FOCO ROUBADO DO INPUT!');
        console.log('  - Novo elemento focado:', e.target.id || e.target.tagName || e.target);
        console.log('  - ClassName:', e.target.className);
        console.trace('  - Stack trace de quem roubou:');
    }
    lastFocusedElement = e.target;
}, true);

// Interceptar blur do input para ver quem causou
document.addEventListener('DOMContentLoaded', () => {
    const input = document.getElementById('capManualValorInline');
    if (input) {
        input.addEventListener('blur', (e) => {
            console.log('🔍 INPUT BLUR - Relacionado a:', e.relatedTarget);
            console.log('  - Novo foco irá para:', e.relatedTarget ? (e.relatedTarget.id || e.relatedTarget.tagName) : 'NENHUM ELEMENTO');
            console.trace('  - Stack trace do blur:');
        });
    }
});

console.log('✅ Debug de roubo de foco ativado');


// ==================== CORREÇÃO - PREVENIR PERDA DE FOCO ====================

document.addEventListener('DOMContentLoaded', () => {
    const cicloSelect = document.getElementById('capManualCicloInline');
    const valorInput = document.getElementById('capManualValorInline');
    
    if (valorInput) {
        // CORREÇÃO 1: Prevenir Enter de tirar foco
        valorInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault(); // Bloqueia comportamento padrão do Enter
                console.log('🛡️ Enter bloqueado - mantendo foco no input');
                
                // Opcional: aplicar CAP ao pressionar Enter
                if (valorInput.value.trim()) {
                    aplicarCapManualInline();
                }
                return false;
            }
        });
        
        // CORREÇÃO 2: Re-focar automaticamente se perder foco sem motivo
        valorInput.addEventListener('blur', (e) => {
            // Se não está indo para outro input/select/button
            if (!e.relatedTarget || (!e.relatedTarget.matches('input, select, button'))) {
                console.log('🔄 Refocando input automaticamente');
                setTimeout(() => {
                    valorInput.focus();
                }, 10);
            }
        });
        
        console.log('✅ Proteção de foco aplicada ao input');
    }
    
    // CORREÇÃO 3: Garantir que select também não cause problemas
    if (cicloSelect) {
        cicloSelect.addEventListener('change', () => {
            // Não focar input automaticamente ao trocar ciclo
            // deixar usuário decidir quando digitar
            console.log('📝 Ciclo alterado, aguardando input do usuário');
        });
    }
});

console.log('✅ Correções de foco aplicadas');


// ==================== SOLUÇÃO DEFINITIVA - CONTENTEDITABLE ====================

function inicializarCapManualContentEditable() {
    const valorDiv = document.getElementById('capManualValorInline');
    
    if (!valorDiv) {
        console.error('❌ Elemento capManualValorInline não encontrado');
        return;
    }
    
    console.log('✅ Inicializando input contenteditable');
    
    // Permitir apenas números
    valorDiv.addEventListener('input', function(e) {
        let texto = this.textContent;
        let apenasNumeros = texto.replace(/\D/g, '');
        
        // Limitar a 6 dígitos
        if (apenasNumeros.length > 6) {
            apenasNumeros = apenasNumeros.substring(0, 6);
        }
        
        if (texto !== apenasNumeros) {
            // Salvar posição do cursor
            const selection = window.getSelection();
            const range = selection.getRangeAt(0);
            const offset = range.startOffset;
            
            // Atualizar texto
            this.textContent = apenasNumeros;
            
            // Restaurar cursor
            try {
                if (this.firstChild) {
                    const newRange = document.createRange();
                    const newOffset = Math.min(offset, apenasNumeros.length);
                    newRange.setStart(this.firstChild, newOffset);
                    newRange.collapse(true);
                    selection.removeAllRanges();
                    selection.addRange(newRange);
                }
            } catch (e) {
                console.warn('Erro ao restaurar cursor:', e);
            }
        }
        
        console.log('✍️ Valor atual:', apenasNumeros);
    });
    
    // Prevenir colagem de texto não-numérico
    valorDiv.addEventListener('paste', function(e) {
        e.preventDefault();
        const texto = (e.clipboardData || window.clipboardData).getData('text');
        const apenasNumeros = texto.replace(/\D/g, '').substring(0, 6);
        
        const selection = window.getSelection();
        if (selection.rangeCount) {
            selection.deleteFromDocument();
            selection.getRangeAt(0).insertNode(document.createTextNode(apenasNumeros));
            selection.collapseToEnd();
        }
        
        console.log('📋 Colado:', apenasNumeros);
    });
    
    // Enter aplica CAP
    valorDiv.addEventListener('keydown', function(e) {
        if (e.key === 'Enter') {
            e.preventDefault();
            aplicarCapManualInline();
            console.log('⏎ Enter pressionado - aplicando CAP');
        }
    });
    
    // Focar ao clicar na área
    valorDiv.addEventListener('click', function() {
        this.focus();
        console.log('🖱️ Div clicada e focada');
    });
    
    console.log('✅ ContentEditable configurado com sucesso');
}

// Atualizar função aplicarCapManualInline para usar textContent
function aplicarCapManualInline() {
    const cicloSelect = document.getElementById('capManualCicloInline');
    const valorDiv = document.getElementById('capManualValorInline');
    
    const ciclo = cicloSelect.value;
    const valor = (valorDiv.textContent || '').replace(/\D/g, '');
    
    console.log('🚀 APLICAR CAP MANUAL (contenteditable)');
    console.log('  - Ciclo:', ciclo);
    console.log('  - Valor:', valor);
    
    if (!ciclo) {
        alert('⚠️ Selecione um ciclo');
        return;
    }
    
    // Se o campo está vazio, REMOVER o CAP Manual
    if (!valor || valor === '0') {
        if (capsManual[ciclo]) {
            const confirma = confirm('🗑️ Remover CAP Manual do ciclo ' + ciclo + '?\n\nO sistema voltará a usar o CAP automático do Google Sheets.');
            if (confirma) {
                removerCapManual(ciclo);
                cicloSelect.value = '';
                valorDiv.textContent = '';
                alert('✅ CAP Manual removido!\n\nSistema voltou ao CAP automático.');
                console.log('🗑️ CAP Manual removido (campo vazio):', ciclo);
            }
        } else {
            alert('⚠️ Digite um valor válido ou remova o CAP Manual existente');
        }
        return;
    }
    
    const capacidade = parseInt(valor);
    
    // Definir CAP Manual (função já seleciona ciclo automaticamente)
    if (definirCapManual(ciclo, capacidade)) {
        // Limpar campo
        valorDiv.textContent = '';
        
        // Mensagem de sucesso
        const msg = '✅ CAP Manual aplicado!\n\n' +
                    'Ciclo: ' + ciclo + '\n' +
                    'Capacidade: ' + capacidade.toLocaleString('pt-BR') + ' pedidos\n\n' +
                    '➡️ Ciclo selecionado automaticamente!\n' +
                    'Clique em "Sugerir Planejamento" para usar.';
        alert(msg);
        
        console.log('✅ CAP Manual aplicado via inline:', ciclo, '=', capacidade);
    }
}

function removerCapManualInline() {
    const cicloSelect = document.getElementById('capManualCicloInline');
    const valorDiv = document.getElementById('capManualValorInline');
    
    const ciclo = cicloSelect.value;
    
    if (!ciclo) {
        alert('⚠️ Selecione um ciclo para remover');
        return;
    }
    
    if (!capsManual[ciclo]) {
        alert('ℹ️ Não há CAP Manual definido para ' + ciclo);
        return;
    }
    
    const confirma = confirm('Deseja remover o CAP Manual do ciclo ' + ciclo + '?\n\nO sistema voltará a usar o CAP automático do Google Sheets.');
    
    if (confirma) {
        removerCapManual(ciclo);
        
        // Limpar campos
        cicloSelect.value = '';
        valorDiv.textContent = '';
        
        alert('✅ CAP Manual removido!\n\nSistema voltou ao CAP automático.');
        
        console.log('🗑️ CAP Manual removido:', ciclo);
    }
}

// Inicializar quando DOM carregar
document.addEventListener('DOMContentLoaded', () => {
    inicializarCapManualContentEditable();
    
    // Atualizar select quando ciclo for selecionado
    const cicloSelect = document.getElementById('capManualCicloInline');
    const valorDiv = document.getElementById('capManualValorInline');
    
    if (cicloSelect && valorDiv) {
        cicloSelect.addEventListener('change', () => {
            const ciclo = cicloSelect.value;
            if (ciclo && capsManual[ciclo]) {
                valorDiv.textContent = capsManual[ciclo];
                console.log('📝 CAP Manual carregado:', ciclo, '=', capsManual[ciclo]);
            } else {
                valorDiv.textContent = '';
            }
        });
    }
});

console.log('✅ Sistema contenteditable carregado');


// ==================== CAPTURA MANUAL DE TECLAS ====================

function setupManualKeyCapture() {
    const valorDiv = document.getElementById('capManualValorInline');
    
    if (!valorDiv) return;
    
    console.log('🔧 Configurando captura manual de teclas');
    
    // Capturar teclas manualmente
    valorDiv.addEventListener('keydown', function(e) {
        console.log('⌨️ Tecla pressionada:', e.key);
        
        // Prevenir comportamento padrão
        e.preventDefault();
        e.stopPropagation();
        
        let textoAtual = this.textContent || '';
        
        // Processar tecla
        if (e.key >= '0' && e.key <= '9') {
            // Número
            if (textoAtual.length < 6) {
                textoAtual += e.key;
                this.textContent = textoAtual;
                console.log('✍️ Adicionado:', e.key, '→ Valor:', textoAtual);
            }
        } else if (e.key === 'Backspace') {
            // Apagar
            textoAtual = textoAtual.slice(0, -1);
            this.textContent = textoAtual;
            console.log('⌫ Apagado → Valor:', textoAtual);
        } else if (e.key === 'Delete') {
            // Limpar tudo
            this.textContent = '';
            console.log('🗑️ Limpo');
        } else if (e.key === 'Enter') {
            // Aplicar
            console.log('⏎ Enter → Aplicando CAP');
            aplicarCapManualInline();
        } else if (e.key === 'Escape') {
            // Limpar
            this.textContent = '';
            console.log('❌ Escape → Limpo');
        }
        
        return false;
    }, true);
    
    // Prevenir paste padrão e implementar manualmente
    valorDiv.addEventListener('paste', function(e) {
        e.preventDefault();
        const texto = (e.clipboardData || window.clipboardData).getData('text');
        const apenasNumeros = texto.replace(/\D/g, '').substring(0, 6);
        this.textContent = apenasNumeros;
        console.log('📋 Colado:', apenasNumeros);
    }, true);
    
    console.log('✅ Captura manual de teclas configurada');
}

// Inicializar
document.addEventListener('DOMContentLoaded', () => {
    setTimeout(() => {
        setupManualKeyCapture();
    }, 100);
});

// ==================== GERAR RELATÓRIO FINAL HTML ====================
document.getElementById('btnGerarRelatorio')?.addEventListener('click', gerarRelatorioFinal);

function gerarRelatorioFinal() {
    try {
        console.log('📊 [RELATÓRIO] Iniciando geração...');
        
        // Pegar dados do planejamento atual
        const stationAtual = document.getElementById('stationSearchInput')?.value || 'Station não selecionada';
        // Usar variável global do ciclo selecionado
        const cicloNome = cicloSelecionado && cicloSelecionado !== 'Todos' ? cicloSelecionado : '';
        
        // IMPORTANTE: Usar variável global dataCicloSelecionada, NÃO o input HTML!
        // A variável global é atualizada quando o usuário muda a data
        let dataExpedicao;
        if (dataCicloSelecionada) {
            // Converter de Date para string YYYY-MM-DD
            const ano = dataCicloSelecionada.getFullYear();
            const mes = String(dataCicloSelecionada.getMonth() + 1).padStart(2, '0');
            const dia = String(dataCicloSelecionada.getDate()).padStart(2, '0');
            dataExpedicao = `${ano}-${mes}-${dia}`;
        } else {
            // Fallback: pegar do input HTML
            dataExpedicao = document.getElementById('dataExpedicaoPlan')?.value || new Date().toISOString().split('T')[0];
        }
        
        console.log('📅 [DEBUG RELATÓRIO] dataCicloSelecionada:', dataCicloSelecionada);
        console.log('📅 [DEBUG RELATÓRIO] dataExpedicao RAW:', dataExpedicao);
        
        // Pegar LHs selecionadas (com classe row-selecionada)
        const linhasSelecionadas = document.querySelectorAll('.planejamento-table tbody tr.row-selecionada');
        
        console.log(`📊 [RELATÓRIO] ${linhasSelecionadas.length} LHs selecionadas`);
        
        if (linhasSelecionadas.length === 0) {
            alert('❌ Nenhuma LH selecionada no planejamento!\n\nPor favor, use "Sugerir Planejamento" ou selecione LHs manualmente.');
            return;
        }
        
        // Coletar dados das LHs para a tabela
        const lhsCompletas = [];
        const tosComplemento = [];
        const lhsInventario = []; // 🆕 LHs com status "Sinalizar Inventário"
        let totalPedidosPlanejados = 0;
        let numeroLH = 1;
        
        linhasSelecionadas.forEach((linha, index) => {
            const cells = linha.querySelectorAll('td');
            
            // DEBUG: Ver estrutura da primeira linha
            if (index === 0) {
                console.log('📊 [DEBUG] Primeira linha com', cells.length, 'colunas:');
                Array.from(cells).forEach((cell, i) => {
                    const texto = cell.textContent.trim();
                    console.log(`  [${i}] = ${texto.substring(0, 40)}`);
                });
            }
            
            // ESTRUTURA: [0]=Checkbox, [1]=Tipo, [2]=Status, [3]=LH TRIP, [4]=Pedidos, [5]=TOs, [6+]=Dinâmicas
            const lhTrip = cells[3]?.textContent.trim() || '';
            
            // 🔍 Verificar se é LH de inventário (coluna Status - índice 2)
            const statusTexto = cells[2]?.textContent.trim() || '';
            // Detecção flexível: aceita com/sem emoji, com/sem espaços extras
            const isInventario = /sinalizar\s*invent[aá]rio/i.test(statusTexto);
            
            // DEBUG: Mostrar status de cada LH
            console.log(`🔍 [DEBUG STATUS] LH ${lhTrip}: "${statusTexto}" -> isInventario: ${isInventario}`);
            
            // Pegar pedidos da coluna de TOs (índice 5) se tiver TOs selecionadas
            const pedidosTOsTexto = cells[5]?.textContent.trim() || '0';
            
            // Extrair número da coluna TOs (pode ter ícone ou não)
            let pedidosTOs = 0;
            const matchTOs = pedidosTOsTexto.match(/(\d[\d.,]*)/);
            if (matchTOs) {
                pedidosTOs = parseInt(matchTOs[1].replace(/\./g, '').replace(/,/g, '')) || 0;
            }
            
            // Se tem pedidos na coluna TOs, é uma LH com TOs parciais
            const temTOs = pedidosTOs > 0;
            
            console.log(`🔍 [DEBUG] LH ${cells[3]?.textContent.trim()}: Coluna TOs = "${pedidosTOsTexto}", Valor = ${pedidosTOs}, temTOs = ${temTOs}`);
            
            let pedidos = 0;
            if (temTOs) {
                // LH com TOs parciais: usar o valor da coluna TOs
                pedidos = pedidosTOs;
                console.log(`📊 [DEBUG] LH ${lhTrip} com TOs: ${pedidos} pedidos`);
            } else {
                // LH completa: pegar da coluna de pedidos (índice 4)
                const pedidosTexto = cells[4]?.textContent.trim() || '0';
                pedidos = parseInt(pedidosTexto.replace(/\./g, '').replace(/,/g, '')) || 0;
                console.log(`📊 [DEBUG] LH ${lhTrip} completa: ${pedidos} pedidos`);
            }
            
            // Buscar Origin nas colunas dinâmicas (geralmente índice 6)
            let origin = '';
            for (let i = 6; i < cells.length; i++) {
                const texto = cells[i]?.textContent.trim() || '';
                // Origin: primeira coluna que não tem ícones
                if (i === 6 && texto && texto !== '-' && !texto.includes('🚛') && !texto.includes('⏰') && !texto.includes('📍')) {
                    origin = texto;
                    break;
                }
            }
            
            // 🆕 Se é LH de inventário, coletar BRs
            if (isInventario) {
                console.log(`✅ [INVENTÁRIO] LH ${lhTrip} DETECTADA!`);
                console.log(`🔍 [INVENTÁRIO] LH ${lhTrip} detectada como "Sinalizar Inventário"`);
                
                // Buscar todos os BRs desta LH
                const pedidosLH = lhTrips[lhTrip] || [];
                
                if (pedidosLH.length > 0) {
                    pedidosLH.forEach(pedido => {
                        // Buscar coluna de BR/Tracking/Shipment ID
                        const colunaBR = Object.keys(pedido).find(col =>
                            col.toLowerCase().includes('tracking') ||
                            col.toLowerCase().includes('br') ||
                            col.toLowerCase().includes('parcel') ||
                            col.toLowerCase().includes('shipment') ||
                            col.toLowerCase().includes('package')
                        );
                        
                        const brId = pedido[colunaBR] || 'N/A';
                        
                        lhsInventario.push({
                            lhTrip: lhTrip,
                            br: brId
                        });
                    });
                    
                    console.log(`✅ [INVENTÁRIO] ${pedidosLH.length} BRs coletados da LH ${lhTrip}`);
                } else {
                    // Se não houver pedidos, adicionar pelo menos uma linha
                    lhsInventario.push({
                        lhTrip: lhTrip,
                        br: 'N/A'
                    });
                    console.log(`⚠️ [INVENTÁRIO] LH ${lhTrip} sem pedidos encontrados`);
                }
            }
            
            // Verificar se tem TOs selecionadas para esta LH (complemento)
            if (temTOs) {
                // LH com TOs parciais (complemento)
                console.log(`🔹 [DEBUG] LH ${lhTrip} tem TOs selecionadas (${pedidos} pedidos)`);
                console.log(`🔹 [DEBUG] typeof tosSelecionadasPorLH:`, typeof tosSelecionadasPorLH);
                console.log(`🔹 [DEBUG] tosSelecionadasPorLH:`, tosSelecionadasPorLH);
                console.log(`🔹 [DEBUG] Chaves em tosSelecionadasPorLH:`, Object.keys(tosSelecionadasPorLH || {}));
                console.log(`🔹 [DEBUG] tosSelecionadasPorLH[${lhTrip}]:`, tosSelecionadasPorLH ? tosSelecionadasPorLH[lhTrip] : 'undefined');
                
                // Verificar se temos acesso aos dados das TOs
                if (tosSelecionadasPorLH && tosSelecionadasPorLH[lhTrip]) {
                    const tosSelecionadas = tosSelecionadasPorLH[lhTrip];
                    console.log(`🔹 [DEBUG] TOs selecionadas encontradas: ${tosSelecionadas.size}`);
                    
                    if (tosSelecionadas && tosSelecionadas.size > 0) {
                        const tosArray = extrairTOsDaLH(lhTrip, dadosPlanilha[lhTrip]);
                        const pedidosLH = lhTrips[lhTrip] || [];
                        
                        tosArray.forEach(to => {
                            if (tosSelecionadas.has(to.toId)) {
                                // Buscar BRs (pedidos) desta TO
                                const brs = [];
                                pedidosLH.forEach(pedido => {
                                    const colunaTO = Object.keys(pedido).find(col =>
                                        col.toLowerCase().includes('to id') ||
                                        col.toLowerCase().includes('to_id') ||
                                        col.toLowerCase().includes('transfer order') ||
                                        col.toLowerCase() === 'to'
                                    );
                                    const toId = pedido[colunaTO] || 'SEM_TO';
                                    
                                    if (toId === to.toId) {
                                        // Debug: mostrar colunas disponíveis
                                        if (brs.length === 0) {
                                            console.log(`🔍 [DEBUG] Colunas disponíveis no pedido:`, Object.keys(pedido));
                                        }
                                        
                                        // Buscar coluna de BR/Tracking/Shipment ID
                                        const colunaBR = Object.keys(pedido).find(col =>
                                            col.toLowerCase().includes('tracking') ||
                                            col.toLowerCase().includes('br') ||
                                            col.toLowerCase().includes('parcel') ||
                                            col.toLowerCase().includes('shipment') ||
                                            col.toLowerCase().includes('package')
                                        );
                                        
                                        console.log(`🔍 [DEBUG] Coluna BR encontrada: "${colunaBR}", Valor: "${pedido[colunaBR]}"`);
                                        const brId = pedido[colunaBR] || 'N/A';
                                        brs.push(brId);
                                    }
                                });
                                
                                console.log(`🔹 [DEBUG] TO ${to.toId}: ${to.qtdPedidos} pedidos, ${brs.length} BRs`);
                                
                                // 📊 FORMATO CORRETO: Agrupar por TO com array de BRs
                                tosComplemento.push({
                                    lhTrip: lhTrip,
                                    origin: origin || '-',
                                    toId: to.toId,
                                    pedidos: to.qtdPedidos || brs.length,
                                    brs: brs.length > 0 ? brs : ['N/A']
                                });
                            }
                        });
                    }
                } else {
                    console.warn(`⚠️ [WARN] LH ${lhTrip} tem TOs mas tosSelecionadasPorLH não está definido ou não tem dados`);
                }
            } else {
                // LH completa
                console.log(`✅ [DEBUG] LH ${lhTrip} completa: ${pedidos} pedidos`);
                lhsCompletas.push({
                    numero: numeroLH++,
                    lhTrip,
                    origin: origin || '-',
                    pedidos
                });
            }
            
            totalPedidosPlanejados += pedidos;
        });
        
        // Calcular backlog (pegar do sistema se disponível)
        let totalBacklog = 0;
        if (typeof pedidosBacklogPorStatus !== 'undefined') {
            totalBacklog = pedidosBacklogPorStatus.length;
        }
        
        // Obter informação do CAP usado (manual ou automático)
        let capUsado = 0;
        let tipoCAP = 'Não especificado';
        if (cicloNome && cicloNome !== 'Todos') {
            const capManual = obtemCapManual(cicloNome);
            if (capManual !== null) {
                capUsado = capManual;
                tipoCAP = 'CAP Manual';
            } else {
                capUsado = obterCapacidadeCiclo(cicloNome);
                tipoCAP = 'CAP Automático';
            }
        }
        
        // Calcular tempo REAL de execução
        let tempoFormatado = '0s';
        if (tempoInicioExecucao && tempoFimExecucao) {
            const segundosTotal = Math.floor((tempoFimExecucao - tempoInicioExecucao) / 1000);
            const minutos = Math.floor(segundosTotal / 60);
            const segundos = segundosTotal % 60;
            
            if (minutos > 0) {
                tempoFormatado = segundos > 0 ? `${minutos}m ${segundos}s` : `${minutos}m`;
            } else {
                tempoFormatado = `${segundos}s`;
            }
        }
        
        // Data formatada (DATA DO PLANEJAMENTO, não data de hoje)
        // Parse seguro para evitar problemas de timezone
        console.log('📅 [DEBUG] Antes do split:', dataExpedicao);
        const [ano, mes, dia] = dataExpedicao.split('-').map(Number);
        console.log('📅 [DEBUG] Após split:', { ano, mes, dia });
        const dataObj = new Date(ano, mes - 1, dia); // mes - 1 porque Date usa 0-11
        console.log('📅 [DEBUG] dataObj:', dataObj);
        const dataFormatada = dataObj.toLocaleDateString('pt-BR');
        console.log('📅 [DEBUG] dataFormatada:', dataFormatada);
        // NÃO limpar o nome da estação - manter completo
        const nomeEstacao = stationAtual;
        
        console.log('📊 [RELATÓRIO] Dados coletados:', {
            totalPedidosPlanejados,
            totalBacklog,
            totalLHs: linhasSelecionadas.length,
            ciclo: cicloNome,
            estacao: nomeEstacao,
            dataExpedicao: dataExpedicao,
            dataFormatada: dataFormatada,
            primeiraLH: lhsCompletas[0] || tosComplemento[0]
        });
        
        // Gerar HTML do relatório
        const htmlRelatorio = `
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relatório de Planejamento - ${nomeEstacao} - ${dataFormatada}</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            background: linear-gradient(135deg, #FFE5DC 0%, #FFD4C4 25%, #FFF5F2 50%, #FFEAE0 75%, #FFE0D1 100%);
            padding: 30px 20px;
            min-height: 100vh;
            -webkit-font-smoothing: antialiased;
            -moz-osx-font-smoothing: grayscale;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 16px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
        }
        
        .header {
            background: white;
            border-bottom: 1px solid #f0f0f0;
            padding: 32px 48px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        
        .header-left {
            display: flex;
            align-items: center;
            gap: 20px;
        }
        
        .shopee-logo {
            width: 48px;
            height: 48px;
            background: linear-gradient(135deg, #EE4D2D 0%, #FF6533 100%);
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 28px;
            font-weight: 900;
            color: white;
            flex-shrink: 0;
            box-shadow: 0 4px 12px rgba(238, 77, 45, 0.25);
        }
        
        .header-content {
            flex: 1;
        }
        
        .header h1 {
            font-size: 28px;
            font-weight: 700;
            color: #1a1a1a;
            margin-bottom: 4px;
            letter-spacing: -0.5px;
        }
        
        .header-subtitle {
            font-size: 14px;
            color: #666;
            font-weight: 500;
        }
        
        .header-badge {
            background: #EE4D2D;
            color: white;
            padding: 8px 16px;
            border-radius: 8px;
            font-size: 13px;
            font-weight: 600;
            letter-spacing: 0.3px;
        }
        
        .info-cards {
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 1px;
            background: #f0f0f0;
            border-top: 1px solid #f0f0f0;
            border-bottom: 1px solid #f0f0f0;
        }
        
        .info-card {
            background: white;
            padding: 24px 20px;
            text-align: center;
            transition: background 0.2s;
        }
        
        .info-card:hover {
            background: #fafafa;
        }
        
        .info-card-label {
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.8px;
            color: #999;
            font-weight: 600;
            margin-bottom: 10px;
        }
        
        .info-card-value {
            font-size: 26px;
            font-weight: 700;
            color: #1a1a1a;
            letter-spacing: -0.5px;
        }
        
        .stats-section {
            padding: 48px;
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 24px;
            background: #fafafa;
        }
        
        .stat-card {
            background: white;
            padding: 32px 24px;
            border-radius: 12px;
            text-align: center;
            border: 1px solid #f0f0f0;
            transition: all 0.3s ease;
            position: relative;
            overflow: hidden;
        }
        
        .stat-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 3px;
            background: linear-gradient(90deg, var(--card-color) 0%, var(--card-color-light) 100%);
        }
        
        .stat-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.08);
            border-color: #e0e0e0;
        }
        
        .stat-card.purple { --card-color: #8B5CF6; --card-color-light: #A78BFA; }
        .stat-card.orange { --card-color: #EE4D2D; --card-color-light: #FF6533; }
        .stat-card.green { --card-color: #10B981; --card-color-light: #34D399; }
        .stat-card.blue { --card-color: #3B82F6; --card-color-light: #60A5FA; }
        
        .stat-card-icon {
            font-size: 40px;
            margin-bottom: 16px;
            opacity: 0.9;
        }
        
        .stat-card-value {
            font-size: 40px;
            font-weight: 800;
            color: #1a1a1a;
            margin-bottom: 8px;
            letter-spacing: -1px;
        }
        
        .stat-card-label {
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.8px;
            color: #666;
            font-weight: 600;
        }
        
        .table-section {
            padding: 48px;
        }
        
        .table-section h2 {
            font-size: 20px;
            color: #1a1a1a;
            margin-bottom: 24px;
            font-weight: 700;
            letter-spacing: -0.3px;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .table-section h2::before {
            content: '';
            width: 4px;
            height: 24px;
            background: linear-gradient(180deg, #EE4D2D 0%, #FF6533 100%);
            border-radius: 2px;
        }
        
        .lhs-table {
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            border: 1px solid #f0f0f0;
            border-radius: 12px;
            overflow: hidden;
        }
        
        .lhs-table thead {
            background: #fafafa;
        }
        
        .lhs-table th {
            padding: 16px 20px;
            text-align: left;
            font-weight: 600;
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.8px;
            color: #666;
            border-bottom: 1px solid #f0f0f0;
        }
        
        .lhs-table tbody tr {
            transition: background 0.2s;
        }
        
        .lhs-table tbody tr:hover {
            background: #fafafa;
        }
        
        .lhs-table tbody tr:not(:last-child) td {
            border-bottom: 1px solid #f5f5f5;
        }
        
        .lhs-table td {
            padding: 18px 20px;
            font-size: 14px;
            color: #333;
        }
        
        .lhs-table td:first-child {
            font-weight: 700;
            color: #999;
            text-align: center;
            width: 80px;
            font-size: 13px;
        }
        
        .lhs-table td:nth-child(2) {
            font-weight: 600;
            font-size: 14px;
            color: #1a1a1a;
            letter-spacing: 0.2px;
        }
        
        .lhs-table td:nth-child(3) {
            color: #666;
        }
        
        .lhs-table td:nth-child(4) {
            font-weight: 700;
            text-align: right;
            color: #EE4D2D;
            font-size: 15px;
        }
        
        .lhs-table td:last-child {
            text-align: left;
        }
        
        .badge-completa {
            display: inline-block;
            padding: 6px 12px;
            background: linear-gradient(135deg, #10B981 0%, #34D399 100%);
            color: white;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 600;
            letter-spacing: 0.3px;
        }
        
        .badge-complemento {
            display: inline-block;
            padding: 6px 12px;
            background: linear-gradient(135deg, #EE4D2D 0%, #FF6533 100%);
            color: white;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 600;
            letter-spacing: 0.3px;
            margin-bottom: 8px;
        }
        
        .tipo-complemento {
            display: flex;
            flex-direction: column;
            gap: 8px;
        }
        
        .tos-list {
            display: flex;
            flex-direction: column;
            gap: 4px;
            padding-left: 8px;
        }
        
        .to-item {
            font-size: 12px;
            color: #666;
            line-height: 1.6;
        }
        
        .to-item strong {
            color: #1a1a1a;
            font-weight: 600;
        }
        
        .tos-table {
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            border: 1px solid #f0f0f0;
            border-radius: 12px;
            overflow: hidden;
        }
        
        .tos-table thead {
            background: #fafafa;
        }
        
        .tos-table th {
            padding: 16px 20px;
            text-align: left;
            font-weight: 600;
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.8px;
            color: #666;
            border-bottom: 1px solid #f0f0f0;
        }
        
        .tos-table tbody tr:not(.brs-row) {
            transition: background 0.2s;
        }
        
        .tos-table tbody tr:not(.brs-row):hover {
            background: #fafafa;
        }
        
        .tos-table tbody tr:not(:last-child):not(.brs-row) td {
            border-bottom: 1px solid #f5f5f5;
        }
        
        .tos-table td {
            padding: 18px 20px;
            font-size: 14px;
            color: #333;
        }
        
        .to-id-cell {
            font-weight: 600;
            color: #1a1a1a;
            font-family: 'SF Mono', 'Monaco', 'Consolas', monospace;
            font-size: 13px;
        }
        
        .tos-table td:nth-child(4) {
            font-weight: 700;
            color: #EE4D2D;
            font-size: 15px;
        }
        
        .btn-expandir {
            background: linear-gradient(135deg, #EE4D2D 0%, #FF6533 100%);
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }
        
        .btn-expandir:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(238, 77, 45, 0.3);
        }
        
        .btn-expandir span {
            font-size: 10px;
        }
        
        .brs-row {
            background: #fafafa;
        }
        
        .brs-container {
            padding: 20px;
        }
        
        .brs-header {
            font-size: 13px;
            font-weight: 600;
            color: #666;
            margin-bottom: 12px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .brs-list {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
        }
        
        .br-item {
            background: white;
            border: 1px solid #e0e0e0;
            padding: 6px 12px;
            border-radius: 6px;
            font-size: 12px;
            font-family: 'SF Mono', 'Monaco', 'Consolas', monospace;
            color: #333;
            transition: all 0.2s;
        }
        
        .br-item:hover {
            border-color: #EE4D2D;
            background: #FFF5F2;
            color: #EE4D2D;
        }
        
        /* Estilos para tabela de inventário */
        .inventario-table {
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            border: 1px solid #f0f0f0;
            border-radius: 12px;
            overflow: hidden;
        }
        
        .inventario-table thead {
            background: linear-gradient(135deg, #8B5CF6 0%, #A78BFA 100%);
        }
        
        .inventario-table th {
            padding: 16px 20px;
            text-align: left;
            font-weight: 600;
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.8px;
            color: white;
            border-bottom: 1px solid #f0f0f0;
        }
        
        .inventario-table tbody tr {
            transition: background 0.2s;
        }
        
        .inventario-table tbody tr:hover {
            background: #fafafa;
        }
        
        .inventario-table tbody tr:not(:last-child) td {
            border-bottom: 1px solid #f5f5f5;
        }
        
        .inventario-table td {
            padding: 18px 20px;
            font-size: 14px;
            color: #333;
        }
        
        .inventario-table td:first-child {
            font-weight: 600;
            color: #8B5CF6;
            font-size: 14px;
            letter-spacing: 0.2px;
        }
        
        .inventario-table td:nth-child(2) {
            font-family: 'SF Mono', 'Monaco', 'Consolas', monospace;
            font-size: 13px;
            color: #1a1a1a;
        }
        
        .footer {
            padding: 32px 48px;
            background: #fafafa;
            border-top: 1px solid #f0f0f0;
            text-align: center;
            color: #999;
            font-size: 12px;
        }
        
        .footer p {
            margin: 4px 0;
            font-weight: 500;
        }
        
        .footer strong {
            color: #666;
            font-weight: 600;
        }
        
        @media print {
            body {
                background: white;
                padding: 0;
            }
            .container {
                box-shadow: none;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="header-left">
                <div class="shopee-logo">S</div>
                <div class="header-content">
                    <h1>Relatório de Planejamento</h1>
                    <p class="header-subtitle">Sistema de Gerenciamento Shopee</p>
                </div>
            </div>
            <div class="header-badge">PLANEJAMENTO HUB</div>
        </div>
        
        <div class="info-cards">
            <div class="info-card">
                <div class="info-card-label">Estação</div>
                <div class="info-card-value">${nomeEstacao}</div>
            </div>
            <div class="info-card">
                <div class="info-card-label">Ciclo</div>
                <div class="info-card-value">${cicloNome || 'Não especificado'}</div>
            </div>
            <div class="info-card">
                <div class="info-card-label">Data de Expedição</div>
                <div class="info-card-value">${dataFormatada}</div>
            </div>
            <div class="info-card">
                <div class="info-card-label">Tempo Total</div>
                <div class="info-card-value">${tempoFormatado}</div>
            </div>
            <div class="info-card">
                <div class="info-card-label">${tipoCAP}</div>
                <div class="info-card-value">${capUsado > 0 ? capUsado.toLocaleString('pt-BR') : '-'}</div>
            </div>
        </div>
        
        <div class="stats-section">
            <div class="stat-card purple">
                <div class="stat-card-icon">📊</div>
                <div class="stat-card-value">${(totalPedidosPlanejados + totalBacklog).toLocaleString('pt-BR')}</div>
                <div class="stat-card-label">Pedidos Totais</div>
            </div>
            
            <div class="stat-card orange">
                <div class="stat-card-icon">📦</div>
                <div class="stat-card-value">${totalPedidosPlanejados.toLocaleString('pt-BR')}</div>
                <div class="stat-card-label">Pedidos Planejados por LH</div>
            </div>
            
            <div class="stat-card green">
                <div class="stat-card-icon">🚚</div>
                <div class="stat-card-value">${linhasSelecionadas.length}</div>
                <div class="stat-card-label">Quantidade de LHs</div>
            </div>
            
            <div class="stat-card blue">
                <div class="stat-card-icon">⬅️</div>
                <div class="stat-card-value">${totalBacklog.toLocaleString('pt-BR')}</div>
                <div class="stat-card-label">Backlog</div>
            </div>
        </div>
        
        <!-- Seção LHs Completas -->
        ${lhsCompletas.length > 0 ? `
        <div class="table-section">
            <h2>📦 LHs Completas (${lhsCompletas.length})</h2>
            <table class="lhs-table">
                <thead>
                    <tr>
                        <th>Nº</th>
                        <th>LH TRIP</th>
                        <th>Origem</th>
                        <th>Pedidos Planejados</th>
                    </tr>
                </thead>
                <tbody>
                    ${lhsCompletas.map(lh => `
                    <tr>
                        <td>${lh.numero}</td>
                        <td>${lh.lhTrip}</td>
                        <td>${lh.origin}</td>
                        <td>${lh.pedidos.toLocaleString('pt-BR')}</td>
                    </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
        ` : ''}
        
        <!-- Seção TOs de Complemento -->
        ${tosComplemento.length > 0 ? `
        <div class="table-section">
            <h2>🔹 TOs de Complemento (${tosComplemento.reduce((sum, to) => sum + to.pedidos, 0)} pedidos)</h2>
            <table class="tos-table">
                <thead>
                    <tr>
                        <th>LH TRIP</th>
                        <th>Origem</th>
                        <th>TO ID</th>
                        <th>Pedidos</th>
                        <th>Ações</th>
                    </tr>
                </thead>
                <tbody>
                    ${tosComplemento.map((to, index) => `
                    <tr>
                        <td>${to.lhTrip}</td>
                        <td>${to.origin}</td>
                        <td class="to-id-cell">${to.toId}</td>
                        <td>${to.pedidos.toLocaleString('pt-BR')}</td>
                        <td>
                            <button class="btn-expandir" onclick="toggleBRs('to-${index}')">
                                <span id="icon-to-${index}">▶</span> Ver BRs (${to.brs.length})
                            </button>
                        </td>
                    </tr>
                    <tr id="to-${index}" class="brs-row" style="display: none;">
                        <td colspan="5">
                            <div class="brs-container">
                                <div class="brs-header">Pedidos (BRs) da TO ${to.toId}:</div>
                                <div class="brs-list">
                                    ${to.brs.map(br => `<span class="br-item">${br}</span>`).join('')}
                                </div>
                            </div>
                        </td>
                    </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
        ` : ''}
        
        <!-- Seção LHs Lixo Sistêmico -->
        ${lhsLixoSistemico.length > 0 ? `
        <div class="table-section" style="background: #fffaeb; border-left: 4px solid #ff9800; padding: 15px; margin-top: 20px;">
            <h2 style="color: #856404;">🗑️ LHs Lixo Sistêmico (${lhsLixoSistemico.length} LHs, ${lhsLixoSistemico.reduce((sum, lh) => sum + (lh.pedidos || 0), 0)} pedidos)</h2>
            <p style="margin: 10px 0; color: #856404; font-size: 14px;">
                ⚠️ LHs sem origin/destination/previsão - Pedidos incluídos no backlog automaticamente
            </p>
            <table class="lhs-table">
                <thead>
                    <tr>
                        <th>LH TRIP</th>
                        <th>Pedidos</th>
                        <th>Origin</th>
                        <th>Destination</th>
                        <th>Previsão</th>
                    </tr>
                </thead>
                <tbody>
                    ${lhsLixoSistemico.map(lh => `
                    <tr style="background: #fff;">
                        <td>${lh.lh_trip || '-'}</td>
                        <td>${lh.pedidos || 0}</td>
                        <td>${lh.origin || '-'}</td>
                        <td>${lh.destination || '-'}</td>
                        <td>${lh.previsao_data || '-'}</td>
                    </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
        ` : ''}
        
        <!-- Seção Análise Inventário -->
        ${lhsInventario.length > 0 ? `
        <div class="table-section">
            <h2>🔍 Análise Inventário (${lhsInventario.length} pedidos)</h2>
            <p style="margin: 10px 0; color: #666; font-size: 14px;">
                📦 LHs com status "Sinalizar Inventário" - Encaminhar para o time de inventário
            </p>
            <table class="inventario-table">
                <thead>
                    <tr>
                        <th>LH TRIP</th>
                        <th>BR (Pedido)</th>
                    </tr>
                </thead>
                <tbody>
                    ${lhsInventario.map((item, index) => `
                    <tr>
                        <td>${item.lhTrip}</td>
                        <td>${item.br}</td>
                    </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
        ` : ''}
        
        <div class="footer">
            <p><strong>Relatório gerado automaticamente pelo Sistema de Gerenciamento Shopee</strong></p>
            <p>Data de geração: ${new Date().toLocaleString('pt-BR')}</p>
        </div>
    </div>
    
    <script>
        function toggleBRs(rowId) {
            const row = document.getElementById(rowId);
            const icon = document.getElementById('icon-' + rowId);
            
            if (row.style.display === 'none') {
                row.style.display = 'table-row';
                icon.textContent = '▼';
            } else {
                row.style.display = 'none';
                icon.textContent = '▶';
            }
        }
    </script>
</body>
</html>
`;
        
        // Criar blob e baixar
        const blob = new Blob([htmlRelatorio], { type: 'text/html;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Relatorio_Planejamento_${nomeEstacao.replace(/\s/g, '_')}_${dataExpedicao}.html`;
        a.style.display = 'none';
        document.body.appendChild(a);
        a.click();
        setTimeout(() => {
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }, 100);
        
        console.log('✅ [RELATÓRIO] HTML gerado com sucesso!');
        alert(`✅ Relatório de Planejamento gerado com sucesso!\n\n📊 ${linhasSelecionadas.length} LHs planejadas\n📦 ${totalPedidosPlanejados.toLocaleString('pt-BR')} pedidos\n\n📄 Arquivo: ${a.download}`);
        
    } catch (error) {
        console.error('❌ [RELATÓRIO] Erro ao gerar:', error);
        alert('❌ Erro ao gerar relatório!\n\n' + error.message);
    }
}

// ======================= SINCRONIZAR LHs SPX =======================

/**
 * Sincroniza LHs visíveis com dados do SPX e gera CSV
 */
async function sincronizarLHsSPX() {
    try {
        // Pegar todas as LHs visíveis na tabela
        const lhsVisiveis = obterLHsVisiveis();
        
        if (lhsVisiveis.length === 0) {
            alert('❌ Nenhuma LH encontrada para sincronizar!');
            return;
        }
        
        console.log(`🔍 [SPX] Sincronizando ${lhsVisiveis.length} LH(s)...`);
        
        // Verificar se tem pasta da station
        if (!pastaStationAtual) {
            alert('❌ Pasta da station não encontrada!\nCarregue os dados primeiro.');
            return;
        }
        
        // Verificar se tem nome da station
        if (!stationAtualNome) {
            alert('❌ Nome da station não identificado!\nCarregue os dados primeiro.');
            return;
        }
        
        // Mostrar loading
        const btnSincronizar = document.getElementById('btnSincronizarLHs');
        const textoOriginal = btnSincronizar.innerHTML;
        btnSincronizar.disabled = true;
        btnSincronizar.innerHTML = '⏳ Sincronizando...';
        
        // Chamar IPC para buscar no SPX
        const resultado = await ipcRenderer.invoke('sincronizar-lhs-spx', {
            lhIds: lhsVisiveis,
            stationFolder: pastaStationAtual,
            currentStationName: stationAtualNome
        });
        
        if (resultado.success) {
            console.log('✅ [SPX] Sincronização concluída:', resultado.data);
            
            // Mostrar resumo
            const msg = `✅ Sincronização SPX concluída!\n\n` +
                  `📊 Total de LHs: ${resultado.data.total}\n` +
                  `✅ Encontradas: ${resultado.data.encontradas}\n` +
                  `❌ Não encontradas: ${resultado.data.erros}\n\n`;
            
            if (resultado.data.csvPath) {
                alert(msg + `📄 Relatório CSV gerado:\n${resultado.data.csvPath}\n\nAbra o arquivo para ver os detalhes completos!`);
                
                // Opcionalmente, abrir o arquivo automaticamente
                if (confirm('Deseja abrir o relatório agora?')) {
                    await ipcRenderer.invoke('abrir-arquivo', resultado.data.csvPath);
                }
                
                // Processar e atualizar visual na tabela
                if (resultado.data.resultados && resultado.data.resultados.length > 0) {
                    processarResultadosSPXComCSV(resultado.data.resultados);
                }
            } else {
                alert(msg + '⚠️ Nenhuma LH foi encontrada no SPX.');
            }
        } else {
            console.error('❌ [SPX] Erro:', resultado.error);
            alert(`❌ Erro na sincronização:\n${resultado.error}`);
        }
        
        // Restaurar botão
        btnSincronizar.disabled = false;
        btnSincronizar.innerHTML = textoOriginal;
        
    } catch (error) {
        console.error('❌ [SPX] Erro fatal:', error);
        alert(`❌ Erro fatal:\n${error.message}`);
    }
}

/**
 * Obtém lista de LHs visíveis na tabela atual
 */
function obterLHsVisiveis() {
    const lhs = [];
    
    // Tentar múltiplas estratégias para encontrar a tabela
    let tbody = null;
    
    // Estratégia 1: Verificar qual aba está ativa
    const abaAtiva = document.querySelector('.tab.active');
    console.log('🔍 [SPX] Aba ativa:', abaAtiva ? abaAtiva.getAttribute('data-tab') : 'nenhuma');
    
    if (abaAtiva) {
        const dataTab = abaAtiva.getAttribute('data-tab');
        if (dataTab === 'planejamento') {
            tbody = document.getElementById('tbodyPlanejamento');
            console.log('📋 [SPX] Usando tabela: Planejamento Hub');
        } else if (dataTab === 'lh-trips') {
            tbody = document.getElementById('tbodyLHTrips');
            console.log('🚚 [SPX] Usando tabela: LH Trips');
        }
    }
    
    // Estratégia 2: Se não encontrou, tenta todas as tabelas visíveis
    if (!tbody) {
        console.log('⚠️ [SPX] Tentando encontrar tabela visível...');
        const tbodies = [
            document.getElementById('tbodyPlanejamento'),
            document.getElementById('tbodyLHTrips')
        ];
        
        for (const tb of tbodies) {
            if (tb && tb.offsetParent !== null) { // Verifica se está visível
                tbody = tb;
                console.log('✅ [SPX] Tabela visível encontrada!');
                break;
            }
        }
    }
    
    if (!tbody) {
        console.error('❌ [SPX] Nenhuma tabela encontrada!');
        return lhs;
    }
    
    const linhas = tbody.querySelectorAll('tr');
    console.log(`🔍 [SPX] Encontradas ${linhas.length} linhas na tabela`);
    
    linhas.forEach((linha, index) => {
        // Procurar pela célula com classe 'lh-trip-cell' ao invés de usar índice fixo
        const celulaLH = linha.querySelector('td.lh-trip-cell');
        
        if (celulaLH) {
            const lhId = celulaLH.textContent.trim();
            console.log(`   🔍 Linha ${index}: LH = "${lhId}"`);
            
            if (lhId && lhId !== '-' && lhId !== '' && !lhs.includes(lhId)) {
                lhs.push(lhId);
                console.log(`   ✅ ${index + 1}. ${lhId}`);
            }
        } else {
            console.log(`   ⚠️ Linha ${index}: Sem célula lh-trip-cell`);
        }
    });
    
    if (lhs.length > 5) {
        console.log(`   ... e mais ${lhs.length - 5} LHs`);
    }
    
    console.log(`✅ [SPX] Total de LHs encontradas: ${lhs.length}`);
    return lhs;
}

/**
 * Processa resultados do SPX e atualiza status visual (versão CSV)
 */
function processarResultadosSPXComCSV(resultados) {
    console.log('📊 [SPX] Processando resultados do CSV e validando status...');
    
    // Encontrar tabela ativa
    let tbody = null;
    const abaAtiva = document.querySelector('.tab.active');
    
    if (abaAtiva) {
        const dataTab = abaAtiva.getAttribute('data-tab');
        if (dataTab === 'planejamento') {
            tbody = document.getElementById('tbodyPlanejamento');
        } else if (dataTab === 'lh-trips') {
            tbody = document.getElementById('tbodyLHTrips');
        }
    }
    
    // Fallback
    if (!tbody) {
        const tbodies = [
            document.getElementById('tbodyPlanejamento'),
            document.getElementById('tbodyLHTrips')
        ];
        for (const tb of tbodies) {
            if (tb && tb.offsetParent !== null) {
                tbody = tb;
                break;
            }
        }
    }
    
    if (!tbody) {
        console.error('❌ [SPX] Nenhuma tabela encontrada para atualizar!');
        return;
    }
    
    let atualizadas = 0;
    let statusAtualizados = 0;
    let horariosAtualizados = 0;
    
    resultados.forEach(resultado => {
        const lhId = resultado.lh_id;
        const dados = resultado.dados;
        
        if (!dados) return;
        
        const linhas = tbody.querySelectorAll('tr');
        linhas.forEach(linha => {
            // Procurar pela célula com classe 'lh-trip-cell'
            const celulaLH = linha.querySelector('td.lh-trip-cell');
            
            if (celulaLH && celulaLH.textContent.trim() === lhId) {
                // Extrair informações do SPX
                const stations = dados.trip_station || [];
                const destino = stations[stations.length - 1] || {};
                
                // Status do SPX
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
                const statusSPX = statusMap[dados.trip_status] || dados.trip_status;
                
                // Chegada Real (ata ou eta)
                const ata = destino.ata && destino.ata > 0 ? new Date(destino.ata * 1000) : null;
                const eta = destino.eta && destino.eta > 0 ? new Date(destino.eta * 1000) : null;
                const chegadaReal = ata || eta;
                const chegadaRealStr = chegadaReal ? chegadaReal.toLocaleString('pt-BR') : "Em trânsito";
                
                // Mapeamento Status SPX → Status Front
                const statusFrontMap = {
                    "Finalizado": { codigo: "P0", texto: "✅ No Piso", classe: "status-p0", icone: "✅" },
                    "Desembarcando": { codigo: "P0D", texto: "🚚 Aguard. Descarregamento", classe: "status-p0-desc", icone: "🚚" },
                    "Chegou no Destino": { codigo: "P0D", texto: "🚚 Aguard. Descarregamento", classe: "status-p0-desc", icone: "🚚" },
                    "Em Trânsito": null, // Mantém status calculado
                    "Embarcando": null,
                    "Criado": null,
                    "Aguardando Motorista": null
                };
                
                // Procurar coluna STATUS LH dinamicamente (corrigido para não pegar coluna LH TRIP)
                const todasColunas = linha.querySelectorAll('td');
                let celulaStatus = null;
                let indexStatus = -1;
                
                // IMPORTANTE: Ignorar a primeira coluna (TIPO) que também tem badges
                // Começar da segunda coluna em diante
                for (let i = 1; i < todasColunas.length; i++) {
                    const celula = todasColunas[i];
                    const badge = celula.querySelector('.badge, .status-badge');
                    const texto = celula.textContent.trim();
                    
                    // IGNORAR se for coluna TIPO (Normal, Backlog)
                    if (texto === 'Normal' || texto === 'Backlog') {
                        continue;
                    }
                    
                    // IGNORAR se for coluna LH TRIP (código da LH)
                    if (celula.classList.contains('lh-trip-cell')) {
                        continue;
                    }
                    
                    // Verificar se contém STATUS conhecidos (não TIPO)
                    const isStatusColumn = 
                        texto.includes('No Piso') ||
                        texto.includes('Aguard') ||
                        texto.includes('Descarregamento') ||
                        texto.includes('Sinalizar') ||
                        texto.includes('Inventário') ||
                        texto.includes('Em transito') ||
                        texto.includes('Em trânsito') ||
                        texto.includes('fora do prazo') ||
                        texto.includes('dentro do prazo') ||
                        texto.includes('No Hub') ||
                        (badge && (
                            badge.classList.contains('status-p0') ||
                            badge.classList.contains('status-p0-desc') ||
                            badge.classList.contains('status-p1') ||
                            badge.classList.contains('status-p2') ||
                            badge.classList.contains('status-p3') ||
                            badge.classList.contains('status-p0i')
                        ));
                    
                    if (isStatusColumn) {
                        celulaStatus = celula;
                        indexStatus = i;
                        console.log(`   🎯 Coluna STATUS encontrada no índice ${i}: "${texto}"`);
                        break;
                    }
                }
                
                // Procurar coluna PREVISÃO HORA
                let celulaPrevisaoHora = null;
                for (let i = 0; i < todasColunas.length; i++) {
                    const celula = todasColunas[i];
                    const texto = celula.textContent.trim();
                    // Verificar se tem formato de hora (HH:MM:SS ou HH:MM)
                    if (/^\d{2}:\d{2}(:\d{2})?$/.test(texto)) {
                        celulaPrevisaoHora = celula;
                        break;
                    }
                }
                
                // Procurar coluna TEMPO P/ CORTE (com debug)
                let celulaTempoCorte = null;
                console.log(`   🔍 Procurando coluna TEMPO... Total colunas: ${todasColunas.length}`);
                
                for (let i = todasColunas.length - 1; i >= 0; i--) {
                    const celula = todasColunas[i];
                    const badge = celula.querySelector('.badge');
                    const texto = celula.textContent.trim();
                    
                    if (i >= todasColunas.length - 3) { // Debug últimas 3 colunas
                        console.log(`      Col ${i}: "${texto}" (badge: ${!!badge})`);
                    }
                    
                    // Verificar se é a coluna de tempo
                    const isTempoColumn = 
                        (badge && (
                            badge.classList.contains('tempo-ok') ||
                            badge.classList.contains('tempo-apertado') ||
                            badge.classList.contains('tempo-atrasado')
                        )) ||
                        texto.includes('min') || 
                        texto.includes('Ciclo') ||
                        texto.includes('encerrado') ||
                        /^\d+h\d+min$/.test(texto) ||
                        /^-?\d+min$/.test(texto) ||
                        texto === '-';
                    
                    if (isTempoColumn) {
                        celulaTempoCorte = celula;
                        console.log(`   ⏰ Coluna TEMPO encontrada no índice ${i}: "${texto}"`);
                        break;
                    }
                }
                
                if (celulaStatus) {
                    const statusAtualTexto = celulaStatus.textContent.trim();
                    const novoStatusFront = statusFrontMap[statusSPX];
                    
                    console.log(`🔍 [DEBUG] LH: ${lhId}`);
                    console.log(`   Status atual (front): "${statusAtualTexto}"`);
                    console.log(`   Status SPX: "${statusSPX}"`);
                    console.log(`   Novo status mapeado:`, novoStatusFront);
                    
                    // Verificar se precisa atualizar o status
                    // REGRA: Se SPX tem status definitivo (Finalizado, Desembarcando), SEMPRE atualiza
                    const statusDefinitivos = ['Finalizado', 'Desembarcando', 'Chegou no Destino'];
                    const ehStatusDefinitivo = statusDefinitivos.includes(statusSPX);
                    
                    // REGRA: Atualizar se tiver novo status E (for definitivo OU status atual for genérico)
                    const statusGenericos = [
                        'Em transito',
                        'Em trânsito', 
                        'fora do prazo',
                        'Sinalizar',
                        'Aguard',
                        'P2',
                        'P3'
                    ];
                    const ehStatusGenerico = statusGenericos.some(s => statusAtualTexto.includes(s));
                    
                    const deveAtualizar = novoStatusFront && (ehStatusDefinitivo || ehStatusGenerico);
                    
                    console.log(`   Deve atualizar? ${deveAtualizar} (definitivo: ${ehStatusDefinitivo}, genérico: ${ehStatusGenerico})`);
                    
                    if (deveAtualizar && novoStatusFront) {
                        // IMPORTANTE: Salvar validação SPX no cache ANTES de atualizar visual
                        cacheSPX.set(lhId, {
                            status: novoStatusFront.texto,
                            statusCodigo: novoStatusFront.codigo,
                            statusSPX: statusSPX,
                            chegadaReal: chegadaReal ? chegadaReal.toISOString() : null,
                            timestamp: new Date().toISOString()
                        });
                        console.log(`   💾 Validação SPX salva no cache: ${lhId} → ${novoStatusFront.codigo}`);
                        
                        // Atualizar o status visual
                        const badgeExistente = celulaStatus.querySelector('.badge, .status-badge');
                        if (badgeExistente) {
                            badgeExistente.textContent = novoStatusFront.texto;
                            badgeExistente.className = `badge ${novoStatusFront.classe}`;
                        } else {
                            celulaStatus.innerHTML = `<span class="badge ${novoStatusFront.classe}">${novoStatusFront.texto}</span>`;
                        }
                        
                        // IMPORTANTE: Atualizar dados subjacentes quando mudar para "No Piso" ou "Aguard. Descarregamento"
                        if (novoStatusFront.codigo === 'P0' || novoStatusFront.codigo === 'P0D') {
                            console.log(`   🔄 Atualizando dados subjacentes para ${lhId}...`);
                            
                            // 1. Atualizar dados da LH no objeto lhTripsPlanejáveis
                            if (lhTripsPlanejáveis && lhTripsPlanejáveis[lhId]) {
                                const pedidosLH = lhTripsPlanejáveis[lhId];
                                
                                // Atualizar o objeto dadosPlanilhaLH (se existir)
                                const dadosLH = buscarDadosPlanilhaPorStation(lhId);
                                if (dadosLH) {
                                    // Marcar que chegou
                                    dadosLH._spx_finalizado = true;
                                    dadosLH._spx_status = statusSPX;
                                    if (chegadaReal) {
                                        dadosLH._spx_chegada_real = chegadaReal;
                                    }
                                }
                                
                                console.log(`   ✅ Dados da LH ${lhId} atualizados no objeto`);
                            }
                            
                            // 2. Desbloquear visualmente
                            linha.classList.remove('bloqueada');
                            linha.style.opacity = '1';
                            linha.style.pointerEvents = 'auto';
                            
                            // 3. Remover atributo data-bloqueada
                            linha.removeAttribute('data-bloqueada');
                            linha.dataset.status = novoStatusFront.codigo;
                            
                            console.log(`   🔓 LH desbloqueada: ${lhId}`);
                            
                            // 4. Limpar tempo de corte COMPLETAMENTE
                            if (celulaTempoCorte) {
                                celulaTempoCorte.innerHTML = '';
                                celulaTempoCorte.textContent = '-';
                                celulaTempoCorte.style.color = '#6b7280';
                                celulaTempoCorte.style.fontWeight = 'normal';
                                celulaTempoCorte.style.textAlign = 'center';
                                console.log(`   ⏰ Tempo de corte limpo: ${lhId} (LH no piso)`);
                            }
                            
                            // 5. Remover qualquer tooltip de bloqueio
                            const celulaLH = linha.querySelector('td.lh-trip-cell');
                            if (celulaLH && celulaLH.title && celulaLH.title.includes('bloqueada')) {
                                celulaLH.title = '';
                            }
                        }
                        
                        statusAtualizados++;
                        console.log(`   ✅ Status atualizado: ${lhId} → ${novoStatusFront.texto} (SPX: ${statusSPX})`);
                    }
                    
                    // Adicionar tooltip com informações completas
                    const tooltipText = `📦 SPX INFO:\n\n` +
                        `Status: ${statusSPX}\n` +
                        `Chegada Real: ${chegadaRealStr}\n` +
                        `Motorista: ${dados.driver_name || 'N/A'}\n` +
                        `Placa: ${dados.vehicle_number || 'N/A'}\n` +
                        `Tipo: ${dados.vehicle_type || 'N/A'}`;
                    
                    celulaStatus.title = tooltipText;
                    celulaStatus.style.cursor = 'help';
                    atualizadas++;
                    
                    // Se status for "Finalizado" E tiver chegada real, atualizar PREVISÃO HORA
                    if (statusSPX === 'Finalizado' && chegadaReal && celulaPrevisaoHora) {
                        const horaChegada = chegadaReal.toLocaleTimeString('pt-BR', { 
                            hour: '2-digit', 
                            minute: '2-digit', 
                            second: '2-digit' 
                        });
                        celulaPrevisaoHora.textContent = horaChegada;
                        celulaPrevisaoHora.style.fontWeight = 'bold';
                        celulaPrevisaoHora.style.color = '#10b981'; // Verde
                        celulaPrevisaoHora.title = `✅ Chegada Real (SPX): ${chegadaRealStr}`;
                        horariosAtualizados++;
                        console.log(`   🕐 Horário atualizado: ${lhId} → ${horaChegada}`);
                    }
                }
            }
        });
    });
    
    console.log(`✅ [SPX] Sincronização concluída:`);
    console.log(`   📊 ${atualizadas} tooltips adicionados`);
    console.log(`   🔄 ${statusAtualizados} status atualizados`);
    console.log(`   🕐 ${horariosAtualizados} horários de chegada atualizados`);
    
    // IMPORTANTE: Re-renderizar sugestão de planejamento se houve atualizações
    if (statusAtualizados > 0) {
        console.log(`   🔄 Re-renderizando sugestão de planejamento...`);
        
        // Re-calcular e atualizar a sugestão
        try {
            // Verificar qual ciclo está selecionado
            const cicloAtivo = document.querySelector('.ciclo-card.ativo');
            const cicloSelecionado = cicloAtivo ? cicloAtivo.dataset.ciclo : 'Todos';
            
            // Re-renderizar a visualização com os novos dados
            renderizarVisualizacao(cicloSelecionado);
            
            console.log(`   ✅ Sugestão re-renderizada com dados atualizados!`);
        } catch (error) {
            console.error(`   ❌ Erro ao re-renderizar:`, error);
        }
    }
    
    // Mostrar notificação visual se houve atualizações
    if (statusAtualizados > 0 || horariosAtualizados > 0) {
        const msg = `✅ Status validado com SPX!\n\n` +
            `🔄 ${statusAtualizados} status atualizado(s)\n` +
            `🕐 ${horariosAtualizados} horário(s) de chegada atualizado(s)`;
        
        // Criar notificação temporária
        const notif = document.createElement('div');
        notif.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            background: #10b981;
            color: white;
            padding: 16px 24px;
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            z-index: 10000;
            font-weight: 500;
            animation: slideIn 0.3s ease-out;
        `;
        notif.textContent = `✅ ${statusAtualizados + horariosAtualizados} atualizações do SPX!`;
        document.body.appendChild(notif);
        
        setTimeout(() => {
            notif.style.animation = 'slideOut 0.3s ease-in';
            setTimeout(() => notif.remove(), 300);
        }, 3000);
    }
}

/**
 * Processa resultados do SPX e atualiza status visual
 */
function processarResultadosSPX(resultados) {
    console.log('📊 [SPX] Processando resultados...');
    
    // Encontrar tabela ativa
    let tbody = null;
    const abaAtiva = document.querySelector('.tab.active');
    
    if (abaAtiva) {
        const dataTab = abaAtiva.getAttribute('data-tab');
        if (dataTab === 'planejamento') {
            tbody = document.getElementById('tbodyPlanejamento');
        } else if (dataTab === 'lh-trips') {
            tbody = document.getElementById('tbodyLHTrips');
        }
    }
    
    // Fallback: tentar encontrar tabela visível
    if (!tbody) {
        const tbodies = [
            document.getElementById('tbodyPlanejamento'),
            document.getElementById('tbodyLHTrips')
        ];
        for (const tb of tbodies) {
            if (tb && tb.offsetParent !== null) {
                tbody = tb;
                break;
            }
        }
    }
    
    if (!tbody) {
        console.error('❌ [SPX] Nenhuma tabela encontrada para atualizar!');
        return;
    }
    
    let divergenciasEncontradas = 0;
    let statusOK = 0;
    
    resultados.forEach(resultado => {
        const lhId = resultado.lh_id;
        
        const linhas = tbody.querySelectorAll('tr');
        linhas.forEach(linha => {
            // Procurar pela célula com classe 'lh-trip-cell'
            const celulaLH = linha.querySelector('td.lh-trip-cell');
            
            if (celulaLH && celulaLH.textContent.trim() === lhId) {
                // Procurar coluna STATUS LH
                const colunas = linha.querySelectorAll('td');
                let celulaStatus = null;
                for (let i = 0; i < colunas.length; i++) {
                    const badge = colunas[i].querySelector('.badge, .status-badge');
                    if (badge || colunas[i].textContent.includes('Sinalizar Inventário') || 
                            colunas[i].textContent.includes('Em trânsito') ||
                            colunas[i].textContent.includes('No Hub')) {
                            celulaStatus = colunas[i];
                            break;
                        }
                    }
                    
                    if (celulaStatus && resultado.encontrada) {
                        const statusAtual = celulaStatus.textContent.trim();
                        
                        // Validar divergência
                        let divergencia = false;
                        const statusLower = statusAtual.toLowerCase();
                        
                        if (resultado.descarregada && (statusLower.includes('em trânsito') || statusLower.includes('em transito'))) {
                            divergencia = true;
                        } else if (resultado.chegou_hub && (statusLower.includes('em trânsito') || statusLower.includes('em transito'))) {
                            divergencia = true;
                        }
                        
                        // Atualizar visual
                        if (divergencia) {
                            divergenciasEncontradas++;
                            celulaStatus.innerHTML = `
                                <div style="display: flex; flex-direction: column; gap: 2px;">
                                    <span style="color: #ff9800; font-weight: 600;">⚠️ ${statusAtual}</span>
                                    <span style="font-size: 11px; color: #4caf50; font-weight: 600;">SPX: ${resultado.status_spx}</span>
                                </div>
                            `;
                            celulaStatus.title = `⚠️ DIVERGÊNCIA DETECTADA!\n\nPlanilha: ${statusAtual}\nSPX: ${resultado.status_spx}\n\nATA (Chegada): ${resultado.ata || 'N/A'}\nUnloaded (Descarregada): ${resultado.unloaded_time || 'N/A'}`;
                            celulaStatus.style.background = '#fff3cd';
                            celulaStatus.style.padding = '8px';
                            celulaStatus.style.borderRadius = '4px';
                            celulaStatus.style.borderLeft = '4px solid #ff9800';
                            console.log(`   ⚠️ ${lhId}: DIVERGÊNCIA - ${resultado.status_spx}`);
                        } else {
                            statusOK++;
                            celulaStatus.title = `✅ Status OK\n\nSPX: ${resultado.status_spx}\nATA: ${resultado.ata || 'N/A'}\nUnloaded: ${resultado.unloaded_time || 'N/A'}`;
                            celulaStatus.style.borderLeft = '3px solid #4caf50';
                            console.log(`   ✅ ${lhId}: OK - ${resultado.status_spx}`);
                        }
                    }
                }
        });
    });
    
    console.log(`\n📊 [SPX] RESUMO:`);
    console.log(`   ✅ Status OK: ${statusOK}`);
    console.log(`   ⚠️ Divergências: ${divergenciasEncontradas}`);
    console.log('✅ [SPX] Resultados processados!');
}

// Inicializar botão e modal após DOM carregar
document.addEventListener('DOMContentLoaded', () => {
    const btnSincronizarLHs = document.getElementById('btnSincronizarLHs');
    if (btnSincronizarLHs) {
        btnSincronizarLHs.addEventListener('click', async () => {
            console.log('🔍 [SPX] Botão Sincronizar LHs SPX clicado');
            await sincronizarLHsSPX();
        });
        console.log('✅ [SPX] Listener do botão Sincronizar LHs SPX registrado');
    } else {
        console.error('❌ [SPX] Botão btnSincronizarLHs não encontrado!');
    }
});

// Fechar modal
function fecharModalSincronizarLHs() {
    document.getElementById('modalSincronizarLHs').style.display = 'none';
}

// Executar sincronização
async function executarSincronizacaoLHs() {
    try {
        // Obter configurações
        const config = getConfigNavegador();
        const diasPendentes = parseInt(document.getElementById('modalDiasPendentes').value);
        const diasFinalizados = parseInt(document.getElementById('modalDiasFinalizados').value);
        
        console.log('🚀 [SINCRONIZAR LHs] Iniciando...');
        console.log(`   - Modo: ${config.headless ? 'Headless' : 'Visível'}`);
        console.log(`   - Dias Pendentes: ${diasPendentes}`);
        console.log(`   - Dias Finalizados: ${diasFinalizados}`);
        
        // Preparar UI do modal
        document.getElementById('btnExecutarSinc').disabled = true;
        document.getElementById('btnCancelarSinc').disabled = true;
        document.getElementById('modalSincProgresso').style.display = 'block';
        document.getElementById('modalSincResultado').style.display = 'none';
        document.getElementById('modalSincProgressBar').style.width = '0%';
        document.getElementById('modalSincStatus').textContent = 'Iniciando sincronização...';
        
        // Chamar IPC
        const result = await ipcRenderer.invoke('exportar-lhs-spx', {
            headless: config.headless,
            diasPendentes,
            diasFinalizados
        });
        
        if (result.success) {
            console.log('✅ [SINCRONIZAR LHs] Concluído!');
            console.log('   Estatísticas:', result.estatisticas);
            
            // Atualizar progresso para 100%
            document.getElementById('modalSincProgressBar').style.width = '100%';
            document.getElementById('modalSincStatus').textContent = 'Consolidando arquivos...';
            
            // Aguardar um pouco para mostrar 100%
            await new Promise(r => setTimeout(r, 500));
            
            // Exibir resultados
            document.getElementById('modalStatPendentes').textContent = result.estatisticas.pendentes.toLocaleString('pt-BR');
            document.getElementById('modalStatExpedidos').textContent = result.estatisticas.expedidos.toLocaleString('pt-BR');
            document.getElementById('modalStatFinalizados').textContent = result.estatisticas.finalizados.toLocaleString('pt-BR');
            document.getElementById('modalStatTotal').textContent = result.estatisticas.total.toLocaleString('pt-BR');
            
            // Extrair apenas o nome do arquivo consolidado
            const nomeArquivo = result.arquivos.consolidado.split('\\').pop();
            const pastaArquivo = result.arquivos.consolidado.replace(nomeArquivo, '');
            
            document.getElementById('modalSincCaminhos').innerHTML = `
                <strong>📂 Arquivos gerados:</strong><br>
                <div style="margin-top: 8px;">
                    <div style="font-size: 12px; color: #666; margin-bottom: 3px;">
                        📁 ${pastaArquivo}
                    </div>
                    <div style="font-size: 13px; font-weight: 600; color: #333;">
                        📄 ${nomeArquivo}
                    </div>
                    <div style="font-size: 11px; color: #999; margin-top: 5px;">
                        ⏱️ ${result.tempo_execucao}
                    </div>
                </div>
            `;
            
            document.getElementById('modalSincProgresso').style.display = 'none';
            document.getElementById('modalSincResultado').style.display = 'block';
            document.getElementById('btnCancelarSinc').disabled = false;
            document.getElementById('btnCancelarSinc').textContent = 'Fechar';
            
        } else {
            throw new Error(result.error || 'Erro desconhecido');
        }
        
    } catch (error) {
        console.error('❌ [SINCRONIZAR LHs] Erro:', error);
        document.getElementById('modalSincProgressBar').style.width = '0%';
        document.getElementById('modalSincStatus').textContent = `❌ Erro: ${error.message}`;
        document.getElementById('modalSincStatus').style.color = '#dc3545';
        alert(`❌ Erro na sincronização: ${error.message}`);
    } finally {
        document.getElementById('btnExecutarSinc').disabled = false;
        document.getElementById('btnCancelarSinc').disabled = false;
    }
}

// Tornar funções globais para onclick
window.fecharModalSincronizarLHs = fecharModalSincronizarLHs;
window.executarSincronizacaoLHs = executarSincronizacaoLHs;

// Listener de progresso (recebe atualizações do main.js)
ipcRenderer.on('exportar-lhs-progresso', (event, data) => {
    const { etapa, pagina, total } = data;
    
    let etapaNome = '';
    let percentual = 0;
    
    switch(etapa) {
        case 1:
            etapaNome = 'Pendentes';
            percentual = 10 + (Math.min(pagina, 20) * 1); // 10-30%
            break;
        case 2:
            etapaNome = 'Expedidos';
            percentual = 35 + (Math.min(pagina, 20) * 1); // 35-55%
            break;
        case 3:
            etapaNome = 'Finalizados';
            percentual = 60 + (Math.min(pagina, 30) * 1); // 60-90%
            break;
    }
    
    // Limitar a 95% para deixar espaço para consolidação
    percentual = Math.min(percentual, 95);
    
    // Atualizar modal se estiver aberto
    const modalProgBar = document.getElementById('modalSincProgressBar');
    const modalStatus = document.getElementById('modalSincStatus');
    if (modalProgBar && modalStatus) {
        modalProgBar.style.width = `${percentual}%`;
        modalStatus.textContent = `[${etapa}/3] Baixando ${etapaNome}... Página ${pagina} (${total.toLocaleString('pt-BR')} itens)`;
        modalStatus.style.color = '#666';
    }
    
    console.log(`📊 [PROGRESSO] Etapa ${etapa}/3 - ${etapaNome} - Pág ${pagina} - Total: ${total}`);
});

console.log('✅ [EXPORTAR LHs] Módulo carregado');

// ==================== MODAL VALIDAR LHs COM SPX ====================
document.addEventListener('DOMContentLoaded', () => {
    const btnValidarLHs = document.getElementById('btnValidarLHs');
    if (btnValidarLHs) {
        btnValidarLHs.addEventListener('click', () => {
            console.log('👍 [MODAL VALID] Botão Validar LHs clicado');
            document.getElementById('modalValidarLHs').style.display = 'flex';
            // Resetar estado do modal
            document.getElementById('modalValidProgresso').style.display = 'none';
            document.getElementById('modalValidResultado').style.display = 'none';
            document.getElementById('btnExecutarValid').disabled = false;
            document.getElementById('btnCancelarValid').disabled = false;
        });
        console.log('✅ [MODAL VALID] Listener do botão Validar LHs registrado');
    } else {
        console.error('❌ [MODAL VALID] Botão btnValidarLHs não encontrado!');
    }
});

// Fechar modal de validação
function fecharModalValidarLHs() {
    document.getElementById('modalValidarLHs').style.display = 'none';
}

// Executar validação
async function executarValidacaoLHs() {
    try {
        console.log('🔍 [VALIDAR LHs] Iniciando...');
        
        // Preparar UI do modal
        document.getElementById('btnExecutarValid').disabled = true;
        document.getElementById('btnCancelarValid').disabled = true;
        document.getElementById('modalValidProgresso').style.display = 'block';
        document.getElementById('modalValidResultado').style.display = 'none';
        document.getElementById('modalValidProgressBar').style.width = '0%';
        document.getElementById('modalValidStatus').textContent = 'Lendo Google Sheets...';
        
        // Simular progresso
        document.getElementById('modalValidProgressBar').style.width = '30%';
        
        // Chamar IPC com filtro de station e sequence_number
        const sequenceFilter = document.getElementById('validSequenceFilter').value;
        const config = {
            stationFiltro: stationAtualNome || null,
            sequenceFilter: sequenceFilter
        };
        console.log('📍 [VALIDAR LHs] Station filtro:', config.stationFiltro);
        console.log('🎯 [VALIDAR LHs] Sequence filtro:', config.sequenceFilter);
        
        const result = await ipcRenderer.invoke('validar-lhs-spx', config);
        
        if (result.success) {
            console.log('✅ [VALIDAR LHs] Concluído!');
            console.log('   Estatísticas:', result.stats);
            
            // Atualizar progresso para 100%
            document.getElementById('modalValidProgressBar').style.width = '100%';
            document.getElementById('modalValidStatus').textContent = 'Validação concluída!';
            
            // Exibir resultado após 500ms
            setTimeout(() => {
                document.getElementById('modalValidProgresso').style.display = 'none';
                document.getElementById('modalValidResultado').style.display = 'block';
                
                // Preencher estatísticas
                document.getElementById('modalStatFaltantes').textContent = result.stats.lhs_com_dados_faltantes.length;
                document.getElementById('modalStatApenasSheets').textContent = result.stats.lhs_apenas_sheets.length;
                document.getElementById('modalStatApenasSPX').textContent = result.stats.lhs_apenas_spx.length;
                document.getElementById('modalStatComparado').textContent = result.stats.total_comparado || 0;
                document.getElementById('modalStatTotal').textContent = result.stats.total_sheets;
                
                // Gerar detalhes
                let detalhesHTML = '';
                
                // LHs com dados faltantes
                if (result.stats.lhs_com_dados_faltantes.length > 0) {
                    detalhesHTML += '<h5 style="margin-bottom: 10px; color: #ff9800;">⚠️ LHs com Dados Faltantes:</h5>';
                    detalhesHTML += '<ul style="margin-bottom: 20px; padding-left: 20px;">';
                    result.stats.lhs_com_dados_faltantes.slice(0, 10).forEach(lh => {
                        detalhesHTML += `<li><strong>${lh.trip_number}</strong> (${lh.destination_station}) - ${lh.campos_vazios.length} campos vazios</li>`;
                    });
                    if (result.stats.lhs_com_dados_faltantes.length > 10) {
                        detalhesHTML += `<li><em>... e mais ${result.stats.lhs_com_dados_faltantes.length - 10} LHs</em></li>`;
                    }
                    detalhesHTML += '</ul>';
                }
                
                // Campos vazios por tipo
                if (Object.keys(result.stats.campos_vazios_por_tipo).length > 0) {
                    detalhesHTML += '<h5 style="margin-bottom: 10px; color: #2196f3;">📊 Campos Vazios por Tipo:</h5>';
                    detalhesHTML += '<ul style="margin-bottom: 20px; padding-left: 20px;">';
                    Object.entries(result.stats.campos_vazios_por_tipo).forEach(([campo, quantidade]) => {
                        detalhesHTML += `<li><strong>${campo}</strong>: ${quantidade} ocorrências</li>`;
                    });
                    detalhesHTML += '</ul>';
                }
                
                // LHs apenas no Sheets
                if (result.stats.lhs_apenas_sheets.length > 0) {
                    detalhesHTML += '<h5 style="margin-bottom: 10px; color: #2196f3;">📋 LHs Apenas no Sheets:</h5>';
                    detalhesHTML += '<ul style="margin-bottom: 20px; padding-left: 20px;">';
                    result.stats.lhs_apenas_sheets.slice(0, 5).forEach(lh => {
                        detalhesHTML += `<li>${lh.trip_number} (${lh.destination_station})</li>`;
                    });
                    if (result.stats.lhs_apenas_sheets.length > 5) {
                        detalhesHTML += `<li><em>... e mais ${result.stats.lhs_apenas_sheets.length - 5} LHs</em></li>`;
                    }
                    detalhesHTML += '</ul>';
                }
                
                // LHs apenas no SPX
                if (result.stats.lhs_apenas_spx.length > 0) {
                    detalhesHTML += '<h5 style="margin-bottom: 10px; color: #9c27b0;">📦 LHs Apenas no SPX:</h5>';
                    detalhesHTML += '<ul style="padding-left: 20px;">';
                    result.stats.lhs_apenas_spx.slice(0, 5).forEach(lh => {
                        detalhesHTML += `<li>${lh.trip_number} (Station: ${lh.destination_station_id})</li>`;
                    });
                    if (result.stats.lhs_apenas_spx.length > 5) {
                        detalhesHTML += `<li><em>... e mais ${result.stats.lhs_apenas_spx.length - 5} LHs</em></li>`;
                    }
                    detalhesHTML += '</ul>';
                }
                
                if (!detalhesHTML) {
                    detalhesHTML = '<p style="text-align: center; color: #28a745;">✅ Nenhuma discrepância encontrada!</p>';
                }
                
                document.getElementById('modalValidDetalhes').innerHTML = detalhesHTML;
                
                // Reabilitar botões
                document.getElementById('btnExecutarValid').disabled = false;
                document.getElementById('btnCancelarValid').disabled = false;
                document.getElementById('btnExecutarValid').textContent = '🔄 Validar Novamente';
            }, 500);
            
        } else {
            console.error('❌ [VALIDAR LHs] Erro:', result.error);
            alert(`Erro na validação: ${result.error}`);
            
            // Resetar UI
            document.getElementById('modalValidProgresso').style.display = 'none';
            document.getElementById('btnExecutarValid').disabled = false;
            document.getElementById('btnCancelarValid').disabled = false;
        }
        
    } catch (error) {
        console.error('❌ [VALIDAR LHs] Erro fatal:', error);
        alert(`Erro fatal na validação: ${error.message}`);
        
        // Resetar UI
        document.getElementById('modalValidProgresso').style.display = 'none';
        document.getElementById('btnExecutarValid').disabled = false;
        document.getElementById('btnCancelarValid').disabled = false;
    }
}


// ======================= ATALHOS CONFIGURAÇÕES =======================
/**
 * Alterna visibilidade da aba Configurações (Ctrl+U)
 */
function toggleAbaConfiguracoes() {
    const tabConfig = document.getElementById('tab-config');
    
    if (tabConfig && tabConfig.classList.contains('active')) {
        // Se já está aberta, fecha
        fecharAbaConfiguracoes();
    } else {
        // Se está fechada, abre
        abrirAbaConfiguracoes();
    }
}

/**
 * Abre a aba Configurações
 */
function abrirAbaConfiguracoes() {
    console.log('⚙️ Abrindo aba Configurações via atalho');
    
    // Desativar todas as tabs
    document.querySelectorAll('.tab').forEach(tab => {
        tab.classList.remove('active');
    });
    
    // Desativar todos os conteúdos
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
    });
    
    // Ativar aba Configurações
    const tabConfig = document.getElementById('tab-config');
    if (tabConfig) {
        tabConfig.classList.add('active');
    }
    
    // Verificar status da sessão
    verificarStatusSessao();
}

/**
 * Fecha a aba Configurações e volta para Planejamento Hub
 */
function fecharAbaConfiguracoes() {
    console.log('⚙️ Fechando aba Configurações');
    
    // Voltar para aba Planejamento Hub
    trocarAba('planejamento');
}


// ======================= EASTER EGG =======================
// ============================================
// TOGGLE MODO HEADLESS (CTRL+U)
// ============================================

/**
 * Alterna entre modo Headless (rápido/invisível) e modo Visual (visível)
 * Salva preferência no localStorage
 */
async function toggleModoHeadless() {
    // Ler estado atual do localStorage (padrão: true - headless ativo)
    const estadoAtual = localStorage.getItem('modoHeadless');
    const headlessAtivo = estadoAtual === null ? true : estadoAtual === 'true';
    
    console.log(`🔍 DEBUG TOGGLE: Estado atual = ${estadoAtual}`);
    console.log(`🔍 DEBUG TOGGLE: Headless ativo? = ${headlessAtivo}`);
    
    // Inverter estado
    const novoEstado = !headlessAtivo;
    
    console.log(`🔍 DEBUG TOGGLE: Novo estado = ${novoEstado}`);
    
    // Salvar no localStorage (para persistência)
    localStorage.setItem('modoHeadless', novoEstado.toString());
    
    console.log(`🔍 DEBUG TOGGLE: Salvo no localStorage`);
    
    // ⚡ NOVO: Atualizar variável global do main.js (muda NA HORA!)
    try {
        const result = await ipcRenderer.invoke('toggle-headless-mode', novoEstado);
        console.log(`✅ Modo alterado no main.js: ${novoEstado ? 'RÁPIDO' : 'VISUAL'}`);
        console.log(`🔍 DEBUG TOGGLE: IPC result =`, result);
    } catch (error) {
        console.error('❌ Erro ao alterar modo:', error);
    }
    
    // Mensagem e emoji baseados no novo estado
    const emoji = novoEstado ? '⚡' : '👁️';
    const modo = novoEstado ? 'RÁPIDO' : 'VISUAL';
    const descricao = novoEstado ? 'invisível' : 'visível';
    const cor = novoEstado ? '#00c853' : '#2196f3';
    
    // Criar notificação visual
    mostrarNotificacaoHeadless(emoji, modo, descricao, cor);
    
    console.log(`${emoji} Modo ${modo} ${novoEstado ? 'ATIVADO' : 'DESATIVADO'} - Efeito IMEDIATO!`);
}

/**
 * Mostra notificação visual na tela
 */
function mostrarNotificacaoHeadless(emoji, modo, descricao, cor) {
    // Remover notificação anterior se existir
    const existente = document.getElementById('notificacao-headless');
    if (existente) {
        existente.remove();
    }
    
    // Criar elemento de notificação
    const notificacao = document.createElement('div');
    notificacao.id = 'notificacao-headless';
    notificacao.innerHTML = `
        <div style="
            position: fixed;
            top: 20px;
            right: 20px;
            background: ${cor};
            color: white;
            padding: 16px 24px;
            border-radius: 12px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.3);
            z-index: 999999;
            font-family: 'Segoe UI', Arial, sans-serif;
            font-size: 16px;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 12px;
            animation: slideInRight 0.3s ease-out;
        ">
            <span style="font-size: 24px;">${emoji}</span>
            <div>
                <div style="font-size: 18px; margin-bottom: 2px;">Modo ${modo}</div>
                <div style="font-size: 13px; opacity: 0.9; font-weight: 400;">
                    Navegador ${descricao}
                </div>
            </div>
        </div>
    `;
    
    // Adicionar CSS de animação se não existir
    if (!document.getElementById('headless-notification-css')) {
        const style = document.createElement('style');
        style.id = 'headless-notification-css';
        style.textContent = `
            @keyframes slideInRight {
                from {
                    transform: translateX(400px);
                    opacity: 0;
                }
                to {
                    transform: translateX(0);
                    opacity: 1;
                }
            }
            @keyframes slideOutRight {
                from {
                    transform: translateX(0);
                    opacity: 1;
                }
                to {
                    transform: translateX(400px);
                    opacity: 0;
                }
            }
        `;
        document.head.appendChild(style);
    }
    
    document.body.appendChild(notificacao);
    
    // Remover após 3 segundos com animação
    setTimeout(() => {
        notificacao.firstElementChild.style.animation = 'slideOutRight 0.3s ease-in';
        setTimeout(() => {
            if (notificacao.parentNode) {
                notificacao.remove();
            }
        }, 300);
    }, 3000);
}

/**
 * Obtém estado atual do modo headless
 * @returns {boolean} true se headless ativo, false se visual
 */
function getModoHeadless() {
    const estado = localStorage.getItem('modoHeadless');
    // Padrão: true (headless ativo)
    return estado === null ? true : estado === 'true';
}

/**
 * Adiciona indicador visual do modo atual na interface
 */
// BADGE FIXO DESABILITADO - Só usa notificação temporária
// function adicionarIndicadorModoHeadless() { ... }

// Adicionar indicador quando DOM carregar - DESABILITADO
// document.addEventListener('DOMContentLoaded', () => {
//     setTimeout(() => {
//         adicionarIndicadorModoHeadless();
//     }, 100);
// });

// ============================================
// EASTER EGG
// ============================================

/**
 * Abre o modal Easter Egg com informações sobre o projeto
 */
function abrirModalEasterEgg() {
    console.log('🎉 Easter Egg descoberto!');
    const modal = document.getElementById('modalEasterEgg');
    if (modal) {
        modal.style.display = 'flex';
    }
}

/**
 * Fecha o modal Easter Egg
 */
function fecharModalEasterEgg() {
    console.log('🎉 Fechando Easter Egg');
    const modal = document.getElementById('modalEasterEgg');
    if (modal) {
        modal.style.display = 'none';
    }
}

// Event listener para fechar modal ao clicar fora
document.addEventListener('DOMContentLoaded', () => {
    const modal = document.getElementById('modalEasterEgg');
    const closeBtn = document.getElementById('closeModalEasterEgg');
    
    if (modal) {
        // Fechar ao clicar fora do modal
        modal.addEventListener('click', (e) => {
            if (e.target.id === 'modalEasterEgg') {
                fecharModalEasterEgg();
            }
        });
    }
    
    if (closeBtn) {
        closeBtn.addEventListener('click', fecharModalEasterEgg);
    }
});

// ============================================
// TOGGLE MODO HEADLESS (CTRL+U)
// ============================================

/**
 * Alterna entre Modo Rápido (headless) e Modo Visual
 * Mostra notificação na tela com feedback visual
 */
// FUNÇÃO ANTIGA REMOVIDA - Agora usa a versão com IPC (linha ~7747)

/**
 * Atualiza o indicador visual no header
 */
// ATUALIZAÇÃO DE BADGE DESABILITADA - Só usa notificação temporária
// function atualizarIndicadorModoHeadless() { ... }

/**
 * Mostra notificação flutuante estilizada
 */
function mostrarNotificacaoModoHeadless(emoji, titulo, descricao, cor) {
    // Remover notificação anterior se existir
    const notifAnterior = document.getElementById('notif-headless');
    if (notifAnterior) {
        notifAnterior.remove();
    }
    
    // Criar elemento de notificação
    const notif = document.createElement('div');
    notif.id = 'notif-headless';
    notif.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: white;
        border-left: 4px solid ${cor};
        border-radius: 8px;
        padding: 16px 20px;
        box-shadow: 0 10px 25px rgba(0,0,0,0.2);
        z-index: 10000;
        min-width: 320px;
        animation: slideInRight 0.3s ease-out;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    `;
    
    notif.innerHTML = `
        <div style="display: flex; align-items: start; gap: 12px;">
            <div style="font-size: 32px; line-height: 1;">${emoji}</div>
            <div style="flex: 1;">
                <div style="font-weight: 600; font-size: 15px; color: #1f2937; margin-bottom: 4px;">
                    ${titulo}
                </div>
                <div style="font-size: 13px; color: #6b7280; line-height: 1.4;">
                    ${descricao}
                </div>
            </div>
            <button onclick="this.parentElement.parentElement.remove()" 
                    style="background: none; border: none; color: #9ca3af; cursor: pointer; 
                           font-size: 20px; padding: 0; width: 24px; height: 24px; line-height: 1;">
                ×
            </button>
        </div>
    `;
    
    // Adicionar animação CSS
    if (!document.getElementById('notif-headless-styles')) {
        const style = document.createElement('style');
        style.id = 'notif-headless-styles';
        style.textContent = `
            @keyframes slideInRight {
                from {
                    transform: translateX(400px);
                    opacity: 0;
                }
                to {
                    transform: translateX(0);
                    opacity: 1;
                }
            }
            @keyframes slideOutRight {
                from {
                    transform: translateX(0);
                    opacity: 1;
                }
                to {
                    transform: translateX(400px);
                    opacity: 0;
                }
            }
        `;
        document.head.appendChild(style);
    }
    
    document.body.appendChild(notif);
    
    // Auto-remover após 4 segundos com animação
    setTimeout(() => {
        notif.style.animation = 'slideOutRight 0.3s ease-in';
        setTimeout(() => {
            if (notif.parentElement) {
                notif.remove();
            }
        }, 300);
    }, 4000);
}

/**
 * Obtém o estado atual do modo headless
 * Usado pelo main.js ao executar o download
 */
function getModoHeadless() {
    const modo = localStorage.getItem('modoHeadless');
    // Default: true (rápido)
    return modo === null ? true : modo === 'true';
}

// Expor função para o main.js
window.getModoHeadless = getModoHeadless;
// ============================================
// SISTEMA DE LICENCIAMENTO - RENDERER
// Adicionar no renderer.js
// ============================================

// ============================================
// INICIALIZAÇÃO DO SISTEMA DE LICENÇA
// ============================================

document.addEventListener('DOMContentLoaded', async () => {
  // Verificar licença ao carregar
  await checkAndShowLicense();
  
  // Setup event listeners
  setupLicenseEventListeners();
  
  // Atalho CTRL + * para painel admin
  document.addEventListener('keydown', (e) => {
    if (e.ctrlKey && e.key === '*') {
      e.preventDefault();
      openAdminPanel();
    }
  });
});

// ============================================
// VERIFICAR E MOSTRAR STATUS DA LICENÇA
// ============================================

async function checkAndShowLicense() {
  try {
    const result = await ipcRenderer.invoke('license-check');
    
    if (!result.valid) {
      // Licença inválida/expirada
      showExpiredModal(result);
      return false;
    }
    
    if (result.warning) {
      // Próximo de expirar (30 dias)
      showWarningModal(result);
    }
    
    return true;
  } catch (error) {
    console.error('❌ Erro ao verificar licença:', error);
    return false;
  }
}

// ============================================
// MODAL: LICENÇA EXPIRADA
// ============================================

function showExpiredModal(result) {
  const modal = document.getElementById('licenseExpiredModal');
  const dateElement = document.getElementById('licenseExpiredDate');
  
  if (result.expiryDate) {
    dateElement.textContent = result.expiryDate;
  }
  
  modal.classList.add('active');
  
  // Bloquear fechar modal (não pode usar app expirado)
  modal.onclick = (e) => e.stopPropagation();
}

// ============================================
// MODAL: AVISO DE EXPIRAÇÃO
// ============================================

function showWarningModal(result) {
  const modal = document.getElementById('licenseWarningModal');
  const dateElement = document.getElementById('licenseWarningDate');
  const daysElement = document.getElementById('licenseWarningDays');
  
  dateElement.textContent = result.expiryDate;
  daysElement.innerHTML = `Faltam <strong>${result.daysRemaining} dias</strong>`;
  
  modal.classList.add('active');
}

// ============================================
// SOLICITAR RENOVAÇÃO
// ============================================

async function requestRenewal() {
  const nameInput = document.getElementById('licenseName');
  const emailInput = document.getElementById('licenseEmail');
  const nameGroup = document.getElementById('licenseNameGroup');
  const emailGroup = document.getElementById('licenseEmailGroup');
  const btn = document.getElementById('btnRequestRenewal');
  
  const nome = nameInput.value.trim();
  const email = emailInput.value.trim();
  
  // Reset erros
  nameGroup.classList.remove('error');
  emailGroup.classList.remove('error');
  
  // Validações
  let hasError = false;
  
  if (nome.length < 3) {
    nameGroup.classList.add('error');
    hasError = true;
  }
  
  if (!email.endsWith('@shopee.com')) {
    emailGroup.classList.add('error');
    hasError = true;
  }
  
  if (hasError) {
    return;
  }
  
  // Desabilitar botão
  btn.disabled = true;
  btn.innerHTML = '<span class="license-loading"></span> Enviando...';
  
  try {
    const result = await ipcRenderer.invoke('license-request-renewal', { nome, email });
    
    if (result.success) {
      // Mostrar modal de sucesso
      showRequestSentModal(result);
    } else {
      alert('❌ Erro: ' + result.error);
    }
  } catch (error) {
    console.error('❌ Erro ao solicitar renovação:', error);
    alert('❌ Erro ao processar solicitação');
  } finally {
    btn.disabled = false;
    btn.textContent = 'Solicitar Renovação';
  }
}

// ============================================
// MODAL: SOLICITAÇÃO ENVIADA
// ============================================

function showRequestSentModal(result) {
  // Fechar modal de expirado
  document.getElementById('licenseExpiredModal').classList.remove('active');
  
  // Abrir modal de enviado
  const modal = document.getElementById('licenseRequestSentModal');
  const codeDisplay = document.getElementById('requestCodeDisplay');
  const messageTextarea = document.getElementById('requestEmailMessage');
  
  codeDisplay.textContent = result.requestCode;
  messageTextarea.value = result.emailMessage;
  
  modal.classList.add('active');
}

// ============================================
// ATIVAR LICENÇA
// ============================================

async function activateLicense() {
  const passwordInput = document.getElementById('licensePassword');
  const passwordGroup = document.getElementById('licensePasswordGroup');
  const btn = document.getElementById('btnActivateLicense');
  
  const password = passwordInput.value.trim().toUpperCase();
  
  // Reset erro
  passwordGroup.classList.remove('error');
  
  if (!password || password.length < 10) {
    passwordGroup.classList.add('error');
    return;
  }
  
  // Desabilitar botão
  btn.disabled = true;
  btn.innerHTML = '<span class="license-loading"></span> Ativando...';
  
  try {
    const result = await ipcRenderer.invoke('license-activate', password);
    
    if (result.success) {
      // Fechar modal expirado
      document.getElementById('licenseExpiredModal').classList.remove('active');
      
      // Mostrar modal de ativado
      showActivatedModal(result);
      
      // Recarregar após 3 segundos
      setTimeout(() => {
        location.reload();
      }, 3000);
    } else {
      passwordGroup.classList.add('error');
      alert('❌ ' + result.error);
    }
  } catch (error) {
    console.error('❌ Erro ao ativar licença:', error);
    alert('❌ Erro ao ativar licença');
  } finally {
    btn.disabled = false;
    btn.textContent = 'Ativar Licença';
  }
}

// ============================================
// MODAL: LICENÇA ATIVADA
// ============================================

function showActivatedModal(result) {
  const modal = document.getElementById('licenseActivatedModal');
  const dateElement = document.getElementById('licenseActivatedDate');
  
  dateElement.textContent = result.expiryDate;
  
  modal.classList.add('active');
}

// ============================================
// PAINEL ADMIN (CTRL + *)
// ============================================

async function openAdminPanel() {
  console.log('🔍 DEBUG: openAdminPanel() chamado!');
  
  const panel = document.getElementById('licenseAdminPanel');
  console.log('🔍 DEBUG: panel =', panel);
  
  if (!panel) {
    console.error('❌ Elemento licenseAdminPanel não encontrado!');
    alert('Erro: Painel admin não encontrado no HTML');
    return;
  }
  
  panel.classList.add('active');
  
  // Carregar status
  await loadAdminStatus();
  
  // Carregar histórico
  await loadAdminHistory();
}

function closeAdminPanel() {
  const panel = document.getElementById('licenseAdminPanel');
  panel.classList.remove('active');
  
  // Limpar formulários
  document.getElementById('adminRequestCode').value = '';
  document.getElementById('adminRequestInfo').style.display = 'none';
  document.getElementById('adminApprovedInfo').style.display = 'none';
}

async function loadAdminStatus() {
  try {
    const result = await ipcRenderer.invoke('license-check');
    
    const statusDiv = document.getElementById('adminLicenseStatus');
    const expiryDateSpan = document.getElementById('adminExpiryDate');
    const daysRemainingSpan = document.getElementById('adminDaysRemaining');
    
    if (result.valid) {
      statusDiv.className = 'license-status active';
      statusDiv.querySelector('.license-status-icon').textContent = '✅';
      statusDiv.querySelector('h3').textContent = 'Licença Ativa';
      expiryDateSpan.textContent = result.expiryDate;
      daysRemainingSpan.textContent = result.daysRemaining + ' dias';
      
      if (result.warning) {
        statusDiv.className = 'license-status warning';
        statusDiv.querySelector('.license-status-icon').textContent = '⚠️';
        statusDiv.querySelector('h3').textContent = 'Próximo de Expirar';
      }
    } else {
      statusDiv.className = 'license-status expired';
      statusDiv.querySelector('.license-status-icon').textContent = '❌';
      statusDiv.querySelector('h3').textContent = 'Licença Expirada';
      expiryDateSpan.textContent = result.expiryDate || 'N/A';
      daysRemainingSpan.textContent = '0 dias';
    }
  } catch (error) {
    console.error('❌ Erro ao carregar status:', error);
  }
}

async function loadAdminHistory() {
  try {
    const history = await ipcRenderer.invoke('license-get-history');
    const listElement = document.getElementById('adminHistoryList');
    
    if (!history || history.length === 0) {
      listElement.innerHTML = '<li style="text-align: center; padding: 20px; color: #999;">Nenhum histórico disponível</li>';
      return;
    }
    
    listElement.innerHTML = history.map(item => {
      const statusBadge = item.status === 'approved' ? 
        '<span class="badge success">Aprovado</span>' :
        item.status === 'pending' ?
        '<span class="badge warning">Pendente</span>' :
        '<span class="badge danger">Negado</span>';
      
      const date = new Date(item.requestedAt).toLocaleString('pt-BR');
      
      return `
        <li class="license-history-item ${item.status}">
          <div class="license-history-item-header">
            <strong>${item.solicitante.name}</strong>
            ${statusBadge}
          </div>
          <div class="license-history-item-details">
            Email: ${item.solicitante.email}<br>
            Data: ${date}<br>
            ${item.approvedBy ? `Aprovado por: ${item.approvedBy}<br>` : ''}
            Código: <code>${item.code}</code>
          </div>
        </li>
      `;
    }).join('');
  } catch (error) {
    console.error('❌ Erro ao carregar histórico:', error);
  }
}

async function searchRequest() {
  const codeInput = document.getElementById('adminRequestCode');
  const code = codeInput.value.trim().toUpperCase();
  const btn = document.getElementById('btnSearchRequest');
  
  if (!code) {
    alert('Digite um código de solicitação');
    return;
  }
  
  btn.disabled = true;
  btn.innerHTML = '<span class="license-loading"></span> Buscando...';
  
  try {
    const result = await ipcRenderer.invoke('license-get-request', code);
    
    if (result.success) {
      showRequestInfo(result.request);
    } else {
      alert('❌ ' + result.error);
    }
  } catch (error) {
    console.error('❌ Erro ao buscar solicitação:', error);
    alert('❌ Erro ao buscar solicitação');
  } finally {
    btn.disabled = false;
    btn.textContent = 'Buscar Solicitação';
  }
}

function showRequestInfo(request) {
  const infoDiv = document.getElementById('adminRequestInfo');
  
  document.getElementById('adminReqName').textContent = request.solicitante.name;
  document.getElementById('adminReqEmail').textContent = request.solicitante.email;
  document.getElementById('adminReqComputer').textContent = request.solicitante.computer;
  document.getElementById('adminReqDate').textContent = new Date(request.requestedAt).toLocaleString('pt-BR');
  
  infoDiv.style.display = 'block';
  document.getElementById('adminApprovedInfo').style.display = 'none';
}

async function approveRequest() {
  const code = document.getElementById('adminRequestCode').value.trim().toUpperCase();
  const btn = document.getElementById('btnApproveRequest');
  
  // Perguntar quem está aprovando
  const approvedBy = prompt('Digite seu email para confirmar aprovação:');
  
  if (!approvedBy || !approvedBy.includes('@')) {
    alert('Email inválido');
    return;
  }
  
  btn.disabled = true;
  btn.innerHTML = '<span class="license-loading"></span> Aprovando...';
  
  try {
    const result = await ipcRenderer.invoke('license-approve-request', { code, approvedBy });
    
    if (result.success) {
      // Esconder info
      document.getElementById('adminRequestInfo').style.display = 'none';
      
      // Mostrar senha gerada
      showApprovedInfo(result);
      
      // Recarregar histórico
      await loadAdminHistory();
    } else {
      alert('❌ ' + result.error);
    }
  } catch (error) {
    console.error('❌ Erro ao aprovar:', error);
    alert('❌ Erro ao aprovar solicitação');
  } finally {
    btn.disabled = false;
    btn.textContent = '✅ Aprovar e Gerar Senha';
  }
}

function showApprovedInfo(result) {
  const approvedDiv = document.getElementById('adminApprovedInfo');
  
  document.getElementById('adminGeneratedPassword').textContent = result.password;
  document.getElementById('adminApprovedEmail').textContent = result.solicitante.email;
  document.getElementById('adminApprovalMessage').value = result.approvalMessage;
  
  approvedDiv.style.display = 'block';
}

async function extendLicense() {
  if (!confirm('Deseja estender a licença por +6 meses?')) {
    return;
  }
  
  try {
    const result = await ipcRenderer.invoke('license-extend');
    
    if (result.success) {
      alert('✅ ' + result.message + '\nNova data: ' + result.newExpiryDate);
      await loadAdminStatus();
    } else {
      alert('❌ ' + result.error);
    }
  } catch (error) {
    console.error('❌ Erro ao estender licença:', error);
    alert('❌ Erro ao estender licença');
  }
}

// ============================================
// UTILITÁRIOS
// ============================================

function copyToClipboard(elementId, btnId) {
  const element = document.getElementById(elementId);
  const btn = document.getElementById(btnId);
  
  element.select();
  document.execCommand('copy');
  
  const originalText = btn.textContent;
  btn.textContent = '✅ Copiado!';
  btn.classList.add('copied');
  
  setTimeout(() => {
    btn.textContent = originalText;
    btn.classList.remove('copied');
  }, 2000);
}

// ============================================
// EVENT LISTENERS
// ============================================

function setupLicenseEventListeners() {
  // Modal Expirado
  document.getElementById('btnRequestRenewal')?.addEventListener('click', requestRenewal);
  document.getElementById('btnActivateLicense')?.addEventListener('click', activateLicense);
  
  // Modal Solicitação Enviada
  document.getElementById('btnCopyRequestMessage')?.addEventListener('click', () => {
    copyToClipboard('requestEmailMessage', 'btnCopyRequestMessage');
  });
  
  document.getElementById('btnCloseRequestSent')?.addEventListener('click', () => {
    document.getElementById('licenseRequestSentModal').classList.remove('active');
    // Voltar para modal expirado
    document.getElementById('licenseExpiredModal').classList.add('active');
  });
  
  // Modal Ativado
  document.getElementById('btnCloseActivated')?.addEventListener('click', () => {
    location.reload();
  });
  
  // Modal Warning
  document.getElementById('btnRenewFromWarning')?.addEventListener('click', () => {
    document.getElementById('licenseWarningModal').classList.remove('active');
    showExpiredModal({});
  });
  
  document.getElementById('btnCloseWarning')?.addEventListener('click', () => {
    document.getElementById('licenseWarningModal').classList.remove('active');
  });
  
  // Painel Admin
  document.getElementById('btnCloseAdminPanel')?.addEventListener('click', closeAdminPanel);
  document.getElementById('btnSearchRequest')?.addEventListener('click', searchRequest);
  document.getElementById('btnApproveRequest')?.addEventListener('click', approveRequest);
  document.getElementById('btnExtendLicense')?.addEventListener('click', extendLicense);
  
  document.getElementById('btnCancelApprove')?.addEventListener('click', () => {
    document.getElementById('adminRequestInfo').style.display = 'none';
    document.getElementById('adminRequestCode').value = '';
  });
  
  document.getElementById('btnCopyApprovalMessage')?.addEventListener('click', () => {
    copyToClipboard('adminApprovalMessage', 'btnCopyApprovalMessage');
  });
  
  // Fechar painel ao clicar fora
  document.getElementById('licenseAdminPanel')?.addEventListener('click', (e) => {
    if (e.target.id === 'licenseAdminPanel') {
      closeAdminPanel();
    }
  });
}