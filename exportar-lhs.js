// ============================================
// EXPORTAR-LHS.JS - Módulo de Exportação de Line Hauls
// Extrai dados de LHs do SPX via API para validação cruzada
// ============================================

const { chromium } = require('playwright-core');
const fs = require('fs-extra');
const path = require('path');
const os = require('os');

class ExportadorLHs {
    constructor(onProgress) {
        this.onProgress = onProgress || (() => {}); // Callback para UI
        
        // Configuração de caminhos
        this.SESSION_FILE = path.join(
            process.env.APPDATA || os.homedir(), 
            'shopee-manager', 
            'shopee_session.json'
        );
        
        // Diretório de saída (mesmo local do executável)
        this.OUTPUT_DIR = path.join(__dirname, 'dados_lhs');
        
        // URLs das APIs
        this.API_PENDENTE = "https://spx.shopee.com.br/api/admin/transportation/trip/list_v2";
        this.API_EXPEDIDO = "https://spx.shopee.com.br/api/admin/transportation/trip/list";
        this.API_FINALIZADO = "https://spx.shopee.com.br/api/admin/transportation/trip/history/list";
    }

    // Detectar navegador do sistema
    async detectSystemBrowser() {
        const paths = [
            'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
            'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'
        ];
        for (const p of paths) {
            if (await fs.pathExists(p)) {
                return p.includes('Chrome') ? 'chrome' : 'msedge';
            }
        }
        return null;
    }

    // Gerar parâmetros para API - Pendentes (30 dias padrão)
    getParamsPendente(dias = 30) {
        const agora = Math.floor(Date.now() / 1000);
        const passado = agora - (dias * 24 * 60 * 60);
        const futuro = agora + (24 * 60 * 60);
        return `station_type=2,3,7,12,14,16,18&query_type=1&tab_type=1&display_range=${passado},${futuro}`;
    }

    // Gerar parâmetros para API - Expedidos (30 dias padrão)
    getParamsExpedido(dias = 30) {
        const agora = Math.floor(Date.now() / 1000);
        const passado = agora - (dias * 24 * 60 * 60);
        const futuro = agora + (24 * 60 * 60);
        return `mtime=${passado},${futuro}&query_type=2`;
    }

    // Gerar parâmetros para API - Finalizados (60 dias padrão)
    getParamsFinalizado(dias = 60) {
        const agora = Math.floor(Date.now() / 1000);
        const passado = agora - (dias * 24 * 60 * 60);
        const futuro = agora + (24 * 60 * 60);
        return `mtime=${passado},${futuro}`;
    }

    // Buscar dados de uma API com paginação
    async buscarDados(page, apiUrl, params, etapaNome, etapaNumero) {
        const lista = [];
        let pagina = 1;
        let continuar = true;

        console.log(`\n📦 [${etapaNumero}/3] Baixando ${etapaNome}...`);

        try {
            while (continuar) {
                const url = `${apiUrl}?${params}&pageno=${pagina}&count=50`;
                
                const res = await page.evaluate(async (u) => {
                    try {
                        return await (await fetch(u)).json();
                    } catch (e) {
                        return null;
                    }
                }, url);

                if (res && res.data && res.data.list && res.data.list.length > 0) {
                    lista.push(...res.data.list);
                    process.stdout.write(`\r   🔄 Pág. ${pagina}: +${res.data.list.length} itens (Total: ${lista.length})`);
                    
                    // Enviar progresso para UI
                    this.onProgress(etapaNumero, pagina, lista.length);
                    
                    pagina++;
                    await new Promise(r => setTimeout(r, 150)); // Delay anti-sobrecarga
                } else {
                    continuar = false;
                }
            }
            
            console.log(`\n   ✅ ${etapaNome}: ${lista.length} registros`);
            
        } catch (error) {
            console.error(`❌ Erro ao buscar ${etapaNome}:`, error.message);
        }

        return lista;
    }

    // Método principal de execução
    async run(headless = true, diasPendentes = 30, diasFinalizados = 60) {
        const inicioExecucao = Date.now();
        
        console.log('');
        console.log('═'.repeat(70));
        console.log('🚀 INICIANDO EXPORTAÇÃO DE LINE HAULS (LHs)');
        console.log(`🌐 Modo: ${headless ? 'Headless (invisível)' : 'Visível'}`);
        console.log(`📅 Período Pendentes/Expedidos: ${diasPendentes} dias`);
        console.log(`📅 Período Finalizados: ${diasFinalizados} dias`);
        console.log('═'.repeat(70));

        // Verificar sessão
        if (!await fs.pathExists(this.SESSION_FILE)) {
            const erro = 'Sessão não encontrada. Faça login manual primeiro via aba Download.';
            console.error(`❌ ${erro}`);
            return {
                success: false,
                error: erro
            };
        }

        // Criar diretório de saída
        await fs.ensureDir(this.OUTPUT_DIR);

        // Iniciar navegador
        const browserChannel = await this.detectSystemBrowser();
        const browser = await chromium.launch({ 
            headless: headless, 
            channel: browserChannel 
        });
        
        const context = await browser.newContext({ 
            storageState: this.SESSION_FILE 
        });
        
        const page = await context.newPage();

        // Navegar para página do SPX (garante contexto correto)
        try {
            await page.goto('https://spx.shopee.com.br/#/hubLinehaulTrips/trip', {
                waitUntil: 'domcontentloaded',
                timeout: 60000 // Aumentado de 30s para 60s
            });
        } catch (e) {
            console.log('⚠️ Timeout na navegação inicial (ignorado)');
        }

        // Listas para armazenar dados
        let listaPendente = [];
        let listaExpedido = [];
        let listaFinalizado = [];

        // =================================================================
        // ETAPA 1: PENDENTES
        // =================================================================
        const params1 = this.getParamsPendente(diasPendentes);
        listaPendente = await this.buscarDados(
            page, 
            this.API_PENDENTE, 
            params1, 
            'PENDENTES', 
            1
        );

        // =================================================================
        // ETAPA 2: EXPEDIDOS
        // =================================================================
        const params2 = this.getParamsExpedido(diasPendentes);
        listaExpedido = await this.buscarDados(
            page, 
            this.API_EXPEDIDO, 
            params2, 
            'EXPEDIDOS', 
            2
        );

        // =================================================================
        // ETAPA 3: FINALIZADOS
        // =================================================================
        const params3 = this.getParamsFinalizado(diasFinalizados);
        listaFinalizado = await this.buscarDados(
            page, 
            this.API_FINALIZADO, 
            params3, 
            'FINALIZADOS', 
            3
        );

        // Fechar navegador
        await browser.close();

        // =================================================================
        // CONSOLIDAÇÃO E SALVAMENTO
        // =================================================================
        console.log('\n' + '═'.repeat(70));
        console.log('🔗 CONSOLIDANDO E SALVANDO ARQUIVOS...');

        const arquivos = {
            pendente: path.join(this.OUTPUT_DIR, 'LH_pendente.json'),
            expedido: path.join(this.OUTPUT_DIR, 'LH_expedido.json'),
            finalizado: path.join(this.OUTPUT_DIR, 'LH_finalizado.json'),
            consolidado: path.join(this.OUTPUT_DIR, 'LH_consolidado_geral.json')
        };

        try {

            // Função para formatar timestamp Unix para DD/MM/YYYY HH:MM
            const formatarTimestamp = (timestamp) => {
                if (!timestamp) return null;
                const data = new Date(timestamp * 1000);
                const dia = String(data.getDate()).padStart(2, '0');
                const mes = String(data.getMonth() + 1).padStart(2, '0');
                const ano = data.getFullYear();
                const hora = String(data.getHours()).padStart(2, '0');
                const minuto = String(data.getMinutes()).padStart(2, '0');
                return `${dia}/${mes}/${ano} ${hora}:${minuto}`;
            };

            // Formatar unloaded_time em todas as listas
            const formatarLista = (lista) => {
                return lista.map(item => {
                    if (item.unloaded_time) {
                        return {
                            ...item,
                            unloaded_time: formatarTimestamp(item.unloaded_time)
                        };
                    }
                    return item;
                });
            };

            const listaPendenteFormatada = formatarLista(listaPendente);
            const listaExpedidoFormatada = formatarLista(listaExpedido);
            const listaFinalizadoFormatada = formatarLista(listaFinalizado);

            // Salvar arquivos individuais com timestamps formatados
            await fs.writeJson(arquivos.pendente, listaPendenteFormatada, { spaces: 2 });
            await fs.writeJson(arquivos.expedido, listaExpedidoFormatada, { spaces: 2 });
            await fs.writeJson(arquivos.finalizado, listaFinalizadoFormatada, { spaces: 2 });

            // Consolidar tudo
            const listaConsolidada = [
                ...listaPendenteFormatada,
                ...listaExpedidoFormatada,
                ...listaFinalizadoFormatada
            ];

            await fs.writeJson(arquivos.consolidado, listaConsolidada, { spaces: 2 });

            const fimExecucao = Date.now();
            const tempoDecorrido = Math.floor((fimExecucao - inicioExecucao) / 1000);
            const minutos = Math.floor(tempoDecorrido / 60);
            const segundos = tempoDecorrido % 60;
            const tempoFormatado = `${minutos}m ${segundos}s`;

            console.log('✅ EXPORTAÇÃO CONCLUÍDA COM SUCESSO!');
            console.log(`📂 Arquivos salvos em: ${this.OUTPUT_DIR}`);
            console.log(`📊 Estatísticas:`);
            console.log(`   - Pendentes:   ${listaPendente.length}`);
            console.log(`   - Expedidos:   ${listaExpedido.length}`);
            console.log(`   - Finalizados: ${listaFinalizado.length}`);
            console.log(`   --------------------------`);
            console.log(`   - TOTAL GERAL: ${listaConsolidada.length}`);
            console.log(`⏱️  Tempo de execução: ${tempoFormatado}`);
            console.log('═'.repeat(70));

            return {
                success: true,
                arquivos: arquivos,
                estatisticas: {
                    pendentes: listaPendente.length,
                    expedidos: listaExpedido.length,
                    finalizados: listaFinalizado.length,
                    total: listaConsolidada.length
                },
                tempo_execucao: tempoFormatado
            };

        } catch (error) {
            console.error('❌ Erro ao salvar arquivos:', error.message);
            return {
                success: false,
                error: `Erro ao salvar arquivos: ${error.message}`
            };
        }
    }
}

module.exports = ExportadorLHs;
