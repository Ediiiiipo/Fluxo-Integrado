// ============================================
// VALIDAR-LHS.JS - Módulo de Validação Cruzada
// Compara dados do Google Sheets com dados do SPX
// ============================================

const fs = require('fs-extra');
const path = require('path');
const { google } = require('googleapis');

class ValidadorLHs {
    constructor() {
        // Configurações do Google Sheets
        this.CREDENCIAIS_FILE = path.join(__dirname, 'credenciais.json');
        this.ID_PLANILHA = '18P9ZPVmNpFOKhWIq9XzZRuSnhtgYR4aJagC8mc3bNa0';
        this.INTERVALO_PLANILHA = 'Página1!A:U';
        
        // Arquivo JSON do SPX
        this.JSON_SPX = path.join(__dirname, 'dados_lhs', 'LH_consolidado_geral.json');
        
        // Mapeamento de campos: Sheets → SPX
        this.MAPEAMENTO_CAMPOS = {
            'destination_id': 'destination_station_id',
            'origin_id': 'origin_station_id',
            'eta_origin_realized': 'opt_origin_realized', // Corrigir nome se necessário
            'opt_origin_realized': 'opt_origin_realized',
            'eta_destination_realized': 'eta_destination_realized',
            'eta_create_date': 'create_time',
            'proposed_departure_datetime': 'proposed_departure_datetime'
        };
    }

    // Autenticar no Google Sheets
    async autenticarGoogleSheets() {
        // Verificar se credenciais existem
        if (!await fs.pathExists(this.CREDENCIAIS_FILE)) {
            throw new Error('Arquivo credenciais.json não encontrado!');
        }
        
        // Usar GoogleAuth (mesmo método do main.js)
        const auth = new google.auth.GoogleAuth({
            keyFile: this.CREDENCIAIS_FILE,
            scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
        });
        
        const client = await auth.getClient();
        return google.sheets({ version: 'v4', auth: client });
    }

    // Ler dados do Google Sheets
    async lerGoogleSheets() {
        console.log('📊 [1/3] Lendo Google Sheets...');
        
        const sheets = await this.autenticarGoogleSheets();
        const resposta = await sheets.spreadsheets.values.get({
            spreadsheetId: this.ID_PLANILHA,
            range: this.INTERVALO_PLANILHA,
        });

        const linhas = resposta.data.values;
        if (!linhas || linhas.length === 0) {
            throw new Error('Nenhum dado encontrado no Google Sheets');
        }

        // Primeira linha = cabeçalhos
        const cabecalhos = linhas[0];
        const dados = [];

        // Converter linhas em objetos
        for (let i = 1; i < linhas.length; i++) {
            const linha = linhas[i];
            const obj = {};
            
            cabecalhos.forEach((cabecalho, index) => {
                obj[cabecalho] = linha[index] || ''; // Vazio se não existir
            });
            
            dados.push(obj);
        }

        console.log(`   ✅ ${dados.length} LHs encontradas no Sheets`);
        return { cabecalhos, dados };
    }

    // Ler dados do JSON do SPX
    async lerJSONSPX() {
        console.log('📦 [2/3] Lendo dados do SPX...');
        
        if (!await fs.pathExists(this.JSON_SPX)) {
            throw new Error('Arquivo JSON do SPX não encontrado. Execute "Sincronizar LHs SPX" primeiro.');
        }

        const dados = await fs.readJson(this.JSON_SPX);
        console.log(`   ✅ ${dados.length} LHs encontradas no SPX`);
        
        // Criar índice por trip_number para busca rápida
        const indice = {};
        dados.forEach(lh => {
            // Usar trip_number como chave única
            const tripNumber = lh.trip_number || lh.trip_id;
            indice[tripNumber] = lh;
        });
        
        return { dados, indice };
    }

    // Converter timestamp Unix para formato legível
    timestampParaData(timestamp) {
        if (!timestamp) return '';
        const data = new Date(timestamp * 1000);
        return data.toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
    }

    // Validar e comparar dados
    async validar(stationFiltro = null, sequenceFilter = 'todos') {
        const inicioExecucao = Date.now();
        
        console.log('');
        console.log('═'.repeat(70));
        console.log('🔍 INICIANDO VALIDAÇÃO CRUZADA');
        if (stationFiltro) {
            console.log(`📍 Filtro: ${stationFiltro}`);
        }
        console.log('═'.repeat(70));

        try {
            // Ler dados
            let sheets = await this.lerGoogleSheets();
            let spx = await this.lerJSONSPX();

            // Aplicar filtro por station se fornecido
            if (stationFiltro) {
                const stationNormalizada = stationFiltro.toLowerCase().replace(/lm\s*hub[_\s]*/gi, '').replace(/[_\s]+/g, '');
                console.log(`📍 Filtrando por station: "${stationNormalizada}"`);
                console.log(`   📝 Station original: "${stationFiltro}"`);
                console.log(`   🎯 Filtro sequence_number: ${sequenceFilter}`);
                
                // Filtrar Sheets
                const totalSheetsAntes = sheets.dados.length;
                console.log(`   🔍 [DEBUG] Primeiras 3 LHs do Sheets (antes do filtro):`);
                sheets.dados.slice(0, 3).forEach((lh, i) => {
                    console.log(`      ${i+1}. trip_number: ${lh.trip_number || 'N/A'}`);
                    console.log(`         destination: "${lh.destination || 'N/A'}"`);
                    console.log(`         origin: "${lh.origin || 'N/A'}"`);
                });
                
                sheets.dados = sheets.dados.filter(lh => {
                    const destination = (lh.destination || '').toLowerCase().replace(/[_\s]+/g, '');
                    const origin = (lh.origin || '').toLowerCase().replace(/[_\s]+/g, '');
                    const match = destination.includes(stationNormalizada) || origin.includes(stationNormalizada);
                    return match;
                });
                console.log(`   ✅ Sheets: ${totalSheetsAntes} → ${sheets.dados.length} LHs`);
                
                if (sheets.dados.length > 0) {
                    console.log(`   🔍 [DEBUG] Primeira LH filtrada do Sheets:`);
                    console.log(`      trip_number: ${sheets.dados[0].trip_number || 'N/A'}`);
                    console.log(`      destination: ${sheets.dados[0].destination || 'N/A'}`);
                    console.log(`      origin: ${sheets.dados[0].origin || 'N/A'}`);
                }
                
                // Filtrar SPX
                const totalSPXAntes = spx.dados.length;
                console.log(`   🔍 [DEBUG] Primeiras 3 LHs do SPX (antes do filtro):`);
                spx.dados.slice(0, 3).forEach((lh, i) => {
                    console.log(`      ${i+1}. trip_number: ${lh.trip_number || 'N/A'}`);
                    if (lh.trip_station && lh.trip_station.length > 0) {
                        lh.trip_station.forEach(station => {
                            console.log(`         seq ${station.sequence_number}: "${station.station_name || 'N/A'}"`);
                        });
                    } else {
                        console.log(`         station_name: "N/A" (sem trip_station)`);
                    }
                });
                
                spx.dados = spx.dados.filter(lh => {
                    // station_name está dentro de trip_station array
                    if (!lh.trip_station || !Array.isArray(lh.trip_station)) {
                        return false;
                    }
                    
                    // Verificar cada station na trip
                    return lh.trip_station.some(station => {
                        const stationName = (station.station_name || '').toLowerCase().replace(/[_\s]+/g, '');
                        const stationMatch = stationName.includes(stationNormalizada);
                        
                        if (!stationMatch) return false;
                        
                        // Aplicar filtro de sequence_number se especificado
                        if (sequenceFilter === 'origem') {
                            return station.sequence_number === 1;
                        } else if (sequenceFilter === 'destino') {
                            return station.sequence_number === 2;
                        } else {
                            return true; // 'todos' - origem OU destino
                        }
                    });
                });
                
                // Recriar índice do SPX com dados filtrados
                spx.indice = {};
                spx.dados.forEach(lh => {
                    const tripNumber = lh.trip_number || lh.trip_id;
                    spx.indice[tripNumber] = lh;
                });
                console.log(`   ✅ SPX: ${totalSPXAntes} → ${spx.dados.length} LHs`);
                
                if (spx.dados.length > 0) {
                    console.log(`   🔍 [DEBUG] Primeira LH filtrada do SPX:`);
                    console.log(`      trip_number: ${spx.dados[0].trip_number || 'N/A'}`);
                    if (spx.dados[0].trip_station && spx.dados[0].trip_station.length > 0) {
                        spx.dados[0].trip_station.forEach(station => {
                            console.log(`      seq ${station.sequence_number}: "${station.station_name || 'N/A'}"`);
                        });
                    }
                }
            }

            console.log('🔄 [3/3] Comparando dados...');
            console.log(`   📊 Total Sheets (filtrado): ${sheets.dados.length}`);
            console.log(`   📊 Total SPX (filtrado): ${spx.dados.length}`);
            console.log(`   📊 Total índice SPX: ${Object.keys(spx.indice).length}`);

            // Estatísticas
            const stats = {
                total_sheets: sheets.dados.length,
                total_spx: spx.dados.length,
                lhs_com_dados_faltantes: [],
                lhs_apenas_sheets: [],
                lhs_apenas_spx: [],
                campos_vazios_por_tipo: {}
            };

            // Comparar cada LH do Sheets
            let lhsComparadas = 0;
            console.log(`   🔍 [DEBUG] Iniciando comparação...`);
            
            sheets.dados.forEach((lhSheets, index) => {
                // Usar trip_number como chave única
                const tripNumber = lhSheets.trip_number || lhSheets.destination_id;
                const lhSPX = spx.indice[tripNumber];
                
                // Debug primeiras 3 comparações
                if (index < 3) {
                    console.log(`   🔍 [DEBUG] Comparação ${index+1}:`);
                    console.log(`      Sheets trip_number: "${tripNumber}"`);
                    console.log(`      Encontrou no SPX: ${lhSPX ? 'SIM' : 'NÃO'}`);
                    if (lhSPX) {
                        console.log(`      SPX trip_number: "${lhSPX.trip_number || lhSPX.trip_id}"`);
                    }
                }

                if (!lhSPX) {
                    // LH existe no Sheets mas não no SPX
                    stats.lhs_apenas_sheets.push({
                        trip_number: tripNumber,
                        destination: lhSheets.destination || 'N/A'
                    });
                    return;
                }
                
                // LH encontrada - incrementar contador
                lhsComparadas++;

                // Verificar campos vazios no Sheets
                const camposVazios = [];
                Object.keys(this.MAPEAMENTO_CAMPOS).forEach(campoSheets => {
                    const valorSheets = lhSheets[campoSheets];
                    
                    if (!valorSheets || valorSheets.trim() === '') {
                        const campoSPX = this.MAPEAMENTO_CAMPOS[campoSheets];
                        const valorSPX = lhSPX[campoSPX];
                        
                        if (valorSPX) {
                            // Converter timestamp se necessário
                            let valorFormatado = valorSPX;
                            if (campoSheets.includes('date') || campoSheets.includes('time')) {
                                valorFormatado = this.timestampParaData(valorSPX);
                            }
                            
                            camposVazios.push({
                                campo_sheets: campoSheets,
                                campo_spx: campoSPX,
                                valor_spx: valorFormatado,
                                valor_original: valorSPX
                            });

                            // Contar por tipo de campo
                            stats.campos_vazios_por_tipo[campoSheets] = 
                                (stats.campos_vazios_por_tipo[campoSheets] || 0) + 1;
                        }
                    }
                });

                if (camposVazios.length > 0) {
                    stats.lhs_com_dados_faltantes.push({
                        trip_number: tripNumber,
                        destination: lhSheets.destination || 'N/A',
                        campos_vazios: camposVazios
                    });
                }
            });
            
            console.log(`   ✅ [DEBUG] Total de LHs comparadas (encontradas em ambos): ${lhsComparadas}`);
            console.log(`   📊 [DEBUG] LHs apenas no Sheets (não encontradas no SPX): ${stats.lhs_apenas_sheets.length}`);
            console.log(`   📊 [DEBUG] LHs com dados faltantes: ${stats.lhs_com_dados_faltantes.length}`);

            // Identificar LHs que existem apenas no SPX
            spx.dados.forEach(lhSPX => {
                const tripNumberSPX = lhSPX.trip_number || lhSPX.trip_id;
                const existeNoSheets = sheets.dados.some(
                    lhSheets => (lhSheets.trip_number || lhSheets.destination_id) === tripNumberSPX
                );
                
                if (!existeNoSheets) {
                    stats.lhs_apenas_spx.push({
                        trip_number: tripNumberSPX,
                        station_name: lhSPX.station_name || 'N/A',
                        status: lhSPX.status
                    });
                }
            });

            // Tempo de execução
            const tempoExecucao = ((Date.now() - inicioExecucao) / 1000).toFixed(2);

            console.log('');
            console.log('═'.repeat(70));
            console.log('✅ VALIDAÇÃO CONCLUÍDA');
            console.log(`⏱️  Tempo: ${tempoExecucao}s`);
            console.log('═'.repeat(70));
            console.log('');
            console.log('📊 ESTATÍSTICAS:');
            console.log(`   🔹 LHs no Sheets: ${stats.total_sheets}`);
            console.log(`   🔹 LHs no SPX: ${stats.total_spx}`);
            console.log(`   ✅ LHs comparadas (encontradas em ambos): ${lhsComparadas}`);
            console.log(`   ⚠️  LHs com dados faltantes: ${stats.lhs_com_dados_faltantes.length}`);
            console.log(`   📋 LHs apenas no Sheets: ${stats.lhs_apenas_sheets.length}`);
            console.log(`   📦 LHs apenas no SPX: ${stats.lhs_apenas_spx.length}`);
            console.log('');
            
            // Adicionar total_comparado nas stats
            stats.total_comparado = lhsComparadas;

            return {
                success: true,
                stats,
                tempo_execucao: `${tempoExecucao}s`
            };

        } catch (error) {
            console.error('❌ Erro na validação:', error.message);
            return {
                success: false,
                error: error.message
            };
        }
    }
}

module.exports = ValidadorLHs;
