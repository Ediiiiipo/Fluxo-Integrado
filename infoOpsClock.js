// --- Importações ---
const { google } = require('googleapis');
const path = require('path');
const fs = require('fs');

// ==============================================================================
// --- ÁREA DE CONFIGURAÇÃO ---
// ==============================================================================

const ARQUIVO_CHAVE = 'credenciais.json';
const ID_PLANILHA = '1BOddilA48UQPD9QsnxAFmVcIs0OYRMYxwhTJ70rrIQ8';
const INTERVALO = "'Página1'!A4:Y";

// ==============================================================================
// --- FUNÇÃO DE LEITURA E SOBRESCRITA ---
// ==============================================================================

async function buscarDadosNoSheets() {
    // Adiciona data/hora no log para você saber quando rodou
    const agora = new Date().toLocaleTimeString();
    console.log(`[${agora}] 📄 Verificando atualizações no Google Sheets...`);

    try {
        const auth = new google.auth.GoogleAuth({
            keyFile: path.resolve(__dirname, ARQUIVO_CHAVE),
            scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
        });

        const client = await auth.getClient();
        const sheets = google.sheets({
            version: 'v4',
            auth: client
        });

        const resposta = await sheets.spreadsheets.values.get({
            spreadsheetId: ID_PLANILHA,
            range: INTERVALO,
        });

        const linhas = resposta.data.values;

        if (!linhas || linhas.length === 0) {
            console.log('⚠️ A planilha está vazia.');
            return;
        }

        // --- FORMATAÇÃO DOS DADOS ---
        const cabecalhos = linhas[0];
        const apenasDados = linhas.slice(1);

        const dadosFormatados = apenasDados.map((linha) => {
            let objeto = {};
            cabecalhos.forEach((coluna, index) => {
                objeto[coluna] = linha[index] || "";
            });
            return objeto;
        });

        // --- SALVANDO O ARQUIVO ---
        const caminhoDoArquivo = path.join(__dirname, 'dados_infoOpsClock.json');
        
        const dadosParaSalvar = {
            ultimaAtualizacao: new Date().toISOString(),
            totalRegistros: dadosFormatados.length,
            dados: dadosFormatados
        };
        
        fs.writeFileSync(caminhoDoArquivo, JSON.stringify(dadosParaSalvar, null, 2));

        console.log(`✅ [${agora}] Arquivo atualizado com ${dadosFormatados.length} linhas.`);
        console.log(`📅 Última atualização: ${dadosParaSalvar.ultimaAtualizacao}`);
        
        return dadosParaSalvar;

    } catch (erro) {
        console.error("❌ ERRO:", erro.message);
    }
}

// ==============================================================================
// --- LOOP DE EXECUÇÃO ---
// ==============================================================================

// Exportar a função para ser usada em outros arquivos
module.exports = { buscarDadosNoSheets };

// Se executado diretamente (node infoOpsClock.js)
if (require.main === module) {
    // 1. Roda a primeira vez imediatamente
    buscarDadosNoSheets();
    
    console.log("Mantenha este terminal aberto para continuar rodando.");
    console.log("---------------------------------------------------");
}
