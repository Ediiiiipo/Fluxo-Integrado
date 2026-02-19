// infoOutboundCapacityDiário

// --- Importações ---
const { google } = require('googleapis');
const path = require('path');
const fs = require('fs');

// ==============================================================================
// --- ÁREA DE CONFIGURAÇÃO ---
// ==============================================================================

const ARQUIVO_CHAVE = 'credenciais.json';
const ID_PLANILHA = '1iJ70tTT_hlUqcWQacHuhP-3CYI8rYNkOdKnBAHXI_eg';
const INTERVALO = "'Resume Out. Capacity'!B5:CE";

// AQUI ESTÁ A MUDANÇA: 24 horas em milissegundos
const TEMPO_ATUALIZACAO = 24 * 60 * 60 * 1000;

// ==============================================================================
// --- FUNÇÃO DE LEITURA E SOBRESCRITA ---
// ==============================================================================

async function buscarDadosNoSheets() {
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
        fs.writeFileSync(caminhoDoArquivo, JSON.stringify(dadosFormatados, null, 2));

        console.log(`✅ [${agora}] Arquivo atualizado com ${dadosFormatados.length} linhas.`);

    } catch (erro) {
        console.error("❌ ERRO:", erro.message);
    }
}

// ==============================================================================
// --- LOOP DE EXECUÇÃO ---
// ==============================================================================

// 1. Roda a primeira vez imediatamente
buscarDadosNoSheets();

console.log("Mantenha este terminal aberto para continuar rodando.");
console.log("---------------------------------------------------");
