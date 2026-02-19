// ============================================
// USER TRACKING - Sistema de Login e Logs
// ============================================

// ===== FUN\u00c7\u00d5ES DE LOGIN =====

// Verificar se usu\u00e1rio j\u00e1 fez login
function verificarLoginUsuario() {
    const emailSalvo = localStorage.getItem('emailUsuario');
    
    if (emailSalvo) {
        // Usu\u00e1rio j\u00e1 fez login
        emailUsuario = emailSalvo;
        console.log(`\u2705 Usu\u00e1rio logado: ${emailUsuario}`);
    } else {
        // Primeira vez - mostrar modal de login
        console.log('\u26a0\ufe0f Usu\u00e1rio n\u00e3o logado - abrindo modal');
        mostrarModalLogin();
    }
}

// Mostrar modal de login
function mostrarModalLogin() {
    const modal = document.getElementById('modalLogin');
    if (modal) {
        modal.style.display = 'flex';
        
        // Focar no input de e-mail
        setTimeout(() => {
            const input = document.getElementById('inputEmailUsuario');
            if (input) input.focus();
        }, 300);
        
        // Permitir Enter para confirmar
        const input = document.getElementById('inputEmailUsuario');
        if (input) {
            input.addEventListener('keypress', (e) => {
                if (e.key === 'Enter') {
                    confirmarEmailUsuario();
                }
            });
        }
    }
}

// Confirmar e-mail do usu\u00e1rio
function confirmarEmailUsuario() {
    const input = document.getElementById('inputEmailUsuario');
    const email = input?.value.trim();
    
    if (!email) {
        alert('\u26a0\ufe0f Por favor, informe seu e-mail.');
        return;
    }
    
    // Validar formato de e-mail
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
        alert('⚠️ Por favor, informe um e-mail válido.');
        return;
    }
    
    // Validar domínio permitido
    const dominiosPermitidos = ['@shopee.com', '@shopeemobile-external.com'];
    const emailLower = email.toLowerCase();
    const dominioValido = dominiosPermitidos.some(dominio => emailLower.endsWith(dominio));
    
    if (!dominioValido) {
        alert('⚠️ Acesso restrito!\n\nApenas e-mails corporativos são permitidos:\n• @shopee.com\n• @shopeemobile-external.com');
        return;
    }
    
    // Salvar e-mail
    emailUsuario = email;
    localStorage.setItem('emailUsuario', email);
    console.log(`\u2705 E-mail salvo: ${email}`);
    
    // Fechar modal
    const modal = document.getElementById('modalLogin');
    if (modal) modal.style.display = 'none';
    
    // Mostrar mensagem de boas-vindas
    alert(`\ud83c\udf89 Bem-vindo, ${email}!\n\nVoc\u00ea j\u00e1 pode come\u00e7ar a usar a ferramenta.`);
}

// ===== FUN\u00c7\u00d5ES DE TRACKING =====

// Enviar log de planejamento para Google Sheets
async function enviarLogPlanejamento(dadosRelatorio) {
    if (!emailUsuario) {
        console.warn('\u26a0\ufe0f E-mail do usu\u00e1rio n\u00e3o definido - log n\u00e3o ser\u00e1 enviado');
        return;
    }
    
    try {
        console.log('\ud83d\udce4 Enviando log de planejamento para Google Sheets...');
        
        // Preparar dados para envio
        const logData = {
            email: emailUsuario,
            dataHora: new Date().toISOString(),
            estacao: dadosRelatorio.estacao,
            ciclo: dadosRelatorio.ciclo,
            dataExpedicao: dadosRelatorio.dataExpedicao,
            capAutomatico: dadosRelatorio.capAutomatico,
            pedidosTotais: dadosRelatorio.pedidosTotais,
            pedidosPlanejados: dadosRelatorio.pedidosPlanejados,
            quantidadeLHs: dadosRelatorio.quantidadeLHs,
            backlog: dadosRelatorio.backlog
        };
        
        console.log('\ud83d\udcca Dados do log:', logData);
        
        // Enviar para o main process
        const resultado = await ipcRenderer.invoke('salvar-log-planejamento', logData);
        
        if (resultado.success) {
            console.log('\u2705 Log salvo com sucesso no Google Sheets!');
        } else {
            console.error('\u274c Erro ao salvar log:', resultado.error);
        }
    } catch (error) {
        console.error('\u274c Erro ao enviar log:', error);
    }
}

// Extrair dados do relat\u00f3rio para enviar ao Sheets
function extrairDadosRelatorio() {
    // Capturar dados da interface
    const estacao = stationAtualNome || stationAtual || 'N/A';
    const ciclo = cicloSelecionado || 'N/A';
    
    // Pegar data do ciclo selecionada
    const dataCiclo = getDataCicloSelecionada();
    const dataExpedicao = dataCiclo.toLocaleDateString('pt-BR');
    
    // CAP Autom\u00e1tico
    const capAutomatico = obterCapacidadeCiclo(ciclo);
    
    // Pedidos totais (planejados + backlog)
    const pedidosPlanejados = calcularTotalPedidosPlanejados();
    const backlog = pedidosBacklogSelecionados.size;
    const pedidosTotais = pedidosPlanejados + backlog;
    
    // Quantidade de LHs
    const quantidadeLHs = lhsSelecionadasPlan.size;
    
    return {
        estacao,
        ciclo,
        dataExpedicao,
        capAutomatico,
        pedidosTotais,
        pedidosPlanejados,
        quantidadeLHs,
        backlog
    };
}

// Calcular total de pedidos planejados (sem backlog)
function calcularTotalPedidosPlanejados() {
    let total = 0;
    
    lhsSelecionadasPlan.forEach(lh => {
        // Verificar se tem TOs parciais
        if (tosSelecionadasPorLH && tosSelecionadasPorLH[lh] && tosSelecionadasPorLH[lh].size > 0) {
            // LH parcial - contar apenas TOs selecionadas
            total += tosSelecionadasPorLH[lh].size;
        } else {
            // LH completa - contar todos os pedidos
            const pedidos = lhTripsPlanej\u00e1veis[lh]?.length || 0;
            total += pedidos;
        }
    });
    
    return total;
}

console.log('\u2705 M\u00f3dulo de User Tracking carregado');
