// ============================================
// VERSION CHECKER - Sistema de Verificação de Versão
// ============================================

// Função para comparar versões (formato: X.Y.Z)
function compararVersoes(versaoAtual, versaoRemota) {
    // Remover espaços e converter para string
    const atual = String(versaoAtual).trim();
    const remota = String(versaoRemota).trim();
    
    // Split por ponto
    const partesAtual = atual.split('.').map(n => parseInt(n) || 0);
    const partesRemota = remota.split('.').map(n => parseInt(n) || 0);
    
    // Garantir que ambas tenham 3 partes (major.minor.patch)
    while (partesAtual.length < 3) partesAtual.push(0);
    while (partesRemota.length < 3) partesRemota.push(0);
    
    // Comparar cada parte
    for (let i = 0; i < 3; i++) {
        if (partesRemota[i] > partesAtual[i]) {
            return 1; // Versão remota é maior
        } else if (partesRemota[i] < partesAtual[i]) {
            return -1; // Versão atual é maior
        }
    }
    
    return 0; // Versões são iguais
}

// Verificar se há nova versão disponível
async function verificarNovaVersao() {
    try {
        console.log('🔍 Verificando se há nova versão disponível...');
        
        // Buscar dados de versão do Google Sheets via IPC
        const resultado = await ipcRenderer.invoke('verificar-versao-sheets');
        
        if (!resultado.success) {
            console.warn('⚠️ Não foi possível verificar versão:', resultado.error);
            return null;
        }
        
        const { versaoLocal, versaoRemota, mostrarPopup, linkDownload } = resultado;
        
        console.log(`📌 Versão local: ${versaoLocal}`);
        console.log(`📌 Versão remota: ${versaoRemota}`);
        console.log(`📌 Mostrar popup: ${mostrarPopup}`);
        
        // Comparar versões
        const comparacao = compararVersoes(versaoLocal, versaoRemota);
        
        if (comparacao === 1 && mostrarPopup) {
            // Há nova versão disponível
            console.log('🆕 Nova versão disponível!');
            return {
                temAtualizacao: true,
                versaoLocal,
                versaoRemota,
                linkDownload
            };
        } else if (comparacao === 1) {
            console.log('🆕 Nova versão disponível, mas popup desabilitado');
            return {
                temAtualizacao: false,
                versaoLocal,
                versaoRemota,
                linkDownload
            };
        } else if (comparacao === 0) {
            console.log('✅ Você está usando a versão mais recente!');
            return {
                temAtualizacao: false,
                versaoLocal,
                versaoRemota
            };
        } else {
            console.log('⚠️ Versão local é maior que a remota (dev mode?)');
            return {
                temAtualizacao: false,
                versaoLocal,
                versaoRemota
            };
        }
        
    } catch (error) {
        console.error('❌ Erro ao verificar versão:', error);
        return null;
    }
}

// Mostrar modal de atualização (MODO BLOQUEANTE)
function mostrarModalAtualizacao(dadosVersao) {
    const modal = document.getElementById('modalAtualizacao');
    if (!modal) {
        console.error('❌ Modal de atualização não encontrado no HTML');
        return;
    }
    
    // Atualizar textos do modal
    const versaoAtualEl = document.getElementById('versaoAtualTexto');
    const versaoNovaEl = document.getElementById('versaoNovaTexto');
    
    if (versaoAtualEl) versaoAtualEl.textContent = dadosVersao.versaoLocal;
    if (versaoNovaEl) versaoNovaEl.textContent = dadosVersao.versaoRemota;
    
    // Configurar botão de download
    const btnBaixar = document.getElementById('btnBaixarAtualizacao');
    if (btnBaixar && dadosVersao.linkDownload) {
        btnBaixar.onclick = () => {
            // Abrir link de download no navegador padrão
            require('electron').shell.openExternal(dadosVersao.linkDownload);
            // NÃO fecha o modal - usuário deve atualizar e reiniciar
            console.log('🚀 Link de download aberto. Aguardando atualização...');
        };
    }
    
    // Mostrar modal
    modal.style.display = 'flex';
    
    // 🔒 BLOQUEAR FECHAMENTO DO MODAL
    bloquearFechamentoModal(modal);
    
    // 🔒 BLOQUEAR TODA A APLICAÇÃO
    bloquearAplicacao();
    
    console.log('🔒 Modal de atualização OBRIGATÓRIA exibido - Aplicação bloqueada');
}

// Fechar modal de atualização (DESABILITADO EM MODO OBRIGATÓRIO)
function fecharModalAtualizacao() {
    // NÃO FAZ NADA - Modal não pode ser fechado em modo obrigatório
    console.warn('⚠️ Tentativa de fechar modal bloqueada - Atualização obrigatória');
}

// Bloquear fechamento do modal (ESC, clique fora, etc.)
function bloquearFechamentoModal(modal) {
    // Bloquear ESC
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            e.preventDefault();
            e.stopPropagation();
            console.warn('⚠️ ESC bloqueado - Atualização obrigatória');
        }
    }, true);
    
    // Bloquear clique fora do modal
    modal.addEventListener('click', (e) => {
        if (e.target === modal) {
            e.preventDefault();
            e.stopPropagation();
            console.warn('⚠️ Clique fora do modal bloqueado - Atualização obrigatória');
        }
    }, true);
    
    console.log('🔒 Fechamento do modal bloqueado');
}

// Bloquear toda a aplicação (overlay sobre tudo)
function bloquearAplicacao() {
    // Criar overlay bloqueante sobre toda a aplicação
    const overlay = document.createElement('div');
    overlay.id = 'overlayBloqueioAtualizacao';
    overlay.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100vw;
        height: 100vh;
        background: rgba(0, 0, 0, 0.95);
        z-index: 9999;
        pointer-events: all;
        cursor: not-allowed;
    `;
    
    // Bloquear todos os eventos
    overlay.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
    }, true);
    
    overlay.addEventListener('keydown', (e) => {
        e.preventDefault();
        e.stopPropagation();
    }, true);
    
    document.body.appendChild(overlay);
    
    // Garantir que o modal fique acima do overlay
    const modal = document.getElementById('modalAtualizacao');
    if (modal) {
        modal.style.zIndex = '10000';
    }
    
    console.log('🔒 Aplicação bloqueada - Overlay ativado');
}

// Função principal: verificar e mostrar atualização se necessário
async function verificarEMostrarAtualizacao() {
    const resultado = await verificarNovaVersao();
    
    if (resultado && resultado.temAtualizacao) {
        mostrarModalAtualizacao(resultado);
        mostrarBadgeNovaVersao();
    }
}

console.log('✅ Módulo de Version Checker carregado');
