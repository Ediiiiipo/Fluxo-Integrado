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

// Mostrar modal de atualização
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
            fecharModalAtualizacao();
        };
    }
    
    // Mostrar modal
    modal.style.display = 'flex';
    
    console.log('✅ Modal de atualização exibido');
}

// Fechar modal de atualização
function fecharModalAtualizacao() {
    const modal = document.getElementById('modalAtualizacao');
    if (modal) {
        modal.style.display = 'none';
    }
}

// Mostrar badge de nova versão no header
function mostrarBadgeNovaVersao() {
    const header = document.querySelector('.app-header');
    if (!header) return;
    
    // Verificar se badge já existe
    if (document.getElementById('badgeNovaVersao')) return;
    
    // Criar badge
    const badge = document.createElement('div');
    badge.id = 'badgeNovaVersao';
    badge.className = 'badge-nova-versao';
    badge.innerHTML = '🆕 Nova versão disponível';
    badge.onclick = () => {
        // Reabrir modal ao clicar no badge
        verificarEMostrarAtualizacao();
    };
    
    header.appendChild(badge);
    
    console.log('✅ Badge de nova versão adicionado ao header');
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
