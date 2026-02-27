// ============================================
// TRADUTOR DE ERROS - Mensagens amigáveis para o usuário
// ============================================

/**
 * Traduz erros técnicos em mensagens amigáveis para o usuário
 * @param {string} errorMessage - Mensagem de erro técnica
 * @returns {string} - Mensagem amigável formatada
 */
function traduzirErroParaUsuario(errorMessage) {
    if (!errorMessage) {
        return '❌ Erro desconhecido\n\nTente novamente ou entre em contato com o suporte.';
    }

    const erro = errorMessage.toLowerCase();

    // ═══════════════════════════════════════════════════════════
    // 1. SEM INTERNET
    // ═══════════════════════════════════════════════════════════
    if (erro.includes('err_internet_disconnected') || erro.includes('err_network_changed')) {
        return `❌ Sem conexão com a internet

Sua conexão foi perdida durante o download.

O que fazer:
• Verifique sua conexão Wi-Fi/Ethernet
• Reconecte à VPN (se usar)
• Tente novamente após estabilizar a conexão`;
    }

    // ═══════════════════════════════════════════════════════════
    // 2. PROBLEMA DE DNS
    // ═══════════════════════════════════════════════════════════
    if (erro.includes('err_name_not_resolved')) {
        return `❌ Não foi possível encontrar o servidor

Problema ao conectar com o SPX Shopee.

O que fazer:
• Verifique se consegue acessar spx.shopee.com.br no navegador
• Desative temporariamente antivírus/firewall
• Tente usar outra rede (4G/5G)`;
    }

    // ═══════════════════════════════════════════════════════════
    // 3. CONEXÃO RECUSADA
    // ═══════════════════════════════════════════════════════════
    if (erro.includes('err_connection_refused')) {
        return `❌ Servidor recusou a conexão

O servidor SPX não está aceitando conexões.

O que fazer:
• Aguarde alguns minutos e tente novamente
• Verifique se o SPX está acessível no navegador
• Pode estar em manutenção`;
    }

    // ═══════════════════════════════════════════════════════════
    // 4. TIMEOUT
    // ═══════════════════════════════════════════════════════════
    if (erro.includes('timeout')) {
        return `❌ Conexão muito lenta

O servidor não respondeu a tempo (mais de 60 segundos).

O que fazer:
• Verifique a velocidade da sua internet
• Feche outros programas que usam internet
• Tente em outro horário (menos congestionado)`;
    }

    // ═══════════════════════════════════════════════════════════
    // 5. FALHA AO TROCAR STATION
    // ═══════════════════════════════════════════════════════════
    if (erro.includes('falha ao trocar para station')) {
        const stationNome = errorMessage.match(/station: (.+)/)?.[1] || 'desconhecida';
        return `❌ Não foi possível trocar para a station

Station: ${stationNome}

O que fazer:
• Verifique se a station existe no SPX
• Tente trocar manualmente para a station no SPX
• Feche e abra a ferramenta novamente`;
    }

    // ═══════════════════════════════════════════════════════════
    // 6. NÃO LOGADO / SESSÃO EXPIRADA
    // ═══════════════════════════════════════════════════════════
    if (erro.includes('não está logado') || erro.includes('sessão expirada') || erro.includes('login')) {
        return `❌ Você não está logado

Sua sessão expirou ou você não fez login.

O que fazer:
• Clique em "Baixar Dados" para fazer login
• Feche e abra a ferramenta novamente
• Entre manualmente no SPX pelo navegador`;
    }

    // ═══════════════════════════════════════════════════════════
    // 7. ERRO GENÉRICO (fallback)
    // ═══════════════════════════════════════════════════════════
    return `❌ Erro no download

${errorMessage}

O que fazer:
• Feche e abra a ferramenta novamente
• Verifique sua conexão com a internet
• Se o problema persistir, entre em contato com o suporte`;
}

// Exportar para uso no renderer.js
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { traduzirErroParaUsuario };
}
