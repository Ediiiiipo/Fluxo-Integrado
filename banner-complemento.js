// 💡 BANNER DE SUGESTÃO DE COMPLEMENTO
// Mostra banner informativo sugerindo LH para complementar o CAP

function mostrarBannerComplemento(lhInfo, faltam) {
    console.log(`💡 [BANNER] Mostrando sugestão de complemento: ${lhInfo.lhTrip}`);
    
    // Criar banner se não existir
    let banner = document.getElementById('bannerComplemento');
    if (!banner) {
        banner = document.createElement('div');
        banner.id = 'bannerComplemento';
        banner.style.cssText = `
            position: fixed;
            bottom: 30px;
            right: 30px;
            background: linear-gradient(135deg, #2196f3 0%, #1976d2 100%);
            color: white;
            padding: 24px 30px;
            border-radius: 16px;
            box-shadow: 0 8px 32px rgba(33, 150, 243, 0.4);
            z-index: 1000;
            max-width: 450px;
            animation: slideInFromRight 0.5s ease-out;
            border: 2px solid rgba(255, 255, 255, 0.3);
        `;
        
        banner.innerHTML = `
            <div style="display: flex; align-items: flex-start; gap: 16px;">
                <div style="font-size: 40px; line-height: 1;">💡</div>
                <div style="flex: 1;">
                    <div style="font-size: 18px; font-weight: 700; margin-bottom: 8px;">
                        Sugestão de Complemento
                    </div>
                    <div style="font-size: 14px; line-height: 1.6; margin-bottom: 16px; opacity: 0.95;">
                        <strong>LH ${lhInfo.lhTrip}</strong> (${lhInfo.qtdPedidos.toLocaleString('pt-BR')} pedidos) pode complementar o CAP.<br>
                        <span style="opacity: 0.8;">Faltam ${faltam.toLocaleString('pt-BR')} pedidos para completar.</span>
                    </div>
                    <div style="display: flex; gap: 12px;">
                        <button id="btnAbrirTOsComplemento" style="
                            background: white;
                            color: #2196f3;
                            border: none;
                            padding: 10px 20px;
                            border-radius: 8px;
                            font-weight: 600;
                            cursor: pointer;
                            font-size: 14px;
                            transition: all 0.3s ease;
                        ">
                            ✅ Sim, abrir TOs
                        </button>
                        <button id="btnFecharBannerComplemento" style="
                            background: rgba(255, 255, 255, 0.2);
                            color: white;
                            border: 1px solid rgba(255, 255, 255, 0.5);
                            padding: 10px 20px;
                            border-radius: 8px;
                            font-weight: 600;
                            cursor: pointer;
                            font-size: 14px;
                            transition: all 0.3s ease;
                        ">
                            ❌ Não, obrigado
                        </button>
                    </div>
                </div>
            </div>
        `;
        
        document.body.appendChild(banner);
        
        // Adicionar animação CSS
        const style = document.createElement('style');
        style.textContent = `
            @keyframes slideInFromRight {
                from {
                    transform: translateX(500px);
                    opacity: 0;
                }
                to {
                    transform: translateX(0);
                    opacity: 1;
                }
            }
            
            @keyframes slideOutToRight {
                from {
                    transform: translateX(0);
                    opacity: 1;
                }
                to {
                    transform: translateX(500px);
                    opacity: 0;
                }
            }
            
            #btnAbrirTOsComplemento:hover {
                background: #f5f5f5 !important;
                transform: scale(1.05);
            }
            
            #btnFecharBannerComplemento:hover {
                background: rgba(255, 255, 255, 0.3) !important;
                transform: scale(1.05);
            }
        `;
        document.head.appendChild(style);
        
        // Event listeners
        document.getElementById('btnAbrirTOsComplemento').addEventListener('click', () => {
            console.log(`💡 [BANNER] Usuário aceitou sugestão, abrindo modal de TOs`);
            fecharBannerComplemento();
            
            // Abrir modal de TOs
            setTimeout(() => {
                abrirModalTOs(lhInfo.lhTrip);
                
                // Mostrar mensagem informativa
                setTimeout(() => {
                    alert('💡 LH de Complemento!\n\n' +
                          `LH: ${lhInfo.lhTrip}\n` +
                          `Pedidos totais: ${lhInfo.qtdPedidos.toLocaleString('pt-BR')}\n` +
                          `Faltam para completar CAP: ${faltam.toLocaleString('pt-BR')}\n\n` +
                          '💡 Sugestão: Selecione TOs para completar exatamente o CAP!\n' +
                          'As TOs serão pré-selecionadas via FIFO.');
                }, 500);
            }, 300);
        });
        
        document.getElementById('btnFecharBannerComplemento').addEventListener('click', () => {
            console.log(`💡 [BANNER] Usuário recusou sugestão`);
            fecharBannerComplemento();
            // Limpar flag de complemento sugerido
            window.lhComplementoSugerida = null;
            // Re-renderizar para remover destaque
            renderizarTabelaPlanejamento();
        });
    }
}

function fecharBannerComplemento() {
    const banner = document.getElementById('bannerComplemento');
    if (banner) {
        banner.style.animation = 'slideOutToRight 0.5s ease-out';
        setTimeout(() => {
            banner.remove();
        }, 500);
    }
}
