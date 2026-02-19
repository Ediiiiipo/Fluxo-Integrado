/**
 * ============================================
 * TROCA DE STATION VIA API - VERSÃO FUNCIONAL
 * ============================================
 * 
 * Baseado no formato real da API do SPX Shopee
 * 
 * ============================================
 */

/**
 * Troca de station usando API direta (RÁPIDO)
 * @param {Page} page - Objeto page do Playwright
 * @param {string} stationId - ID da station (número)
 * @returns {Promise<boolean>} true se sucesso
 */
async function trocarStationViaAPI(page, stationId) {
  try {
    console.log(`🚀 Trocando para station ID: ${stationId} via API...`);
    
    // Fazer a chamada POST para trocar de station (URL relativa como na extensão)
    const response = await page.evaluate(async (id) => {
      try {
        const getCookie = (name) => {
          const match = document.cookie.match(new RegExp('(?:^|;\\s*)' + name + '=([^;]*)'));
          return match ? match[1] : '';
        };
        
        const res = await fetch('/api/admin/basicserver/change_station/', {
          method: 'POST',
          headers: {
            'accept': 'application/json',
            'app': 'FMS Portal',
            'content-type': 'application/json;charset=UTF-8',
            'x-csrftoken': getCookie('csrftoken'),
            'device-id': getCookie('spx-admin-device-id')
          },
          credentials: 'include',
          body: JSON.stringify({ station_id: parseInt(id) })
        });
        
        const data = await res.json();
        console.log('📦 DEBUG: Resposta da API:', data);
        
        // Verificar retcode como na extensão
        if (data.retcode === 0) {
          return { success: true, data };
        } else {
          return { success: false, error: data.message || 'Erro desconhecido', data };
        }
      } catch (error) {
        return { success: false, error: error.message };
      }
    }, stationId);
    
    if (!response.success) {
      console.error('❌ Erro na API:', response.error || response.data);
      return false;
    }
    
    console.log('✅ API respondeu com sucesso!');
    
    // Recarregar a página para aplicar a mudança
    console.log('🔄 Recarregando página...');
    await page.reload({ waitUntil: 'networkidle' });
    
    console.log('✅ Station trocada com sucesso!');
    return true;
    
  } catch (error) {
    console.error('❌ Erro ao trocar station:', error.message);
    return false;
  }
}

/**
 * Busca o ID de uma station pelo nome
 * @param {Page} page - Objeto page do Playwright
 * @param {string} stationName - Nome da station
 * @returns {Promise<number|null>} ID da station ou null
 */
async function buscarStationIdPorNome(page, stationName) {
  try {
    console.log(`🔍 Buscando ID da station: ${stationName}`);
    
    const result = await page.evaluate(async (name) => {
      try {
        // Buscar lista de stations (URL relativa)
        const res = await fetch('/api/admin/basicserver/current_user/station_list/?count=999&status_list=0');
        const data = await res.json();
        
        // Formato da API: { retcode: 0, message: "success", data: { station_list: [...], role_list: [...], email: "..." } }
        if (!data.data || !data.data.station_list) {
          console.error('❌ Formato inesperado:', Object.keys(data.data || {}));
          return { success: false, error: 'Formato de resposta inválido' };
        }
        
        const stations = data.data.station_list;
        console.log('✅ Stations encontradas:', stations.length);
        
        // Procurar pela station - usar station_name (com underscore)
        const station = stations.find(s => s.station_name === name);
        
        if (station) {
          console.log('✅ Station encontrada:', station.station_name, '(ID:', station.id, ')');
          // Campo é "id" (não station_id)
          return { success: true, id: station.id, name: station.station_name };
        }
        
        // Se não encontrar, listar primeiras 5
        const primeiras = stations.slice(0, 5).map(s => s.station_name);
        console.log('⚠️ Station não encontrada. Primeiras disponíveis:', primeiras);
        
        return { success: false, error: 'Station não encontrada', availableStations: primeiras };
      } catch (error) {
        console.error('❌ Erro na avaliação:', error.message);
        console.error('❌ Stack:', error.stack);
        return { success: false, error: error.message, stack: error.stack };
      }
    }, stationName);
    
    if (!result.success) {
      console.error(`❌ Erro: ${result.error}`);
      if (result.availableStations) {
        console.error('🔍 Primeiras stations disponíveis:', result.availableStations);
      }
      return null;
    }
    
    console.log(`✅ Station encontrada: ${result.name} (ID: ${result.id})`);
    return result.id;
    
  } catch (error) {
    console.error('❌ Erro ao buscar station:', error.message);
    return null;
  }
}

/**
 * Troca de station completa: busca ID + troca via API
 * @param {Page} page - Objeto page do Playwright
 * @param {string} stationName - Nome da station
 * @returns {Promise<boolean>} true se sucesso
 */
async function trocarStationCompleto(page, stationName) {
  console.log('');
  console.log('═'.repeat(70));
  console.log('🔄 TROCA DE STATION VIA API (MÉTODO RÁPIDO)');
  console.log('═'.repeat(70));
  
  // 1. Buscar ID da station
  const stationId = await buscarStationIdPorNome(page, stationName);
  
  if (!stationId) {
    console.error('❌ Não foi possível encontrar a station');
    return false;
  }
  
  // 2. Trocar via API
  const sucesso = await trocarStationViaAPI(page, stationId);
  
  if (sucesso) {
    console.log('');
    console.log('✅ Station trocada com sucesso!');
    console.log('═'.repeat(70));
  } else {
    console.error('');
    console.error('❌ Falha ao trocar station');
    console.error('═'.repeat(70));
  }
  
  return sucesso;
}

// Exportar funções
module.exports = {
  trocarStationViaAPI,
  buscarStationIdPorNome,
  trocarStationCompleto
};
