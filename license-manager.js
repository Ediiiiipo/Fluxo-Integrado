/**
 * ============================================
 * SISTEMA DE LICENCIAMENTO - v1.0
 * ============================================
 * 
 * Protege o app com:
 * - Validade de 6 meses
 * - Validação de email @shopee.com
 * - Notificação para admins via Gmail
 * - Aprovação via CTRL+*
 * 
 * Admins:
 * - ediiiiipo@gmail.com
 * - petternds@gmail.com
 * ============================================
 */

const crypto = require('crypto');
const fs = require('fs-extra');
const path = require('path');
const os = require('os');

// ============================================
// CONFIGURAÇÕES
// ============================================

const LICENSE_CONFIG = {
  // Admins que podem aprovar
  ADMINS: [
    'ediiiiipo@gmail.com',
    'petternds@gmail.com'
  ],
  
  // Validade padrão (6 meses em dias)
  VALIDITY_DAYS: 180,
  
  // Aviso de expiração (30 dias antes)
  WARNING_DAYS: 30,
  
  // Diretório de dados (oculto)
  DATA_DIR: path.join(os.homedir(), '.shopee-license'),
  
  // Arquivos
  LICENSE_FILE: 'license.json',
  REQUESTS_FILE: 'requests.json',
  
  // SMTP Gmail - CONFIGURAR DEPOIS
  SMTP: {
    enabled: false, // Ativar quando configurar
    service: 'gmail',
    auth: {
      user: 'SEU_EMAIL@gmail.com',     // ← CONFIGURAR
      pass: 'SUA_APP_PASSWORD_AQUI'    // ← CONFIGURAR
    }
  }
};

// ============================================
// CLASSE DE LICENCIAMENTO
// ============================================

class LicenseManager {
  constructor() {
    this.dataDir = LICENSE_CONFIG.DATA_DIR;
    this.licenseFile = path.join(this.dataDir, LICENSE_CONFIG.LICENSE_FILE);
    this.requestsFile = path.join(this.dataDir, LICENSE_CONFIG.REQUESTS_FILE);
  }
  
  // ============================================
  // INICIALIZAÇÃO
  // ============================================
  
  async initialize() {
    // Criar diretório se não existir
    await fs.ensureDir(this.dataDir);
    
    // Inicializar arquivos se não existirem
    if (!await fs.pathExists(this.licenseFile)) {
      await this.createInitialLicense();
    }
    
    if (!await fs.pathExists(this.requestsFile)) {
      await fs.writeJson(this.requestsFile, { requests: [] });
    }
  }
  
  async createInitialLicense() {
    const initialLicense = {
      version: '1.0',
      status: 'active',
      createdAt: new Date().toISOString(),
      expiresAt: this.calculateExpiryDate(LICENSE_CONFIG.VALIDITY_DAYS),
      activatedBy: {
        name: 'Sistema',
        email: 'system@internal',
        date: new Date().toISOString()
      },
      approvedBy: 'system',
      passwordHash: null
    };
    
    await fs.writeJson(this.licenseFile, initialLicense, { spaces: 2 });
  }
  
  // ============================================
  // VALIDAÇÃO DE LICENÇA
  // ============================================
  
  async checkLicense() {
    try {
      const license = await this.getLicense();
      
      if (!license) {
        return { valid: false, reason: 'no_license' };
      }
      
      const now = new Date();
      const expiryDate = new Date(license.expiresAt);
      
      if (now > expiryDate) {
        return { 
          valid: false, 
          reason: 'expired',
          expiryDate: expiryDate.toLocaleDateString('pt-BR')
        };
      }
      
      // Verificar se está próximo de expirar (30 dias)
      const daysUntilExpiry = Math.floor((expiryDate - now) / (1000 * 60 * 60 * 24));
      
      if (daysUntilExpiry <= LICENSE_CONFIG.WARNING_DAYS) {
        return {
          valid: true,
          warning: true,
          daysRemaining: daysUntilExpiry,
          expiryDate: expiryDate.toLocaleDateString('pt-BR')
        };
      }
      
      return { 
        valid: true,
        daysRemaining: daysUntilExpiry,
        expiryDate: expiryDate.toLocaleDateString('pt-BR')
      };
      
    } catch (error) {
      console.error('❌ Erro ao verificar licença:', error);
      return { valid: false, reason: 'error' };
    }
  }
  
  // ============================================
  // SOLICITAÇÃO DE RENOVAÇÃO
  // ============================================
  
  async requestRenewal(nome, email) {
    try {
      // Validar email @shopee.com
      if (!this.validateShopeeEmail(email)) {
        return {
          success: false,
          error: 'Email deve ser @shopee.com'
        };
      }
      
      // Validar nome
      if (!nome || nome.trim().length < 3) {
        return {
          success: false,
          error: 'Nome muito curto (mínimo 3 caracteres)'
        };
      }
      
      // Gerar código de solicitação
      const requestCode = this.generateRequestCode();
      
      // Coletar informações do sistema
      const systemInfo = this.getSystemInfo();
      
      // Criar solicitação
      const request = {
        code: requestCode,
        status: 'pending',
        requestedAt: new Date().toISOString(),
        solicitante: {
          name: nome.trim(),
          email: email.trim().toLowerCase(),
          computer: systemInfo.hostname,
          user: systemInfo.username,
          ip: systemInfo.ip
        },
        approvedBy: null,
        approvedAt: null,
        password: null
      };
      
      // Salvar solicitação
      await this.saveRequest(request);
      
      // Criar mensagem para copiar/colar
      const emailMessage = this.generateEmailMessage(request);
      
      return {
        success: true,
        requestCode: requestCode,
        emailMessage: emailMessage,
        adminEmails: LICENSE_CONFIG.ADMINS,
        message: 'Solicitação criada! Copie a mensagem e envie para os admins.'
      };
      
    } catch (error) {
      console.error('❌ Erro ao solicitar renovação:', error);
      return {
        success: false,
        error: 'Erro ao processar solicitação'
      };
    }
  }
  
  // ============================================
  // APROVAÇÃO DE SOLICITAÇÃO
  // ============================================
  
  async approveRequest(requestCode, approvedBy) {
    try {
      // Buscar solicitação
      const request = await this.getRequest(requestCode);
      
      if (!request) {
        return {
          success: false,
          error: 'Solicitação não encontrada'
        };
      }
      
      if (request.status === 'approved') {
        return {
          success: false,
          error: 'Solicitação já foi aprovada anteriormente',
          password: request.password || '[SENHA JÁ ENVIADA]'
        };
      }
      
      // Gerar senha
      const password = this.generatePassword();
      
      // Atualizar solicitação
      request.status = 'approved';
      request.approvedBy = approvedBy;
      request.approvedAt = new Date().toISOString();
      request.password = password;
      
      await this.updateRequest(request);
      
      // Gerar mensagem de email para copiar
      const approvalMessage = this.generateApprovalMessage(request);
      
      return {
        success: true,
        password: password,
        solicitante: request.solicitante,
        approvalMessage: approvalMessage,
        message: 'Solicitação aprovada! Copie a mensagem e envie para o solicitante.'
      };
      
    } catch (error) {
      console.error('❌ Erro ao aprovar solicitação:', error);
      return {
        success: false,
        error: 'Erro ao aprovar solicitação'
      };
    }
  }
  
  // ============================================
  // ATIVAÇÃO DE LICENÇA
  // ============================================
  
  async activateLicense(password) {
    try {
      const passwordHash = this.hashPassword(password);
      
      // Verificar se senha existe em alguma solicitação aprovada
      const requests = await this.getAllRequests();
      const validRequest = requests.find(r => 
        r.status === 'approved' && 
        r.password &&
        this.hashPassword(r.password) === passwordHash
      );
      
      if (!validRequest) {
        return {
          success: false,
          error: 'Senha inválida ou não encontrada'
        };
      }
      
      // Ativar licença
      const expiryDate = this.calculateExpiryDate(LICENSE_CONFIG.VALIDITY_DAYS);
      const license = {
        version: '1.0',
        status: 'active',
        createdAt: new Date().toISOString(),
        expiresAt: expiryDate,
        activatedBy: validRequest.solicitante,
        approvedBy: validRequest.approvedBy,
        passwordHash: passwordHash,
        requestCode: validRequest.code
      };
      
      await fs.writeJson(this.licenseFile, license, { spaces: 2 });
      
      return {
        success: true,
        expiryDate: new Date(expiryDate).toLocaleDateString('pt-BR'),
        daysValid: LICENSE_CONFIG.VALIDITY_DAYS,
        message: 'Licença ativada com sucesso! Válida por 6 meses.'
      };
      
    } catch (error) {
      console.error('❌ Erro ao ativar licença:', error);
      return {
        success: false,
        error: 'Erro ao ativar licença'
      };
    }
  }
  
  // ============================================
  // MENSAGENS DE EMAIL (COPIAR/COLAR)
  // ============================================
  
  generateEmailMessage(request) {
    return `
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🔐 SOLICITAÇÃO DE RENOVAÇÃO - SHOPEE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📅 Data: ${new Date(request.requestedAt).toLocaleString('pt-BR')}
👤 Nome: ${request.solicitante.name}
📧 Email: ${request.solicitante.email}
💻 Computador: ${request.solicitante.computer}
🪟 Usuário: ${request.solicitante.user}
🌐 IP: ${request.solicitante.ip}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
CÓDIGO DE APROVAÇÃO
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

${request.code}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
INSTRUÇÕES PARA APROVAR
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. Abra o app Shopee - Planejamento Integrado
2. Pressione CTRL + *
3. Cole o código: ${request.code}
4. Clique em "Gerar Senha"
5. Copie e envie a senha para: ${request.solicitante.email}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
`.trim();
  }
  
  generateApprovalMessage(request) {
    const expiryDate = this.calculateExpiryDate(LICENSE_CONFIG.VALIDITY_DAYS);
    const formattedDate = new Date(expiryDate).toLocaleDateString('pt-BR');
    
    return `
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
✅ RENOVAÇÃO APROVADA - SHOPEE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Olá ${request.solicitante.name},

Sua solicitação de renovação foi APROVADA!
Aprovado por: ${request.approvedBy}

🔑 SENHA DE ATIVAÇÃO:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

${request.password}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
⏰ Validade: 6 meses (até ${formattedDate})
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📋 INSTRUÇÕES:

1. Abra o app Shopee - Planejamento Integrado
2. Digite a senha acima
3. Clique em "Ativar Licença"
4. Pronto! ✅

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
⚠️ IMPORTANTE:
• Guarde esta senha em local seguro
• Não compartilhe com outras pessoas
• Esta senha só pode ser usada uma vez
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Dúvidas? Entre em contato:
ediiiiipo@gmail.com ou petternds@gmail.com
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
`.trim();
  }
  
  // ============================================
  // UTILITÁRIOS
  // ============================================
  
  validateShopeeEmail(email) {
    if (!email || typeof email !== 'string') return false;
    
    const emailLower = email.toLowerCase().trim();
    
    // Deve terminar com @shopee.com
    if (!emailLower.endsWith('@shopee.com')) return false;
    
    // Formato básico de email
    const regex = /^[^\s@]+@shopee\.com$/;
    return regex.test(emailLower);
  }
  
  generateRequestCode() {
    const data = new Date();
    const ano = data.getFullYear();
    const mes = String(data.getMonth() + 1).padStart(2, '0');
    const dia = String(data.getDate()).padStart(2, '0');
    const random = Math.random().toString(36).substring(2, 8).toUpperCase();
    
    return `REQ-${ano}-${mes}-${dia}-${random}`;
  }
  
  generatePassword() {
    const ano = new Date().getFullYear();
    const random1 = Math.random().toString(36).substring(2, 6).toUpperCase();
    const random2 = Math.random().toString(36).substring(2, 6).toUpperCase();
    const random3 = Math.random().toString(36).substring(2, 6).toUpperCase();
    
    return `SHOP-${ano}-${random1}-${random2}-${random3}`;
  }
  
  hashPassword(password) {
    return crypto
      .createHash('sha256')
      .update(password)
      .digest('hex');
  }
  
  calculateExpiryDate(days) {
    const date = new Date();
    date.setDate(date.getDate() + days);
    return date.toISOString();
  }
  
  getSystemInfo() {
    return {
      hostname: os.hostname(),
      username: os.userInfo().username,
      platform: os.platform(),
      ip: this.getLocalIP()
    };
  }
  
  getLocalIP() {
    const nets = os.networkInterfaces();
    for (const name of Object.keys(nets)) {
      for (const net of nets[name]) {
        if (net.family === 'IPv4' && !net.internal) {
          return net.address;
        }
      }
    }
    return 'N/A';
  }
  
  // ============================================
  // PERSISTÊNCIA
  // ============================================
  
  async getLicense() {
    try {
      return await fs.readJson(this.licenseFile);
    } catch (error) {
      return null;
    }
  }
  
  async saveRequest(request) {
    const data = await fs.readJson(this.requestsFile);
    data.requests.push(request);
    await fs.writeJson(this.requestsFile, data, { spaces: 2 });
  }
  
  async getRequest(code) {
    const data = await fs.readJson(this.requestsFile);
    return data.requests.find(r => r.code === code);
  }
  
  async updateRequest(request) {
    const data = await fs.readJson(this.requestsFile);
    const index = data.requests.findIndex(r => r.code === request.code);
    if (index !== -1) {
      data.requests[index] = request;
      await fs.writeJson(this.requestsFile, data, { spaces: 2 });
    }
  }
  
  async getAllRequests() {
    try {
      const data = await fs.readJson(this.requestsFile);
      return data.requests || [];
    } catch (error) {
      return [];
    }
  }
  
  async getRecentRequests(limit = 10) {
    const requests = await this.getAllRequests();
    return requests
      .sort((a, b) => new Date(b.requestedAt) - new Date(a.requestedAt))
      .slice(0, limit);
  }
  
  // ============================================
  // ADMIN - EXTENSÃO MANUAL
  // ============================================
  
  async extendLicenseManually(months = 6) {
    try {
      const license = await this.getLicense();
      if (!license) {
        return {
          success: false,
          error: 'Nenhuma licença encontrada'
        };
      }
      
      const currentExpiry = new Date(license.expiresAt);
      const newExpiry = new Date(currentExpiry);
      newExpiry.setMonth(newExpiry.getMonth() + months);
      
      license.expiresAt = newExpiry.toISOString();
      license.lastExtendedAt = new Date().toISOString();
      
      await fs.writeJson(this.licenseFile, license, { spaces: 2 });
      
      return {
        success: true,
        newExpiryDate: newExpiry.toLocaleDateString('pt-BR'),
        message: `Licença estendida por ${months} meses`
      };
    } catch (error) {
      return {
        success: false,
        error: 'Erro ao estender licença'
      };
    }
  }
}

// ============================================
// EXPORTAR
// ============================================

module.exports = LicenseManager;