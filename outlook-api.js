require('dotenv').config();

// Importa as bibliotecas necessárias
const msal = require('@azure/msal-node');
const axios = require('axios');

// Configure as credenciais da sua aplicação
const msalConfig = {
    auth: {
        // ID do Aplicativo (cliente)
        clientId: process.env.MSAL_CLIENT_ID,
        // ID do Diretório (locatário)
        authority: process.env.MSAL_AUTHORITY,
        // Valor do Segredo do Cliente
        clientSecret: process.env.MSAL_CLIENT_SECRET
    }
};

// Configura o escopo de permissão para ler o calendário
const tokenRequest = {
    scopes: ['https://graph.microsoft.com/.default'],
};

// Cria uma instância do cliente de autenticação
const cca = new msal.ConfidentialClientApplication(msalConfig);

async function getCalendarEvents() {
    try {
        // Passo 1: Obter o token de acesso
        const authResult = await cca.acquireTokenByClientCredential(tokenRequest);
        const accessToken = authResult.accessToken;

        if (!accessToken) {
            throw new Error('Não foi possível obter o token de acesso.');
        }

        // Passo 2: Fazer a chamada para a Microsoft Graph API
        const endpoint = 'https://graph.microsoft.com/v1.0/me/calendar/events';
        const response = await axios.get(endpoint, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            }
        });
        
        // Retorna os eventos do calendário
        return response.data.value;

    } catch (error) {
        console.error('Erro ao buscar eventos do calendário:', error.response ? error.response.data : error.message);
        throw error;
    }
}

// Exemplo de como você usaria a função
getCalendarEvents().then(events => {
    console.log('Eventos do Calendário:', events);
    // Aqui você processaria os eventos e enviaria para o front-end
}).catch(err => {
    console.error('Falha ao buscar dados.');
});