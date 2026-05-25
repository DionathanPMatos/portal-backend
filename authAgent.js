// Puxe a mesma chave do seu server.js ou .env
const DCA_AGENT_SECRET_KEY = process.env.DCA_AGENT_SECRET_KEY || "dca@2025#agent";

module.exports = function(req, res, next) {
    const agentSecret = req.header('X-Agent-Secret');

    if (!agentSecret) {
        return res.status(401).json({ msg: 'Nenhum segredo de agente, autorização negada' });
    }

    if (agentSecret !== DCA_AGENT_SECRET_KEY) {
        return res.status(401).json({ msg: 'Segredo de agente inválido' });
    }

    // Segredo válido, pode prosseguir
    next();
};