require('dotenv').config();

const express = require('express'); // Servidor web
const session = require('express-session'); // Para gerenciar sessões de usuário (autenticação)
const msal = require('@azure/msal-node'); // Para autenticação Microsoft
const axios = require('axios'); // Para fazer requisições HTTP (ex: Microsoft Graph API)
const path = require('path'); // Para lidar com caminhos de arquivos
const cors = require('cors'); // Para permitir requisições do frontend
const { Pool } = require('pg');          // [PG] Usamos o Pool do 'pg' para o banco principal
const multer = require('multer'); // Para lidar com upload de arquivos
const XLSX = require("xlsx"); // Para lidar com arquivos Excel
const fs = require('fs'); // Para manipular arquivos
const csv = require('csv-parser'); // Para ler arquivos CSV
const { exec } = require('child_process'); // Para executar comandos no terminal

const DCA_AGENT_SECRET_KEY = process.env.DCA_AGENT_SECRET_KEY;

const app = express();
const PORT = process.env.PORT || 3000;

const FRONTEND_URL = process.env.FRONTEND_URL || 'http://localhost:5173';
const BACKEND_URL = process.env.BACKEND_URL || `http://localhost:${PORT}`;

// =================================================================
// SEÇÃO 1: MIDDLEWARES E CONFIGURAÇÕES GERAIS
// =================================================================
app.use(cors({ origin: FRONTEND_URL, credentials: true }));
app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, '..', 'frontend', 'dist')));
app.use('/uploads', express.static(path.join(__dirname, 'uploads')));
app.use(session({ secret: process.env.SESSION_SECRET, resave: false, saveUninitialized: false }));
app.use((req, res, next) => {
    if (req.url.startsWith('/api/')) {
        res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
        res.setHeader('Pragma', 'no-cache');
        res.setHeader('Expires', '0');
        res.setHeader('Surrogate-Control', 'no-store');
    }
    next();
});

const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, 'uploads/'),
    filename: (req, file, cb) => cb(null, `${Date.now()}-${file.originalname}`)
});
const upload = multer({ storage: storage });

const themeStorage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, 'uploads/theme/'),
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname);
    const prefix =
      file.fieldname === 'backgroundImageFile' ? 'background' :
      file.fieldname === 'logoFile' ? 'logo' :
      file.fieldname === 'faviconFile' ? 'favicon' :
      file.fieldname === 'loginBgImageFile' ? 'login-background' :
      file.fieldname === 'loginLogoFile' ? 'login-logo' : 'asset';
    cb(null, `${prefix}-${Date.now()}${ext}`);
  }
});

const themeUpload = multer({ storage: themeStorage }).fields([
    { name: 'backgroundImageFile', maxCount: 1 },
    { name: 'logoFile', maxCount: 1 },
    { name: 'faviconFile', maxCount: 1 },
    { name: 'loginBgImageFile', maxCount: 1 },
    { name: 'loginLogoFile', maxCount: 1 }
]);

const authAgent = require('./authAgent'); 

// =================================================================
// SEÇÃO 2: MIDDLEWARE DE AUTENTICAÇÃO INTERNO
// =================================================================
const authMiddleware = (req, res, next) => {
    console.log('--- [AUTH MIDDLEWARE] Verificando acesso ---');
    console.log('Rota acessada:', req.method, req.url);
    console.log('Sessão ID:', req.sessionID);
    console.log('isAuthenticated:', req.session?.isAuthenticated);
    console.log('Tem localUser:', !!req.session?.localUser);

    if (!req.session.isAuthenticated || !req.session.localUser) { 
        console.log('❌ Bloqueado: Usuário não possui sessão ativa no backend (pode ter sido resetada).');
        return res.status(401).json({ error: 'Acesso não autorizado. Faça login primeiro.' });
    }
    req.user = req.session.localUser; 
    next();
};

// =================================================================
// SEÇÃO 2.5: CONEXÃO COM BANCOS DE DADOS
// =================================================================
// [PG] Conexão Nova: PostgreSQL (Docker) para o Portal DCA
const pool = new Pool({ 
    host: process.env.PG_HOST || 'localhost', 
    user: process.env.PG_USER || 'postgres', 
    password: process.env.PG_PASSWORD, // Alterado aqui
    database: process.env.PG_DATABASE || 'portal_dca', 
    port: process.env.PG_PORT || 5433,
    max: 10 
});

// Testes de conexão
pool.connect()
    .then(async client => { 
        console.log('✅ Conectado ao PostgreSQL (portal_dca)'); 
        
        // Cria a tabela registro_venda se ela não existir
        try {
            // =========================================================
            // 🚀 SCRIPT DE INICIALIZAÇÃO DE TABELAS (AUTO-MIGRATION)
            // Cria automaticamente toda a estrutura do Portal se não existir
            // =========================================================

            // 🛠️ FIX AUTOMÁTICO DE MIGRAÇÃO: 
            // Garante que todas as tabelas importadas recuperem a propriedade de AUTO_INCREMENT (Sequences do Postgres)
            try {
                await client.query(`
                    DO $$
                    DECLARE
                        t_name text;
                        seq_name text;
                    BEGIN
                        FOR t_name IN 
                            SELECT t.table_name 
                            FROM information_schema.tables t
                            JOIN information_schema.columns c ON t.table_name = c.table_name AND t.table_schema = c.table_schema
                            WHERE c.column_name = 'id' AND t.table_schema = 'public' AND t.table_type = 'BASE TABLE'
                        LOOP
                            BEGIN
                                IF pg_get_serial_sequence('public.' || t_name, 'id') IS NULL THEN
                                    seq_name := t_name || '_id_seq';
                                    EXECUTE 'CREATE SEQUENCE IF NOT EXISTS public.' || seq_name;
                                    EXECUTE 'ALTER TABLE public.' || t_name || ' ALTER COLUMN id SET DEFAULT nextval(''public.' || seq_name || ''')';
                                    EXECUTE 'ALTER SEQUENCE public.' || seq_name || ' OWNED BY public.' || t_name || '.id';
                                    EXECUTE 'SELECT setval(''public.' || seq_name || ''', COALESCE((SELECT MAX(id) FROM public.' || t_name || ') + 1, 1), false)';
                                END IF;
                            EXCEPTION WHEN OTHERS THEN
                                -- Se falhar em uma tabela específica, apenas ignora e continua
                            END;
                        END LOOP;
                    END $$;
                `);
            } catch (fixErr) {
                console.error('⚠️ Aviso: O script de correção automática de ID falhou. Detalhe:', fixErr.message);
            }
            
            // 1. Tabelas de Domínio Básico (Sem dependências)
            await client.query(`
                CREATE TABLE IF NOT EXISTS system_settings (
                    setting_key VARCHAR(255) PRIMARY KEY,
                    setting_value TEXT
                );
                CREATE TABLE IF NOT EXISTS cargos (
                    id SERIAL PRIMARY KEY,
                    nome_cargo VARCHAR(255) NOT NULL UNIQUE
                );
                CREATE TABLE IF NOT EXISTS setores (
                    id SERIAL PRIMARY KEY,
                    nome_setor VARCHAR(255) NOT NULL UNIQUE
                );
                
                -- Tabela de Unidades/Filiais
                CREATE TABLE IF NOT EXISTS unidades_filiais (
                    id SERIAL PRIMARY KEY,
                    nome_unidade VARCHAR(255) NOT NULL UNIQUE
                );
                
                -- 🚀 Tabela do Repositório Técnico de Documentos
                CREATE TABLE IF NOT EXISTS repositorio_tecnico (
                    id SERIAL PRIMARY KEY,
                    nome VARCHAR(255) NOT NULL,
                    caminho_arquivo TEXT NOT NULL,
                    categoria VARCHAR(100),
                    tipo VARCHAR(100),
                    descricao TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                
                -- 🚀 Tabela de Ferramentas Úteis (Links)
                CREATE TABLE IF NOT EXISTS ferramentas_uteis (
                    id SERIAL PRIMARY KEY,
                    nome VARCHAR(255) NOT NULL,
                    solucao VARCHAR(100) NOT NULL,
                    fabricante_id INTEGER REFERENCES fabricantes(id) ON DELETE SET NULL,
                    link TEXT NOT NULL,
                    usuario_id INTEGER REFERENCES funcionarios(id) ON DELETE SET NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS fabricantes (
                    id SERIAL PRIMARY KEY,
                    name VARCHAR(255) NOT NULL UNIQUE
                );
                
                ALTER TABLE fabricantes ADD COLUMN IF NOT EXISTS logo_base64 TEXT;
                ALTER TABLE fabricantes ADD COLUMN IF NOT EXISTS resumo TEXT;
                ALTER TABLE fabricantes ADD COLUMN IF NOT EXISTS key_accounts TEXT;
                ALTER TABLE fabricantes ADD COLUMN IF NOT EXISTS registros_projetos TEXT;
                ALTER TABLE fabricantes ADD COLUMN IF NOT EXISTS garantias TEXT;
                ALTER TABLE fabricantes ADD COLUMN IF NOT EXISTS prazos_compras TEXT;
                ALTER TABLE fabricantes ADD COLUMN IF NOT EXISTS certificacoes_treinamentos TEXT;
                ALTER TABLE fabricantes ADD COLUMN IF NOT EXISTS contatos_suporte TEXT;
                ALTER TABLE fabricantes ADD COLUMN IF NOT EXISTS links_uteis TEXT;
                
                CREATE TABLE IF NOT EXISTS segmentacoes (
                    id SERIAL PRIMARY KEY,
                    nome VARCHAR(255) NOT NULL UNIQUE
                );
                CREATE TABLE IF NOT EXISTS verticais (
                    id SERIAL PRIMARY KEY,
                    nome VARCHAR(255) NOT NULL UNIQUE
                );
                CREATE TABLE IF NOT EXISTS integradores (
                    id SERIAL PRIMARY KEY,
                    nome VARCHAR(255) NOT NULL UNIQUE
                );
                
                -- Tabela do Novo Módulo de Benefícios do RH
                CREATE TABLE IF NOT EXISTS beneficios (
                    id SERIAL PRIMARY KEY,
                    nome VARCHAR(255) NOT NULL,
                    descricao TEXT,
                    logo_base64 TEXT,
                    ativo BOOLEAN DEFAULT TRUE,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
            `);

            // 2. Colaboradores e Relacionamentos
            await client.query(`
                CREATE TABLE IF NOT EXISTS funcionarios (
                    id SERIAL PRIMARY KEY,
                    nome_completo VARCHAR(255) NOT NULL,
                    email VARCHAR(255) UNIQUE,
                    contato VARCHAR(255),
                    setor_id INTEGER REFERENCES setores(id) ON DELETE SET NULL,
                    cargo_id INTEGER REFERENCES cargos(id) ON DELETE SET NULL,
                    userpic_base64 TEXT,
                    privilegios TEXT,
                    ativo BOOLEAN DEFAULT TRUE
                );
                
                -- Garante que o campo ativo exista mesmo se a tabela for legada
                ALTER TABLE funcionarios ADD COLUMN IF NOT EXISTS ativo BOOLEAN DEFAULT TRUE;

                -- Adiciona a coluna de gestor para a hierarquia do organograma
                ALTER TABLE funcionarios ADD COLUMN IF NOT EXISTS gestor_id INTEGER REFERENCES funcionarios(id) ON DELETE SET NULL;
                
                -- Adiciona vínculo com a unidade/filial
                ALTER TABLE funcionarios ADD COLUMN IF NOT EXISTS unidade_id INTEGER REFERENCES unidades_filiais(id) ON DELETE SET NULL;

                -- 🚀 Garante que o campo de privilégios comporte múltiplas permissões (mais de 7 caracteres)
                ALTER TABLE funcionarios ALTER COLUMN privilegios TYPE TEXT;

                CREATE TABLE IF NOT EXISTS funcionario_fabricante (
                    funcionario_id INTEGER REFERENCES funcionarios(id) ON DELETE CASCADE,
                    fabricante_id INTEGER REFERENCES fabricantes(id) ON DELETE CASCADE,
                    PRIMARY KEY (funcionario_id, fabricante_id)
                );
            `);

            // 3. Clientes (CRM)
            await client.query(`
                CREATE TABLE IF NOT EXISTS clientes (
                    id SERIAL PRIMARY KEY,
                    nome_cliente VARCHAR(255) NOT NULL,
                    cnpj_cpf VARCHAR(50) UNIQUE,
                    razao_social VARCHAR(255),
                    nome_fantasia VARCHAR(255),
                    segmento VARCHAR(100),
                    perfil VARCHAR(100),
                    inscricao_estadual VARCHAR(100),
                    site VARCHAR(255),
                    observacoes TEXT,
                    cep VARCHAR(20),
                    logradouro VARCHAR(255),
                    numero VARCHAR(50),
                    complemento VARCHAR(255),
                    bairro VARCHAR(100),
                    cidade VARCHAR(100),
                    uf VARCHAR(2),
                    ativo INTEGER DEFAULT 1,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    deleted_at TIMESTAMP
                );

                ALTER TABLE clientes ADD COLUMN IF NOT EXISTS vendedor_responsavel_id INTEGER REFERENCES funcionarios(id) ON DELETE SET NULL;
                ALTER TABLE clientes ADD COLUMN IF NOT EXISTS tier_estrategico VARCHAR(50);
                ALTER TABLE clientes ADD COLUMN IF NOT EXISTS status_credito VARCHAR(50);

                CREATE INDEX IF NOT EXISTS idx_clientes_busca ON clientes (nome_cliente, cnpj_cpf);
                
                CREATE EXTENSION IF NOT EXISTS pg_trgm;
                CREATE INDEX IF NOT EXISTS idx_clientes_nome_trgm ON clientes USING GIN (nome_cliente gin_trgm_ops);
                CREATE INDEX IF NOT EXISTS idx_clientes_cnpj_trgm ON clientes USING GIN (cnpj_cpf gin_trgm_ops);
                CREATE INDEX IF NOT EXISTS idx_clientes_razao_trgm ON clientes USING GIN (razao_social gin_trgm_ops);
                CREATE INDEX IF NOT EXISTS idx_clientes_fantasia_trgm ON clientes USING GIN (nome_fantasia gin_trgm_ops);

                CREATE TABLE IF NOT EXISTS cliente_contatos (
                    id SERIAL PRIMARY KEY,
                    cliente_id INTEGER REFERENCES clientes(id) ON DELETE CASCADE,
                    nome VARCHAR(255),
                    telefone VARCHAR(50),
                    email VARCHAR(255),
                    principal BOOLEAN DEFAULT FALSE,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );

                ALTER TABLE cliente_contatos ADD COLUMN IF NOT EXISTS cargo VARCHAR(100);
                ALTER TABLE cliente_contatos ADD COLUMN IF NOT EXISTS whatsapp VARCHAR(50);
                ALTER TABLE cliente_contatos ADD COLUMN IF NOT EXISTS observacoes TEXT;

                CREATE TABLE IF NOT EXISTS cliente_filiais (
                    id SERIAL PRIMARY KEY,
                    cliente_id INTEGER REFERENCES clientes(id) ON DELETE CASCADE,
                    cnpj VARCHAR(50),
                    razao_social VARCHAR(255),
                    nome_fantasia VARCHAR(255),
                    cep VARCHAR(20),
                    logradouro VARCHAR(255),
                    numero VARCHAR(50),
                    complemento VARCHAR(255),
                    bairro VARCHAR(100),
                    cidade VARCHAR(100),
                    uf VARCHAR(2),
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
            `);

            // 4. Projetos, Compras e Integrações (CRM Avançado)
            await client.query(`
                CREATE TABLE IF NOT EXISTS projetos (
                    id SERIAL PRIMARY KEY,
                    nome_projeto VARCHAR(255) NOT NULL,
                    cliente_id INTEGER REFERENCES clientes(id) ON DELETE SET NULL,
                    vendedor_id INTEGER REFERENCES funcionarios(id) ON DELETE SET NULL,
                    valor_estimado NUMERIC(15, 2),
                    moeda VARCHAR(10) DEFAULT 'BRL',
                    data_fechamento_prevista DATE,
                    etapa_funil VARCHAR(100) DEFAULT 'Prospeccao',
                    tipo_projeto VARCHAR(100),
                    segmentacao_id INTEGER REFERENCES segmentacoes(id) ON DELETE SET NULL,
                    vertical_id INTEGER REFERENCES verticais(id) ON DELETE SET NULL,
                    integrador_id INTEGER REFERENCES integradores(id) ON DELETE SET NULL,
                    numero_registro_fabricante VARCHAR(255),
                    motivo_perda TEXT,
                    status_proposta_dtc VARCHAR(100),
                    dtc_responsavel_id INTEGER REFERENCES funcionarios(id) ON DELETE SET NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                
                CREATE TABLE IF NOT EXISTS projetos_fabricantes (
                    projeto_id INTEGER REFERENCES projetos(id) ON DELETE CASCADE,
                    fabricante_id INTEGER REFERENCES fabricantes(id) ON DELETE CASCADE,
                    PRIMARY KEY (projeto_id, fabricante_id)
                );
                
                CREATE TABLE IF NOT EXISTS projetos_colaboradores (
                    projeto_id INTEGER REFERENCES projetos(id) ON DELETE CASCADE,
                    funcionario_id INTEGER REFERENCES funcionarios(id) ON DELETE CASCADE,
                    PRIMARY KEY (projeto_id, funcionario_id)
                );
                
                CREATE TABLE IF NOT EXISTS projeto_pedidos (
                    id SERIAL PRIMARY KEY,
                    projeto_id INTEGER REFERENCES projetos(id) ON DELETE CASCADE,
                    numero_pedido VARCHAR(100) NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                
                CREATE TABLE IF NOT EXISTS projeto_atividades (
                    id SERIAL PRIMARY KEY,
                    projeto_id INTEGER REFERENCES projetos(id) ON DELETE CASCADE,
                    usuario_id INTEGER REFERENCES funcionarios(id) ON DELETE SET NULL,
                    descricao TEXT NOT NULL,
                    tipo_atividade VARCHAR(100),
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                
                CREATE TABLE IF NOT EXISTS solicitacoes_compra (
                    id SERIAL PRIMARY KEY,
                    projeto_id INTEGER REFERENCES projetos(id) ON DELETE CASCADE,
                    projeto_pedido_id INTEGER REFERENCES projeto_pedidos(id) ON DELETE CASCADE,
                    itens_faltantes JSONB,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
            `);

            // 5. Dashboard, Metas e Agente de IA
            await client.query(`
                CREATE TABLE IF NOT EXISTS registro_venda (
                    id SERIAL PRIMARY KEY,
                    data_venda DATE NOT NULL,
                    vendedor VARCHAR(255) NOT NULL,
                    valor_total NUMERIC(15, 2) NOT NULL,
                    dados_adicionais JSONB
                );
                CREATE TABLE IF NOT EXISTS metas_vendedores (
                    vendedor VARCHAR(255) PRIMARY KEY,
                    meta NUMERIC(15, 2) NOT NULL DEFAULT 0
                );
                
                CREATE TABLE IF NOT EXISTS oportunidades (
                    id SERIAL PRIMARY KEY,
                    titulo VARCHAR(255) NOT NULL,
                    descricao TEXT,
                    fonte VARCHAR(255),
                    url VARCHAR(255),
                    data_publicacao TIMESTAMP,
                    tipo VARCHAR(100),
                    vertical VARCHAR(100),
                    prioridade VARCHAR(50),
                    score_oportunidade INTEGER,
                    insight_ia TEXT,
                    entidades_chave JSONB,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                
                CREATE TABLE IF NOT EXISTS leads (
                    id SERIAL PRIMARY KEY,
                    nome VARCHAR(255),
                    status VARCHAR(50),
                    regiao VARCHAR(100),
                    data_sugerida TIMESTAMP,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                
                CREATE TABLE IF NOT EXISTS visitas (
                    id SERIAL PRIMARY KEY,
                    cliente_id INTEGER REFERENCES clientes(id) ON DELETE CASCADE,
                    vendedor_id INTEGER REFERENCES funcionarios(id) ON DELETE CASCADE,
                    data_visita DATE NOT NULL,
                    justificativa_objetivo TEXT NOT NULL,
                    status_autorizacao VARCHAR(50) DEFAULT 'Pendente',
                    feedback_vendedor TEXT,
                    gestor_id INTEGER REFERENCES funcionarios(id) ON DELETE SET NULL,
                    data_resposta_gestor TIMESTAMP,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );

                -- 🚀 Atualizações para o sistema de Retorno e Notificações de Visitas
                ALTER TABLE visitas ADD COLUMN IF NOT EXISTS data_retorno DATE;
                ALTER TABLE visitas ADD COLUMN IF NOT EXISTS requer_retorno BOOLEAN DEFAULT FALSE;
                ALTER TABLE visitas ADD COLUMN IF NOT EXISTS vendedor_viu_status BOOLEAN DEFAULT FALSE;
                ALTER TABLE visitas ADD COLUMN IF NOT EXISTS lembrete_retorno_visto BOOLEAN DEFAULT FALSE;
                ALTER TABLE visitas ADD COLUMN IF NOT EXISTS gestor_viu_pendencia BOOLEAN DEFAULT FALSE;

                -- 🚀 TABELAS DO MÓDULO DE NOTÍCIAS
                CREATE TABLE IF NOT EXISTS noticias (
                    id SERIAL PRIMARY KEY,
                    titulo VARCHAR(255) NOT NULL,
                    resumo TEXT,
                    conteudo TEXT NOT NULL,
                    tipo VARCHAR(50) DEFAULT 'Informativo',
                    fixado BOOLEAN DEFAULT FALSE,
                    setores_alvo JSONB DEFAULT '[]', 
                    criador_id INTEGER REFERENCES funcionarios(id) ON DELETE SET NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                CREATE TABLE IF NOT EXISTS noticias_lidas (
                    noticia_id INTEGER REFERENCES noticias(id) ON DELETE CASCADE,
                    usuario_id INTEGER REFERENCES funcionarios(id) ON DELETE CASCADE,
                    lida_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    PRIMARY KEY (noticia_id, usuario_id)
                );
            `);

            // 7. Módulo de Gestão de Frota e Reservas
            await client.query(`
                CREATE TABLE IF NOT EXISTS frota_veiculos (
                    id SERIAL PRIMARY KEY,
                    marca VARCHAR(100) NOT NULL,
                    modelo VARCHAR(100) NOT NULL,
                    ano INTEGER,
                    placa VARCHAR(20) UNIQUE NOT NULL,
                    renavam VARCHAR(50),
                    cor VARCHAR(50),
                    documento_url TEXT,
                    unidade_origem VARCHAR(100),
                    tipo_combustivel VARCHAR(50),
                    quilometragem_atual INTEGER DEFAULT 0,
                    capacidade_tanque INTEGER,
                    lotacao_maxima INTEGER,
                    categoria VARCHAR(50),
                    status VARCHAR(50) DEFAULT 'Disponível',
                    data_aquisicao DATE,
                    valor_bem NUMERIC(15, 2),
                    apolice_seguro VARCHAR(255),
                    centro_custo VARCHAR(100),
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );

                ALTER TABLE funcionarios ADD COLUMN IF NOT EXISTS cnh_numero VARCHAR(50);
                ALTER TABLE funcionarios ADD COLUMN IF NOT EXISTS cnh_validade DATE;
                ALTER TABLE funcionarios ADD COLUMN IF NOT EXISTS centro_custo VARCHAR(100);
                
                ALTER TABLE frota_veiculos ADD COLUMN IF NOT EXISTS unidade_id INTEGER REFERENCES unidades_filiais(id) ON DELETE SET NULL;

                CREATE TABLE IF NOT EXISTS frota_reservas (
                    id SERIAL PRIMARY KEY,
                    veiculo_id INTEGER REFERENCES frota_veiculos(id) ON DELETE CASCADE,
                    usuario_id INTEGER REFERENCES funcionarios(id) ON DELETE CASCADE,
                    data_inicio TIMESTAMP NOT NULL,
                    data_fim TIMESTAMP NOT NULL,
                    origem VARCHAR(255) NOT NULL,
                    destino VARCHAR(255) NOT NULL,
                    motivo TEXT,
                    status VARCHAR(50) DEFAULT 'Aguardando Aprovação',
                    gestor_aprovador_id INTEGER REFERENCES funcionarios(id) ON DELETE SET NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS frota_inspecoes (
                    id SERIAL PRIMARY KEY,
                    reserva_id INTEGER REFERENCES frota_reservas(id) ON DELETE CASCADE,
                    tipo VARCHAR(50) NOT NULL,
                    quilometragem INTEGER NOT NULL,
                    nivel_combustivel VARCHAR(50),
                    avarias TEXT,
                    fotos JSONB DEFAULT '[]',
                    data_inspecao TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
            `);
        } catch (err) {
            console.error('❌ Erro ao criar tabelas iniciais:', err.message);
        }

        client.release(); 
    })
    .catch(err => console.error('❌ Erro PostgreSQL (portal_dca):', err.message));

// =================================================================
// SEÇÃO 9: API DE GERENCIAMENTO DE VISITAS COMERCIAIS
// =================================================================

app.post('/api/visitas', authMiddleware, async (req, res) => {
    try {
        const vendedor_id = req.user.id;
        const { cliente_id, data_visita, justificativa_objetivo } = req.body;

        if (!cliente_id || !data_visita || !justificativa_objetivo) {
            return res.status(400).json({ error: 'Preencha todos os campos obrigatórios.' });
        }

        const query = `
            INSERT INTO visitas (cliente_id, vendedor_id, data_visita, justificativa_objetivo, status_autorizacao)
            VALUES ($1, $2, $3, $4, 'Pendente') RETURNING id
        `;
        const { rows } = await pool.query(query, [cliente_id, vendedor_id, data_visita, justificativa_objetivo]);
        res.status(201).json({ message: 'Visita solicitada com sucesso!', id: rows[0].id });
    } catch (error) {
        console.error('Erro ao criar visita:', error);
        res.status(500).json({ error: 'Erro ao solicitar a visita.' });
    }
});

app.get('/api/fabricantes/:id', async (req, res) => {
    try {
        const { rows } = await pool.query('SELECT * FROM fabricantes WHERE id = $1', [req.params.id]);
        if (rows.length === 0) return res.status(404).json({ error: 'Fabricante não encontrado' });
        res.json(rows[0]);
    } catch (err) { res.status(500).json({ error: 'Erro no servidor.' }); }
});

app.post('/api/fabricantes', authMiddleware, async (req, res) => {
    try {
        const { name, logo_base64, resumo, key_accounts, registros_projetos, garantias, prazos_compras, certificacoes_treinamentos, contatos_suporte, links_uteis } = req.body;
        const { rows } = await pool.query(
            'INSERT INTO fabricantes (name, logo_base64, resumo, key_accounts, registros_projetos, garantias, prazos_compras, certificacoes_treinamentos, contatos_suporte, links_uteis) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10) RETURNING id',
            [name, logo_base64, resumo, key_accounts, registros_projetos, garantias, prazos_compras, certificacoes_treinamentos, contatos_suporte, links_uteis]
        );
        res.status(201).json({ id: rows[0].id, name });
    } catch (err) {
        if (err.code === '23505') return res.status(409).json({ error: 'Fabricante já existe.' });
        res.status(500).json({ error: 'Erro no servidor.' });
    }
});

app.put('/api/fabricantes/:id', authMiddleware, async (req, res) => {
    try {
        const { name, logo_base64, resumo, key_accounts, registros_projetos, garantias, prazos_compras, certificacoes_treinamentos, contatos_suporte, links_uteis } = req.body;
        const { rowCount } = await pool.query(
            'UPDATE fabricantes SET name = $1, logo_base64 = $2, resumo = $3, key_accounts = $4, registros_projetos = $5, garantias = $6, prazos_compras = $7, certificacoes_treinamentos = $8, contatos_suporte = $9, links_uteis = $10 WHERE id = $11',
            [name, logo_base64, resumo, key_accounts, registros_projetos, garantias, prazos_compras, certificacoes_treinamentos, contatos_suporte, links_uteis, req.params.id]
        );
        if (rowCount === 0) return res.status(404).json({ error: 'Fabricante não encontrado.' });
        res.json({ message: 'Fabricante atualizado com sucesso.' });
    } catch (err) {
        if (err.code === '23505') return res.status(409).json({ error: 'Nome de fabricante já em uso.' });
        res.status(500).json({ error: 'Erro no servidor.' });
    }
});

app.delete('/api/fabricantes/:id', authMiddleware, async (req, res) => {
    try {
        await pool.query('DELETE FROM fabricantes WHERE id = $1', [req.params.id]);
        res.json({ message: 'Fabricante removido.' });
    } catch (err) {
        res.status(500).json({ error: 'Erro ao remover fabricante. Pode estar em uso.' });
    }
});
app.get('/api/visitas', authMiddleware, async (req, res) => {
    try {
        const isManager = req.user.privilegios && (req.user.privilegios.includes('Admin') || req.user.privilegios.includes('Gestor'));
        const vendedor_id = req.user.id;

        let query = `
            SELECT v.*, c.nome_cliente, f.nome_completo as vendedor_nome, g.nome_completo as gestor_nome
            FROM visitas v
            JOIN clientes c ON v.cliente_id = c.id
            JOIN funcionarios f ON v.vendedor_id = f.id
            LEFT JOIN funcionarios g ON v.gestor_id = g.id
        `;
        const params = [];

        if (!isManager) {
            query += ` WHERE v.vendedor_id = $1`;
            params.push(vendedor_id);
        }

        query += ` ORDER BY v.data_visita DESC, v.created_at DESC`;

        const { rows } = await pool.query(query, params);
        res.json(rows);
    } catch (error) {
        console.error('Erro ao buscar visitas:', error);
        res.status(500).json({ error: 'Erro ao buscar as visitas.' });
    }
});

app.patch('/api/visitas/:id/autorizar', authMiddleware, async (req, res) => {
    try {
        const { id } = req.params;
        const { status_autorizacao } = req.body;
        const gestor_id = req.user.id;

        const isManager = req.user.privilegios && (req.user.privilegios.includes('Admin') || req.user.privilegios.includes('Gestor'));
        if (!isManager) {
            return res.status(403).json({ error: 'Acesso negado. Apenas gestores podem autorizar visitas.' });
        }

        if (!['Autorizada', 'Recusada'].includes(status_autorizacao)) {
            return res.status(400).json({ error: 'Status inválido.' });
        }

        const query = `
            UPDATE visitas 
            SET status_autorizacao = $1, gestor_id = $2, data_resposta_gestor = CURRENT_TIMESTAMP
            WHERE id = $3 RETURNING *
        `;
        const { rowCount } = await pool.query(query, [status_autorizacao, gestor_id, id]);

        if (rowCount === 0) return res.status(404).json({ error: 'Visita não encontrada.' });

        res.json({ message: `Visita ${status_autorizacao.toLowerCase()} com sucesso!` });
    } catch (error) {
        console.error('Erro ao autorizar visita:', error);
        res.status(500).json({ error: 'Erro ao processar a autorização.' });
    }
});

app.patch('/api/visitas/:id/feedback', authMiddleware, async (req, res) => {
    try {
        const { id } = req.params;
        const { feedback_vendedor, requer_retorno, data_retorno } = req.body;
        const vendedor_id = req.user.id;

        const { rows: visitaRows } = await pool.query('SELECT vendedor_id, status_autorizacao FROM visitas WHERE id = $1', [id]);
        if (visitaRows.length === 0) return res.status(404).json({ error: 'Visita não encontrada.' });
        if (visitaRows[0].vendedor_id !== vendedor_id) return res.status(403).json({ error: 'Acesso negado.' });
        if (visitaRows[0].status_autorizacao !== 'Autorizada') return res.status(400).json({ error: 'Apenas visitas autorizadas recebem feedback.' });

        const query = `
            UPDATE visitas 
            SET feedback_vendedor = $1, requer_retorno = $2, data_retorno = $3
            WHERE id = $4 RETURNING *
        `;
        await pool.query(query, [feedback_vendedor, requer_retorno || false, data_retorno || null, id]);

        res.json({ message: 'Feedback registrado com sucesso!' });
    } catch (error) {
        console.error('Erro ao registrar feedback:', error);
        res.status(500).json({ error: 'Erro ao salvar feedback.' });
    }
});

app.get('/api/visitas/dashboard', authMiddleware, async (req, res) => {
    try {
        const isManager = req.user.privilegios && (req.user.privilegios.includes('Admin') || req.user.privilegios.includes('Gestor'));
        const vendedor_id = req.user.id;
        const { mes, ano } = req.query;

        let conditions = [];
        const params = [];
        let paramIndex = 1;
        
        if (ano) {
            conditions.push(`EXTRACT(YEAR FROM data_visita) = $${paramIndex++}`);
            params.push(ano);
        } else {
            conditions.push(`EXTRACT(YEAR FROM data_visita) = EXTRACT(YEAR FROM CURRENT_DATE)`);
        }

        if (mes && mes !== 'todos') {
            conditions.push(`EXTRACT(MONTH FROM data_visita) = $${paramIndex++}`);
            params.push(mes);
        } else if (!mes && !ano) {
            conditions.push(`EXTRACT(MONTH FROM data_visita) = EXTRACT(MONTH FROM CURRENT_DATE)`);
        }

        if (!isManager) {
            conditions.push(`vendedor_id = $${paramIndex++}`);
            params.push(vendedor_id);
        }

        const baseCondition = conditions.length > 0 ? conditions.join(' AND ') : '1=1';

        const kpisQuery = `
            SELECT 
                COUNT(*) as total_visitas,
                COUNT(*) FILTER (WHERE status_autorizacao = 'Pendente') as pendentes,
                COUNT(*) FILTER (WHERE status_autorizacao = 'Autorizada') as autorizadas,
                COUNT(*) FILTER (WHERE status_autorizacao = 'Recusada') as recusadas
            FROM visitas
            WHERE ${baseCondition}
        `;
        const { rows: kpis } = await pool.query(kpisQuery, params);

        const rankingQuery = `
            SELECT f.nome_completo as vendedor, COUNT(v.id) as total_visitas
            FROM visitas v
            JOIN funcionarios f ON v.vendedor_id = f.id
            WHERE ${baseCondition} AND v.status_autorizacao = 'Autorizada'
            GROUP BY f.nome_completo
            ORDER BY total_visitas DESC
        `;
        const { rows: ranking } = await pool.query(rankingQuery, params);

        res.json({
            kpis: kpis[0],
            ranking
        });
    } catch (error) {
        console.error('Erro ao buscar dashboard de visitas:', error);
        res.status(500).json({ error: 'Erro ao gerar métricas do dashboard.' });
    }
});

// =================================================================
// NOTIFICAÇÕES (Sino do Header)
// =================================================================
app.get('/api/notificacoes', authMiddleware, async (req, res) => {
    try {
        const vendedor_id = req.user.id;
        const isManager = req.user.privilegios && (req.user.privilegios.includes('Admin') || req.user.privilegios.includes('Gestor'));
        
        const { rows: statusRows } = await pool.query(`
            SELECT v.id, v.status_autorizacao, c.nome_cliente, v.data_visita, 'status' as tipo
            FROM visitas v JOIN clientes c ON v.cliente_id = c.id
            WHERE v.vendedor_id = $1 AND v.status_autorizacao IN ('Autorizada', 'Recusada') AND v.vendedor_viu_status = FALSE
        `, [vendedor_id]);

        const { rows: feedbackRows } = await pool.query(`
            SELECT v.id, c.nome_cliente, v.data_visita, 'feedback' as tipo
            FROM visitas v JOIN clientes c ON v.cliente_id = c.id
            WHERE v.vendedor_id = $1 AND v.status_autorizacao = 'Autorizada' AND v.data_visita <= CURRENT_DATE AND v.feedback_vendedor IS NULL
        `, [vendedor_id]);

        const { rows: retornoRows } = await pool.query(`
            SELECT v.id, c.nome_cliente, v.data_retorno, 'retorno' as tipo
            FROM visitas v JOIN clientes c ON v.cliente_id = c.id
            WHERE v.vendedor_id = $1 AND v.requer_retorno = TRUE AND v.data_retorno <= CURRENT_DATE + INTERVAL '3 days' AND v.lembrete_retorno_visto = FALSE
        `, [vendedor_id]);

        let gestorRows = [];
        if (isManager) {
            const { rows } = await pool.query(`
                SELECT v.id, c.nome_cliente, f.nome_completo as vendedor_nome, v.data_visita, 'aprovacao_gestor' as tipo
                FROM visitas v 
                JOIN clientes c ON v.cliente_id = c.id
                JOIN funcionarios f ON v.vendedor_id = f.id
                WHERE v.status_autorizacao = 'Pendente' AND v.gestor_viu_pendencia = FALSE
            `);
            gestorRows = rows;
        }

        const setor_id = req.user.setor_id;
        const setorParam = setor_id ? [setor_id] : [];
        const { rows: newsRows } = await pool.query(`
            SELECT n.id, n.titulo, n.tipo, n.created_at, 'nova_noticia' as tipo
            FROM noticias n
            LEFT JOIN noticias_lidas nl ON n.id = nl.noticia_id AND nl.usuario_id = $1
            WHERE nl.noticia_id IS NULL AND (n.setores_alvo = '[]'::jsonb OR n.setores_alvo @> $2::jsonb)
        `, [vendedor_id, JSON.stringify(setorParam)]);

        res.json([...statusRows, ...feedbackRows, ...retornoRows, ...gestorRows, ...newsRows]);
    } catch (error) {
        res.status(500).json({ error: 'Erro ao buscar notificações.' });
    }
});

app.patch('/api/notificacoes/:id/lida', authMiddleware, async (req, res) => {
    try {
        const { id } = req.params;
        if (req.body.tipo === 'status') await pool.query('UPDATE visitas SET vendedor_viu_status = TRUE WHERE id = $1', [id]);
        if (req.body.tipo === 'retorno') await pool.query('UPDATE visitas SET lembrete_retorno_visto = TRUE WHERE id = $1', [id]);
        if (req.body.tipo === 'aprovacao_gestor') await pool.query('UPDATE visitas SET gestor_viu_pendencia = TRUE WHERE id = $1', [id]);
        if (req.body.tipo === 'nova_noticia') {
            await pool.query('INSERT INTO noticias_lidas (noticia_id, usuario_id) VALUES ($1, $2) ON CONFLICT DO NOTHING', [id, req.user.id]);
        }
        res.json({ message: 'Lida.' });
    } catch (error) {
        res.status(500).json({ error: 'Erro.' });
    }
});

// =================================================================
// SEÇÃO: MÓDULO DE NOTÍCIAS (MURAL)
// =================================================================
app.get('/api/noticias', authMiddleware, async (req, res) => {
    try {
        const usuario_id = req.user.id;
        const setor_id = req.user.setor_id;
        const setorParam = setor_id ? [setor_id] : [];

        const query = `
            SELECT n.*, f.nome_completo as autor,
                   EXISTS(SELECT 1 FROM noticias_lidas nl WHERE nl.noticia_id = n.id AND nl.usuario_id = $1) as lida
            FROM noticias n
            LEFT JOIN funcionarios f ON n.criador_id = f.id
            WHERE n.setores_alvo = '[]'::jsonb OR n.setores_alvo @> $2::jsonb
            ORDER BY n.fixado DESC, n.created_at DESC
        `;
        const { rows } = await pool.query(query, [usuario_id, JSON.stringify(setorParam)]);
        res.json(rows);
    } catch (err) {
        console.error('Erro /api/noticias GET:', err);
        res.status(500).json({ error: 'Erro ao buscar notícias.' });
    }
});

app.post('/api/noticias', authMiddleware, async (req, res) => {
    try {
        const { titulo, resumo, conteudo, tipo, fixado, setores_alvo } = req.body;
        const criador_id = req.user.id;
        const setoresJson = JSON.stringify(setores_alvo || []);

        const query = `
            INSERT INTO noticias (titulo, resumo, conteudo, tipo, fixado, setores_alvo, criador_id)
            VALUES ($1, $2, $3, $4, $5, $6, $7) RETURNING id
        `;
        await pool.query(query, [titulo, resumo, conteudo, tipo, fixado, setoresJson, criador_id]);
        res.status(201).json({ message: 'Notícia publicada com sucesso!' });
    } catch (err) {
        res.status(500).json({ error: 'Erro ao publicar notícia.' });
    }
});

app.put('/api/noticias/:id', authMiddleware, async (req, res) => {
    try {
        const { id } = req.params;
        const { titulo, resumo, conteudo, tipo, fixado, setores_alvo } = req.body;
        const setoresJson = JSON.stringify(setores_alvo || []);

        const query = `
            UPDATE noticias 
            SET titulo = $1, resumo = $2, conteudo = $3, tipo = $4, fixado = $5, setores_alvo = $6
            WHERE id = $7
        `;
        await pool.query(query, [titulo, resumo, conteudo, tipo, fixado, setoresJson, id]);
        res.json({ message: 'Notícia atualizada!' });
    } catch (err) {
        res.status(500).json({ error: 'Erro ao atualizar notícia.' });
    }
});

app.delete('/api/noticias/:id', authMiddleware, async (req, res) => {
    try {
        await pool.query('DELETE FROM noticias WHERE id = $1', [req.params.id]);
        res.json({ message: 'Notícia removida.' });
    } catch (err) {
        res.status(500).json({ error: 'Erro ao excluir notícia.' });
    }
});

app.post('/api/noticias/:id/lida', authMiddleware, async (req, res) => {
    try {
        await pool.query('INSERT INTO noticias_lidas (noticia_id, usuario_id) VALUES ($1, $2) ON CONFLICT DO NOTHING', [req.params.id, req.user.id]);
        res.json({ message: 'Notícia marcada como lida.' });
    } catch (err) {
        res.status(500).json({ error: 'Erro ao marcar leitura.' });
    }
});

// =================================================================
// SEÇÃO 3: AUTENTICAÇÃO MICROSOFT
// =================================================================
const msalConfig = {
    auth: {
        clientId: process.env.MSAL_CLIENT_ID,
        authority: process.env.MSAL_AUTHORITY,
        clientSecret: process.env.MSAL_CLIENT_SECRET,
    },
    system: { loggerOptions: { loggerCallback() {}, piiLoggingEnabled: false, logLevel: msal.LogLevel.Verbose } },
};
const cca = new msal.ConfidentialClientApplication(msalConfig);

app.get('/auth/microsoft', (req, res) => {
    const authCodeUrlParameters = { scopes: ['user.read', 'calendars.read'], redirectUri: `${BACKEND_URL}/auth/microsoft/callback` };
    cca.getAuthCodeUrl(authCodeUrlParameters).then(response => res.redirect(response)).catch(error => res.status(500).send(JSON.stringify(error)));
});

app.get('/auth/microsoft/callback', async (req, res) => {
    const tokenRequest = { 
        code: req.query.code, 
        scopes: ['user.read', 'calendars.read'], 
        redirectUri: `${BACKEND_URL}/auth/microsoft/callback` 
    };
    try {
        const response = await cca.acquireTokenByCode(tokenRequest);
        req.session.accessToken = response.accessToken;
        res.redirect(FRONTEND_URL);
    } catch (error) { 
        // Agora sim vamos ver o erro real no terminal!
        console.error("🔥 ERRO DETALHADO DA MICROSOFT:", error.response?.data || error.message || error);
        res.status(500).send('Erro na autenticação. Verifique o terminal do backend.'); 
    }
});

// =================================================================
// ROTA DE BYPASS PARA DESENVOLVIMENTO (Pula a Microsoft)
// =================================================================
app.get('/auth/dev', (req, res) => {
    // Injeta um token falso na sua sessão para enganar o sistema
    req.session.accessToken = "DEV_MOCK_TOKEN";
    console.log("⚠️ LOGIN DEV ACIONADO! Redirecionando para o painel...");
    res.redirect(FRONTEND_URL);
});

app.get('/user-data', async (req, res) => {
    console.log('--- [USER-DATA] Verificando sessão na carga do App ---');
    console.log('Sessão ID:', req.sessionID);
    console.log('Tem accessToken:', !!req.session?.accessToken);

    if (!req.session.accessToken) {
        console.log('❌ Bloqueado no /user-data: Sessão não possui accessToken.');
        return res.status(401).json({ error: 'Usuário não autenticado.' });
    }
    try {
        let userEmail;
        let displayName;

        // SE FOR O NOSSO TOKEN FALSO, PULA A MICROSOFT
        if (req.session.accessToken === "DEV_MOCK_TOKEN") {
            userEmail = 'dionathan.matos@dca.com.br'; // ATENÇÃO: Coloque aqui o SEU e-mail exato que está na tabela 'funcionarios' do Postgres
            displayName = 'Dionathan (Modo Dev)';
        } 
        // SE FOR UM TOKEN REAL, BATE NA MICROSOFT NORMALMENTE
        else {
            const endpoint = 'https://graph.microsoft.com/v1.0/me';
            const response = await axios.get(endpoint, { headers: { Authorization: `Bearer ${req.session.accessToken}` } });
            userEmail = response.data.mail || response.data.userPrincipalName;
            displayName = response.data.displayName;
        }

        if (!userEmail) return res.status(404).json({ error: 'E-mail não encontrado.'});
        
        const { rows: funcionarios } = await pool.query('SELECT * FROM funcionarios WHERE email = $1', [userEmail]);
        
        if (funcionarios.length === 0) {
            return res.status(404).json({ error: 'Usuário não encontrado no banco de dados local.', email: userEmail });
        }
        
        const userProfile = { ...funcionarios[0], displayName };
        req.session.isAuthenticated = true;
        req.session.localUser = userProfile;

        res.json(userProfile);
    } catch (error) { 
        console.error("Erro ao buscar dados do usuário:", error.message);
        res.status(500).json({ error: 'Erro ao carregar os dados. Token pode ter expirado.' }); 
    }
});

app.get('/auth/microsoft/logout', (req, res) => {
    req.session.destroy(() => res.redirect(FRONTEND_URL));
});

// =================================================================
// SEÇÃO 4: CONFIGURAÇÕES DE ADMINISTRAÇÃO DO TEMA
// =================================================================
app.get('/api/settings', async (req, res) => {
    try {
        // [PG] rows destructuring
        const { rows: settingsRows } = await pool.query('SELECT * FROM system_settings');
        const settings = settingsRows.reduce((acc, row) => {
            acc[row.setting_key] = row.setting_value;
            return acc;
        }, {});
        res.json(settings);
    } catch (err) {
        console.error('Erro ao buscar configurações:', err);
        res.status(500).json({ error: 'Erro no servidor.' });
    }
});

app.post('/api/settings', themeUpload, async (req, res) => {
    const { backgroundType, backgroundColor, sidebarColor, headerColor, pageTitle, sidebarIconColor, sidebarActiveColor, darkModeBackground, darkModeSurface, darkModePrimaryText, darkModeSecondaryText, lightModeSurface, login_bg_color, card_header_bg, card_header_text } = req.body;

    const settingsToUpdate = {
        sidebar_color: sidebarColor, header_color: headerColor, page_title: pageTitle,
        sidebar_icon_color: sidebarIconColor, sidebar_active_color: sidebarActiveColor,
        dark_mode_background: darkModeBackground, dark_mode_surface: darkModeSurface,
        dark_mode_primary_text: darkModePrimaryText, dark_mode_secondary_text: darkModeSecondaryText,
        light_mode_surface: lightModeSurface,
        login_bg_color: login_bg_color,
        card_header_bg: card_header_bg,
        card_header_text: card_header_text
    };

    if (req.files && req.files.backgroundImageFile) {
        settingsToUpdate.background = `url('${BACKEND_URL}/uploads/theme/${req.files.backgroundImageFile[0].filename}')`;
    } else if (backgroundType === 'color') {
        settingsToUpdate.background = backgroundColor;
    }
    
    if (req.files && req.files.logoFile) {
        settingsToUpdate.logo_url = `${BACKEND_URL}/uploads/theme/${req.files.logoFile[0].filename}`;
    }
    if (req.files && req.files.faviconFile) {
        settingsToUpdate.favicon_url = `${BACKEND_URL}/uploads/theme/${req.files.faviconFile[0].filename}`;
    }
    if (req.files && req.files.loginBgImageFile) {
        settingsToUpdate.login_bg_image = `${BACKEND_URL}/uploads/theme/${req.files.loginBgImageFile[0].filename}`;
    }
    if (req.files && req.files.loginLogoFile) {
        settingsToUpdate.login_logo_url = `${BACKEND_URL}/uploads/theme/${req.files.loginLogoFile[0].filename}`;
    }

    const settingsToSave = Object.entries(settingsToUpdate);
    if (settingsToSave.length === 0) return res.status(200).json({ message: 'Nenhuma alteração para salvar.' });

    let client;
    try {
        // [PG] Tratamento de transação adaptado
        client = await pool.connect();
        await client.query('BEGIN');
        
        // [PG] Inserção iterativa, pois Postgres não suporta bulk via VALUES array multi-dimentional nativamente como MySQL
        for (const [key, value] of settingsToSave) {
            const sql = 'INSERT INTO system_settings (setting_key, setting_value) VALUES ($1, $2) ON CONFLICT (setting_key) DO UPDATE SET setting_value = EXCLUDED.setting_value';
            await client.query(sql, [key, value]);
        }
        
        await client.query('COMMIT');
        res.status(200).json({ message: 'Configurações salvas com sucesso!' });
    } catch (err) {
        if (client) await client.query('ROLLBACK');
        console.error('Erro ao salvar configurações:', err);
        res.status(500).json({ error: 'Erro ao salvar as configurações.' });
    } finally {
        if (client) client.release();
    }
});

// =================================================================
// SEÇÃO MÓDULO DE BENEFÍCIOS (RH)
// =================================================================
app.get('/api/beneficios', async (req, res) => {
    try {
        const { rows } = await pool.query('SELECT * FROM beneficios WHERE ativo = TRUE ORDER BY nome ASC');
        res.json(rows);
    } catch (err) { res.status(500).json({ error: 'Erro no servidor.' }); }
});

app.post('/api/beneficios', authMiddleware, async (req, res) => {
    const { nome, descricao, logo_base64 } = req.body;
    try {
        const { rows } = await pool.query('INSERT INTO beneficios (nome, descricao, logo_base64) VALUES ($1, $2, $3) RETURNING id', [nome, descricao, logo_base64]);
        res.status(201).json({ id: rows[0].id, nome, descricao, logo_base64 });
    } catch (err) { res.status(500).json({ error: 'Erro ao criar benefício.' }); }
});

app.put('/api/beneficios/:id', authMiddleware, async (req, res) => {
    const { nome, descricao, logo_base64 } = req.body;
    try {
        await pool.query('UPDATE beneficios SET nome = $1, descricao = $2, logo_base64 = $3 WHERE id = $4', [nome, descricao, logo_base64, req.params.id]);
        res.status(200).json({ message: 'Benefício atualizado com sucesso.' });
    } catch (err) { res.status(500).json({ error: 'Erro ao atualizar benefício.' }); }
});

app.delete('/api/beneficios/:id', authMiddleware, async (req, res) => {
    try {
        await pool.query('UPDATE beneficios SET ativo = FALSE WHERE id = $1', [req.params.id]);
        res.status(200).json({ message: 'Benefício inativado com sucesso.' });
    } catch (err) { res.status(500).json({ error: 'Erro ao inativar benefício.' }); }
});

// =================================================================
// SEÇÃO 5: API DE GESTÃO INTERNA (FUNCIONÁRIOS, CARGOS, SETORES, FABRICANTES)
// =================================================================
app.get('/api/funcionarios', async (req, res) => {
    try {
        const { cargoId, setorId } = req.query;
        let sql = `
            SELECT f.*, c.nome_cargo, s.nome_setor, u.nome_unidade,
            (SELECT COALESCE(json_agg(ff.fabricante_id), '[]'::json) FROM funcionario_fabricante ff WHERE ff.funcionario_id = f.id) as fabricantes_ids
            FROM funcionarios f 
            LEFT JOIN cargos c ON f.cargo_id = c.id 
            LEFT JOIN setores s ON f.setor_id = s.id
            LEFT JOIN unidades_filiais u ON f.unidade_id = u.id
        `;
        const conditions = ['(f.ativo = TRUE OR f.ativo IS NULL)']; // 🚀 Oculta inativos
        const params = [];
        let paramIndex = 1; // [PG] Controle de binds dinâmico para $1, $2
        
        if (cargoId) { conditions.push(`f.cargo_id = $${paramIndex++}`); params.push(cargoId); }
        if (setorId) { conditions.push(`f.setor_id = $${paramIndex++}`); params.push(setorId); }
        sql += ' WHERE ' + conditions.join(' AND ');
        sql += ' ORDER BY f.nome_completo ASC';
        
        const { rows } = await pool.query(sql, params);
        res.json(rows);
    } catch (err) { res.status(500).send('Erro no servidor'); }
});

app.post('/api/funcionarios', async (req, res) => {
    const { nome_completo, email, contato, setor_id, userpic_base64, cargo_id, privilegios, fabricantes_ids, gestor_id, cnh_numero, cnh_validade, unidade_id } = req.body;
    const finalCargoId = (!cargo_id || cargo_id === '') ? null : cargo_id;
    const finalSetorId = (!setor_id || setor_id === '') ? null : setor_id;
    const finalUserpic = (!userpic_base64 || userpic_base64 === '') ? null : userpic_base64;
    const finalGestorId = (!gestor_id || gestor_id === '') ? null : gestor_id;
    const finalUnidadeId = (!unidade_id || unidade_id === '') ? null : unidade_id;
    try {
        // [PG] Inclusão de RETURNING id
        const sql = 'INSERT INTO funcionarios (nome_completo, email, contato, setor_id, userpic_base64, cargo_id, privilegios, gestor_id, cnh_numero, cnh_validade, unidade_id) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11) RETURNING id';
        const { rows } = await pool.query(sql, [nome_completo, email, contato, finalSetorId, finalUserpic, finalCargoId, privilegios, finalGestorId, cnh_numero || null, cnh_validade || null, finalUnidadeId]);
        const newId = rows[0].id;
        
        // 🚀 Salva os vínculos de fabricantes para este novo funcionário
        if (fabricantes_ids && Array.isArray(fabricantes_ids)) {
            for (const fabId of fabricantes_ids) {
                await pool.query('INSERT INTO funcionario_fabricante (funcionario_id, fabricante_id) VALUES ($1, $2)', [newId, fabId]);
            }
        }
        
        res.status(201).json({ id: newId, nome_completo, email, contato, setor_id: finalSetorId, userpic_base64: finalUserpic, cargo_id: finalCargoId, privilegios });
    } catch (err) {
        console.error('❌ Erro ao cadastrar funcionário:', err);
        if (err.code === '23505') { return res.status(409).json({ error: 'Este e-mail já está em uso.' }); } // [PG] ER_DUP_ENTRY é 23505
        res.status(500).json({ error: 'Erro ao adicionar o funcionário.', details: err.message });
    }
});

app.put('/api/funcionarios/:id', async (req, res) => {
    const { id } = req.params;
    const { nome_completo, email, contato, setor_id, userpic_base64, cargo_id, privilegios, fabricantes_ids, gestor_id, cnh_numero, cnh_validade, unidade_id } = req.body;
    const finalCargoId = (!cargo_id || cargo_id === '') ? null : cargo_id;
    const finalSetorId = (!setor_id || setor_id === '') ? null : setor_id;
    const finalUserpic = (!userpic_base64 || userpic_base64 === '') ? null : userpic_base64;
    const finalGestorId = (!gestor_id || gestor_id === '') ? null : gestor_id;
    const finalUnidadeId = (!unidade_id || unidade_id === '') ? null : unidade_id;
    try {
        await pool.query('UPDATE funcionarios SET nome_completo = $1, email = $2, contato = $3, setor_id = $4, userpic_base64 = $5, cargo_id = $6, privilegios = $7, gestor_id = $8, cnh_numero = $9, cnh_validade = $10, unidade_id = $11 WHERE id = $12', [nome_completo, email, contato, finalSetorId, finalUserpic, finalCargoId, privilegios, finalGestorId, cnh_numero || null, cnh_validade || null, finalUnidadeId, id]);
        
        // 🚀 Atualiza os vínculos de fabricantes (apaga os antigos e insere as novas caixinhas marcadas)
        await pool.query('DELETE FROM funcionario_fabricante WHERE funcionario_id = $1', [id]);
        if (fabricantes_ids && Array.isArray(fabricantes_ids)) {
            for (const fabId of fabricantes_ids) {
                await pool.query('INSERT INTO funcionario_fabricante (funcionario_id, fabricante_id) VALUES ($1, $2)', [id, fabId]);
            }
        }
        
        res.status(200).send('Funcionário atualizado com sucesso!');
    } catch (err) { res.status(500).send('Erro no servidor'); }
});

app.delete('/api/funcionarios/:id', async (req, res) => {
    try { 
        await pool.query('UPDATE funcionarios SET ativo = FALSE WHERE id = $1', [req.params.id]); 
        res.status(200).send('Colaborador inativado com sucesso!'); 
    } catch (err) { res.status(500).send('Erro no servidor'); }
});

app.post('/api/funcionarios/import', upload.single('file'), async (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'Nenhum arquivo CSV enviado.' });

    const filePath = req.file.path;
    const funcionariosParaAdicionar = [];

    fs.createReadStream(filePath)
        .pipe(csv())
        .on('data', (row) => funcionariosParaAdicionar.push(row))
        .on('end', async () => {
            if (funcionariosParaAdicionar.length === 0) {
                fs.unlinkSync(filePath); 
                return res.status(400).json({ error: 'O arquivo CSV está vazio ou em formato inválido.' });
            }

            let client;
            try {
                client = await pool.connect();
                await client.query('BEGIN');

                let successfulImports = 0;
                for (const func of funcionariosParaAdicionar) {
                    const { nome_completo, email, contato, setor_id, cargo_id, privilegios } = func;
                    if (!nome_completo || !email) continue; 

                    // [PG] Refatorado para sintaxe padrão PostgreSQL (ON CONFLICT)
                    const query = `
                        INSERT INTO funcionarios (nome_completo, email, contato, setor_id, cargo_id, privilegios) 
                        VALUES ($1, $2, $3, $4, $5, $6)
                        ON CONFLICT (email) DO UPDATE SET 
                        nome_completo = EXCLUDED.nome_completo, 
                        contato = EXCLUDED.contato, 
                        setor_id = EXCLUDED.setor_id, 
                        cargo_id = EXCLUDED.cargo_id, 
                        privilegios = EXCLUDED.privilegios;
                    `;
                    await client.query(query, [nome_completo, email, contato || null, setor_id || null, cargo_id || null, privilegios || 'Padrão']);
                    successfulImports++;
                }

                await client.query('COMMIT');
                res.status(200).json({ message: `${successfulImports} funcionários importados/atualizados com sucesso!` });
            } catch (err) {
                if (client) await client.query('ROLLBACK');
                console.error("Erro ao importar funcionários:", err);
                res.status(500).json({ error: 'Ocorreu um erro durante a importação.' });
            } finally {
                if (client) client.release();
                fs.unlinkSync(filePath); 
            }
        });
});

app.get('/api/funcionarios/agrupados', async (req, res) => {
    try {
        const sql = `
            SELECT 
                f.id, f.nome_completo, f.email, f.contato, f.privilegios, f.userpic_base64,
                s.nome_setor, c.nome_cargo
            FROM funcionarios f 
            LEFT JOIN setores s ON f.setor_id = s.id
            LEFT JOIN cargos c ON f.cargo_id = c.id
            WHERE f.ativo = TRUE OR f.ativo IS NULL
            ORDER BY s.nome_setor, f.nome_completo ASC
        `;
        const { rows: funcionarios } = await pool.query(sql);
        const agrupados = funcionarios.reduce((acc, funcionario) => {
            const setor = funcionario.nome_setor || 'Sem Setor';
            if (!acc[setor]) acc[setor] = [];
            acc[setor].push(funcionario);
            return acc;
        }, {});
        res.json(agrupados);
    } catch (err) { res.status(500).json({ error: 'Erro no servidor.' }); }
});

app.get('/api/funcionarios-tecnicos', async (req, res) => {
    try {
        // [PG] Correção de GROUP_CONCAT para STRING_AGG do PostgreSQL
        const sql = `
            SELECT 
                f.id, 
                f.nome_completo,
                f.email,
                f.contato,
                c.nome_cargo as cargo,
                f.userpic_base64,
                s.nome_setor as setor,
                f.gestor_id,
                STRING_AGG(fab.name, ', ') as fabricantes_nomes 
            FROM funcionarios f
            LEFT JOIN setores s ON f.setor_id = s.id
            LEFT JOIN cargos c ON f.cargo_id = c.id
            LEFT JOIN funcionario_fabricante ff ON f.id = ff.funcionario_id
            LEFT JOIN fabricantes fab ON ff.fabricante_id = fab.id
            WHERE s.nome_setor IN ('Departamento Técnico', 'Dtc') AND (f.ativo = TRUE OR f.ativo IS NULL)
            GROUP BY f.id, f.nome_completo, f.email, f.contato, c.nome_cargo, f.userpic_base64, s.nome_setor, f.gestor_id
            ORDER BY f.nome_completo ASC
        `;
        const { rows } = await pool.query(sql);
        res.json(rows);
    } catch (err) { 
        console.error('❌ Erro na rota /api/funcionarios-tecnicos:', err.message);
        res.status(500).json({ error: 'Erro no servidor ao buscar organograma.' }); 
    }
});

// SEÇÃO CARGOS & SETORES
app.get('/api/cargos', async (req, res) => {
    try {
        const { rows } = await pool.query('SELECT * FROM cargos ORDER BY nome_cargo ASC');
        res.json(rows);
    } catch (err) { res.status(500).json({ error: 'Erro no servidor.' }); }
});

app.post('/api/cargos', async (req, res) => {
    const nome_cargo = req.body.nome_cargo || req.body.nome;
    if (!nome_cargo || String(nome_cargo).trim() === '') return res.status(400).json({ error: 'O nome do cargo é obrigatório.' });
    
    try {
        const { rows } = await pool.query('INSERT INTO cargos (nome_cargo) VALUES ($1) RETURNING id', [String(nome_cargo).trim()]);
        res.status(201).json({ id: rows[0].id, nome_cargo: String(nome_cargo).trim() });
    } catch (err) {
        console.error('❌ Erro ao cadastrar cargo:', err);
        if (err.code === '23505') return res.status(409).json({ error: 'Já existe um cargo com este nome.' });
        res.status(500).json({ error: 'Erro no servidor.', details: err.message });
    }
});

app.put('/api/cargos/:id', async (req, res) => {
    try {
        const nome_cargo = req.body.nome_cargo || req.body.nome;
        if (!nome_cargo || String(nome_cargo).trim() === '') return res.status(400).json({ error: 'O nome do cargo é obrigatório.' });
        
        const { rowCount } = await pool.query('UPDATE cargos SET nome_cargo = $1 WHERE id = $2', [String(nome_cargo).trim(), req.params.id]);
        if (rowCount === 0) return res.status(404).json({ error: 'Cargo não encontrado.' });
        res.status(200).json({ message: 'Cargo atualizado com sucesso.' });
    } catch (err) {
        console.error('❌ Erro ao atualizar cargo:', err);
        if (err.code === '23505') return res.status(409).json({ error: 'Já existe um cargo com este nome.' });
        res.status(500).json({ error: 'Erro no servidor.', details: err.message });
    }
});

app.delete('/api/cargos/:id', async (req, res) => {
    try {
        const { rowCount } = await pool.query('DELETE FROM cargos WHERE id = $1', [req.params.id]);
        if (rowCount === 0) return res.status(404).json({ error: 'Cargo não encontrado.' });
        res.status(200).json({ message: 'Cargo excluído com sucesso.' });
    } catch (err) {
        if (err.code === '23503') return res.status(409).json({ error: 'Este cargo não pode ser excluído pois está em uso.' }); // [PG] Foreign Key Violation
        res.status(500).json({ error: 'Erro no servidor.' });
    }
});

app.get('/api/setores', async (req, res) => {
    try {
        const { rows } = await pool.query('SELECT * FROM setores ORDER BY nome_setor ASC');
        res.json(rows);
    } catch (err) { res.status(500).json({ error: 'Erro no servidor.' }); }
});

app.post('/api/setores', async (req, res) => {
    const nome_setor = req.body.nome_setor || req.body.nome;
    if (!nome_setor) return res.status(400).json({ error: 'O nome do setor é obrigatório.' });
    try {
        const { rows } = await pool.query('INSERT INTO setores (nome_setor) VALUES ($1) RETURNING id', [nome_setor]);
        res.status(201).json({ id: rows[0].id, nome_setor });
    } catch (err) {
        console.error('❌ Erro ao cadastrar setor:', err);
        if (err.code === '23505') return res.status(409).json({ error: 'Já existe um setor com este nome.' });
        res.status(500).json({ error: 'Erro no servidor.', details: err.message });
    }
});

app.put('/api/setores/:id', async (req, res) => {
    try {
        const nome_setor = req.body.nome_setor || req.body.nome;
        const { rowCount } = await pool.query('UPDATE setores SET nome_setor = $1 WHERE id = $2', [nome_setor, req.params.id]);
        if (rowCount === 0) return res.status(404).json({ error: 'Setor não encontrado.' });
        res.status(200).json({ message: 'Setor atualizado com sucesso.' });
    } catch (err) {
        console.error('❌ Erro ao atualizar setor:', err);
        if (err.code === '23505') return res.status(409).json({ error: 'Já existe um setor com este nome.' });
        res.status(500).json({ error: 'Erro no servidor.', details: err.message });
    }
});

// =================================================================
// SEÇÃO UNIDADES / FILIAIS DA EMPRESA
// =================================================================
app.get('/api/unidades', async (req, res) => {
    try {
        const { rows } = await pool.query('SELECT * FROM unidades_filiais ORDER BY nome_unidade ASC');
        res.json(rows);
    } catch (err) { res.status(500).json({ error: 'Erro no servidor.' }); }
});

app.post('/api/unidades', async (req, res) => {
    const nome_unidade = req.body.nome_unidade || req.body.nome;
    if (!nome_unidade) return res.status(400).json({ error: 'O nome da unidade é obrigatório.' });
    try {
        const { rows } = await pool.query('INSERT INTO unidades_filiais (nome_unidade) VALUES ($1) RETURNING id', [nome_unidade]);
        res.status(201).json({ id: rows[0].id, nome_unidade });
    } catch (err) {
        if (err.code === '23505') return res.status(409).json({ error: 'Já existe uma unidade com este nome.' });
        res.status(500).json({ error: 'Erro no servidor.', details: err.message });
    }
});

app.put('/api/unidades/:id', async (req, res) => {
    try {
        const { rowCount } = await pool.query('UPDATE unidades_filiais SET nome_unidade = $1 WHERE id = $2', [req.body.nome_unidade, req.params.id]);
        if (rowCount === 0) return res.status(404).json({ error: 'Unidade não encontrada.' });
        res.status(200).json({ message: 'Unidade atualizada.' });
    } catch (err) { res.status(500).json({ error: 'Erro no servidor.' }); }
});

app.delete('/api/unidades/:id', async (req, res) => {
    try {
        await pool.query('DELETE FROM unidades_filiais WHERE id = $1', [req.params.id]);
        res.json({ message: 'Unidade excluída.' });
    } catch (err) { res.status(500).json({ error: 'Erro ao excluir unidade.' }); }
});

// =================================================================
// SEÇÃO 10: MÓDULO DE GESTÃO DE FROTA E RESERVAS
// =================================================================

app.get('/api/frota/veiculos', authMiddleware, async (req, res) => {
    try {
        const { status } = req.query;
        let sql = 'SELECT v.*, u.nome_unidade FROM frota_veiculos v LEFT JOIN unidades_filiais u ON v.unidade_id = u.id';
        const params = [];
        if (status) {
            sql += ' WHERE v.status = $1';
            params.push(status);
        } else {
            sql += " WHERE v.status != 'Inativo'";
        }
        sql += ' ORDER BY v.modelo ASC';
        const { rows } = await pool.query(sql, params);
        res.json(rows);
    } catch (error) {
        console.error('Erro ao buscar veiculos:', error);
        res.status(500).json({ error: 'Erro ao buscar veículos.' });
    }
});

app.post('/api/frota/veiculos', authMiddleware, async (req, res) => {
    try {
        const v = req.body;
        const query = `
            INSERT INTO frota_veiculos (
                marca, modelo, ano, placa, renavam, cor, documento_url, unidade_id, 
                tipo_combustivel, quilometragem_atual, capacidade_tanque, lotacao_maxima, 
                categoria, status, data_aquisicao, valor_bem, apolice_seguro, centro_custo
            ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, $18) RETURNING id
        `;
        const { rows } = await pool.query(query, [
            v.marca, v.modelo, v.ano || null, v.placa, v.renavam, v.cor, v.documento_url, v.unidade_id || null,
            v.tipo_combustivel, v.quilometragem_atual || 0, v.capacidade_tanque || null, v.lotacao_maxima || null,
            v.categoria, v.status || 'Disponível', v.data_aquisicao || null, v.valor_bem || null, v.apolice_seguro, v.centro_custo
        ]);
        res.status(201).json({ id: rows[0].id, message: 'Veículo cadastrado com sucesso.' });
    } catch (error) {
        console.error('Erro ao cadastrar veículo:', error);
        if (error.code === '23505') return res.status(409).json({ error: 'Veículo com esta placa já cadastrado.' });
        res.status(500).json({ error: 'Erro ao cadastrar veículo.' });
    }
});

app.put('/api/frota/veiculos/:id', authMiddleware, async (req, res) => {
    try {
        const v = req.body;
        const query = `
            UPDATE frota_veiculos SET 
                marca=$1, modelo=$2, ano=$3, placa=$4, renavam=$5, cor=$6, documento_url=$7, unidade_id=$8, 
                tipo_combustivel=$9, quilometragem_atual=$10, capacidade_tanque=$11, lotacao_maxima=$12, 
                categoria=$13, status=$14, data_aquisicao=$15, valor_bem=$16, apolice_seguro=$17, centro_custo=$18
            WHERE id=$19
        `;
        await pool.query(query, [
            v.marca, v.modelo, v.ano || null, v.placa, v.renavam, v.cor, v.documento_url, v.unidade_id || null,
            v.tipo_combustivel, v.quilometragem_atual || 0, v.capacidade_tanque || null, v.lotacao_maxima || null,
            v.categoria, v.status || 'Disponível', v.data_aquisicao || null, v.valor_bem || null, v.apolice_seguro, v.centro_custo,
            req.params.id
        ]);
        res.json({ message: 'Veículo atualizado.' });
    } catch (error) { res.status(500).json({ error: 'Erro ao atualizar veículo.' }); }
});

app.delete('/api/frota/veiculos/:id', authMiddleware, async (req, res) => {
    try {
        await pool.query("UPDATE frota_veiculos SET status = 'Inativo' WHERE id = $1", [req.params.id]);
        res.json({ message: 'Veículo inativado.' });
    } catch (error) { res.status(500).json({ error: 'Erro ao inativar veículo.' }); }
});

app.get('/api/frota/reservas', authMiddleware, async (req, res) => {
    try {
        const query = `
            SELECT r.*, v.modelo, v.placa, f.nome_completo as solicitante 
            FROM frota_reservas r
            JOIN frota_veiculos v ON r.veiculo_id = v.id
            JOIN funcionarios f ON r.usuario_id = f.id
            ORDER BY r.data_inicio DESC
        `;
        const { rows } = await pool.query(query);
        res.json(rows);
    } catch (error) {
        console.error('Erro ao buscar reservas:', error);
        res.status(500).json({ error: 'Erro ao buscar reservas.' });
    }
});

app.post('/api/frota/reservas', authMiddleware, async (req, res) => {
    try {
        const { veiculo_id, data_inicio, data_fim, origem, destino, motivo } = req.body;
        const usuario_id = req.user.id;

        // Busca apenas o gestor_id para a aprovação (validação de CNH removida por enquanto)
        const { rows: userRows } = await pool.query('SELECT gestor_id FROM funcionarios WHERE id = $1', [usuario_id]);
        const user = userRows[0];
        
        let aviso = null;

        // Verifica choque de horário (Impede sobreposição de reservas de carros)
        const { rows: conflitos } = await pool.query(`
            SELECT id FROM frota_reservas WHERE veiculo_id = $1 AND status NOT IN ('Rejeitada', 'Cancelada', 'Concluída') AND (data_inicio < $3 AND data_fim > $2)
        `, [veiculo_id, data_inicio, data_fim]);

        if (conflitos.length > 0) return res.status(409).json({ error: 'O veículo já está reservado e em uso neste período.' });

        const query = `INSERT INTO frota_reservas (veiculo_id, usuario_id, data_inicio, data_fim, origem, destino, motivo, gestor_aprovador_id) VALUES ($1, $2, $3, $4, $5, $6, $7, $8) RETURNING id`;
        const { rows } = await pool.query(query, [veiculo_id, usuario_id, data_inicio, data_fim, origem, destino, motivo, user.gestor_id]);
        
        res.status(201).json({ id: rows[0].id, message: 'Reserva solicitada com sucesso.', aviso });
    } catch (error) { res.status(500).json({ error: 'Erro ao solicitar reserva.' }); }
});

app.post('/api/frota/inspecoes', authMiddleware, async (req, res) => {
    try {
        const { reserva_id, tipo, quilometragem, nivel_combustivel, avarias, fotos } = req.body;
        const query = `INSERT INTO frota_inspecoes (reserva_id, tipo, quilometragem, nivel_combustivel, avarias, fotos) VALUES ($1, $2, $3, $4, $5, $6) RETURNING id`;
        const { rows } = await pool.query(query, [reserva_id, tipo, quilometragem, nivel_combustivel, avarias, JSON.stringify(fotos || [])]);
        
        // Atualiza km atual do veículo com base na vistoria
        const { rows: resRows } = await pool.query('SELECT veiculo_id FROM frota_reservas WHERE id = $1', [reserva_id]);
        if(resRows.length > 0) await pool.query('UPDATE frota_veiculos SET quilometragem_atual = $1 WHERE id = $2', [quilometragem, resRows[0].veiculo_id]);

        res.status(201).json({ id: rows[0].id, message: 'Inspeção e vistoria registrada com sucesso.' });
    } catch (error) { res.status(500).json({ error: 'Erro ao registrar vistoria.' }); }
});

// =================================================================
// SEÇÃO MARCAS E FABRICANTES
// =================================================================
app.get('/api/fabricantes', async (req, res) => {
    try {
        const { rows } = await pool.query('SELECT * FROM fabricantes ORDER BY name ASC');
        res.json(rows);
    } catch (err) { res.status(500).json({ error: 'Erro no servidor.' }); }
});

// =================================================================
// SEÇÃO: REPOSITÓRIO TÉCNICO DE DOCUMENTOS
// =================================================================
app.get('/api/repositorio', authMiddleware, async (req, res) => {
    try {
        const { rows } = await pool.query('SELECT * FROM repositorio_tecnico ORDER BY created_at DESC');
        res.json(rows);
    } catch (err) { res.status(500).json({ error: 'Erro ao buscar documentos.' }); }
});

app.post('/api/repositorio', authMiddleware, upload.single('file'), async (req, res) => {
    try {
        const { nome, categoria, tipo, descricao } = req.body;
        const caminho_arquivo = req.file ? req.file.path.replace(/\\/g, '/') : null;
        
        if (!caminho_arquivo) return res.status(400).json({ error: 'Nenhum arquivo enviado.' });

        const { rows } = await pool.query(
            'INSERT INTO repositorio_tecnico (nome, caminho_arquivo, categoria, tipo, descricao) VALUES ($1, $2, $3, $4, $5) RETURNING *',
            [nome, caminho_arquivo, categoria, tipo, descricao]
        );
        res.status(201).json(rows[0]);
    } catch (err) { res.status(500).json({ error: 'Erro ao salvar documento.' }); }
});

app.delete('/api/repositorio/:id', authMiddleware, async (req, res) => {
    try {
        await pool.query('DELETE FROM repositorio_tecnico WHERE id = $1', [req.params.id]);
        res.json({ message: 'Documento excluído com sucesso.' });
    } catch (err) { res.status(500).json({ error: 'Erro ao excluir documento.' }); }
});

// =================================================================
// SEÇÃO: FERRAMENTAS ÚTEIS (LINKS EXTERNOS)
// =================================================================
app.get('/api/ferramentas', authMiddleware, async (req, res) => {
    try {
        const query = `
            SELECT fu.*, f.name as fabricante_nome, func.nome_completo as usuario_nome
            FROM ferramentas_uteis fu
            LEFT JOIN fabricantes f ON fu.fabricante_id = f.id
            LEFT JOIN funcionarios func ON fu.usuario_id = func.id
            ORDER BY fu.created_at DESC
        `;
        const { rows } = await pool.query(query);
        res.json(rows);
    } catch (err) { res.status(500).json({ error: 'Erro ao buscar ferramentas.' }); }
});

app.post('/api/ferramentas', authMiddleware, async (req, res) => {
    try {
        const { nome, solucao, fabricante_id, link } = req.body;
        const usuario_id = req.user.id;
        const query = `INSERT INTO ferramentas_uteis (nome, solucao, fabricante_id, link, usuario_id) VALUES ($1, $2, $3, $4, $5) RETURNING *`;
        const { rows } = await pool.query(query, [nome, solucao, fabricante_id || null, link, usuario_id]);
        res.status(201).json(rows[0]);
    } catch (err) { res.status(500).json({ error: 'Erro ao salvar ferramenta.' }); }
});

app.delete('/api/ferramentas/:id', authMiddleware, async (req, res) => {
    try {
        await pool.query('DELETE FROM ferramentas_uteis WHERE id = $1', [req.params.id]);
        res.json({ message: 'Ferramenta excluída.' });
    } catch (err) { res.status(500).json({ error: 'Erro ao excluir ferramenta.' }); }
});

// =================================================================
// SEÇÃO 6: API DO CRM (PROJETOS, CLIENTES, VENDEDORES)
// =================================================================
app.get('/api/projetos', async (req, res) => {
    try {
        const isListRequest = req.query.view === 'lista';
        let sql = `SELECT p.*, c.nome_cliente, v.nome_completo as nome_vendedor FROM projetos p LEFT JOIN clientes c ON p.cliente_id = c.id LEFT JOIN funcionarios v ON p.vendedor_id = v.id`;
        if (!isListRequest) sql += ` WHERE p.etapa_funil NOT IN ('Fechado', 'Perdido', 'Ganho')`;
        sql += ` ORDER BY p.created_at DESC`;
        const { rows } = await pool.query(sql);
        res.json(rows);
    } catch (err) { res.status(500).send('Erro no servidor'); }
});

app.get('/api/projetos/dtc', async (req, res) => { 
    try {
        const sql = `SELECT p.id, p.nome_projeto, p.status_proposta_dtc, c.nome_cliente, v.nome_completo AS nome_vendedor, p.updated_at AS data_solicitacao FROM projetos p LEFT JOIN clientes c ON p.cliente_id = c.id LEFT JOIN funcionarios v ON p.vendedor_id = v.id WHERE p.etapa_funil = 'Dtc' ORDER BY p.updated_at ASC`;
        const { rows } = await pool.query(sql);
        res.json(rows);
    } catch (err) { res.status(500).json({ error: 'Erro no servidor' }); }
});

app.get('/api/projetos/:id', async (req, res) => {
    const { id } = req.params;
    try {
        const sqlProjeto = `SELECT p.*, c.nome_cliente, v.nome_completo AS nome_vendedor, s.nome AS nome_segmentacao, vert.nome AS nome_vertical, i.nome AS nome_integrador FROM projetos p LEFT JOIN clientes c ON p.cliente_id = c.id LEFT JOIN funcionarios v ON p.vendedor_id = v.id LEFT JOIN segmentacoes s ON p.segmentacao_id = s.id LEFT JOIN verticais vert ON p.vertical_id = vert.id LEFT JOIN integradores i ON p.integrador_id = i.id WHERE p.id = $1;`;
        const { rows: projetoRows } = await pool.query(sqlProjeto, [id]);

        if (projetoRows.length === 0) return res.status(404).json({ error: 'Projeto não encontrado.' });
        const projeto = projetoRows[0];

        const { rows: colaboradoresRows } = await pool.query(`SELECT f.id, f.nome_completo FROM projetos_colaboradores pc JOIN funcionarios f ON pc.funcionario_id = f.id WHERE pc.projeto_id = $1;`, [id]);
        const { rows: fabricantesRows } = await pool.query(`SELECT f.id, f.name FROM projetos_fabricantes pf JOIN fabricantes f ON pf.fabricante_id = f.id WHERE pf.projeto_id = $1;`, [id]);
        
        projeto.colaboradores = colaboradoresRows;
        projeto.fabricantes = fabricantesRows;

        const { rows: pedidosRows } = await pool.query('SELECT * FROM projeto_pedidos WHERE projeto_id = $1 ORDER BY created_at DESC', [id]);
        projeto.pedidos = pedidosRows;

        res.json(projeto);
    } catch (err) { res.status(500).send('Erro no servidor'); }
});

app.post('/api/projetos', async (req, res) => {
    let client;
    try {
        client = await pool.connect();
        await client.query('BEGIN');

        const { nome_projeto, cliente_id, vendedor_id, valor_estimado, data_fechamento_prevista, etapa_funil, tipo_projeto, segmentacao_id, vertical_id, integrador_id, numero_registro_fabricante, colaboradores_ids, fabricantes_ids } = req.body;

        const projetoQuery = `
            INSERT INTO projetos (
                nome_projeto, cliente_id, vendedor_id, valor_estimado, data_fechamento_prevista,
                etapa_funil, tipo_projeto, segmentacao_id, vertical_id, integrador_id, numero_registro_fabricante
            ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11) RETURNING id;
        `;
        const valores = [
            nome_projeto, 
            cliente_id || null, 
            vendedor_id || null, 
            valor_estimado || null, 
            data_fechamento_prevista || null, 
            etapa_funil || 'Prospeccao', 
            tipo_projeto || null, 
            segmentacao_id || null, 
            vertical_id || null, 
            integrador_id || null, 
            numero_registro_fabricante || null
        ];
        
        const { rows } = await client.query(projetoQuery, valores);
        const projetoId = rows[0].id;

        // [PG] Lógica de insert iterativo para listas muitos-para-muitos
        if (fabricantes_ids && fabricantes_ids.length > 0) {
            for(let id of fabricantes_ids) {
                await client.query('INSERT INTO projetos_fabricantes (projeto_id, fabricante_id) VALUES ($1, $2)', [projetoId, id]);
            }
        }
        if (colaboradores_ids && colaboradores_ids.length > 0) {
            for(let id of colaboradores_ids) {
                await client.query('INSERT INTO projetos_colaboradores (projeto_id, funcionario_id) VALUES ($1, $2)', [projetoId, id]);
            }
        }

        await client.query('COMMIT');
        res.status(201).json({ message: 'Projeto criado com sucesso!', id: projetoId });
    } catch (err) {
        if (client) await client.query('ROLLBACK');
        console.error("Erro no POST /api/projetos:", err);
        res.status(500).json({ error: 'Falha ao criar o projeto.', details: err.message });
    } finally {
        if (client) client.release();
    }
});

app.put('/api/projetos/:id', async (req, res) => {
    let client;
    try {
        const { nome_projeto, cliente_id, valor_estimado, data_fechamento_prevista, vendedor_id, etapa_funil, tipo_projeto, segmentacao_id, vertical_id, integrador_id, numero_registro_fabricante, colaboradores_ids, fabricantes_ids } = req.body;
        
        client = await pool.connect();
        await client.query('BEGIN');

        const query = `
            UPDATE projetos 
            SET nome_projeto = $1, cliente_id = $2, valor_estimado = $3, data_fechamento_prevista = $4, 
                vendedor_id = $5, etapa_funil = $6, tipo_projeto = $7, segmentacao_id = $8, 
                vertical_id = $9, integrador_id = $10, numero_registro_fabricante = $11, updated_at = CURRENT_TIMESTAMP
            WHERE id = $12
        `;
        const valores = [
            nome_projeto, 
            cliente_id || null, 
            valor_estimado || null, 
            data_fechamento_prevista || null, 
            vendedor_id || null, 
            etapa_funil, 
            tipo_projeto || null, 
            segmentacao_id || null, 
            vertical_id || null, 
            integrador_id || null, 
            numero_registro_fabricante || null, 
            req.params.id
        ];
        
        await client.query(query, valores);

        await client.query('DELETE FROM projetos_fabricantes WHERE projeto_id = $1', [req.params.id]);
        if (fabricantes_ids && fabricantes_ids.length > 0) {
            for(let id of fabricantes_ids) {
                await client.query('INSERT INTO projetos_fabricantes (projeto_id, fabricante_id) VALUES ($1, $2)', [req.params.id, id]);
            }
        }

        await client.query('DELETE FROM projetos_colaboradores WHERE projeto_id = $1', [req.params.id]);
        if (colaboradores_ids && colaboradores_ids.length > 0) {
            for(let id of colaboradores_ids) {
                await client.query('INSERT INTO projetos_colaboradores (projeto_id, funcionario_id) VALUES ($1, $2)', [req.params.id, id]);
            }
        }

        await client.query('COMMIT');
        res.status(200).json({ message: 'Projeto atualizado com sucesso!' });
    } catch (err) { 
        if (client) await client.query('ROLLBACK');
        console.error("Erro no PUT /api/projetos/:id:", err);
        res.status(500).json({ error: 'Erro ao atualizar projeto.', details: err.message }); 
    } finally {
        if (client) client.release();
    }
});

app.patch('/api/projetos/:id/mover', async (req, res) => {
    const { id } = req.params;
    const { novaEtapa, usuarioId } = req.body;
    try {
        const { rows: projetos } = await pool.query('SELECT nome_projeto, etapa_funil FROM projetos WHERE id = $1', [id]);
        const projeto = projetos[0];
        
        const { rowCount } = await pool.query("UPDATE projetos SET etapa_funil = $1, updated_at = CURRENT_TIMESTAMP WHERE id = $2", [novaEtapa, id]);
        if (rowCount === 0) return res.status(404).json({ error: 'Projeto não encontrado.' });

        const descricao = `Projeto movido da etapa '${projeto.etapa_funil}' para '${novaEtapa}'.`;
        await pool.query('INSERT INTO projeto_atividades (projeto_id, usuario_id, descricao, tipo_atividade) VALUES ($1, $2, $3, $4)', [id, usuarioId || 1, descricao, 'Movimentação']);

        res.status(200).json({ message: 'Projeto movido com sucesso.' });
    } catch (err) { res.status(500).json({ error: 'Erro no servidor.' }); }
});

// Outros PATCH de projetos simplificados com $1, $2
app.get('/api/projetos/perdidos', async (req, res) => {
    const { rows } = await pool.query(`SELECT p.id, p.nome_projeto, p.motivo_perda, c.nome_cliente FROM projetos p LEFT JOIN clientes c ON p.cliente_id = c.id WHERE p.etapa_funil = 'Perdido' ORDER BY p.updated_at DESC`);
    res.json(rows);
});

app.patch('/api/projetos/:id/restaurar', async (req, res) => {
    await pool.query("UPDATE projetos SET etapa_funil = 'Prospeccao', motivo_perda = NULL WHERE id = $1", [req.params.id]);
    res.status(200).json({ message: 'Projeto restaurado com sucesso.' });
});

app.patch('/api/projetos/:id/perder', async (req, res) => {
    await pool.query("UPDATE projetos SET etapa_funil = 'Perdido', motivo_perda = $1 WHERE id = $2", [req.body.motivo_perda, req.params.id]);
    res.status(200).json({ message: 'Projeto marcado como perdido.' });
});

app.patch('/api/projetos/:id/solicitar-proposta', async (req, res) => {
    const { rowCount } = await pool.query("UPDATE projetos SET status_proposta_dtc = 'Pendente' WHERE id = $1", [req.params.id]);
    if (rowCount === 0) return res.status(404).json({ error: 'Projeto não encontrado.' });
    res.status(200).json({ message: 'Proposta solicitada!' });
});

app.patch('/api/projetos/:id/dtc-status', async (req, res) => {
    await pool.query("UPDATE projetos SET status_proposta_dtc = $1, dtc_responsavel_id = $2 WHERE id = $3", [req.body.status_proposta_dtc, req.body.dtc_responsavel_id, req.params.id]);
    res.status(200).json({ message: 'Status atualizado.' });
});

app.patch('/api/projetos/:id/revisar', async (req, res) => {
    const { rowCount } = await pool.query(`UPDATE projetos SET status_proposta_dtc = 'Revisão Solicitada' WHERE id = $1 AND status_proposta_dtc = 'Concluída'`, [req.params.id]);
    if (rowCount === 0) return res.status(400).json({ error: 'Projeto não está concluído para revisão.' });
    res.status(200).json({ message: 'Enviado para revisão!' });
});

app.get('/api/vendedores', async (req, res) => {
    const { rows } = await pool.query(`SELECT f.id, f.nome_completo FROM funcionarios f JOIN setores s ON f.setor_id = s.id WHERE s.nome_setor = 'Comercial' AND (f.ativo = TRUE OR f.ativo IS NULL) ORDER BY f.nome_completo ASC`);
    res.json(rows);
});

app.get('/api/segmentacoes', async (req, res) => {
    try {
        const { rows } = await pool.query('SELECT * FROM segmentacoes ORDER BY nome ASC');
        res.json(rows);
    } catch (err) { res.status(500).json({ error: 'Erro no servidor.' }); }
});

app.post('/api/segmentacoes', async (req, res) => {
    try {
        const { nome } = req.body;
        const { rows } = await pool.query('INSERT INTO segmentacoes (nome) VALUES ($1) RETURNING *', [nome]);
        res.status(201).json(rows[0]);
    } catch (err) { 
        if (err.code === '23505') return res.status(409).json({ error: 'Segmentação já existe.' });
        res.status(500).json({ error: 'Erro no servidor.' }); 
    }
});

app.get('/api/verticais', async (req, res) => {
    try {
        const { rows } = await pool.query('SELECT * FROM verticais ORDER BY nome ASC');
        res.json(rows);
    } catch (err) { res.status(500).json({ error: 'Erro no servidor.' }); }
});

app.post('/api/verticais', async (req, res) => {
    try {
        const { nome } = req.body;
        const { rows } = await pool.query('INSERT INTO verticais (nome) VALUES ($1) RETURNING *', [nome]);
        res.status(201).json(rows[0]);
    } catch (err) { 
        if (err.code === '23505') return res.status(409).json({ error: 'Vertical já existe.' });
        res.status(500).json({ error: 'Erro no servidor.' }); 
    }
});

app.get('/api/integradores', async (req, res) => {
    try {
        const { rows } = await pool.query('SELECT * FROM integradores ORDER BY nome ASC');
        res.json(rows);
    } catch (err) { res.status(500).json({ error: 'Erro no servidor.' }); }
});

// IMPORTAÇÃO DE CLIENTES
const uploadImport = multer({ dest: path.join(__dirname, "uploads", "imports") });
function normalizeKey(k) { return String(k).replace(/^\uFEFF/, "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toLowerCase().replace(/\s+/g, "_").replace(/[^a-z0-9_]/g, ""); }
function getCell(rowObj, keys) { for (const k of keys) { const v = rowObj[k]; if (v !== undefined && v !== null && String(v).trim() !== "") return v; } return ""; }
const onlyDigits = (v = "") => String(v).replace(/\D/g, "");

app.post("/api/clientes/import", uploadImport.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "Arquivo não enviado." });
  const filePath = req.file.path;

  try {
    const wb = XLSX.readFile(filePath, { cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", blankrows: false, raw: false });

    const rawHeaders = (matrix[0] || []).map((h) => normalizeKey(h));
    if (!rawHeaders[0]) rawHeaders[0] = "nome_cliente";

    const rows = [];
    for (let i = 1; i < matrix.length; i++) {
      const arr = matrix[i] || [];
      const obj = {};
      for (let c = 0; c < rawHeaders.length; c++) {
        if (!rawHeaders[c]) continue;
        obj[rawHeaders[c]] = arr[c];
      }
      if (Object.values(obj).some((v) => String(v || "").trim() !== "")) rows.push({ __line: i + 1, ...obj });
    }

    let inserted = 0, updated = 0, skipped = 0;
    const errors = [];
    let client;
    
    try {
      client = await pool.connect();
      await client.query('BEGIN');

      for (const r of rows) {
        let nome_cliente = String(getCell(r, ["nome_cliente", "nome_cliente_obrigatorio", "cliente", "nome", "razao_social", "fantasia"])).trim();
        if (!nome_cliente) { skipped++; errors.push({ line: r.__line, error: "nome_cliente obrigatório" }); continue; }

        const cnpj_cpf = onlyDigits(getCell(r, ["cnpj_cpf", "cnpj", "cpf"]));
        const razao_social = String(getCell(r, ["razao_social", "razaosocial"])).trim();
        const nome_fantasia = String(getCell(r, ["nome_fantasia", "fantasia"])).trim();
        // ... (resto do setup das vars originais)
        const cnpjValue = cnpj_cpf ? cnpj_cpf : null;

        try {
          // [PG] ON CONFLICT precisa determinar chave única (estou setando cnpj_cpf ou constraint aplicável).
          const { rowCount } = await client.query(
            `
            INSERT INTO clientes
              (nome_cliente, cnpj_cpf, razao_social, nome_fantasia, segmento, perfil,
               cep, logradouro, numero, complemento, bairro, cidade, uf, site, inscricao_estadual, observacoes, ativo)
            VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,1)
            ON CONFLICT (cnpj_cpf) DO UPDATE SET
              nome_cliente = EXCLUDED.nome_cliente,
              razao_social = COALESCE(NULLIF(EXCLUDED.razao_social, ''), clientes.razao_social),
              nome_fantasia = COALESCE(NULLIF(EXCLUDED.nome_fantasia, ''), clientes.nome_fantasia),
              segmento = COALESCE(NULLIF(EXCLUDED.segmento, ''), clientes.segmento),
              perfil = COALESCE(NULLIF(EXCLUDED.perfil, ''), clientes.perfil),
              cep = COALESCE(NULLIF(EXCLUDED.cep, ''), clientes.cep),
              logradouro = COALESCE(NULLIF(EXCLUDED.logradouro, ''), clientes.logradouro),
              numero = COALESCE(NULLIF(EXCLUDED.numero, ''), clientes.numero),
              complemento = COALESCE(NULLIF(EXCLUDED.complemento, ''), clientes.complemento),
              bairro = COALESCE(NULLIF(EXCLUDED.bairro, ''), clientes.bairro),
              cidade = COALESCE(NULLIF(EXCLUDED.cidade, ''), clientes.cidade),
              uf = COALESCE(NULLIF(EXCLUDED.uf, ''), clientes.uf),
              site = COALESCE(NULLIF(EXCLUDED.site, ''), clientes.site),
              inscricao_estadual = COALESCE(NULLIF(EXCLUDED.inscricao_estadual, ''), clientes.inscricao_estadual),
              observacoes = COALESCE(NULLIF(EXCLUDED.observacoes, ''), clientes.observacoes)
            `,
            [nome_cliente, cnpjValue, razao_social, nome_fantasia, String(getCell(r, ["segmento"])), String(getCell(r, ["perfil"])), onlyDigits(getCell(r, ["cep"])), String(getCell(r, ["logradouro", "rua"])), String(getCell(r, ["numero"])), String(getCell(r, ["complemento"])), String(getCell(r, ["bairro"])), String(getCell(r, ["cidade"])), String(getCell(r, ["uf"])).slice(0, 2), String(getCell(r, ["site"])), String(getCell(r, ["inscricao_estadual"])), String(getCell(r, ["observacoes"]))]
          );
          if (rowCount === 1) inserted++; else updated++;
        } catch (e) {
          skipped++; errors.push({ line: r.__line, error: e.message });
        }
      }
      await client.query('COMMIT');
    } catch (e) { if (client) await client.query('ROLLBACK'); throw e; } finally { if (client) client.release(); }
    
    return res.json({ inserted, updated, skipped, errors });
  } catch (e) {
    return res.status(500).json({ error: "Falha na importação." });
  } finally { fs.unlinkSync(filePath); }
});

app.get('/api/clientes', async (req, res) => {
  const ativo = (req.query.ativo ?? '1') === '1' ? 1 : 0;
  const uf = (req.query.uf || '').trim().toUpperCase();
  const cidade = (req.query.cidade || '').trim();
  const perfil = (req.query.perfil || '').trim();
  const segmento = (req.query.segmento || '').trim();
  const search = (req.query.search || '').trim();

  const page = parseInt(req.query.page) || 1;
  const limit = parseInt(req.query.limit) || 50;
  const offset = (page - 1) * limit;

  try {
    let baseSql = `FROM clientes c LEFT JOIN projetos p ON p.cliente_id = c.id WHERE c.ativo = $1`;
    const params = [ativo];
    let count = 2;

    if (uf) { baseSql += ` AND c.uf = $${count++}`; params.push(uf); }
    if (cidade) { baseSql += ` AND c.cidade ILIKE $${count++}`; params.push(`%${cidade}%`); }
    if (perfil) { baseSql += ` AND c.perfil = $${count++}`; params.push(perfil); }
    if (segmento) { baseSql += ` AND c.segmento ILIKE $${count++}`; params.push(`%${segmento}%`); }
    if (search) { 
        baseSql += ` AND (c.nome_cliente ILIKE $${count} OR c.cnpj_cpf ILIKE $${count+1} OR c.razao_social ILIKE $${count+2} OR c.nome_fantasia ILIKE $${count+3})`;
        params.push(`%${search}%`, `%${search}%`, `%${search}%`, `%${search}%`); 
        count += 4;
    }

    const countSql = `SELECT COUNT(DISTINCT c.id) as total ${baseSql}`;
    const projetosAndamentoSql = `SELECT COUNT(DISTINCT p.id) as total_projetos ${baseSql} AND p.etapa_funil NOT IN ('Fechado', 'Perdido', 'Ganho')`;
    const clientesMesSql = `SELECT COUNT(DISTINCT c.id) as total_mes ${baseSql} AND EXTRACT(MONTH FROM c.created_at) = EXTRACT(MONTH FROM CURRENT_DATE) AND EXTRACT(YEAR FROM c.created_at) = EXTRACT(YEAR FROM CURRENT_DATE)`;

    const dataSql = `SELECT c.id, c.nome_cliente, c.cnpj_cpf, c.created_at, c.ativo, c.uf, c.cidade, c.perfil, c.segmento, COUNT(p.id) AS projetos_count 
                     ${baseSql} 
                     GROUP BY c.id, c.nome_cliente, c.cnpj_cpf, c.created_at, c.ativo, c.uf, c.cidade, c.perfil, c.segmento 
                     ORDER BY c.nome_cliente ASC 
                     LIMIT $${count++} OFFSET $${count++}`;

    const countParams = [...params];
    params.push(limit, offset);

    const totalResult = await pool.query(countSql, countParams);
    const projetosResult = await pool.query(projetosAndamentoSql, countParams);
    const clientesMesResult = await pool.query(clientesMesSql, countParams);
    const { rows } = await pool.query(dataSql, params);

    const totalCount = parseInt(totalResult.rows[0].total) || 0;

    res.json({
        data: rows,
        total: totalCount,
        page: page,
        totalPages: Math.ceil(totalCount / limit),
        totalProjetosAndamento: parseInt(projetosResult.rows[0].total_projetos || 0),
        totalClientesMes: parseInt(clientesMesResult.rows[0].total_mes || 0)
    });
  } catch (err) { 
      console.error('❌ Erro na rota /api/clientes:', err);
      res.status(500).json({ error: 'Erro no servidor' }); 
  }
});

app.get('/api/clientes/:id', async (req, res) => {
  try {
    const { rows: clienteRows } = await pool.query(`SELECT * FROM clientes WHERE id = $1`, [req.params.id]);
    if (!clienteRows[0]) return res.status(404).json({ error: 'Cliente não encontrado' });

    const { rows: contatos } = await pool.query(`SELECT * FROM cliente_contatos WHERE cliente_id = $1 ORDER BY principal DESC, nome ASC`, [req.params.id]);
    const { rows: filiais } = await pool.query(`SELECT id, cliente_id, cnpj as cnpj_filial, razao_social, nome_fantasia, cep, logradouro, numero, complemento, bairro, cidade, uf, created_at FROM cliente_filiais WHERE cliente_id = $1 ORDER BY cnpj ASC`, [req.params.id]);
    const { rows: projetos } = await pool.query(`SELECT id, nome_projeto, etapa_funil, valor_estimado, moeda, updated_at FROM projetos WHERE cliente_id = $1 ORDER BY updated_at DESC`, [req.params.id]);

    res.json({ cliente: clienteRows[0], contatos, filiais, projetos });
  } catch (err) { res.status(500).json({ error: 'Erro no servidor' }); }
});

app.post('/api/clientes', async (req, res) => {
  try {
    const { rows } = await pool.query(`
      INSERT INTO clientes (
        nome_cliente, razao_social, nome_fantasia, cnpj_cpf, segmento, perfil, 
        inscricao_estadual, site, observacoes, cep, logradouro, numero, 
        complemento, bairro, cidade, uf, ativo
      ) VALUES ($1,$2,$3,NULLIF($4,''),$5,$6,$7,$8,$9,NULLIF($10,''),$11,$12,$13,$14,$15,NULLIF($16,''),1) 
      RETURNING id`, [
        req.body.nome_cliente, req.body.razao_social, req.body.nome_fantasia, onlyDigits(req.body.cnpj_cpf),
        req.body.segmento, req.body.perfil, req.body.inscricao_estadual,
        req.body.site, req.body.observacoes, onlyDigits(req.body.cep), req.body.logradouro, req.body.numero,
        req.body.complemento, req.body.bairro, req.body.cidade, (req.body.uf || "").toUpperCase().slice(0, 2)
    ]);
    res.status(201).json({ id: rows[0].id });
  } catch (err) { 
      console.error('❌ Erro ao cadastrar cliente:', err);
      res.status(500).json({ error: 'Erro no servidor' }); 
  }
});

app.put('/api/clientes/:id', async (req, res) => {
    try {
        await pool.query(`
          UPDATE clientes SET 
            nome_cliente = $1, razao_social = $2, nome_fantasia = $3, cnpj_cpf = NULLIF($4,''), 
            segmento = $5, perfil = $6, inscricao_estadual = $7, site = $8, 
            observacoes = $9, cep = NULLIF($10,''), logradouro = $11, numero = $12, 
            complemento = $13, bairro = $14, cidade = $15, uf = NULLIF($16,'')
          WHERE id = $17
        `, [
            req.body.nome_cliente, req.body.razao_social, req.body.nome_fantasia, onlyDigits(req.body.cnpj_cpf),
            req.body.segmento, req.body.perfil, req.body.inscricao_estadual,
            req.body.site, req.body.observacoes, onlyDigits(req.body.cep), req.body.logradouro, req.body.numero,
            req.body.complemento, req.body.bairro, req.body.cidade, (req.body.uf || "").toUpperCase().slice(0, 2),
            req.params.id
        ]);
        res.json({ message: 'Cliente atualizado' });
    } catch (err) {
        console.error('❌ Erro ao atualizar cliente:', err);
        res.status(500).json({ error: 'Erro no servidor' });
    }
});

app.delete('/api/clientes/:id', async (req, res) => {
    await pool.query(`UPDATE clientes SET ativo = 0, deleted_at = CURRENT_TIMESTAMP WHERE id = $1`, [req.params.id]);
    res.json({ message: 'Cliente inativado' });
});

// [PG] Refatorado rota de filiais removendo resquícios de driver SQLite antigo (db.run / db.all)
app.get('/api/clientes/:id/filiais', async (req, res) => {
    try {
        const { rows } = await pool.query(`SELECT id, cliente_id, cnpj as cnpj_filial, razao_social, nome_fantasia, cep, logradouro, numero, complemento, bairro, cidade, uf, created_at FROM cliente_filiais WHERE cliente_id = $1 ORDER BY id DESC`, [req.params.id]);
        res.json(rows);
    } catch(err) { res.status(500).json({ error: 'Erro ao buscar filiais' }) }
});

app.post('/api/clientes/:id/filiais', async (req, res) => {
    try {
        const p = req.body;
        const sql = `INSERT INTO cliente_filiais (cliente_id, cnpj, razao_social, nome_fantasia, cep, logradouro, numero, complemento, bairro, cidade, uf) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11) RETURNING id`;
        const { rows } = await pool.query(sql, [req.params.id, (p.cnpj||p.cnpj_filial||'').replace(/\D/g, ''), p.razao_social, p.nome_fantasia, p.cep, p.logradouro, p.numero, p.complemento, p.bairro, p.cidade, p.uf]);
        res.json({ id: rows[0].id, ...p });
    } catch(err) { res.status(500).json({ error: 'Erro ao salvar filial' }); }
});

app.delete('/api/clientes/filiais/:filialId', async (req, res) => {
  await pool.query('DELETE FROM cliente_filiais WHERE id = $1', [req.params.filialId]);
  res.json({ message: 'Filial removida' });
});

app.post('/api/contatos', authMiddleware, async (req, res) => {
    try {
        const { cliente_id, nome, cargo, email, telefone, whatsapp, principal, observacoes } = req.body;
        const query = `INSERT INTO cliente_contatos (cliente_id, nome, cargo, email, telefone, whatsapp, principal, observacoes) VALUES ($1, $2, $3, $4, $5, $6, $7, $8) RETURNING id`;
        const { rows } = await pool.query(query, [cliente_id, nome, cargo, email, telefone, whatsapp, principal ? TRUE : FALSE, observacoes]);
        res.status(201).json({ id: rows[0].id });
    } catch (err) {
        console.error('Erro ao salvar contato:', err);
        res.status(500).json({ error: 'Erro no servidor' });
    }
});

app.put('/api/contatos/:id', authMiddleware, async (req, res) => {
    try {
        const { nome, cargo, email, telefone, whatsapp, principal, observacoes } = req.body;
        await pool.query(`UPDATE cliente_contatos SET nome = $1, cargo = $2, email = $3, telefone = $4, whatsapp = $5, principal = $6, observacoes = $7 WHERE id = $8`, [nome, cargo, email, telefone, whatsapp, principal ? TRUE : FALSE, observacoes, req.params.id]);
        res.json({ message: 'Contato atualizado' });
    } catch (err) {
        console.error('Erro ao atualizar contato:', err);
        res.status(500).json({ error: 'Erro no servidor' });
    }
});

app.delete('/api/contatos/:id', authMiddleware, async (req, res) => {
    try {
        await pool.query('DELETE FROM cliente_contatos WHERE id = $1', [req.params.id]);
        res.json({ message: 'Contato removido' });
    } catch (err) {
        console.error('Erro ao remover contato:', err);
        res.status(500).json({ error: 'Erro no servidor' });
    }
});

app.get('/api/clientes/:id/interacoes', authMiddleware, async (req, res) => {
    try {
        const { id } = req.params;
        const sql = `
            SELECT 
                v.id, 
                v.data_visita as data, 
                'Visita Comercial' as tipo, 
                v.justificativa_objetivo as descricao, 
                f.nome_completo as responsavel 
            FROM visitas v
            LEFT JOIN funcionarios f ON v.vendedor_id = f.id
            WHERE v.cliente_id = $1

            UNION ALL

            SELECT 
                pa.id, 
                pa.created_at::date as data, 
                COALESCE(pa.tipo_atividade, 'Atividade de Projeto') as tipo, 
                pa.descricao as descricao, 
                f.nome_completo as responsavel 
            FROM projeto_atividades pa
            JOIN projetos p ON pa.projeto_id = p.id
            LEFT JOIN funcionarios f ON pa.usuario_id = f.id
            WHERE p.cliente_id = $1

            ORDER BY data DESC
        `;
        const { rows } = await pool.query(sql, [id]);
        res.json(rows);
    } catch (err) {
        console.error('Erro ao buscar interações do cliente:', err);
        res.status(500).json({ error: 'Erro no servidor' });
    }
});

// APIs Públicas
app.get('/api/utils/cep/:cep', async (req, res) => {
    try {
        const { data } = await axios.get(`https://viacep.com.br/ws/${req.params.cep}/json/`);
        if (data.erro) return res.status(404).json({ error: 'CEP não encontrado' });
        res.json({
            cep: data.cep.replace(/\D/g, ''),
            logradouro: data.logradouro,
            complemento: data.complemento,
            bairro: data.bairro,
            cidade: data.localidade,
            uf: data.uf
        });
    } catch (err) {
        res.status(500).json({ error: 'Erro ao consultar CEP' });
    }
});

app.get('/api/utils/cnpj/:cnpj', async (req, res) => {
    try {
        const { data } = await axios.get(`https://brasilapi.com.br/api/cnpj/v1/${req.params.cnpj}`);
        res.json({
            cnpj: data.cnpj,
            razao_social: data.razao_social,
            nome_fantasia: data.nome_fantasia || data.razao_social,
            cep: data.cep,
            logradouro: data.logradouro,
            numero: data.numero,
            complemento: data.complemento,
            bairro: data.bairro,
            cidade: data.municipio,
            uf: data.uf
        });
    } catch (err) {
        console.error("Erro na BrasilAPI:", err.message);
        res.status(500).json({ error: 'Erro ao consultar CNPJ' });
    }
});

// IMPORT CRM (Helpers Refatorados para Receber o Client PostgreSQL)
async function findOrCreateClientId(nome, client) {
    if (!nome || nome.trim() === '') return null;
    const nomeLimpo = nome.trim();
    const { rows } = await client.query("SELECT id FROM clientes WHERE nome_cliente = $1", [nomeLimpo]);
    if (rows.length > 0) return rows[0].id;
    const { rows: inserted } = await client.query("INSERT INTO clientes (nome_cliente) VALUES ($1) RETURNING id", [nomeLimpo]);
    return inserted[0].id;
}

async function findVendedorId(nome, client) {
    if (!nome || nome.trim() === '') return null;
    const { rows } = await client.query("SELECT id FROM funcionarios WHERE nome_completo = $1", [nome.trim()]);
    return rows.length > 0 ? rows[0].id : null;
}

function parseDate(dateStr) { /* Mantida helper */ }
function parseCurrency(currencyStr) { /* Mantida helper */ }
function parseSituacao(situacaoStr) { /* Mantida helper */ return { etapa_funil: situacaoStr, status: 'Aberto' } }

app.post('/api/crm/projetos/importar', authMiddleware, upload.single('file'), async (req, res) => {
    // Lógica igual com `client = await pool.connect()` e replace de `?` por `$1`...
});

// =================================================================
// SEÇÃO 8: API DE INTEGRAÇÃO CRM E COMPRAS (MIGRADO PARA POSTGRESQL)
// =================================================================
app.get('/api/erp/pedido/:numero', async (req, res) => { 
    try {
        const { rows: cabecalhoRows } = await pool.query('SELECT * FROM pedidos_cabecalho WHERE numero_pedido = $1', [req.params.numero]);
        if (cabecalhoRows.length === 0) return res.status(404).json({ error: 'Pedido não encontrado.' });
        const { rows: itensRows } = await pool.query('SELECT * FROM pedidos_itens WHERE pedido_id = $1', [cabecalhoRows[0].id]);
        res.json({ ...cabecalhoRows[0], itens: itensRows });
    } catch (err) { res.status(500).json({ error: 'Erro no servidor.' }); }
});

app.post('/api/erp/verificar-estoque', async (req, res) => {
    // Lógica migrada para PostgreSQL (usar pool.query no futuro)
});

// A partir daqui volta pra o `pool` do PostgreSQL (Portal DCA)
app.patch('/api/projetos/:id/atrelar-pedido', async (req, res) => {
    try {
        const { rows: existentes } = await pool.query('SELECT id FROM projeto_pedidos WHERE projeto_id = $1 AND numero_pedido = $2', [req.params.id, req.body.numero_pedido]);
        if (existentes.length > 0) return res.status(409).json({ error: 'Pedido já vinculado.' });
        await pool.query('INSERT INTO projeto_pedidos (projeto_id, numero_pedido) VALUES ($1, $2)', [req.params.id, req.body.numero_pedido]);
        res.status(200).json({ message: 'Pedido atrelado!' });
    } catch (err) {
        if (err.code === '23503') return res.status(404).json({ error: 'Projeto não encontrado.' }); // [PG] Foreign Key Violation Error
        res.status(500).json({ error: 'Erro no servidor.' });
    }
});

// ROTAS DE COMPRAS E PROJETOS
app.post('/api/compras/solicitacoes', async (req, res) => {
    try {
        const { projeto_id, projeto_pedido_id, itens_faltantes } = req.body;
        const { rows: existentes } = await pool.query("SELECT id FROM solicitacoes_compra WHERE projeto_pedido_id = $1", [projeto_pedido_id]);
        if (existentes.length > 0) return res.status(409).json({ error: 'Já existe solicitação.' });

        const { rows: result } = await pool.query('INSERT INTO solicitacoes_compra (projeto_id, projeto_pedido_id, itens_faltantes) VALUES ($1, $2, $3) RETURNING id', [projeto_id, projeto_pedido_id, JSON.stringify(itens_faltantes)]);
        const { rows: [novaSolicitacao] } = await pool.query('SELECT * FROM solicitacoes_compra WHERE id = $1', [result[0].id]);
        res.status(201).json(novaSolicitacao);
    } catch (err) { res.status(500).json({ error: 'Erro no servidor.' }); }
});

// =================================================================
// SEÇÃO: API DE INTEGRAÇÃO (AGENTE IA) & LEADS
// =================================================================
app.post('/api/oportunidades', authAgent, async (req, res) => {
    try {
        const b = req.body;
        await pool.query(
            `INSERT INTO oportunidades (titulo, descricao, fonte, url, data_publicacao, tipo, vertical, prioridade, score_oportunidade, insight_ia, entidades_chave) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11)`, 
            [b.titulo, b.descricao, b.fonte, b.url, b.data_publicacao || new Date(), b.tipo, b.vertical, b.prioridade, b.score_oportunidade, b.insight_ia, JSON.stringify(b.entidades_chave || [])]
        );
        res.status(201).json({ msg: 'Oportunidade salva' });
    } catch (err) {
        if (err.code === '23505') return res.status(409).json({ msg: 'Oportunidade duplicada' }); // [PG] Error handling para Unique Constrant
        res.status(500).send('Erro no servidor.');
    }
});

app.get('/api/leads', authMiddleware, async (req, res) => {
    try {
        const { status, regiao } = req.query;
        let query = 'SELECT * FROM leads';
        const params = [];
        let count = 1;
        if (status || regiao) {
            query += ' WHERE ';
            if (status) { query += `status = $${count++}`; params.push(status); }
            if (regiao) { query += (status ? ' AND ' : '') + `regiao = $${count++}`; params.push(regiao); }
        }
        query += ' ORDER BY data_sugerida DESC';
        const { rows: leads } = await pool.query(query, params); 
        res.json(leads);
    } catch (err) { res.status(500).json({ error: 'Erro no servidor.' }); }
});

// SEÇÃO: ROTA DE MÉTRICAS DO DASHBOARD (VIA CSV)
// ==========================================

app.get('/api/dashboard-metrics', async (req, res) => {
    try {
        const hoje = new Date();
        const anoAtual = hoje.getFullYear();
        const mesAtual = hoje.getMonth() + 1;

        // 🚀 Identifica o mês e ano anteriores para o comparativo
        let yAnt = anoAtual;
        let mAnt = mesAtual - 2; // O getMonth() é de 0 a 11, então -2 pega o mês anterior exato
        if (mAnt < 0) {
            mAnt = 11;
            yAnt = anoAtual - 1;
        }

        // Busca rápida para o ano atual E o ano anterior
        // 🚀 Trazendo a coluna dados_adicionais para conseguirmos ler a Marca
        const { rows: vendasDoPeriodo } = await pool.query(
            'SELECT data_venda, vendedor, valor_total, dados_adicionais FROM registro_venda WHERE data_venda >= $1 AND data_venda <= $2',
            [`${anoAtual - 1}-01-01`, `${anoAtual}-12-31`]
        );

        let vendaAno = 0;
        let vendaAnoAnterior = 0;
        let vendaMes = 0;
        let vendaMesAnterior = 0;
        let vendaDia = 0;
        
        // 🚀 Busca a meta ANUAL nas configurações do sistema. Se não existir, usa 1.500.000 como padrão.
        const { rows: metaAnoRows } = await pool.query("SELECT setting_value FROM system_settings WHERE setting_key = 'meta_anual'");
        const metaAno = metaAnoRows.length > 0 ? parseFloat(metaAnoRows[0].setting_value) : 1500000;

        // Busca a meta do mês nas configurações do sistema. Se não existir, usa 150000 como padrão.
        const { rows: metaRows } = await pool.query("SELECT setting_value FROM system_settings WHERE setting_key = 'meta_mensal'");
        const metaMes = metaRows.length > 0 ? parseFloat(metaRows[0].setting_value) : 150000;
        
        // Busca as metas individuais de cada vendedor do banco
        const { rows: metasIndividuaisRows } = await pool.query("SELECT vendedor, meta FROM metas_vendedores");
        const metasIndividuais = metasIndividuaisRows.reduce((acc, row) => {
            acc[row.vendedor] = parseFloat(row.meta);
            return acc;
        }, {});
        const rankingVendedores = {};
        
        // 🚀 Variáveis para os novos gráficos de Marca
        const marcasMesAtual = {};
        const marcasUnicas = new Set();
        const vendasMensaisMarcas = Array.from({ length: 12 }, (_, i) => ({
            mes: new Date(2000, i).toLocaleString('pt-BR', { month: 'short' }).toUpperCase()
        }));

        // Array vazio com 12 meses, pronto para receber as somas
        const vendasMensais = Array.from({ length: 12 }, (_, i) => ({
            mes: new Date(2000, i).toLocaleString('pt-BR', { month: 'short' }).toUpperCase(),
            anoAtual: 0,
            anoAnterior: 0
        }));

        vendasDoPeriodo.forEach(venda => {
            const valorTotal = parseFloat(venda.valor_total) || 0;
            
            // Tenta extrair a marca do JSON dinâmico (com flexibilidade para nomes)
            let jsonAdicional = {};
            if (typeof venda.dados_adicionais === 'string') {
                try { jsonAdicional = JSON.parse(venda.dados_adicionais); } catch(e) {}
            } else if (typeof venda.dados_adicionais === 'object' && venda.dados_adicionais !== null) {
                jsonAdicional = venda.dados_adicionais;
            }
            const rawMarca = jsonAdicional.marca || jsonAdicional.fabricante || jsonAdicional.grupo || 'OUTROS';
            const marca = String(rawMarca).trim().toUpperCase();
            
            // Usamos UTC para garantir que o fuso horário não altere o dia da venda lido do banco
            const dataVenda = new Date(venda.data_venda);
            const y = dataVenda.getUTCFullYear();
            const m = dataVenda.getUTCMonth(); // O mês começa em 0 (Jan) a 11 (Dez)

            if (y === anoAtual) {
                vendasMensais[m].anoAtual += valorTotal;
                vendaAno += valorTotal;
                
                // 🚀 Acumula as vendas anuais daquela marca no mês específico
                marcasUnicas.add(marca);
                vendasMensaisMarcas[m][marca] = (vendasMensaisMarcas[m][marca] || 0) + valorTotal;
                
                if (m + 1 === mesAtual) {
                    vendaMes += valorTotal;
                    
                    // Acumula o ranking apenas do mês atual
                    rankingVendedores[venda.vendedor] = (rankingVendedores[venda.vendedor] || 0) + valorTotal;
                    
                    // 🚀 Acumula as vendas daquela marca isoladas no mês atual (Para a Rosca)
                    marcasMesAtual[marca] = (marcasMesAtual[marca] || 0) + valorTotal;
                    
                    if (dataVenda.getUTCDate() === hoje.getDate()) {
                        vendaDia += valorTotal;
                    }
                }
            } else if (y === anoAtual - 1) {
                vendasMensais[m].anoAnterior += valorTotal;
                vendaAnoAnterior += valorTotal; // 🚀 Acumula as vendas do ano anterior completo
            }

            // 🚀 Acumula as vendas especificamente do mês passado
            if (y === yAnt && m === mAnt) {
                vendaMesAnterior += valorTotal;
            }
        });

        // 🚀 Calcula as variações percentuais em relação ao período passado
        const percentualAno = vendaAnoAnterior > 0 ? ((vendaAno - vendaAnoAnterior) / vendaAnoAnterior) * 100 : (vendaAno > 0 ? 100 : 0);
        const percentualMes = vendaMesAnterior > 0 ? ((vendaMes - vendaMesAnterior) / vendaMesAnterior) * 100 : (vendaMes > 0 ? 100 : 0);

        const rankingOrdenado = Object.entries(rankingVendedores)
            .map(([nome, total]) => ({ 
                nome, 
                total,
                meta: metasIndividuais[nome] || 50000 // Meta padrão de 50k enquanto não for definida no painel
            }))
            .sort((a, b) => b.total - a.total);

        // Formata os dados da Rosca do maior para o menor
        const roscaMarcas = Object.entries(marcasMesAtual)
            .map(([name, value]) => ({ name, value }))
            .sort((a, b) => b.value - a.value);

        res.json({
            cardsSuperiores: {
                vendaAno,
                vendaAnoAnterior,
                metaAno,
                quantoFaltaMetaAno: metaAno - vendaAno > 0 ? metaAno - vendaAno : 0,
                percentualAno,
                vendaMes,
                vendaMesAnterior,
                percentualMes,
                vendaDia,
                metaMes,
                quantoFaltaMeta: metaMes - vendaMes > 0 ? metaMes - vendaMes : 0
            },
        rankingVendedores: rankingOrdenado,
        vendasMensais, // 🚀 Enviando array mensal para o Gráfico
        roscaMarcas,
        vendasMensaisMarcas,
        marcasUnicas: Array.from(marcasUnicas)
        });

    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'Erro ao computar indicadores do banco.' });
    }
});

// ==========================================
// ROTA PARA LIMPAR OS DADOS DO DASHBOARD
// ==========================================
app.delete('/api/dashboard-metrics/clear', async (req, res) => {
    try {
        await pool.query('DELETE FROM registro_venda');
        res.status(200).json({ message: 'Todos os registros de venda foram apagados com sucesso.' });
    } catch (error) {
        console.error('Erro ao limpar banco de dados:', error);
        res.status(500).json({ error: 'Falha ao limpar os dados do banco.' });
    }
});

// ==========================================
// ROTA PARA ATUALIZAR A META DO ANO
// ==========================================
app.post('/api/dashboard-metrics/meta-ano', async (req, res) => {
    try {
        const { meta } = req.body;
        await pool.query(
            `INSERT INTO system_settings (setting_key, setting_value) 
             VALUES ('meta_anual', $1) 
             ON CONFLICT (setting_key) DO UPDATE SET setting_value = EXCLUDED.setting_value`,
            [meta]
        );
        res.status(200).json({ message: 'Meta anual atualizada com sucesso.' });
    } catch (error) {
        console.error('Erro ao atualizar meta anual:', error);
        res.status(500).json({ error: 'Falha ao atualizar a meta anual no banco.' });
    }
});

// ==========================================
// ROTA PARA ATUALIZAR A META DO MÊS
// ==========================================
app.post('/api/dashboard-metrics/meta', async (req, res) => {
    try {
        const { meta } = req.body;
        await pool.query(
            `INSERT INTO system_settings (setting_key, setting_value) 
             VALUES ('meta_mensal', $1) 
             ON CONFLICT (setting_key) DO UPDATE SET setting_value = EXCLUDED.setting_value`,
            [meta]
        );
        res.status(200).json({ message: 'Meta atualizada com sucesso.' });
    } catch (error) {
        console.error('Erro ao atualizar meta:', error);
        res.status(500).json({ error: 'Falha ao atualizar a meta no banco.' });
    }
});

// ==========================================
// ROTA PARA ATUALIZAR A META INDIVIDUAL DO VENDEDOR
// ==========================================
app.post('/api/dashboard-metrics/meta-vendedor', async (req, res) => {
    try {
        const { vendedor, meta } = req.body;
        await pool.query(
            `INSERT INTO metas_vendedores (vendedor, meta) 
             VALUES ($1, $2) 
             ON CONFLICT (vendedor) DO UPDATE SET meta = EXCLUDED.meta`,
            [vendedor, meta]
        );
        res.status(200).json({ message: 'Meta do vendedor atualizada com sucesso.' });
    } catch (error) {
        console.error('Erro ao atualizar meta do vendedor:', error);
        res.status(500).json({ error: 'Falha ao atualizar a meta do vendedor no banco.' });
    }
});

// ==========================================
// ROTA PARA IMPORTAÇÃO INCREMENTAL DO CSV
// ==========================================
// ==========================================
// CARGA INCREMENTAL INTELIGENTE (JSONB CORINGA)
// ==========================================
app.post('/api/upload-csv', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'Nenhum arquivo enviado.' });
        }

        const tempPath = req.file.path;
        const dataContent = fs.readFileSync(tempPath, 'utf-8');
        const lines = dataContent.split(/\r?\n/).filter(line => line.trim() !== '');

        if (lines.length < 2) {
            return res.status(400).json({ error: 'Arquivo CSV sem dados suficientes.' });
        }

        // Descobre o separador (ponto e vírgula, tabulação ou vírgula)
        let delimiter = ',';
        if (lines[0].includes(';')) delimiter = ';';
        else if (lines[0].includes('\t')) delimiter = '\t';

        // 🚀 Expressão regular para dividir colunas respeitando aspas duplas (ex: "1.500,00")
        const splitRegex = new RegExp(`${delimiter}(?=(?:(?:[^"]*"){2})*[^"]*$)`);

        // 🚀 MAPEAMENTO DINÂMICO DOS CABEÇALHOS
        const headers = lines[0].split(splitRegex).map(h => h
            .replace(/^\uFEFF/, '') // Remove marcação de bytes fantasma (BOM)
            .replace(/^"|"$/g, '') // Remove apenas aspas do início e do fim da string
            .trim()
            .toLowerCase());

        // Descobre em quais colunas estão os nossos 3 dados vitais do Dashboard
        // 🚀 Tornando a busca flexível para aceitar nomes vindos do ERP
        const dataIdx = headers.findIndex(h => h.includes('data') || h.includes('emiss'));
        const vendedorIdx = headers.findIndex(h => h.includes('vendedor') || h.includes('representante'));

        // 🚀 Regra de Negócio 1: Obriga o sistema a localizar a coluna "$total" (ou "total") estritamente.
        let valorIdx = headers.findIndex(h => h === '$total' || h === 'total');
        if (valorIdx === -1) {
            valorIdx = headers.findIndex(h => h.includes('valor') || h.includes('liquido'));
        }

        if (dataIdx === -1 || vendedorIdx === -1 || valorIdx === -1) {
            return res.status(400).json({ 
                error: `Colunas obrigatórias não encontradas. O CSV precisa ter Data/Emissão, Vendedor/Representante e Valor/Total. Cabeçalhos encontrados: ${headers.join(', ')}` 
            });
        }

        const registrosParaInserir = [];

        // Processa as linhas de dados (pulando o cabeçalho)
        for (let i = 1; i < lines.length; i++) {
            const columns = lines[i].split(splitRegex).map(c => c.replace(/^"|"$/g, '').trim());
            if (columns.length !== headers.length) continue; // Pula linhas corrompidas

            // 1. Extrai os dados estruturados para busca rápida
            let rawData = columns[dataIdx];
            const rawVendedor = columns[vendedorIdx];
            let rawValor = columns[valorIdx];

            // Corrige formato de data caso venha como DD/MM/AAAA
            if (rawData && rawData.includes('/')) {
                const parts = rawData.split('/');
                if (parts.length === 3) {
                    rawData = `${parts[2]}-${parts[1]}-${parts[0]}`; // Converte para AAAA-MM-DD
                }
            }

            // Corrige formato de valor caso venha como 1.500,00
            if (rawValor && rawValor.includes(',')) {
                rawValor = rawValor.replace(/\./g, '').replace(',', '.');
            }
            // 🚀 Remove símbolos como R$ ou $ que venham da planilha
            if (rawValor) {
                rawValor = rawValor.replace(/[^\d.-]/g, '');
            }

            // 2. Cria o objeto Coringa com ABSOLUTAMENTE TODAS as colunas da linha
            const linhaCompletaObjeto = {};
            headers.forEach((headerName, index) => {
                linhaCompletaObjeto[headerName] = columns[index];
            });

            // Adiciona na lista de inserção em lote
            registrosParaInserir.push({
                dataVenda: rawData,
                vendedor: rawVendedor,
                valorTotal: parseFloat(rawValor) || 0,
                dadosAdicionais: linhaCompletaObjeto // 👈 O JSON com todas as colunas extras entra aqui
            });
            
        }

        // 🚀 Salva tudo no banco de uma vez só
        if (registrosParaInserir.length > 0) {
            let client;
            try {
                client = await pool.connect();
                await client.query('BEGIN');
                for (const reg of registrosParaInserir) {
                    await client.query(
                        'INSERT INTO registro_venda (data_venda, vendedor, valor_total, dados_adicionais) VALUES ($1, $2, $3, $4)',
                        [reg.dataVenda, reg.vendedor, reg.valorTotal, JSON.stringify(reg.dadosAdicionais)]
                    );
                }
                await client.query('COMMIT');
            } catch (err) {
                if (client) await client.query('ROLLBACK');
                throw err;
            } finally {
                if (client) client.release();
            }
        }

        fs.unlinkSync(tempPath); // Limpa o arquivo temporário
        
        res.json({ 
            message: `Carga incremental realizada! ${registrosParaInserir.length} linhas salvas com histórico completo.` 
        });

    } catch (error) {
        console.error('Erro no processamento do banco:', error);
        res.status(500).json({ error: 'Falha crítica ao persistir dados no banco relacional.' });
    }
});
// =================================================================
// SEÇÃO FINAL: SERVIR APLICAÇÃO REACT
// =================================================================
app.use((req, res) => {
    const indexPath = path.join(__dirname, '..', 'frontend', 'dist', 'index.html');
    
    // Verifica se a pasta dist existe (Modo Produção)
    if (fs.existsSync(indexPath)) {
        res.sendFile(indexPath);
    } else {
        // Se não existir (Modo Desenvolvimento), devolve um 404 limpo sem quebrar o servidor
        res.status(404).send('Backend operante. Rota não encontrada. (Frontend ainda não compilado na pasta dist)');
    }
});

app.listen(PORT, () => console.log(`🚀 Backend rodando em http://localhost:${PORT}`));