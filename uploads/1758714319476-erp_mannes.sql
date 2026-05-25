-- phpMyAdmin SQL Dump
-- version 5.2.1
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Tempo de geração: 24/09/2025 às 08:22
-- Versão do servidor: 10.4.32-MariaDB
-- Versão do PHP: 8.2.12

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Banco de dados: `erp_mannes`
--

-- --------------------------------------------------------

--
-- Estrutura para tabela `estoque_produtos`
--

CREATE TABLE `estoque_produtos` (
  `id` int(11) NOT NULL,
  `codigo_produto` varchar(50) NOT NULL,
  `descricao_produto` varchar(255) DEFAULT NULL,
  `quantidade_disponivel` int(11) DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Despejando dados para a tabela `estoque_produtos`
--

INSERT INTO `estoque_produtos` (`id`, `codigo_produto`, `descricao_produto`, `quantidade_disponivel`) VALUES
(1, 'DHCALSSF-022-0', 'CABO ÓPTICO ASU 12F S/S COG', 1500),
(2, 'DHCALPSF-015-0', 'ALÇA P/ FIO ÓPTICO FIG. 8 (01F)', 200),
(3, 'DHCALASF-015-0', 'LAÇO P/ FIO ÓPTICO (01F)', 100),
(4, 'DHCFSCAF-003-0', 'CONECTOR FAST SC/APC', 100),
(5, 'DHCACEO0-001-0', 'CAIXA DE EMENDA ÓPTICA 24F', 10),
(6, 'DHCADSCF-001-0', 'ADAPTADOR ÓPTICO SC/APC SM SIMPLEX', 0);

-- --------------------------------------------------------

--
-- Estrutura para tabela `pedidos_cabecalho`
--

CREATE TABLE `pedidos_cabecalho` (
  `id` int(11) NOT NULL,
  `numero_pedido` varchar(50) NOT NULL,
  `nome_cliente_erp` varchar(255) DEFAULT NULL,
  `cnpj_cliente_erp` varchar(20) DEFAULT NULL,
  `data_pedido` date DEFAULT NULL,
  `valor_total_pedido` decimal(15,2) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Despejando dados para a tabela `pedidos_cabecalho`
--

INSERT INTO `pedidos_cabecalho` (`id`, `numero_pedido`, `nome_cliente_erp`, `cnpj_cliente_erp`, `data_pedido`, `valor_total_pedido`) VALUES
(1, 'PED-799469', 'CARMO ENERGY S.A.', '41.955.491/0002-92', '2025-01-31', 8142.18);

-- --------------------------------------------------------

--
-- Estrutura para tabela `pedidos_itens`
--

CREATE TABLE `pedidos_itens` (
  `id` int(11) NOT NULL,
  `pedido_id` int(11) NOT NULL,
  `codigo_produto` varchar(50) DEFAULT NULL,
  `descricao_produto` varchar(255) DEFAULT NULL,
  `quantidade` int(11) DEFAULT NULL,
  `valor_unitario` decimal(15,2) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Despejando dados para a tabela `pedidos_itens`
--

INSERT INTO `pedidos_itens` (`id`, `pedido_id`, `codigo_produto`, `descricao_produto`, `quantidade`, `valor_unitario`) VALUES
(1, 1, 'DHCALSSF-022-0', 'CABO ÓPTICO ASU 12F S/S COG', 2000, 3.09),
(2, 1, 'DHCALPSF-015-0', 'ALÇA P/ FIO ÓPTICO FIG. 8 (01F)', 150, 3.08),
(3, 1, 'DHCALASF-015-0', 'LAÇO P/ FIO ÓPTICO (01F)', 150, 3.08),
(4, 1, 'DHCFSCAF-003-0', 'CONECTOR FAST SC/APC', 24, 6.42),
(5, 1, 'DHCACEO0-001-0', 'CAIXA DE EMENDA ÓPTICA 24F', 6, 139.95),
(6, 1, 'DHCADSCF-001-0', 'ADAPTADOR ÓPTICO SC/APC SM SIMPLEX', 24, 1.85);

--
-- Índices para tabelas despejadas
--

--
-- Índices de tabela `estoque_produtos`
--
ALTER TABLE `estoque_produtos`
  ADD PRIMARY KEY (`id`),
  ADD UNIQUE KEY `codigo_produto` (`codigo_produto`);

--
-- Índices de tabela `pedidos_cabecalho`
--
ALTER TABLE `pedidos_cabecalho`
  ADD PRIMARY KEY (`id`),
  ADD UNIQUE KEY `numero_pedido` (`numero_pedido`);

--
-- Índices de tabela `pedidos_itens`
--
ALTER TABLE `pedidos_itens`
  ADD PRIMARY KEY (`id`),
  ADD KEY `pedido_id` (`pedido_id`);

--
-- AUTO_INCREMENT para tabelas despejadas
--

--
-- AUTO_INCREMENT de tabela `estoque_produtos`
--
ALTER TABLE `estoque_produtos`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=7;

--
-- AUTO_INCREMENT de tabela `pedidos_cabecalho`
--
ALTER TABLE `pedidos_cabecalho`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=2;

--
-- AUTO_INCREMENT de tabela `pedidos_itens`
--
ALTER TABLE `pedidos_itens`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=7;

--
-- Restrições para tabelas despejadas
--

--
-- Restrições para tabelas `pedidos_itens`
--
ALTER TABLE `pedidos_itens`
  ADD CONSTRAINT `pedidos_itens_ibfk_1` FOREIGN KEY (`pedido_id`) REFERENCES `pedidos_cabecalho` (`id`) ON DELETE CASCADE;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
