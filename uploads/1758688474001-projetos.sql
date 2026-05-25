-- phpMyAdmin SQL Dump
-- version 5.2.1
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Tempo de geração: 24/09/2025 às 04:49
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
-- Banco de dados: `portal_dca`
--

-- --------------------------------------------------------

--
-- Estrutura para tabela `projetos`
--

CREATE TABLE `projetos` (
  `id` int(11) NOT NULL,
  `nome_projeto` varchar(255) NOT NULL,
  `cliente_id` int(11) DEFAULT NULL,
  `valor_estimado` decimal(15,2) DEFAULT NULL,
  `moeda` varchar(10) DEFAULT 'BRL',
  `data_fechamento_prevista` date DEFAULT NULL,
  `etapa_funil` enum('Prospeccao','Dtc','Poc','Negociacao','Aprovação','Fechado','Perdido','Ganho') NOT NULL DEFAULT 'Prospeccao',
  `tipo_projeto` varchar(50) DEFAULT 'Privado',
  `segmentacao_id` int(11) DEFAULT NULL,
  `vertical_id` int(11) DEFAULT NULL,
  `integrador_id` int(11) DEFAULT NULL,
  `numero_registro_fabricante` varchar(255) DEFAULT NULL,
  `vendedor_id` int(11) DEFAULT NULL,
  `dtc_responsavel_id` int(11) DEFAULT NULL,
  `compras_responsavel_id` int(11) DEFAULT NULL,
  `financeiro_responsavel_id` int(11) DEFAULT NULL,
  `status_proposta_dtc` enum('Pendente','Em Elaboração','Concluída','Revisão Solicitada') DEFAULT NULL,
  `status_analise_credito` enum('Pendente','Concluído','Negado') DEFAULT NULL,
  `justificativa_credito` text DEFAULT NULL,
  `status_compra` enum('Pendente','Em Cotação','Compra Realizada','Material Recebido') DEFAULT NULL,
  `pedido_erp_id` varchar(100) DEFAULT NULL,
  `motivo_perda` varchar(255) DEFAULT NULL,
  `created_at` timestamp NOT NULL DEFAULT current_timestamp(),
  `updated_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Despejando dados para a tabela `projetos`
--

INSERT INTO `projetos` (`id`, `nome_projeto`, `cliente_id`, `valor_estimado`, `moeda`, `data_fechamento_prevista`, `etapa_funil`, `tipo_projeto`, `segmentacao_id`, `vertical_id`, `integrador_id`, `numero_registro_fabricante`, `vendedor_id`, `dtc_responsavel_id`, `compras_responsavel_id`, `financeiro_responsavel_id`, `status_proposta_dtc`, `status_analise_credito`, `justificativa_credito`, `status_compra`, `pedido_erp_id`, `motivo_perda`, `created_at`, `updated_at`) VALUES
(1, 'Laserway', 2, 100000.00, 'BRL', '2025-09-23', '', 'Privado', NULL, NULL, NULL, NULL, 3, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-18 00:29:28', '2025-09-20 03:03:44'),
(2, 'Teste 1', 1, 100000.00, 'BRL', '2025-10-02', '', 'Privado', NULL, NULL, NULL, NULL, 3, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-18 01:15:26', '2025-09-18 02:04:42'),
(3, 'Test 2', 4, 1.00, 'BRL', '2025-09-26', '', 'Privado', NULL, NULL, NULL, NULL, 27, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-18 01:29:58', '2025-09-20 02:58:41'),
(4, 'Teste 3', 3, 5.00, 'BRL', '2025-10-03', '', 'Privado', NULL, NULL, NULL, NULL, 31, 27, NULL, NULL, 'Concluída', NULL, NULL, NULL, NULL, NULL, '2025-09-18 01:55:38', '2025-09-20 03:03:28'),
(5, 'test 5', 3, 14.00, 'BRL', '2025-09-19', '', 'Privado', NULL, NULL, NULL, NULL, 3, 27, NULL, NULL, 'Concluída', NULL, NULL, NULL, NULL, NULL, '2025-09-18 01:56:02', '2025-09-20 02:58:00'),
(6, 'Teste 5', 1, 50.00, 'BRL', '2025-09-20', '', 'Privado', NULL, NULL, NULL, NULL, 35, NULL, NULL, NULL, 'Concluída', NULL, NULL, NULL, NULL, NULL, '2025-09-18 01:59:37', '2025-09-20 03:09:25'),
(7, 'TESTE 6', 1, 10.00, 'BRL', '2025-09-13', '', 'Privado', NULL, NULL, NULL, NULL, 27, 3, NULL, NULL, 'Revisão Solicitada', NULL, NULL, NULL, NULL, NULL, '2025-09-18 02:38:34', '2025-09-20 02:04:13'),
(8, 'TESTE 8', 1, 500000.00, 'BRL', '2025-09-05', '', 'Privado', NULL, NULL, NULL, NULL, 3, 27, NULL, NULL, 'Concluída', NULL, NULL, NULL, NULL, NULL, '2025-09-18 03:43:39', '2025-09-20 02:59:12'),
(9, 'PROJETO DO GIGIO', 3, 1.00, 'BRL', '2025-10-10', '', 'Privado', NULL, NULL, NULL, NULL, 27, 34, NULL, NULL, 'Concluída', NULL, NULL, NULL, NULL, NULL, '2025-09-18 11:32:31', '2025-09-20 02:57:53'),
(10, 'JBS', 4, 1.00, 'BRL', '2025-09-19', '', 'Privado', NULL, NULL, NULL, NULL, 36, 34, NULL, NULL, 'Concluída', NULL, NULL, NULL, NULL, 'Teste', '2025-09-18 11:41:01', '2025-09-22 01:35:15'),
(11, 'ATA FIESC 2026', 3, 1.00, 'BRL', '2025-09-25', '', 'Privado', NULL, NULL, NULL, NULL, 36, 27, NULL, NULL, 'Pendente', NULL, NULL, NULL, NULL, NULL, '2025-09-18 11:49:40', '2025-09-21 00:11:37'),
(12, 'dasdas', 1, 1.00, 'BRL', '2025-09-04', '', 'Privado', NULL, NULL, NULL, NULL, 27, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-20 03:20:28', '2025-09-20 03:43:31'),
(13, 'TESTE21', 1, 1.00, 'BRL', '2025-09-06', '', 'Privado', NULL, NULL, NULL, NULL, 36, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-20 03:21:39', '2025-09-20 03:48:39'),
(14, 'teste1', 1, 1.00, 'BRL', '2025-09-13', '', 'Privado', NULL, NULL, NULL, NULL, 27, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-20 03:31:00', '2025-09-20 03:31:00'),
(15, '51', 1, 4.00, 'BRL', NULL, '', 'Privado', NULL, NULL, NULL, NULL, 31, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-20 03:33:28', '2025-09-20 03:33:57'),
(16, 'dasdasd', 3, 3.00, 'BRL', NULL, '', 'Privado', NULL, NULL, NULL, NULL, 27, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-20 03:47:47', '2025-09-20 04:41:45'),
(17, 'rtearaas', 1, 6.00, 'BRL', '2025-09-29', '', 'Privado', NULL, NULL, NULL, NULL, 34, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-20 04:42:01', '2025-09-20 23:46:24'),
(18, '1', 3, 1.00, 'BRL', '2025-09-10', '', 'Privado', NULL, NULL, NULL, NULL, 27, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-20 11:14:46', '2025-09-20 23:52:12'),
(19, 'TESTE22', 2, 1.00, 'BRL', '2025-09-25', '', 'Privado', NULL, NULL, NULL, NULL, 31, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-20 23:52:06', '2025-09-21 00:01:28'),
(20, 'TESTE2', 3, 32.00, 'BRL', '2025-09-10', '', 'Privado', NULL, NULL, NULL, NULL, 27, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-21 00:00:57', '2025-09-21 00:02:53'),
(21, 'TESTE2', 1, 3.00, 'BRL', '2025-09-09', '', 'Privado', NULL, NULL, NULL, NULL, 34, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-21 00:02:37', '2025-09-21 00:02:37'),
(22, 'teste 55', 1, 1.00, 'BRL', '2025-09-12', '', 'Privado', NULL, NULL, NULL, NULL, 31, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-21 00:13:49', '2025-09-21 23:22:42'),
(23, 'TESTE99', 3, 3.00, 'BRL', '2025-09-12', '', 'Privado', NULL, NULL, NULL, NULL, 3, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-21 23:27:29', '2025-09-22 00:54:25'),
(24, 'TESTE99', 3, 3.00, 'BRL', '2025-09-12', '', 'Privado', NULL, NULL, NULL, NULL, 3, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-21 23:27:30', '2025-09-22 00:45:55'),
(25, 'TESTE99', 3, 3.00, 'BRL', '2025-09-12', '', 'Privado', NULL, NULL, NULL, NULL, 3, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-21 23:27:30', '2025-09-22 00:50:16'),
(26, 'TESTE99', 3, 3.00, 'BRL', '2025-09-12', '', 'Privado', NULL, NULL, NULL, NULL, 3, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-21 23:27:31', '2025-09-22 01:35:14'),
(27, 'TESTE99', 3, 3.00, 'BRL', '2025-09-12', '', 'Privado', NULL, NULL, NULL, NULL, 3, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-21 23:27:31', '2025-09-22 00:54:33'),
(28, 'TESTE100', 3, 1.00, 'BRL', '2025-09-17', '', 'Privado', NULL, NULL, NULL, NULL, 27, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-21 23:28:54', '2025-09-21 23:28:54'),
(29, 'tratas', 1, 1.00, 'BRL', '2025-10-03', '', 'Privado', NULL, NULL, NULL, NULL, 31, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-21 23:31:52', '2025-09-22 01:23:58'),
(30, 'teste5516', 1, 1.00, 'BRL', '2025-09-05', '', 'Privado', NULL, NULL, NULL, NULL, 34, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-22 00:36:49', '2025-09-22 01:35:15'),
(31, 'dasda', 3, 4.00, 'BRL', '2025-09-04', '', 'Privado', NULL, NULL, NULL, NULL, 27, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-22 00:58:07', '2025-09-22 01:35:14'),
(32, 'dasdas', 2, 1.00, 'BRL', '2025-10-01', '', 'Privado', NULL, NULL, NULL, NULL, 27, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-22 01:04:00', '2025-09-22 01:35:14'),
(33, 'asdasd', 1, 1.00, 'BRL', '2025-10-10', '', 'Privado', NULL, NULL, NULL, NULL, 31, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-22 01:24:16', '2025-09-22 01:42:26'),
(34, 'dsd', 3, 1.00, 'BRL', '2025-09-25', '', 'Privado', NULL, NULL, NULL, NULL, 27, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-22 01:42:40', '2025-09-22 03:06:11'),
(35, '12', 1, 1.00, 'BRL', '2025-10-09', '', 'Privado', NULL, NULL, NULL, NULL, 27, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-22 02:52:33', '2025-09-22 23:00:00'),
(36, 'teste62', 3, 0.00, 'BRL', '1899-11-30', 'Negociacao', 'Privado', NULL, NULL, NULL, NULL, 3, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-22 03:33:34', '2025-09-22 23:30:31'),
(37, 'DEFESA CIVIL', 1, 1000000.00, 'BRL', '2025-10-10', 'Dtc', 'Público', 3, 3, 1, '56169511', 3, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, '2025-09-24 02:25:29', '2025-09-24 02:37:09');

--
-- Índices para tabelas despejadas
--

--
-- Índices de tabela `projetos`
--
ALTER TABLE `projetos`
  ADD PRIMARY KEY (`id`),
  ADD KEY `cliente_id` (`cliente_id`),
  ADD KEY `vendedor_id` (`vendedor_id`),
  ADD KEY `dtc_responsavel_id` (`dtc_responsavel_id`),
  ADD KEY `compras_responsavel_id` (`compras_responsavel_id`),
  ADD KEY `financeiro_responsavel_id` (`financeiro_responsavel_id`),
  ADD KEY `fk_projetos_segmentacao` (`segmentacao_id`),
  ADD KEY `fk_projetos_vertical` (`vertical_id`),
  ADD KEY `fk_projetos_integrador` (`integrador_id`);

--
-- AUTO_INCREMENT para tabelas despejadas
--

--
-- AUTO_INCREMENT de tabela `projetos`
--
ALTER TABLE `projetos`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=38;

--
-- Restrições para tabelas despejadas
--

--
-- Restrições para tabelas `projetos`
--
ALTER TABLE `projetos`
  ADD CONSTRAINT `fk_projetos_integrador` FOREIGN KEY (`integrador_id`) REFERENCES `integradores` (`id`) ON DELETE SET NULL ON UPDATE NO ACTION,
  ADD CONSTRAINT `fk_projetos_segmentacao` FOREIGN KEY (`segmentacao_id`) REFERENCES `segmentacoes` (`id`) ON DELETE SET NULL ON UPDATE NO ACTION,
  ADD CONSTRAINT `fk_projetos_vertical` FOREIGN KEY (`vertical_id`) REFERENCES `verticais` (`id`) ON DELETE SET NULL ON UPDATE NO ACTION,
  ADD CONSTRAINT `projetos_ibfk_1` FOREIGN KEY (`cliente_id`) REFERENCES `clientes` (`id`),
  ADD CONSTRAINT `projetos_ibfk_2` FOREIGN KEY (`vendedor_id`) REFERENCES `funcionarios` (`id`),
  ADD CONSTRAINT `projetos_ibfk_3` FOREIGN KEY (`dtc_responsavel_id`) REFERENCES `funcionarios` (`id`),
  ADD CONSTRAINT `projetos_ibfk_4` FOREIGN KEY (`compras_responsavel_id`) REFERENCES `funcionarios` (`id`),
  ADD CONSTRAINT `projetos_ibfk_5` FOREIGN KEY (`financeiro_responsavel_id`) REFERENCES `funcionarios` (`id`);
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
