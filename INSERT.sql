INSERT INTO status_ordem VALUES
(1,'Planejada'),
(2,'Em andamento'),
(3,'Finalizada'),
(4,'Cancelada'),
(5,'Aguardando maquina');

INSERT INTO status_maquina VALUES
(1,'Operador'),
(2,'em manutenção '),
(3,'parada'),
(4,'setup'),
(5,'Desativada');

INSERT INTO funcao_operador VALUES
(1,'Operador CNC'),
(2,'Ajustador'),
(3,'inspetor de qualidade'),
(4,'Operador de montagem'),
(5,'tecnico de producão');


INSERT INTO produto VALUES
(1,'eixo Automotivo','PRD-001','Eixo para transmissão',1,'2026-02-01'),
(2,'Pistão','PRD-002','Pistão para motor',1,'2026-02-02'),
(3,'Engrenagem','PRD-003','Engrenagem de aço',1,'2026-02-03'),
(4,'Disco de Freio','PRD-004','Disco ventilado',1,'2026-02-04'),
(5,'Suporte Metálico','PRD-005','Suporte estrutural',0,'2026-02-05');

INSERT INTO maquina (id_maquina, descricao, setor, patrimonio, status)VALUES
(1,'Torno CNC 01','Usinagem','PAT-001','Operando'),
(2,'Fresa CNC 02','Usinagem','PAT-002','Em manutenção'),
(3,'Prensa Hidráulica','Estamparia','PAT-003','Operando'),
(4,'Linha Montagem A','Montagem','PAT-004','Setup'),
(5,'Forno Térmico','Tratamento Térmico','PAT-005','Parada');

INSERT INTO operador(id_operador,nome,registro_funcional,funcao,data_admissao)VALUES
(1,'Carlos Silva','RF001','Operador CNC','2024-01-10'),
(2,'Mariana Souza','RF002','Ajustador','2023-11-05'),
(3,'Rafael Lima','RF003','Inspetor de Qualidade','2022-08-20'),
(4,'Ana Pereira','RF004','Operador de Montagem','2024-03-01'),
(5,'João Mendes','RF005','Técnico de Produção','2021-06-15');

INSERT INTO turno VALUES
(1,'Manhã','06:00:00','14:00:00'),
(2,'Tarde','14:00:00','22:00:00'),
(3,'Noite','22:00:00','06:00:00'),
(4,'Administrativo','08:00:00','17:00:00'),
(5,'Extra','18:00:00','22:00:00');

INSERT INTO ordem_producao VALUES
(1,'2026-03-01','2026-03-02',500,480,'Finalizada',1,1,1,1,'Produção concluída'),
(2,'2026-03-03',NULL,300,0,'Em andamento',2,2,2,2,'Em produção'),
(3,'2026-03-04',NULL,200,0,'Aguardando máquina',3,5,5,3,'Aguardando liberação'),
(4,'2026-03-05','2026-03-06',150,150,'Finalizada',4,3,4,1,'Produção normal'),
(5,'2026-03-07',NULL,400,0,'Planejada',1,4,1,5,'Ordem planejada');
