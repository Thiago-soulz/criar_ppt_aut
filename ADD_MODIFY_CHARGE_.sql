ALTER TABLE produto
ADD data_cadastro DATE NOT NULL;

ALTER TABLE maquina
ADD patrimonio VARCHAR(30) NOT NULL;

ALTER TABLE operador
ADD data_admissao DATE NOT NULL;

ALTER TABLE ordem_producao
ADD observacao VARCHAR(200);

ALTER TABLE status_ordem
MODIFY descricao VARCHAR(80) NOT NULL;

ALTER TABLE produto
MODIFY descricao VARCHAR(300);

ALTER TABLE ordem_producao
MODIFY quantidade_produzida INT NOT NULL;

ALTER TABLE turno
MODIFY descricao VARCHAR(60) NOT NULL;

ALTER TABLE ordem_producao
CHANGE quantidade_planejada qtd_planejada INT NOT NULL;

ALTER TABLE operador
CHANGE matricula registro_funcional VARCHAR(30) NOT NULL;

ALTER TABLE maquina
CHANGE nome descricao VARCHAR(100) NOT NULL;

ALTER TABLE produto
CHANGE codigo codigo_interno VARCHAR(50) NOT NULL;