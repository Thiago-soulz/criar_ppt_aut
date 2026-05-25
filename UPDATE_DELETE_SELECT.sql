UPDATE maquina
SET status = 'Operando'
WHERE descricao = 'Fresa CNC 02';

UPDATE ordem_producao
SET quantidade_produzida = 280
WHERE id_ordem = 2;

UPDATE ordem_producao
SET status = 'em andamento'
WHERE id_ordem = 5;

UPDATE operador
SET funcao = 'Técnico de Produção'
WHERE nome = 'Ana Pereira';

DELETE FROM produto
WHERE nome = 'Suporte Metálico';

DELETE FROM ordem_producao
WHERE id_ordem = 3;

DELETE FROM operador
WHERE registro_funcional = 'RF003';

SELECT * FROM produto 
WHERE nome LIKE 'D%';

SELECT * FROM operador 
WHERE nome LIKE '%Silva%';

SELECT * FROM maquina
WHERE setor LIKE '%agem%';

SELECT * FROM ordem_producao 
WHERE observacao LIKE '%produção%';

SELECT * FROM maquina
WHERE descricao LIKE '%co';

SELECT * FROM ordem_producao
WHERE qtd_planejada > 300;

SELECT * FROM produto 
WHERE ativo = 0;

SELECT * FROM operador 
Where data_admissao < '2023-01-01';

SELECT * FROM ordem_producao
WHERE data_inicio BETWEEN '2026-03-01' AND '2026-03-05';


SELECT COUNT(*) AS total_ordens
FROM ordem_producao;

SELECT SUM(qtd_planejada) AS soma_planejada
FROM ordem_producao;
 
SELECT MIN(qtd_planejada) AS menor_quantidade
FROM ordem_producao;

SELECT AVG(quantidade_produzida) AS media_producao
FROM ordem_producao;

SELECT COUNT(*) AS ordens_finalizadas
FROM ordem_producao
where status = 'Finalizada';




SELECT
    op.id_ordem,
    p.nome AS produto,
    op.status
FROM ordem_producao op
INNER JOIN produto p
ON op.id_produto = p.id_produto;


SELECT
    op.id_ordem,
    o.nome AS operador,
    op.data_inicio
FROM ordem_producao op
INNER JOIN operador o
ON op.id_operador = o.id_operador;



SELECT op.*
FROM ordem_producao op
INNER JOIN maquina m
ON op.id_maquina = m.id_maquina
WHERE m.setor = 'Usinagem';


SELECT
    op.id_ordem,
    t.descricao AS turno,
    op.qtd_planejada
FROM ordem_producao op
INNER JOIN turno t
ON op.id_turno = t.id_turno
WHERE t.descricao = 'Manhã';



SELECT op.*
FROM ordem_producao op
INNER JOIN maquina m
ON op.id_maquina = m.id_maquina
WHERE m.status = 'Em manutenção';



SELECT COUNT(*) AS total_finalizadas
FROM ordem_producao
WHERE status = 'Finalizada';


SELECT
    op.id_ordem,
    p.nome AS produto,
    o.nome AS operador,
    m.descricao AS maquina
FROM ordem_producao op
INNER JOIN produto p
ON op.id_produto = p.id_produto
INNER JOIN operador o
ON op.id_operador = o.id_operador
INNER JOIN maquina m
ON op.id_maquina = m.id_maquina
ORDER BY p.nome;


SELECT
    status,
    COUNT(*) AS quantidade
FROM ordem_producao
GROUP BY status;



SELECT
    p.nome,
    SUM(op.qtd_planejada) AS total_planejado
FROM ordem_producao op
INNER JOIN produto p
ON op.id_produto = p.id_produto
GROUP BY p.nome;



SELECT
    p.nome AS produto,
    o.nome AS operador,
    m.descricao AS maquina,
    t.descricao AS turno
FROM ordem_producao op
INNER JOIN produto p
ON op.id_produto = p.id_produto
INNER JOIN operador o
ON op.id_operador = o.id_operador
INNER JOIN maquina m
ON op.id_maquina = m.id_maquina
INNER JOIN turno t
ON op.id_turno = t.id_turno
WHERE op.status = 'Finalizada';



SELECT *
FROM operador
WHERE data_admissao BETWEEN '2022-01-01' AND '2024-12-31';


SELECT
    m.descricao AS maquina,
    AVG(op.quantidade_produzida) AS media_producao
FROM ordem_producao op
INNER JOIN maquina m
ON op.id_maquina = m.id_maquina
GROUP BY m.descricao;






