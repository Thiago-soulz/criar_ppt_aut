CREATE DATABASE industria_automotiva;
USE industria_automotiva;


CREATE TABLE status_ordem (
    id_status_ordem INT PRIMARY KEY,
    descricao VARCHAR(50) NOT NULL
);

CREATE TABLE status_maquina (
    id_status_maquina INT PRIMARY KEY,
    descricao VARCHAR(50) NOT NULL
);

CREATE TABLE funcao_operador (
    id_funcao INT PRIMARY KEY,
    descricao VARCHAR(50) NOT NULL
);


CREATE TABLE produto (
    id_produto INT PRIMARY KEY,
    nome VARCHAR(100) NOT NULL,
    codigo VARCHAR(50) NOT NULL,
    descricao VARCHAR(200),
    ativo BOOLEAN NOT NULL
);

CREATE TABLE maquina (
    id_maquina INT PRIMARY KEY,
    nome VARCHAR(100) NOT NULL,
    setor VARCHAR(50) NOT NULL,
    status VARCHAR(30) NOT NULL
);

CREATE TABLE operador (
    id_operador INT PRIMARY KEY,
    nome VARCHAR(100) NOT NULL,
    matricula VARCHAR(30) NOT NULL,
    funcao VARCHAR(50) NOT NULL
);
CREATE TABLE turno (
    id_turno INT PRIMARY KEY,
    descricao VARCHAR(50) NOT NULL,
    hora_inicio TIME NOT NULL,
    hora_fim TIME NOT NULL
);
CREATE TABLE ordem_producao (
    id_ordem INT PRIMARY KEY,
    data_inicio DATE NOT NULL,
    data_fim DATE,
    quantidade_planejada INT NOT NULL,
    quantidade_produzida INT,
    status VARCHAR(30) NOT NULL,
    id_produto INT NOT NULL,
    id_maquina INT NOT NULL,
    id_operador INT NOT NULL,
    id_turno INT NOT NULL,

    FOREIGN KEY (id_produto) REFERENCES produto(id_produto),
    FOREIGN KEY (id_maquina) REFERENCES maquina(id_maquina),
    FOREIGN KEY (id_operador) REFERENCES operador(id_operador),
    FOREIGN KEY (id_turno) REFERENCES turno(id_turno)
);
