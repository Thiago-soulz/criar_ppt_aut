import json
import os

from config import CONFIG_PATH


def salvar(x1, y1, x2, y2):
    """Salva as coordenadas do recorte em JSON."""

    dados = {"x1": x1, "y1": y1, "x2": x2, "y2": y2}

    with open(CONFIG_PATH, "w") as f:
        json.dump(dados, f)


def carregar():
    """
    Carrega as coordenadas do recorte do JSON.
    Retorna (x1, y1, x2, y2) ou valores padrão se não existir.
    """

    if not os.path.exists(CONFIG_PATH):
        return 0, 0, 500, 500

    with open(CONFIG_PATH, "r") as f:
        dados = json.load(f)

    return dados["x1"], dados["y1"], dados["x2"], dados["y2"]
