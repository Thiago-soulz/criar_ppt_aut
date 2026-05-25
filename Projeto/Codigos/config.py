import os
import sys

# Detecta se está rodando como .exe ou como script Python
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(
        os.path.dirname(os.path.abspath(__file__))
    )

# Caminhos de arquivos
TEMPLATE_PATH = os.path.join(BASE_DIR, "template", "template.pptx")
CONFIG_PATH = os.path.join(BASE_DIR, "crop_config.json")
OUTPUT_PATH = os.path.join(BASE_DIR, "Analise_Pallets_Final.pptx")
PNG_DIR = os.path.join(BASE_DIR, "Codigos", "PNG")

# Zoom do preview — 1x para coordenadas baterem com o PDF
PREVIEW_ZOOM = 1

# Zoom da conversão final do PDF para imagem
PDF_ZOOM = 2

# ── Configurações modo PÁGINA INTEIRA ─────────────────────────────────────────
PAGINA_GRID_X_INICIO    = 0.30
PAGINA_GRID_Y_INICIO    = 1.50
PAGINA_GRID_MAX_COLUNAS = 5
PAGINA_ESPACAMENTO_X    = 0.10
PAGINA_ESPACAMENTO_Y    = 0.10
PAGINA_IMAGEM_LARGURA   = 2.1
PAGINA_IMAGEM_ALTURA    = 2.9

# ── Configurações modo RECORTE ────────────────────────────────────────────────
RECORTE_GRID_X_INICIO    = 0.30
RECORTE_GRID_Y_INICIO    = 1.50
RECORTE_GRID_MAX_COLUNAS = 3
RECORTE_ESPACAMENTO_X    = 0.15
RECORTE_ESPACAMENTO_Y    = 0.30
RECORTE_IMAGEM_LARGURA   = 3.6
RECORTE_IMAGEM_ALTURA    = 3.7

# ── Estilo do título no slide ─────────────────────────────────────────────────
TITULO_COR_FUNDO     = (145, 185, 225)
TITULO_COR_TEXTO     = (0, 0, 0)
TITULO_FONTE_TAMANHO = 16
TITULO_X             = 0.55
TITULO_Y             = 1.0
TITULO_LARGURA       = 11.0
TITULO_ALTURA        = 0.45