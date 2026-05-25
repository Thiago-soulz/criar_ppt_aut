import os
import sys
import tempfile

import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.font as tkfont

from PIL import Image, ImageTk

import crop_config
import excel_reader
import pdf_converter
import pptx_builder
from config import PREVIEW_ZOOM, PNG_DIR


# ── Estado da aplicação ────────────────────────────────────────────────────────

class Estado:
    excel_file = ""
    pdf_folder = ""
    crop_x1, crop_y1, crop_x2, crop_y2 = crop_config.carregar()
    modo = "pagina"
    cfg_pagina = {}
    cfg_recorte = {}
    canvas = None
    janela_crop = None
    preview_scale = 1.0
    start_x = 0
    start_y = 0
    rect = None


estado = Estado()

# ── Cores e estilos ────────────────────────────────────────────────────────────

COR_BG          = "#F5F6FA"
COR_SIDEBAR     = "#FFFFFF"
COR_PRIMARIA    = "#1A3C6E"
COR_ACENTO      = "#E30613"
COR_CARD        = "#FFFFFF"
COR_CARD_SEL    = "#EBF2FF"
COR_BORDA       = "#E2E8F0"
COR_BORDA_SEL   = "#1A3C6E"
COR_TEXTO       = "#1E293B"
COR_SUBTEXTO    = "#64748B"
COR_BTN_VERDE   = "#16A34A"
COR_BTN_VERDE_H = "#15803D"
COR_STATUS      = "#2563EB"

# Pasta onde está o logo (vem do config.py que já resolve o caminho dinamicamente)
_LOGO_DIR = PNG_DIR


def _encontrar_logo():
    nomes = ["Bosch-logo.png"]
    for nome in nomes:
        caminho = os.path.join(_LOGO_DIR, nome)
        if os.path.exists(caminho):
            return caminho
    try:
        print("[DEBUG] Arquivos na pasta PNG:", os.listdir(_LOGO_DIR))
    except Exception as e:
        print("[DEBUG] Pasta nao encontrada:", e)
    return None


def _carregar_logo(altura=40):
    """Carrega o PNG da Bosch (já transparente) mantendo proporção."""
    caminho = _encontrar_logo()
    if not caminho:
        print("[DEBUG] Logo nao encontrado!")
        return None
    try:
        print("[DEBUG] Carregando:", caminho)
        img = Image.open(caminho).convert("RGBA")
        ratio = altura / img.height
        largura = int(img.width * ratio)
        img = img.resize((largura, altura), Image.LANCZOS)
        return ImageTk.PhotoImage(img)
    except Exception as e:
        print("[DEBUG] Erro ao carregar logo:", e)
        return None


def _carregar_config_padrao():
    import config as cfg
    estado.cfg_pagina = {
        "x_inicio":    tk.DoubleVar(value=cfg.PAGINA_GRID_X_INICIO),
        "y_inicio":    tk.DoubleVar(value=cfg.PAGINA_GRID_Y_INICIO),
        "colunas":     tk.IntVar(value=cfg.PAGINA_GRID_MAX_COLUNAS),
        "esp_x":       tk.DoubleVar(value=cfg.PAGINA_ESPACAMENTO_X),
        "esp_y":       tk.DoubleVar(value=cfg.PAGINA_ESPACAMENTO_Y),
        "largura_max": tk.DoubleVar(value=cfg.PAGINA_IMAGEM_LARGURA),
        "altura_max":  tk.DoubleVar(value=cfg.PAGINA_IMAGEM_ALTURA),
    }
    estado.cfg_recorte = {
        "x_inicio":    tk.DoubleVar(value=cfg.RECORTE_GRID_X_INICIO),
        "y_inicio":    tk.DoubleVar(value=cfg.RECORTE_GRID_Y_INICIO),
        "colunas":     tk.IntVar(value=cfg.RECORTE_GRID_MAX_COLUNAS),
        "esp_x":       tk.DoubleVar(value=cfg.RECORTE_ESPACAMENTO_X),
        "esp_y":       tk.DoubleVar(value=cfg.RECORTE_ESPACAMENTO_Y),
        "largura_max": tk.DoubleVar(value=cfg.RECORTE_IMAGEM_LARGURA),
        "altura_max":  tk.DoubleVar(value=cfg.RECORTE_IMAGEM_ALTURA),
    }


# ── Janela de configuração ────────────────────────────────────────────────────

def abrir_configuracoes():
    cfg = estado.cfg_pagina if estado.modo == "pagina" else estado.cfg_recorte
    titulo_modo = "Página Inteira" if estado.modo == "pagina" else "Recorte"

    janela_cfg = tk.Toplevel()
    janela_cfg.title(f"Configurações — {titulo_modo}")
    janela_cfg.geometry("420x380")
    janela_cfg.resizable(False, False)
    janela_cfg.configure(bg=COR_BG)
    janela_cfg.grab_set()

    tk.Label(janela_cfg, text=f"Configurações — {titulo_modo}",
             font=("Segoe UI", 13, "bold"), bg=COR_BG, fg=COR_TEXTO).pack(pady=16)

    frame = tk.Frame(janela_cfg, bg=COR_BG, padx=24)
    frame.pack(fill=tk.BOTH, expand=True)

    campos = [
        ("Número de colunas",         "colunas",     "int", "Imagens por linha"),
        ("Tamanho da imagem (cm)",     "largura_max", "cm",  "Largura — altura ajusta proporcional"),
        ("Espaço entre colunas (cm)",  "esp_x",       "cm",  "Espaço horizontal"),
        ("Espaço entre lines (cm)",   "esp_y",       "cm",  "Espaço vertical"),
        ("Posição inicial X (cm)",     "x_inicio",    "cm",  "Distância da borda esquerda"),
        ("Posição inicial Y (cm)",     "y_inicio",    "cm",  "Distância do topo"),
    ]

    entradas = {}

    for label_text, chave, tipo, dica in campos:
        row = tk.Frame(frame, bg=COR_BG)
        row.pack(fill=tk.X, pady=5)

        tk.Label(row, text=label_text, width=26, anchor="w",
                 font=("Segoe UI", 9), bg=COR_BG, fg=COR_TEXTO).pack(side=tk.LEFT)

        if tipo == "int":
            var = tk.IntVar(value=cfg[chave].get())
        else:
            var = tk.DoubleVar(value=round(cfg[chave].get() * 2.54, 2))

        entry = tk.Entry(row, textvariable=var, width=7,
                         font=("Segoe UI", 9), relief="solid", bd=1)
        entry.pack(side=tk.LEFT, padx=6)

        tk.Label(row, text=dica, fg=COR_SUBTEXTO,
                 font=("Segoe UI", 8), bg=COR_BG).pack(side=tk.LEFT)

        entradas[chave] = (var, tipo)

    def restaurar_padrao():
        _carregar_config_padrao()
        cfg2 = estado.cfg_pagina if estado.modo == "pagina" else estado.cfg_recorte
        for chave, (var, tipo) in entradas.items():
            if tipo == "int":
                var.set(cfg2[chave].get())
            else:
                var.set(round(cfg2[chave].get() * 2.54, 2))

    def salvar():
        try:
            for chave, (var, tipo) in entradas.items():
                if tipo == "int":
                    cfg[chave].set(int(var.get()))
                else:
                    cfg[chave].set(round(var.get() / 2.54, 4))
            largura = cfg["largura_max"].get()
            cfg["altura_max"].set(round(largura * 1.414, 4))
            janela_cfg.destroy()
        except ValueError:
            messagebox.showerror("Erro", "Verifique os valores digitados.")

    frame_btns = tk.Frame(janela_cfg, bg=COR_BG)
    frame_btns.pack(pady=14)

    tk.Button(frame_btns, text="Restaurar Padrão", width=16,
              font=("Segoe UI", 9), bg=COR_BORDA, fg=COR_TEXTO,
              relief="flat", cursor="hand2",
              command=restaurar_padrao).pack(side=tk.LEFT, padx=6)

    tk.Button(frame_btns, text="Salvar", width=16,
              font=("Segoe UI", 9, "bold"), bg=COR_PRIMARIA, fg="white",
              relief="flat", cursor="hand2",
              command=salvar).pack(side=tk.LEFT, padx=6)


# ── Callbacks do recorte ──────────────────────────────────────────────────────

def on_mouse_down(event):
    estado.start_x = event.x
    estado.start_y = event.y
    if estado.rect:
        estado.canvas.delete(estado.rect)
    estado.rect = estado.canvas.create_rectangle(
        estado.start_x, estado.start_y, estado.start_x, estado.start_y,
        outline="red", width=2)


def on_mouse_move(event):
    estado.canvas.coords(estado.rect, estado.start_x, estado.start_y, event.x, event.y)


def on_mouse_up(event):
    x1 = min(estado.start_x, event.x)
    y1 = min(estado.start_y, event.y)
    x2 = max(estado.start_x, event.x)
    y2 = max(estado.start_y, event.y)

    estado.crop_x1 = int((x1 / estado.preview_scale) / PREVIEW_ZOOM)
    estado.crop_y1 = int((y1 / estado.preview_scale) / PREVIEW_ZOOM)
    estado.crop_x2 = int((x2 / estado.preview_scale) / PREVIEW_ZOOM)
    estado.crop_y2 = int((y2 / estado.preview_scale) / PREVIEW_ZOOM)

    crop_config.salvar(estado.crop_x1, estado.crop_y1, estado.crop_x2, estado.crop_y2)
    label_recorte_info.config(
        text=f"Área definida  ✔  ({estado.crop_x2 - estado.crop_x1}×{estado.crop_y2 - estado.crop_y1} px)")
    messagebox.showinfo("Recorte salvo", "Área de recorte definida com sucesso!")
    estado.janela_crop.destroy()


def definir_recorte():
    arquivo = filedialog.askopenfilename(title="Selecionar PDF", filetypes=[("PDF", "*.pdf")])
    if not arquivo:
        return

    imagem_path = pdf_converter.converter_preview(arquivo)
    estado.janela_crop = tk.Toplevel()
    estado.janela_crop.title("Definir Recorte")
    estado.janela_crop.state("zoomed")

    imagem_original = Image.open(imagem_path)
    largura_tela = estado.janela_crop.winfo_screenwidth() - 100
    altura_tela = estado.janela_crop.winfo_screenheight() - 150
    escala_x = largura_tela / imagem_original.width
    escala_y = altura_tela / imagem_original.height
    estado.preview_scale = min(escala_x, escala_y, 1)

    nova_largura = int(imagem_original.width * estado.preview_scale)
    nova_altura = int(imagem_original.height * estado.preview_scale)
    tk_img = ImageTk.PhotoImage(imagem_original.resize((nova_largura, nova_altura)))

    frame = tk.Frame(estado.janela_crop)
    frame.pack(fill=tk.BOTH, expand=True)
    estado.canvas = tk.Canvas(frame, width=nova_largura, height=nova_altura, bg="gray")
    estado.canvas.pack(expand=True)
    estado.canvas.create_image(0, 0, anchor=tk.NW, image=tk_img)
    estado.canvas.image = tk_img
    estado.canvas.bind("<ButtonPress-1>", on_mouse_down)
    estado.canvas.bind("<B1-Motion>", on_mouse_move)
    estado.canvas.bind("<ButtonRelease-1>", on_mouse_up)


# ── Ações ─────────────────────────────────────────────────────────────────────

def selecionar_modo(modo):
    estado.modo = modo

    for m, btn, card in modo_cards:
        if m == modo:
            card.config(bg=COR_CARD_SEL, highlightbackground=COR_BORDA_SEL, highlightthickness=2)
            for w in card.winfo_children():
                if isinstance(w, (tk.Label, tk.Frame, tk.Canvas)):
                    w.config(bg=COR_CARD_SEL)
                    for child in w.winfo_children():
                        if isinstance(child, (tk.Label, tk.Frame, tk.Canvas)):
                            child.config(bg=COR_CARD_SEL)
        else:
            card.config(bg=COR_CARD, highlightbackground=COR_BORDA, highlightthickness=1)
            for w in card.winfo_children():
                if isinstance(w, (tk.Label, tk.Frame, tk.Canvas)):
                    w.config(bg=COR_CARD)
                    for child in w.winfo_children():
                        if isinstance(child, (tk.Label, tk.Frame, tk.Canvas)):
                            child.config(bg=COR_CARD)

    if modo == "recorte":
        # MODIFICADO: Agora insere após o frame invisível externo da pasta (fora do quadrado branco)
        frame_recorte_info.pack(fill=tk.X, pady=(6, 4), after=frame_pasta_geral)
        janela.geometry("450x640")  # Modificado: Altura reduzida com recorte
    else:
        frame_recorte_info.pack_forget()
        janela.geometry("450x560")  # Modificado: Largura e Altura reduzidas


def selecionar_planilha():
    arquivo = filedialog.askopenfilename(
        title="Selecionar Planilha", filetypes=[("Excel", "*.xlsx *.xls")])
    if arquivo:
        estado.excel_file = arquivo
        label_planilha.config(text=os.path.basename(arquivo), fg=COR_TEXTO)


def selecionar_pasta():
    pasta = filedialog.askdirectory(title="Selecionar Pasta PDFs")
    if pasta:
        estado.pdf_folder = pasta
        label_pasta.config(text=pasta, fg=COR_TEXTO)


def gerar():
    if not estado.excel_file:
        messagebox.showerror("Erro", "Selecione a planilha")
        return
    if not estado.pdf_folder:
        messagebox.showerror("Erro", "Selecione a pasta PDFs")
        return

    try:
        label_status.config(text="⏳  Lendo planilha...", fg=COR_STATUS)
        janela.update()

        pallets = excel_reader.ler_pallets(estado.excel_file)
        pallets_com_imagens = {}
        nao_encontrados = []

        for numero_pallet, numeros_pdf in pallets.items():
            arquivos_pdf = []
            for numero in numeros_pdf:
                pdf_path = pdf_converter.buscar_pdf(numero, estado.pdf_folder)
                if not pdf_path:
                    nao_encontrados.append(f"Pallet {numero_pallet:02d} → PDF {numero:03d}.pdf")
                    continue

                nome = os.path.splitext(os.path.basename(pdf_path))[0]
                imagem_temp = os.path.join(tempfile.gettempdir(), f"{nome}.png")

                label_status.config(text=f"⏳  Convertendo {os.path.basename(pdf_path)}", fg=COR_STATUS)
                janela.update()

                if estado.modo == "pagina":
                    pdf_converter.converter_pagina_inteira(pdf_path, imagem_temp)
                else:
                    estado.crop_x1, estado.crop_y1, estado.crop_x2, estado.crop_y2 = crop_config.carregar()
                    pdf_converter.converter_recorte(
                        pdf_path, imagem_temp,
                        estado.crop_x1, estado.crop_y1,
                        estado.crop_x2, estado.crop_y2)

                arquivos_pdf.append(pdf_path)

            if arquivos_pdf:
                pallets_com_imagens[numero_pallet] = arquivos_pdf

        cfg_atual = estado.cfg_pagina if estado.modo == "pagina" else estado.cfg_recorte
        cfg = {k: v.get() for k, v in cfg_atual.items()}

        def atualizar_status(texto):
            label_status.config(text=f"⏳  {texto}", fg=COR_STATUS)
            janela.update()

        label_status.config(text="⏳  Montando PowerPoint...", fg=COR_STATUS)
        janela.update()

        output = pptx_builder.montar_apresentacao(
            pallets_com_imagens, modo=estado.modo,
            nome_excel=estado.excel_file,
            cfg_override=cfg, callback_status=atualizar_status)

        label_status.config(text="✅  Finalizado com sucesso!", fg="#16A34A")
        janela.update()

        messagebox.showinfo("Sucesso", f"PowerPoint criado:\n\n{output}")

        if nao_encontrados:
            lista = "\n".join(nao_encontrados)
            messagebox.showwarning("Arquivos não encontrados",
                                   f"Os seguintes PDFs não foram encontrados:\n\n{lista}")

    except Exception as erro:
        messagebox.showerror("Erro", str(erro))
        label_status.config(text="❌  Erro durante o processamento", fg=COR_ACENTO)


# ══════════════════════════════════════════════════════════════════════════════
# ── Interface principal ───────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

janela = tk.Tk()
janela.title("Sistema de Pallets")
janela.geometry("540x620")  # Modificado: Inicialmente menor
janela.resizable(False, False)
janela.configure(bg=COR_BG)

_carregar_config_padrao()

# ── Topbar com logo ───────────────────────────────────────────────────────────

topbar = tk.Frame(janela, bg=COR_SIDEBAR, height=54)  # Modificado: altura reduzida de 64 para 54
topbar.pack(side=tk.TOP, fill=tk.X)
topbar.pack_propagate(False)

tk.Frame(topbar, bg=COR_ACENTO, width=4).pack(side=tk.LEFT, fill=tk.Y)

frame_logo_top = tk.Frame(topbar, bg=COR_SIDEBAR, padx=14)
frame_logo_top.pack(side=tk.LEFT, fill=tk.Y)

_tk_logo_top = _carregar_logo(altura=32)  # Modificado: Reduzido de 38 para 32
if _tk_logo_top:
    lbl_logo_top = tk.Label(frame_logo_top, image=_tk_logo_top, bg=COR_SIDEBAR)
    lbl_logo_top.image = _tk_logo_top
    lbl_logo_top.pack(expand=True)
else:
    frame_fallback = tk.Frame(frame_logo_top, bg=COR_SIDEBAR)
    frame_fallback.pack(expand=True)
    tk.Label(frame_fallback, text="⊕", font=("Segoe UI", 16),
             bg=COR_SIDEBAR, fg=COR_ACENTO).pack(side=tk.LEFT)
    tk.Label(frame_fallback, text=" BOSCH", font=("Segoe UI", 14, "bold"),
             bg=COR_SIDEBAR, fg=COR_ACENTO).pack(side=tk.LEFT)

frame_cfg_btn = tk.Frame(topbar, bg=COR_SIDEBAR, padx=12)
frame_cfg_btn.pack(side=tk.RIGHT, fill=tk.Y)

_cfg_img_path = os.path.join(_LOGO_DIR, "config.png")
_tk_cfg_icon = None
try:
    _cfg_img = Image.open(_cfg_img_path).convert("RGBA")
    _cfg_ratio = 20 / _cfg_img.height  # Modificado: Reduzido de 26 para 20
    _cfg_img = _cfg_img.resize((int(_cfg_img.width * _cfg_ratio), 20), Image.LANCZOS)
    _tk_cfg_icon = ImageTk.PhotoImage(_cfg_img)
except Exception:
    pass

if _tk_cfg_icon:
    btn_cfg_top = tk.Button(frame_cfg_btn, image=_tk_cfg_icon,
                            bg=COR_SIDEBAR, relief="flat", cursor="hand2",
                            bd=0, highlightthickness=0,
                            command=abrir_configuracoes)
    btn_cfg_top.image = _tk_cfg_icon
    btn_cfg_top.pack(side=tk.TOP, pady=(6, 1))
else:
    btn_cfg_top = tk.Button(frame_cfg_btn, text="⚙",
                            font=("Segoe UI", 14), bg=COR_SIDEBAR,
                            fg=COR_SUBTEXTO, relief="flat", cursor="hand2",
                            bd=0, highlightthickness=0,
                            command=abrir_configuracoes)
    btn_cfg_top.pack(side=tk.TOP, pady=(6, 1))

tk.Label(frame_cfg_btn, text="Configurar", font=("Segoe UI", 7),
         bg=COR_SIDEBAR, fg=COR_SUBTEXTO, cursor="hand2").pack(side=tk.TOP)

tk.Frame(janela, bg=COR_BORDA, height=1).pack(fill=tk.X)

# ── Conteúdo principal ────────────────────────────────────────────────────────

conteudo = tk.Frame(janela, bg=COR_BG, padx=24, pady=12)  # Modificado: Padding reduzido
conteudo.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

tk.Label(conteudo, text="SISTEMA DE PALLETS",
         font=("Segoe UI", 16, "bold"), bg=COR_BG, fg=COR_TEXTO).pack(anchor="w")  # Modificado: Fonte 20 -> 16

tk.Frame(conteudo, bg=COR_ACENTO, height=3, width=42).pack(anchor="w", pady=(2, 10))

# ── Cards de modo ─────────────────────────────────────────────────────────────

tk.Label(conteudo, text="Modo de geração", font=("Segoe UI", 9, "bold"),
         bg=COR_BG, fg=COR_TEXTO).pack(anchor="w", pady=(0, 6))

frame_cards = tk.Frame(conteudo, bg=COR_BG)
frame_cards.pack(fill=tk.X, pady=(0, 2))

modo_cards = []


def criar_card(parent, modo, canvas_draw, titulo, descricao):
    card = tk.Frame(parent, bg=COR_CARD, cursor="hand2",
                    highlightbackground=COR_BORDA, highlightthickness=1,
                    padx=0, pady=0)
    card.pack(side=tk.LEFT, padx=(0, 10), ipadx=0, ipady=0)
    card.config(width=145, height=135)  # Modificado: Cards menores (170x170 -> 145x135)
    card.pack_propagate(False)

    inner = tk.Frame(card, bg=COR_CARD, padx=10, pady=10)  # Modificado: padding interno menor
    inner.pack(expand=True, fill=tk.BOTH)

    cnv = tk.Canvas(inner, width=32, height=32, bg=COR_CARD,
                    highlightthickness=0, bd=0)  # Modificado: canvas menor
    cnv.pack()
    canvas_draw(cnv, "#000000")

    tk.Label(inner, text=titulo, font=("Segoe UI", 9, "bold"),
             bg=COR_CARD, fg="#000000").pack(pady=(4, 2))
    tk.Label(inner, text=descricao, font=("Segoe UI", 8),
             bg=COR_CARD, fg="#000000", wraplength=125, justify="center").pack()

    def bind_click(widget):
        widget.bind("<Button-1>", lambda e, m=modo: selecionar_modo(m))
        for child in widget.winfo_children():
            bind_click(child)

    bind_click(card)
    return card


def draw_crop(cnv, cor):
    s = 28  # Modificado: Ajustado para o novo tamanho de canvas
    t = 3
    m = 2
    L = 10
    cnv.create_rectangle(m, m, m + L, m + t, fill=cor, outline="")
    cnv.create_rectangle(m, m, m + t, m + L, fill=cor, outline="")
    cnv.create_rectangle(s - L, m, s, m + t, fill=cor, outline="")
    cnv.create_rectangle(s - t, m, s, m + L, fill=cor, outline="")
    cnv.create_rectangle(m, s - t, m + L, s, fill=cor, outline="")
    cnv.create_rectangle(m, s - L, m + t, s, fill=cor, outline="")
    cnv.create_rectangle(s - L, s - t, s, s, fill=cor, outline="")
    cnv.create_rectangle(s - t, s - L, s, s, fill=cor, outline="")


def draw_pagina_png(cnv, cor):
    caminho_png = os.path.join(_LOGO_DIR, "pagina.png")
    try:
        img = Image.open(caminho_png).convert("RGBA")
        img = img.resize((28, 28), Image.LANCZOS)  # Modificado: Reduzido de 36 para 28
        photo = ImageTk.PhotoImage(img)
        cnv._pagina_img = photo
        cnv.create_image(2, 2, anchor="nw", image=photo)
    except Exception:
        cnv.create_rectangle(6, 2, 26, 30, fill=cor, outline="")
        cnv.create_rectangle(6, 2, 18, 8, fill=cor, outline="")
        cnv.create_polygon(18, 2, 26, 8, 18, 8, fill=cor, outline="")


def draw_crop_preto(cnv, cor):
    draw_crop(cnv, "#000000")


card_pagina  = criar_card(frame_cards, "pagina",  draw_pagina_png, "Página Inteira",
                           "Gera a apresentação\ncom todas as páginas.")
card_recorte = criar_card(frame_cards, "recorte", draw_crop_preto, "Recorte",
                           "Gera apenas a área\nselecionada.")

modo_cards = [("pagina", None, card_pagina), ("recorte", None, card_recorte)]

# ── Separador ─────────────────────────────────────────────────────────────────

tk.Frame(conteudo, bg=COR_BORDA, height=1).pack(fill=tk.X, pady=8)

# ── Seleção de arquivos ───────────────────────────────────────────────────────

def _draw_file_icon(parent, bg):
    cnv = tk.Canvas(parent, width=20, height=20, bg=bg,
                    highlightthickness=0, bd=0)
    cnv.create_rectangle(3, 1, 17, 19, fill="#CBD5E1", outline="")
    cnv.create_rectangle(3, 1, 12, 6,  fill="#94A3B8", outline="")
    cnv.create_polygon(12, 1, 17, 6, 12, 6, fill="#94A3B8", outline="")
    cnv.create_rectangle(5, 8,  15, 9, fill="#64748B", outline="")
    cnv.create_rectangle(5, 11, 15, 12, fill="#64748B", outline="")
    cnv.create_rectangle(5, 14, 11, 15, fill="#64748B", outline="")
    return cnv


def _draw_folder_icon(parent, bg):
    cnv = tk.Canvas(parent, width=20, height=20, bg=bg,
                    highlightthickness=0, bd=0)
    cnv.create_rectangle(1, 7, 19, 19, fill="#94A3B8", outline="")
    cnv.create_rectangle(1, 4, 9, 8,  fill="#CBD5E1", outline="")
    return cnv


def arquivo_row(parent, label_titulo, icone_fn, comando, placeholder):
    # Criado um frame invisível geral para podermos organizar com segurança o 'after' do pack
    frame_geral = tk.Frame(parent, bg=COR_BG)
    frame_geral.pack(fill=tk.X, pady=4)

    frame = tk.Frame(frame_geral, bg=COR_CARD,
                     highlightbackground=COR_BORDA, highlightthickness=1,
                     padx=12, pady=8)  # Modificado: Padding reduzido
    frame.pack(fill=tk.X)

    tk.Label(frame, text=label_titulo, font=("Segoe UI", 9, "bold"),
             bg=COR_CARD, fg=COR_TEXTO).pack(anchor="w", pady=(0, 4))

    row = tk.Frame(frame, bg=COR_CARD)
    row.pack(fill=tk.X)

    icone = icone_fn(row, COR_CARD)
    icone.pack(side=tk.LEFT, padx=(0, 6))

    lbl = tk.Label(row, text=placeholder, font=("Segoe UI", 8),
                   bg=COR_CARD, fg=COR_SUBTEXTO, anchor="w")
    lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

    btn = tk.Button(row, text="Selecionar", font=("Segoe UI", 8),
                    bg=COR_BG, fg=COR_PRIMARIA, relief="flat", cursor="hand2",
                    padx=10, pady=2, bd=0,
                    highlightbackground=COR_BORDA, highlightthickness=1,
                    command=comando)
    btn.pack(side=tk.RIGHT)

    return lbl, frame_geral


label_planilha, frame_planilha_geral = arquivo_row(conteudo, "Selecionar Planilha",
                                              _draw_file_icon, selecionar_planilha,
                                              "Nenhuma planilha selecionada")
label_pasta, frame_pasta_geral = arquivo_row(conteudo, "Selecionar Pasta PDFs",
                                           _draw_folder_icon, selecionar_pasta,
                                           "Nenhuma pasta selecionada")

# ── Frame recorte info — MODIFICADO: Fora do card branco, inserido no 'conteudo' ───

frame_recorte_info = tk.Frame(conteudo, bg="#FFF7ED",
                               highlightbackground="#1348F8", highlightthickness=1,
                               padx=10, pady=6)  # Modificado: Padding menor

frame_recorte_esq = tk.Frame(frame_recorte_info, bg="#FFF7ED")
frame_recorte_esq.pack(side=tk.LEFT, fill=tk.X, expand=True)

tk.Label(frame_recorte_esq, text="✂  Área de Recorte",
         font=("Segoe UI", 8, "bold"), bg="#FFF7ED", fg="#1348F8").pack(anchor="w")

label_recorte_info = tk.Label(frame_recorte_esq,
    text="Nenhuma área definida — clique em Definir Área",
    font=("Segoe UI", 8), bg="#FFF7ED", fg="#1348F8")
label_recorte_info.pack(anchor="w", pady=(1, 0))

tk.Button(frame_recorte_info, text="✂  Definir Área",
          font=("Segoe UI", 8, "bold"), bg="#1348F8", fg="white",
          relief="flat", cursor="hand2", padx=10, pady=4,
          command=definir_recorte).pack(side=tk.RIGHT)

# ── Separador ─────────────────────────────────────────────────────────────────

tk.Frame(conteudo, bg=COR_BORDA, height=1).pack(fill=tk.X, pady=8)

# ── Botão gerar ───────────────────────────────────────────────────────────────

frame_btn = tk.Frame(conteudo, bg=COR_BG)
frame_btn.pack(fill=tk.X)

btn_gerar = tk.Button(frame_btn,
                      text="  GERAR POWERPOINT   →",
                      font=("Segoe UI", 11, "bold"),  # Modificado: Fonte 12 -> 11
                      bg=COR_BTN_VERDE, fg="white", relief="flat",
                      cursor="hand2", pady=10, anchor="center",  # Modificado: pady 14 -> 10
                      command=gerar)
btn_gerar.pack(fill=tk.X)

btn_gerar.bind("<Enter>", lambda e: btn_gerar.config(bg=COR_BTN_VERDE_H))
btn_gerar.bind("<Leave>", lambda e: btn_gerar.config(bg=COR_BTN_VERDE))

# ── Status ────────────────────────────────────────────────────────────────────

frame_status = tk.Frame(conteudo, bg=COR_BG)
frame_status.pack(fill=tk.X, pady=(8, 0))

cnv_status = tk.Canvas(frame_status, width=16, height=16, bg=COR_BG,
                        highlightthickness=0)
cnv_status.pack(side=tk.LEFT, padx=(0, 6))
cnv_status.create_oval(1, 1, 15, 15, outline=COR_STATUS, width=1.5)
cnv_status.create_line(8, 8, 8, 4,  fill=COR_STATUS, width=1.5)
cnv_status.create_line(8, 8, 12, 8, fill=COR_STATUS, width=1.5)

label_status = tk.Label(frame_status, text="Aguardando processamento...",
                         font=("Segoe UI", 8), bg=COR_BG, fg=COR_STATUS)
label_status.pack(side=tk.LEFT)

# ── Inicializa modo padrão ────────────────────────────────────────────────────

selecionar_modo("pagina")

janela.mainloop()