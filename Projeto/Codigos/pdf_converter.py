import os
import tempfile

import fitz

from config import PREVIEW_ZOOM, PDF_ZOOM


def converter_preview(pdf_path):
    """
    Renderiza a primeira página do PDF em zoom 1x.
    Usado para exibir o PDF na janela de recorte.
    """

    documento = fitz.open(pdf_path)
    pagina = documento.load_page(0)

    pix = pagina.get_pixmap(
        matrix=fitz.Matrix(PREVIEW_ZOOM, PREVIEW_ZOOM)
    )

    temp = os.path.join(tempfile.gettempdir(), "preview.png")
    pix.save(temp)
    documento.close()

    return temp


def converter_pagina_inteira(pdf_path, png_path):
    """
    Converte a página inteira do PDF em PNG.
    Usado no modo Página Inteira.
    """

    documento = fitz.open(pdf_path)
    pagina = documento.load_page(0)

    matriz = fitz.Matrix(PDF_ZOOM, PDF_ZOOM)
    pix = pagina.get_pixmap(matrix=matriz, alpha=False)
    pix.save(png_path)
    documento.close()


def converter_recorte(pdf_path, png_path, crop_x1, crop_y1, crop_x2, crop_y2):
    """
    Converte a área recortada do PDF em PNG.
    Usado no modo Recorte.
    """

    documento = fitz.open(pdf_path)
    pagina = documento.load_page(0)

    largura_pdf = pagina.rect.width
    altura_pdf = pagina.rect.height

    x1 = max(0, crop_x1)
    y1 = max(0, crop_y1)
    x2 = min(largura_pdf, crop_x2)
    y2 = min(altura_pdf, crop_y2)

    largura_crop = max(x2 - x1, 20)
    altura_crop = max(y2 - y1, 20)

    clip = fitz.Rect(x1, y1, x1 + largura_crop, y1 + altura_crop)
    matriz = fitz.Matrix(PDF_ZOOM, PDF_ZOOM)

    pix = pagina.get_pixmap(matrix=matriz, clip=clip, alpha=False)
    pix.save(png_path)
    documento.close()


def buscar_pdf(numero, pdf_folder):
    """
    Procura o arquivo PDF pelo número.
    Tenta os formatos: 1.pdf, 01.pdf, 001.pdf
    """

    nomes = [
        f"{numero}.pdf",
        f"{numero:02d}.pdf",
        f"{numero:03d}.pdf"
    ]

    for nome in nomes:
        caminho = os.path.join(pdf_folder, nome)
        if os.path.exists(caminho):
            return caminho

    return None