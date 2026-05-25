import pandas as pd


def ler_pallets(excel_file):
    """
    Lê a planilha Excel e retorna um dicionário:
        { numero_pallet: [lista_de_numeros_pdf] }

    Exemplo:
        { 1: [101, 102], 2: [201, 202, 203] }
    """

    df = pd.read_excel(
        excel_file,
        header=None,
        engine="openpyxl"
    )

    linha_pallet = _encontrar_linha_pallet(df)

    if linha_pallet is None:
        raise ValueError("Linha dos pallets não encontrada na planilha.")

    pallets = {}

    for coluna in df.columns:

        numero_pallet = _ler_pallet(df, linha_pallet, coluna)

        if numero_pallet is None:
            continue

        numeros_pdf = _ler_pdfs_da_coluna(df, linha_pallet, coluna)

        if numeros_pdf:
            pallets[numero_pallet] = numeros_pdf

    return pallets


def _encontrar_linha_pallet(df):
    """Retorna o índice da linha que contém a palavra 'Pallet'."""

    for i in range(len(df)):
        if "Pallet" in df.iloc[i].astype(str).to_string():
            return i

    return None


def _ler_pallet(df, linha_pallet, coluna):
    """Tenta converter o valor da célula para número de pallet."""

    try:
        return int(float(df.iloc[linha_pallet, coluna]))
    except (ValueError, TypeError):
        return None


def _ler_pdfs_da_coluna(df, linha_pallet, coluna):
    """Retorna lista de números de PDF abaixo da linha do pallet."""

    numeros = []

    for linha in range(linha_pallet + 1, len(df)):

        valor = df.iloc[linha, coluna]

        if pd.isna(valor):
            continue

        try:
            numeros.append(int(float(valor)))
        except (ValueError, TypeError):
            pass

    return numeros
