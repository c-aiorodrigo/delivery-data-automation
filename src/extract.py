import pandas as pd

def extract_and_clean (file_path: str) -> pd.DataFrame:
    """
    Read the sheet and clean datas.
    """
    print("Extraindo e limpando dados...")
    
    # 1. Carrega os dados
    df = pd.read_excel(file_path, sheet_name='Dados Gerais')

    # 2. Limpeza de strings
    df.columns = df.columns.str.strip().str.upper()
    df['LOJAS'] = df['LOJAS'].str.strip().str.upper()

    # 3. Conversão de tipos
    df['DATA DE ENTREGA'] = pd.to_datetime(df['DATA DE ENTREGA'], errors='coerce')
    df['VALOR'] = pd.to_numeric(df['VALOR'], errors='coerce')

    return df