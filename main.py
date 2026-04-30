import os
from src.extract import extract_and_clean

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
SHEET_PATH = os.path.join(ROOT_DIR, "delivery_datas.xlsx")

def run_pipeline():
    print("Iniciando Dados do Delivery")
    
    # ETAPA 1: Extração
    cleaned_df = extract_and_clean(SHEET_PATH)
    
    # Só para testar se deu certo:
    print(cleaned_df.head())
    print("\nExtração concluída com sucesso!")

if __name__ == "__main__":
    run_pipeline()