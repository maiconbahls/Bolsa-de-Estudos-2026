import pandas as pd
import os
import sqlite3

def debug_import_logic():
    file_path = "BASES.BOLSAS/BASE.PAGAMENTOS.xlsx"
    if not os.path.exists(file_path):
        print(f"Arquivo {file_path} não encontrado.")
        return

    print(f"Lendo {file_path}...")
    df = pd.read_excel(file_path)
    
    # Simular normalização do app.py
    cols_orig = list(df.columns)
    df.columns = [str(c).upper().strip() for c in df.columns]
    cols_norm = list(df.columns)
    
    print(f"Colunas Originais: {cols_orig}")
    print(f"Colunas Normalizadas: {cols_norm}")
    
    map_cols = {
        'MATRÍCULA': 'matricula', 'MATRICULA': 'matricula', 'ID': 'matricula',
        'NOME': 'nome', 'COLABORADOR': 'nome', 'NOMES': 'nome',
        'VALOR': 'valor', 'VALOR LIQUIDO': 'valor', 'LÍQUIDO': 'valor', 'TOTAL': 'valor', 'VLR. LIQUIDO': 'valor',
        'ANO': 'ano',
        'COD. LOCAL': 'cod_local', 'COD LOCAL': 'cod_local', 'CODIGO LOCAL': 'cod_local', 'CÓDIGO LOCAL': 'cod_local'
    }
    
    print("\nVerificando mapeamento:")
    for key, val in map_cols.items():
        if key in df.columns:
            print(f"Mapeado: {key} -> {val}")
    
    # Testar extração da primeira linha
    row = df.iloc[0]
    dados = {'cod_local': None}
    for col_name, col_db in map_cols.items():
        if col_name in df.columns:
            val = row[col_name]
            if pd.notna(val):
                dados[col_db] = val
    
    print(f"\nDados extraídos da linha 0: {dados}")
    
    # Verificar Organograma
    org_path = "BASES.BOLSAS/ORGANOGRAMA.xlsx"
    if os.path.exists(org_path):
        df_org = pd.read_excel(org_path)
        print(f"\nOrganograma Colunas: {list(df_org.columns)}")
        print(f"Linha 0 do Organograma:\n{df_org.iloc[0]}")
        
if __name__ == "__main__":
    debug_import_logic()
