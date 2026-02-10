import sqlite3
import pandas as pd
import os

def check_na_bucket():
    conn = sqlite3.connect('bolsas.db')
    
    # 1. Ver quem são as pessoas no N/A
    # Vamos simular a lógica do Dashboard
    df_pag = pd.read_sql_query("SELECT matricula, nome, valor, cod_local, diretoria FROM historico_pagamentos", conn)
    df_bol = pd.read_sql_query("SELECT matricula, diretoria as diretoria_cad, cod_local as cod_local_cad FROM bolsistas", conn)
    
    df_merged = pd.merge(df_pag, df_bol, on='matricula', how='left')
    
    # Fallback
    df_merged['cod_local_final'] = df_merged['cod_local_cad'].fillna(df_merged['cod_local'])
    df_merged['diretoria_final'] = df_merged['diretoria_cad'].fillna(df_merged['diretoria'])
    
    # Normalizar (o que o app faz)
    df_merged['diretoria_final'] = df_merged['diretoria_final'].fillna('N/A').astype(str).str.upper().str.strip()
    df_merged.loc[df_merged['diretoria_final'] == 'NAN', 'diretoria_final'] = 'N/A'
    df_merged.loc[df_merged['diretoria_final'] == 'NONE', 'diretoria_final'] = 'N/A'
    
    na_data = df_merged[df_merged['diretoria_final'] == 'N/A']
    
    print(f"--- Colaboradores no Bucket N/A (Total: R$ {na_data['valor'].sum():,.2f}) ---")
    if not na_data.empty:
        # Agrupar por pessoa para não listar 5000 linhas
        resumo = na_data.groupby(['matricula', 'nome', 'cod_local_final']).agg({'valor': 'sum'}).reset_index()
        print(resumo.sort_values('valor', ascending=False).head(30))
        print(f"\nTotal de pessoas no N/A: {len(resumo)}")
    else:
        print("Gráfico deveria estar limpo!")

    # 2. Verificar se esses cod_local_final existem no Organograma
    if os.path.exists('BASES.BOLSAS/ORGANOGRAMA.xlsx'):
        df_org = pd.read_excel('BASES.BOLSAS/ORGANOGRAMA.xlsx')
        df_org['Cod. Local'] = df_org['Cod. Local'].astype(str).str.strip()
        
        print("\n--- Verificando Códigos Locais do N/A no Organograma ---")
        for cl in na_data['cod_local_final'].unique():
            if pd.isna(cl) or cl == '':
                print(f"Código: [VAZIO] -> N/A")
                continue
            
            match = False
            cl_str = str(cl).strip()
            for ocl in df_org['Cod. Local']:
                if cl_str.startswith(ocl):
                    match = True
                    # Pegar a diretoria que seria mapeada
                    dir_match = df_org[df_org['Cod. Local'] == ocl]['Diretoria'].values[0]
                    print(f"Código: {cl_str} -> DEVERIA SER: {dir_match}")
                    break
            if not match:
                print(f"Código: {cl_str} -> NÃO ENCONTRADO NO ORGANOGRAMA")
                
    conn.close()

if __name__ == "__main__":
    check_na_bucket()
