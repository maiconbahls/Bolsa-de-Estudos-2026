import sqlite3
import pandas as pd
import os

def check_data():
    conn = sqlite3.connect('bolsas.db')
    
    # 1. Pagamentos sem Bolsistas
    pag = pd.read_sql('SELECT DISTINCT matricula, nome FROM historico_pagamentos', conn)
    bol = pd.read_sql('SELECT matricula FROM bolsistas', conn)
    missing = pag[~pag['matricula'].isin(bol['matricula'])]
    
    print("--- 1. MATRÍCULAS NO PAGAMENTO MAS NÃO NO CADASTRO BOLSISTAS ---")
    if not missing.empty:
        print(missing.head(20))
        print(f"Total: {len(missing)}")
    else:
        print("Nenhuma!")
        
    # 2. Bolsistas sem Cod_Local
    cl_null = pd.read_sql('SELECT matricula, nome FROM bolsistas WHERE cod_local IS NULL OR cod_local = ""', conn)
    print("\n--- 2. CADASTRO DE BOLSISTAS SEM CÓDIGO LOCAL ---")
    if not cl_null.empty:
        print(cl_null.head(10))
        print(f"Total: {len(cl_null)}")
    else:
        print("Todos têm Código Local!")

    # 3. Bolsistas com Cod_Local que não batem com Organograma
    if os.path.exists('BASES.BOLSAS/ORGANOGRAMA.xlsx'):
        df_org = pd.read_excel('BASES.BOLSAS/ORGANOGRAMA.xlsx')
        df_org['Cod. Local'] = df_org['Cod. Local'].astype(str).str.strip()
        org_codes = df_org['Cod. Local'].tolist()
        
        has_cl = pd.read_sql('SELECT matricula, nome, cod_local FROM bolsistas WHERE cod_local IS NOT NULL AND cod_local != ""', conn)
        no_match = []
        for _, row in has_cl.iterrows():
            cl = str(row['cod_local']).strip()
            match = False
            for ocl in org_codes:
                if cl.startswith(ocl):
                    match = True
                    break
            if not match:
                no_match.append((row['matricula'], row['nome'], cl))
        
        print("\n--- 3. CADASTRO COM CÓDIGO LOCAL QUE NÃO COMBINA COM ORGANOGRAMA ---")
        if no_match:
            print(pd.DataFrame(no_match, columns=['Matrícula', 'Nome', 'Cod_Local']).head(20))
            print(f"Total: {len(no_match)}")
        else:
            print("Todos os Códigos Locais batem com o Organograma!")
            
    conn.close()

if __name__ == "__main__":
    check_data()
