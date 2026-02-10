import sqlite3
import pandas as pd
import os

def fix_database_na():
    conn = sqlite3.connect('bolsas.db')
    
    # 1. Carregar Organograma
    if not os.path.exists('BASES.BOLSAS/ORGANOGRAMA.xlsx'):
        print("Organograma não encontrado.")
        return
        
    df_org = pd.read_excel('BASES.BOLSAS/ORGANOGRAMA.xlsx')
    df_org.columns = [str(c).strip() for c in df_org.columns]
    df_org['Cod. Local'] = df_org['Cod. Local'].astype(str).str.strip()
    # Ordenar por tamanho para match mais específico
    df_org = df_org.sort_values('Cod. Local', key=lambda s: s.str.len(), ascending=False)
    
    # Criar mapping
    mapping = {}
    for _, row in df_org.iterrows():
        mapping[str(row['Cod. Local']).strip()] = str(row.get('Diretoria', 'N/A')).strip().upper()

    # 2. Buscar pagamentos N/A
    pagamentos = pd.read_sql_query("SELECT id, matricula, cod_local, diretoria FROM historico_pagamentos", conn)
    
    print(f"Total de pagamentos: {len(pagamentos)}")
    
    updates = []
    rescued_count = 0
    
    for _, row in pagamentos.iterrows():
        cl = str(row['cod_local']).strip() if row['cod_local'] else ""
        current_dir = str(row['diretoria']).strip().upper() if row['diretoria'] else "N/A"
        
        if current_dir in ["N/A", "NAN", "NONE", ""]:
            # Tentar resgatar
            if cl:
                for prefix, diretoria in mapping.items():
                    if cl.startswith(prefix):
                        updates.append((diretoria, row['id']))
                        rescued_count += 1
                        break
    
    if updates:
        print(f"Atualizando {len(updates)} registros...")
        conn.executemany("UPDATE historico_pagamentos SET diretoria = ? WHERE id = ?", updates)
        conn.commit()
    
    print(f"Resgatados: {rescued_count}")
    
    # 3. Também atualizar a tabela bolsistas se estiverem sem diretoria
    bolsistas = pd.read_sql_query("SELECT matricula, cod_local, diretoria FROM bolsistas", conn)
    bol_updates = []
    for _, row in bolsistas.iterrows():
        cl = str(row['cod_local']).strip() if row['cod_local'] else ""
        current_dir = str(row['diretoria']).strip().upper() if row['diretoria'] else "N/A"
        
        if current_dir in ["N/A", "NAN", "NONE", ""]:
            if cl:
                for prefix, diretoria in mapping.items():
                    if cl.startswith(prefix):
                        bol_updates.append((diretoria, row['matricula']))
                        break
    
    if bol_updates:
        print(f"Atualizando {len(bol_updates)} bolsistas...")
        conn.executemany("UPDATE bolsistas SET diretoria = ? WHERE matricula = ?", bol_updates)
        conn.commit()
        
    conn.close()

if __name__ == "__main__":
    fix_database_na()
