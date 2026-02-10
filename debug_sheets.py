import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
import os

# Configuração da página
st.set_page_config(layout="wide")
st.title("🧪 Teste de Conexão Google Sheets")

# URLs (Copiadas do app.py para isolamento)
GSHEETS_URLS = {
    "ORGANOGRAMA": "https://docs.google.com/spreadsheets/d/1LUcoB0TUTfrSK2TPXNxilMi3uKp-QQyACblBh1kmzTU/edit?gid=1170896878#gid=1170896878",
    "PAGAMENTOS": "https://docs.google.com/spreadsheets/d/1-wJFyFdnB3CbpKJbvA-qKfwE97XbmHoI8YYJwTXQ6As/edit?gid=1763008282#gid=1763008282",
    "BOLSAS": "https://docs.google.com/spreadsheets/d/134-AMd93Db0gXv3NNVkPiYPJxYi6e6qZdNtorGgDelk/edit?gid=1113624519#gid=1113624519"
}

def teste_conexao(nome_base, url):
    st.subheader(f"Testando: {nome_base}")
    try:
        # Tenta conectar
        conn = st.connection("gsheets", type=GSheetsConnection)
        
        # Tenta ler (usando cache do st.connection, mas aqui queremos forçar o teste)
        st.write(f"Tentando ler URL: {url}...")
        df = conn.read(spreadsheet=url)
        
        if not df.empty:
            st.success(f"✅ SUCESSO! {len(df)} registros encontrados.")
            st.dataframe(df.head())
        else:
            st.warning("⚠️ Conectou, mas a tabela está vazia.")
            
    except Exception as e:
        st.error(f"❌ ERRO ao conectar:")
        st.code(str(e))

# Executar testes
col1, col2, col3 = st.columns(3)

with col1:
    teste_conexao("ORGANOGRAMA", GSHEETS_URLS["ORGANOGRAMA"])

with col2:
    teste_conexao("PAGAMENTOS", GSHEETS_URLS["PAGAMENTOS"])

with col3:
    teste_conexao("BOLSAS", GSHEETS_URLS["BOLSAS"])
