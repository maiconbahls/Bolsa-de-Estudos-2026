#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sistema de Gestao de Bolsas de Estudos - COCAL
Com super tabela interativa
"""

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import sqlite3
import os
import warnings
import logging
import hashlib
from pathlib import Path
from streamlit_option_menu import option_menu
warnings.filterwarnings('ignore')

# ---------------------------------------------------------------------------
# Logging configuration
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
DB_PATH = "bolsas.db"

MESES = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
]

MESES_SAFRA = [
    "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro",
    "Outubro", "Novembro", "Dezembro", "Janeiro", "Fevereiro", "Março",
]

MESES_SAFRA_NUM = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]

# ---------------------------------------------------------------------------
# Configuração Global Plotly (Download Transparente)
# ---------------------------------------------------------------------------
PLOTLY_CONFIG = {
    'displayModeBar': True,
    'toImageButtonOptions': {
        'format': 'png',
        'filename': 'grafico_bolsas_cocal',
        'height': None,
        'width': None,
        'scale': 2 # Alta resolução
    }
}

# ---------------------------------------------------------------------------
# Paleta de Cores (Inspirada no Layout Institucional)
# ---------------------------------------------------------------------------
APP_COLORS = {
    'primary': '#78c045',   # Verde Vibrante (Cargas/Auxilio)
    'secondary': '#2c3e50', # Azul Petróleo Escuro (Contraste)
    'accent': '#1f2937',    # Cinza Escuro
    'grid': '#f3f4f6',      # Cinza Claro para grades
    'text': '#374151'       # Texto Padrão
}

# ---------------------------------------------------------------------------
# Sistema de Autenticação
# ---------------------------------------------------------------------------
def hash_password(password):
    """Gera hash SHA256 da senha"""
    return hashlib.sha256(password.encode()).hexdigest()

def check_authentication():
    """Verifica se o usuário está autenticado"""
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    return st.session_state.authenticated

def login_page():
    """Renderiza a página de login"""
    # CSS customizado para a página de login
    st.markdown("""
        <style>
            .login-container {
                max-width: 400px;
                margin: 100px auto;
                padding: 40px;
                background: white;
                border-radius: 12px;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            }
            .login-header {
                text-align: center;
                margin-bottom: 30px;
            }
            .login-header h1 {
                color: #1e293b;
                font-size: 1.8rem;
                margin-bottom: 10px;
            }
            .login-header p {
                color: #64748b;
                font-size: 0.9rem;
            }
            .stButton > button {
                width: 100%;
                background: linear-gradient(135deg, #78c045 0%, #5a9033 100%);
                color: white;
                border: none;
                padding: 12px;
                font-weight: 600;
                border-radius: 8px;
                margin-top: 20px;
            }
            .stButton > button:hover {
                background: linear-gradient(135deg, #5a9033 0%, #78c045 100%);
            }
        </style>
    """, unsafe_allow_html=True)
    
    # Container centralizado
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("""
            <div class="login-header">
                <h1>🎓 Sistema de Bolsas</h1>
                <p>COCAL - Recursos Humanos</p>
            </div>
        """, unsafe_allow_html=True)
        
        # Formulário de login
        with st.form("login_form"):
            username = st.text_input("👤 Usuário", placeholder="Digite seu usuário")
            password = st.text_input("🔒 Senha", type="password", placeholder="Digite sua senha")
            submit = st.form_submit_button("🔓 Entrar")
            
            if submit:
                # Credenciais: gestao / gestao
                correct_username = "gestao"
                correct_password_hash = hash_password("gestao")
                
                if username == correct_username and hash_password(password) == correct_password_hash:
                    st.session_state.authenticated = True
                    st.session_state.username = username
                    st.success("✅ Login realizado com sucesso!")
                    st.rerun()
                else:
                    st.error("❌ Usuário ou senha incorretos!")
        
        # Rodapé
        st.markdown("---")
        st.caption("🔐 Acesso restrito - Sistema interno COCAL")

def logout():
    """Realiza logout do usuário"""
    st.session_state.authenticated = False
    if 'username' in st.session_state:
        del st.session_state.username
    st.rerun()

# ---------------------------------------------------------------------------
# Backup Automático e Integração
# ---------------------------------------------------------------------------
def backup_database():
    """Cria um backup do banco de dados antes de alterações críticas"""
    import shutil
    import os
    from datetime import datetime
    
    if not os.path.exists("backups"):
        os.makedirs("backups")
        
    db_file = "bolsas.db"
    if os.path.exists(db_file):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = f"backups/bolsas_{timestamp}.db"
        try:
            shutil.copy2(db_file, backup_file)
            # Manter apenas os ultimos 10 backups para não lotar disco
            backups = sorted([os.path.join("backups", f) for f in os.listdir("backups") if f.endswith(".db")])
            while len(backups) > 10:
                os.remove(backups.pop(0))
            return True, f"Backup criado: {backup_file}"
        except Exception as e:
            return False, f"Falha no backup: {e}"
    return False, "Banco não encontrado"

# ---------------------------------------------------------------------------
# Data Connection Layer (Google Sheets + Local Fallback)
# ---------------------------------------------------------------------------
from streamlit_gsheets import GSheetsConnection

# Helper para ler Excel localmente de forma segura
def safe_read_excel(file_path, **kwargs):
    """Lê um arquivo Excel mesmo que ele esteja aberto em outro programa (copia para temp)"""
    import shutil
    import tempfile
    
    if not os.path.exists(file_path):
        return pd.DataFrame()
        
    try:
        # Tenta ler diretamente primeiro
        return pd.read_excel(file_path, **kwargs)
    except PermissionError:
        # Se falhar por permissão (arquivo aberto), copia para local temporário
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                shutil.copy2(file_path, tmp.name)
                df = pd.read_excel(tmp.name, **kwargs)
            os.unlink(tmp.name)
            return df
        except Exception as e:
            logger.error(f"Erro crítico ao ler arquivo (mesmo via temp): {file_path} -> {e}")
            return pd.DataFrame()
    except Exception as e:
        logger.error(f"Erro ao ler arquivo: {file_path} -> {e}")
        return pd.DataFrame()

# URLs das Planilhas Google (Substitua pelos seus links reais)
GSHEETS_URLS = {
    "ORGANOGRAMA": "https://docs.google.com/spreadsheets/d/1LUcoB0TUTfrSK2TPXNxilMi3uKp-QQyACblBh1kmzTU/edit?gid=1170896878#gid=1170896878",
    "PAGAMENTOS": "https://docs.google.com/spreadsheets/d/1-wJFyFdnB3CbpKJbvA-qKfwE97XbmHoI8YYJwTXQ6As/edit?gid=1763008282#gid=1763008282",
    "BOLSAS": "https://docs.google.com/spreadsheets/d/134-AMd93Db0gXv3NNVkPiYPJxYi6e6qZdNtorGgDelk/edit?gid=1113624519#gid=1113624519"
}

# Caminhos locais para Fallback
LOCAL_PATHS = {
    "ORGANOGRAMA": "BASES.BOLSAS/ORGANOGRAMA.xlsx",
    "PAGAMENTOS": "BASES.BOLSAS/BASE.PAGAMENTOS.xlsx",
    "BOLSAS": "BASES.BOLSAS/BASE.BOLSAS.2025.xlsx"
}

@st.cache_data(ttl=300)
def get_dataset(source_key):
    """
    Carrega dados de uma fonte (Google Sheets ou Local).
    Prioridade: Google Sheets > Excel Local.
    """
    df = pd.DataFrame()
    source_name = f"Google Sheets ({source_key})"
    
    # 1. Tentar Google Sheets
    try:
        url = GSHEETS_URLS.get(source_key)
        if url:
             logger.info(f"Tentando conectar ao Google Sheets: {source_key}...")
             # Conexão usa st.secrets automaticamente se configurado, ou o public URL se for público
             conn = st.connection("gsheets", type=GSheetsConnection)
             # Tenta ler. Se falhar, vai pro except.
             df = conn.read(spreadsheet=url)
             
             if not df.empty:
                 logger.info(f"Dados carregados do Google Sheets: {source_key} | Shape: {df.shape}")
             else:
                 logger.warning(f"Google Sheets retornou dados vazios para: {source_key}")
    except Exception as e:
        logger.warning(f"Falha ao conectar Google Sheets [{source_key}]: {e}. Tentando local...")
        # st.warning(f"Falha na conexão com Google Sheets ({source_key}): {e}") # Opcional: mostrar pro usuário
        df = pd.DataFrame() # Reset para garantir fallback

    # 2. Fallback para Excel Local se falhou ou vazio
    if df.empty:
        local_path = LOCAL_PATHS.get(source_key)
        if local_path and os.path.exists(local_path):
            source_name = f"Excel Local ({local_path})"
            df = safe_read_excel(local_path)
            logger.info(f"Dados carregados localmente: {local_path} | Shape: {df.shape}")
        else:
            logger.warning(f"Arquivo local não encontrado: {local_path}")
    
    # 3. Limpeza Padrão (Trim headers) e validação
    if not df.empty:
        # Converter colunas para string e remover espaços
        df.columns = [str(c).strip() for c in df.columns]
        # Remover colunas vazias se houver
        df = df.dropna(how='all', axis=1)
    
    return df

@st.cache_data(ttl=300)
def carregar_organograma():
    """Carrega o organograma com mapeamento Cod. Local -> Diretoria, Gestor N3, Gestor N4"""
    return get_dataset("ORGANOGRAMA")

def get_organograma_mapping(df_org):
    """Cria um cache do organograma focado em Cod. Local -> Diretoria (Coluna C FÍSICA)"""
    if df_org.empty:
        return {}
    mapping = {}
    
    for i in range(len(df_org)):
        try:
            row = df_org.iloc[i]
            # Manter o código original, apenas limpando espaços
            # Mantemos o ponto conforme solicitado pelo usuário
            cod = str(row.iloc[0]).strip()
            
            # Remover o .0 que o Excel às vezes coloca em números inteiros,
            # mas manter o ponto se ele fizer parte de um código composto (ex: 10.20)
            if cod.endswith('.0') and len(cod) > 2:
                # Só removemos se for um "inteiro float" (ex: 123.0 -> 123)
                # Se for algo como 1.0, mantemos por segurança ou tratamos com cuidado
                pass 
            
            # Diretoria (Col C / índice 2)
            diretoria = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) and str(row.iloc[2]).strip() != "" else str(row.iloc[1]).strip()
            
            if cod and cod.lower() != 'nan' and diretoria and diretoria.lower() != 'nan':
                mapping[cod] = {
                    'diretoria': diretoria.upper(),
                    'g3': row.iloc[3] if len(row) > 3 else "N/A",
                    'g4': row.iloc[6] if len(row) > 6 else "N/A"
                }
        except:
            continue
            
    return dict(sorted(mapping.items(), key=lambda x: len(x[0]), reverse=True))

def buscar_info_organograma_fast(cod_local, mapping):
    """Versão otimizada usando cache de dicionário mantendo a estrutura original do código"""
    if not cod_local or not mapping:
        return "N/D", "N/D", "N/D"
    
    # Limpar apenas espaços, mantendo pontos e zeros conforme solicitado
    cl = str(cod_local).strip()
    
    # Tratar códigos inválidos explicitamente
    if cl.upper() in ['SEM CODIGO LOCAL', 'SEM CÓDIGO LOCAL', 'N/A', 'N/D', 'NAN', 'NONE', '']:
        return "N/D", "N/D", "N/D"

    # 1. Busca Exata
    if cl in mapping:
        info = mapping[cl]
        return info['diretoria'], info['g3'], info['g4']
        
    # 2. Busca por Prefixo (se o código for mais longo, ex: 10.20.30 buscando 10.20)
    for prefix, info in mapping.items():
        if cl.startswith(prefix):
            return info['diretoria'], info['g3'], info['g4']
    
    # 3. Normalizar prefixo: Pagamentos usam '01.' ou '02.', Organograma usa '1.' ou '2.'
    #    Ex: '01.1.02.001' -> '1.1.02.001', '02.1.04' -> '2.1.04'
    import re
    cl_normalizado = re.sub(r'^0+(\d)', r'\1', cl)  # Remove zeros à esquerda do primeiro número
    
    if cl_normalizado != cl:
        # Tentar busca exata com código normalizado
        if cl_normalizado in mapping:
            info = mapping[cl_normalizado]
            return info['diretoria'], info['g3'], info['g4']
        
        # Tentar busca por prefixo com código normalizado
        for prefix, info in mapping.items():
            if cl_normalizado.startswith(prefix):
                return info['diretoria'], info['g3'], info['g4']
            
    return "N/D", "N/D", "N/D"

def enriquecer_com_organograma(df, df_org):
    """Enriquece um DataFrame de bolsistas com dados do organograma usando Cod. Local"""
    if df_org.empty or len(df) == 0:
        return df
    
    # Garantir cod_local existe
    if 'cod_local' not in df.columns:
        df['cod_local'] = None

    # Normalizar diretoria existente (tratar N/D, N/A, None, etc como NaN para o combine_first)
    if 'diretoria' in df.columns:
        df['diretoria'] = df['diretoria'].astype(str).replace(['N/D', 'N/A', 'None', 'nan', '', 'nan'], None)
    
    mapping = get_organograma_mapping(df_org)
    
    diretorias_org = []
    gestores_n3 = []
    gestores_n4 = []
    
    for _, row in df.iterrows():
        cl = str(row.get('cod_local', '')).strip()
        if cl and cl not in ['None', 'nan', '']:
            d, g3, g4 = buscar_info_organograma_fast(cl, mapping)
        else:
            d, g3, g4 = None, None, None
            
        diretorias_org.append(d)
        gestores_n3.append(g3)
        gestores_n4.append(g4)
    
    df['diretoria_org'] = diretorias_org
    df['gestor_n3'] = gestores_n3
    df['gestor_n4'] = gestores_n4
    
    # Usar diretoria do organograma como prioridade se encontrada
    if 'diretoria' in df.columns:
        # Se diretoria_org tiver valor, ele ganha. Se não, mantém a diretoria original.
        df['diretoria'] = df['diretoria_org'].fillna(df['diretoria'])
    else:
        df['diretoria'] = df['diretoria_org']
    
    # Finalização: Preencher quem sobrou com N/D e normalizar (tratar strings de erro)
    df['diretoria'] = df['diretoria'].fillna('N/D').astype(str).str.upper().str.strip()
    df['diretoria'] = df['diretoria'].replace(['NAN', 'NONE', '', 'nan'], 'N/D')
    
    df = df.drop(columns=['diretoria_org'], errors='ignore')
    return df

def get_conn() -> sqlite3.Connection:
    """Abre e devolve uma conexão SQLite."""
    return sqlite3.connect(DB_PATH)

# ---------------------------------------------------------------------------
# UI Components
# ---------------------------------------------------------------------------
def format_br_currency(val):
    """Formata valor para moeda BRL"""
    if pd.isna(val) or val == "": return "R$ 0,00"
    try:
        val = float(val)
    except:
        return "R$ 0,00"
    return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def format_br_number(val):
    """Formata número para padrão BR"""
    if pd.isna(val) or val == "": return "0"
    try:
        val = float(val)
    except:
        return str(val)
    # Se for inteiro, não mostra casas decimais
    if val == int(val):
        return f"{int(val):,}".replace(",", ".")
    return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def load_css():
    """Carrega o CSS customizado e força o tema claro."""
    # Estilo base para forçar fundo branco e texto escuro em TUDO
    st.markdown("""
        <style>
            /* Reset Global de Streamlit */
            .stApp, .main, .block-container, [data-testid="stAppViewContainer"] {
                background-color: #ffffff !important;
                color: #1e293b !important;
            }
            [data-testid="stHeader"] {
                background-color: #ffffff !important;
                border-bottom: 1px solid #e2e8f0;
            }
            
            /* Remover barras escuras e fundos de widgets */
            [data-testid="stHorizontalBlock"], .stColumn, div[data-testid="stVerticalBlock"] {
                background-color: transparent !important;
            }
            
            /* Forçar texto escuro em todos os elementos */
            .stMarkdown, .stText, label, .stMetric, h1, h2, h3, h4, h5, h6, p, span {
                color: #1e293b !important;
            }
            
            /* AgGrid Themes - Forçar Tema Claro Radical (Total) */
            .ag-theme-balham, .ag-theme-alpine, .ag-theme-streamlit,
            .ag-theme-balham-dark, .ag-theme-alpine-dark,
            .stAgGrid, [class*="ag-theme-"] {
                background-color: transparent !important; /* Deixar o container transparente e forçar as classes internas */
                --ag-background-color: #ffffff !important;
                --ag-header-background-color: #f8fafc !important;
                --ag-odd-row-background-color: #ffffff !important;
                --ag-row-background-color: #ffffff !important;
                --ag-foreground-color: #1e293b !important;
                --ag-header-foreground-color: #475569 !important;
                --ag-border-color: #e2e8f0 !important;
                --ag-row-border-color: #f1f5f9 !important;
                --ag-control-panel-background-color: #ffffff !important;
                --ag-widget-container-background-color: #ffffff !important;
            }
            
            /* Target direto em elementos internos do AgGrid */
            .ag-root-wrapper, .ag-root, .ag-header, .ag-header-row, 
            .ag-header-cell, .ag-row, .ag-cell, .ag-body-viewport,
            .ag-center-cols-viewport, .ag-body-horizontal-scroll-viewport {
                background-color: #ffffff !important;
                color: #1e293b !important;
                border-color: #e2e8f0 !important;
            }

            .ag-row, .ag-row-even, .ag-row-odd {
                background-color: #ffffff !important;
                color: #1e293b !important;
            }
            
            /* Forçar Header a ser claro */
            .ag-header {
                background-color: #f8fafc !important;
                border-bottom: 2px solid #e2e8f0 !important;
            }
            .ag-header-cell-label {
                color: #475569 !important;
                font-weight: bold !important;
            }
            
            /* Tabelas Nativas (st.dataframe) */
            div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] {
                background-color: #ffffff !important;
            }
            div[data-testid="stDataFrame"] div, div[data-testid="stDataEditor"] div {
                background-color: transparent !important;
            }

            /* OptionMenu Fix: Forçar o fundo do menu a ser sempre branco */
            div.container-fluid, .nav-item, .nav-link-selected {
                background-color: #ffffff !important;
            }
            .nav-link {
                background-color: transparent !important;
                color: #475569 !important;
            }
            
            /* Plotly Force Light */
            .js-plotly-plot, .plotly {
                background-color: #ffffff !important;
                border: 1px solid #f1f5f9;
                border-radius: 8px;
            }
            
            /* Toolbar e Decorações */
            div[data-testid="stToolbar"] { visibility: hidden; }
            div[data-testid="stDecoration"] { display: none; }
        </style>
    """, unsafe_allow_html=True)
    
    css_path = Path(__file__).parent / "static/style.css"
    try:
        with open(css_path, "r", encoding="utf-8") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        st.warning("⚠️ Arquivo CSS não encontrado.")

def render_header(stats):
    """Renderiza o cabeçalho principal com estatísticas."""
    st.markdown(f"""
    <div class="top-header">
        <div>
            <h1>🎓 Sistema de Bolsas de Estudos</h1>
            <span class="stats">COCAL | Controle Mensal</span>
        </div>
        <div style="text-align: right;">
            <span class="stats"><strong>{stats['ativos']}</strong> bolsas ativas</span><br>
            <span class="stats"><strong>R$ {stats['investimento']:,.0f}</strong> Safra 25/26</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

def render_stat_card(value, title, color="#1e293b"):
    """Renderiza um card de estatística individual."""
    st.markdown(f"""
    <div class="stat-card">
        <h2 style="color: {color};">{value}</h2>
        <p>{title}</p>
    </div>
    """, unsafe_allow_html=True)

def render_modern_metric(icon, label, value, color="#2563eb", bg_gradient="linear-gradient(135deg, #667eea 0%, #764ba2 100%)"):
    """Renderiza um card de métrica moderno sem ícone, apenas com borda colorida."""
    st.markdown(f"""
    <div style="
        background: white;
        padding: 24px;
        border-radius: 8px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        border: 1px solid #e5e7eb;
        border-left: 5px solid {color};
        margin-bottom: 20px;
        height: 100%;
    ">
        <p style="
            margin: 0 0 10px 0;
            color: #6b7280;
            font-size: 0.75rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        ">{label}</p>
        <h2 style="
            margin: 0;
            color: #111827;
            font-size: 2rem;
            font-weight: 800;
            line-height: 1;
        ">{value}</h2>
    </div>
    """, unsafe_allow_html=True)



def render_area_chart(df, x_col, y_col, title, label_y="Valor"):
    """Renderiza um gráfico de área com Plotly."""
    # Verificar se há dados
    if df is None or len(df) == 0:
        st.info("📊 Sem dados para exibir no gráfico.")
        return
    
    # Configuração de Cores (Novo Padrão)
    line_color = APP_COLORS['primary']
    fill_color = 'rgba(120, 192, 69, 0.15)' # Primary com transparência
    
    fig = px.area(
        df, 
        x=x_col, 
        y=y_col,
        title=title,
        labels={x_col: "", y_col: label_y},
        text=y_col
    )
    
    fig.update_traces(
        line=dict(color=line_color, width=3),
        fillcolor=fill_color,
        mode='lines+markers+text',
        textposition="top center",
        texttemplate='R$ %{text:,.0f}' # Valor simplificado
    )
    
    fig.update_layout(
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(color="#374151", size=11),
        title=dict(
            text=title,
            font=dict(size=14, color="#111827", weight="bold"),
            x=0
        ),
        margin=dict(l=20, r=20, t=60, b=20),
        xaxis=dict(
            showgrid=False, 
            linecolor='#e5e7eb',
            tickfont=dict(color='#6b7280')
        ),
        yaxis=dict(
            showgrid=True, 
            gridcolor="#f3f4f6",
            tickfont=dict(color='#6b7280'),
            tickprefix="R$ "
        ),
        hovermode="x unified"
    )
    st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)

def render_bar_chart(df, x_col, y_col, title, label_y="Valor", currency=False):
    """Renderiza um gráfico de barras com Plotly."""
    # Verificar se há dados
    if df is None or len(df) == 0:
        st.info("📊 Sem dados para exibir no gráfico.")
        return
    
    # Preparar texto formatado se for moeda
    df_chart = df.copy()
    if currency:
        df_chart['chart_text'] = df_chart[y_col].apply(lambda x: format_br_currency(x))
    else:
        df_chart['chart_text'] = df_chart[y_col]

    # Altura dinâmica baseada na quantidade de barras (mínimo 400px)
    n_bars = len(df_chart)
    chart_height = max(400, n_bars * 35)

    # Bar Chart com cor sólida e design clean
    fig = px.bar(
        df_chart, 
        x=y_col, 
        y=x_col, 
        orientation='h', 
        title=title,
        text='chart_text',
        color_discrete_sequence=[APP_COLORS['primary']]
    )
    
    fig.update_traces(
        textposition='outside',
        texttemplate='%{text}', 
        textfont=dict(color='#374151', size=11, weight='bold'),
        cliponaxis=False 
    )
    
    fig.update_layout(
        height=chart_height,
        coloraxis_showscale=False,
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(color="#374151", size=11),
        title=dict(
            text=title,
            font=dict(size=14, color="#111827", weight="bold"),
            x=0
        ),
        margin=dict(l=20, r=120, t=40, b=40), # Aumentar margem direita para não cortar texto
        xaxis=dict(
            showgrid=True, 
            gridcolor="#f3f4f6", 
            title=None,
            tickfont=dict(color='#6b7280')
        ),
        yaxis=dict(
            showgrid=False,
            title=None,
            tickfont=dict(color='#4b5563', size=11),
            categoryorder='total ascending' # Ordenar do maior para o menor
        ),
        separators=',.'
    )
    st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)


# ---------------------------------------------------------------------------
# AgGrid
# ---------------------------------------------------------------------------
try:
    from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode
    AGGRID = True
except ImportError:
    AGGRID = False
except Exception:
    AGGRID = False

st.set_page_config(page_title="Bolsas COCAL", page_icon="🎓", layout="wide")

# Carregar estilo visual
load_css()

def init_database():
    conn = get_conn()
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS bolsistas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            matricula TEXT UNIQUE NOT NULL,
            nome TEXT NOT NULL,
            cpf TEXT,
            diretoria TEXT,
            curso TEXT,
            instituicao TEXT,
            inicio_curso DATE,
            fim_curso DATE,
            ano_referencia INTEGER,
            mensalidade REAL DEFAULT 0,
            porcentagem REAL DEFAULT 0.5,
            valor_reembolso REAL DEFAULT 0,
            situacao TEXT DEFAULT 'ATIVO',
            checagem TEXT DEFAULT 'REGULAR',
            observacao TEXT,
            data_cadastro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Migração para adicionar colunas se não existirem
    try:
        cursor.execute("ALTER TABLE bolsistas ADD COLUMN inicio_curso DATE")
    except: pass
    try:
        cursor.execute("ALTER TABLE bolsistas ADD COLUMN fim_curso DATE")
    except: pass
    try:
        cursor.execute("ALTER TABLE bolsistas ADD COLUMN ano_referencia INTEGER")
    except: pass
    try:
        cursor.execute("ALTER TABLE bolsistas ADD COLUMN instituicao TEXT")
    except: pass
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS pagamentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bolsista_id INTEGER NOT NULL,
            mes INTEGER NOT NULL,
            ano INTEGER NOT NULL,
            valor REAL,
            status TEXT DEFAULT 'PENDENTE',
            observacao TEXT,
            FOREIGN KEY (bolsista_id) REFERENCES bolsistas(id),
            UNIQUE(bolsista_id, mes, ano)
        )
    ''')
    
    # Tabela histórico de pagamentos importados (caso não exista)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS historico_pagamentos (
           id INTEGER PRIMARY KEY AUTOINCREMENT,
           matricula TEXT,
           nome TEXT,
           mes INTEGER,
           ano INTEGER,
           mes_referencia TEXT,
           valor REAL,
           data_pagamento DATE,
           cod_local TEXT,
           diretoria TEXT
        )
    ''')
    
    # Migração para adicionar colunas extras ao historico_pagamentos
    try:
        cursor.execute("ALTER TABLE historico_pagamentos ADD COLUMN cod_local TEXT")
    except:
        pass
    try:
        cursor.execute("ALTER TABLE historico_pagamentos ADD COLUMN diretoria TEXT")
    except:
        pass
    try:
        cursor.execute("ALTER TABLE historico_pagamentos ADD COLUMN safra TEXT")
    except:
        pass

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS observacoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bolsista_id INTEGER NOT NULL,
            data TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            texto TEXT,
            anexo_blob BLOB,
            nome_anexo TEXT,
            FOREIGN KEY(bolsista_id) REFERENCES bolsistas(id)
        )
    ''')
    
    try:
        cursor.execute("ALTER TABLE observacoes ADD COLUMN anexo_blob BLOB")
    except: pass
    try:
        cursor.execute("ALTER TABLE observacoes ADD COLUMN nome_anexo TEXT")
    except: pass
    
    try:
        cursor.execute("ALTER TABLE bolsistas ADD COLUMN tipo TEXT")
    except: pass
    try:
        cursor.execute("ALTER TABLE bolsistas ADD COLUMN modalidade TEXT")
    except: pass
    try:
        cursor.execute("ALTER TABLE bolsistas ADD COLUMN cod_local TEXT")
    except: pass
    
    # NOVA TABELA: ORÇAMENTO
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS orcamento (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            diretoria TEXT NOT NULL,
            ano INTEGER NOT NULL,
            valor_mensal_meta REAL DEFAULT 0,
            UNIQUE(diretoria, ano)
        )
    ''')
    
    conn.commit()
    conn.close()

init_database()

# get_conn já definida acima

def df_to_excel(df):
    """Converte DataFrame para bytes Excel para download"""
    from io import BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Dados')
    return output.getvalue()

def cadastrar_bolsista(dados):
    conn = get_conn()
    try:
        # Backup antes de escrever
        backup_database()
        
        # Prepara valores usando .get() para suportar novos campos
        campos = ['matricula', 'nome', 'cpf', 'diretoria', 'cod_local', 'curso', 'instituicao', 'tipo', 'modalidade',
                  'inicio_curso', 'fim_curso', 'ano_referencia',
                  'mensalidade', 'porcentagem', 'valor_reembolso', 'situacao', 'checagem', 'observacao']
        
        # Cria lista de valores na ordem correta
        valores = [dados.get(k) for k in campos]
        
        conn.execute('''
            INSERT INTO bolsistas (matricula, nome, cpf, diretoria, cod_local, curso, instituicao, tipo, modalidade,
                                   inicio_curso, fim_curso, ano_referencia,
                                   mensalidade, porcentagem, valor_reembolso, situacao, checagem, observacao)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', valores)
        conn.commit()
        return True, "Cadastrado com sucesso!"
    except sqlite3.IntegrityError:
        return False, f"Matrícula {dados.get('matricula')} já existe!"
    except Exception as e:
        return False, f"Erro ao cadastrar: {str(e)}"
    finally:
        conn.close()

def upsert_bolsista(dados, preserve_status=False):
    conn = get_conn()
    try:
        # Backup antes de escrever
        backup_database()
        
        campos = ['matricula', 'nome', 'cpf', 'diretoria', 'cod_local', 'curso', 'instituicao', 'tipo', 'modalidade',
                  'inicio_curso', 'fim_curso', 'ano_referencia',
                  'mensalidade', 'porcentagem', 'valor_reembolso', 'situacao', 'checagem', 'observacao']
        
        # Tenta update primeiro (assim preservamos ID)
        # Verifica se existe
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM bolsistas WHERE matricula = ?", (dados.get('matricula'),))
        existe = cursor.fetchone()
        
        valores = [dados.get(k) for k in campos]
        
        if existe:
            # Update
            update_fields = []
            update_vals = []
            for k in campos:
                # Se preserve_status for True, pular campos sensíveis se já existirem
                if preserve_status and k in ['situacao', 'checagem', 'observacao']:
                    continue
                    
                if k != 'matricula' and dados.get(k) is not None:
                     update_fields.append(f"{k} = ?")
                     update_vals.append(dados.get(k))
            
            if update_fields:
                update_vals.append(dados.get('matricula'))
                sql = f"UPDATE bolsistas SET {', '.join(update_fields)} WHERE matricula = ?"
                conn.execute(sql, update_vals)
                conn.commit()
                return True, "Atualizado"
            else:
                return True, "Sem mudanças"
        else:
            # Insert
            conn.execute(f'''
                INSERT INTO bolsistas ({', '.join(campos)})
                VALUES ({', '.join(['?']*len(campos))})
            ''', valores)
            conn.commit()
            return True, "Cadastrado"

    except Exception as e:
        return False, f"Erro: {str(e)}"
    finally:
        conn.close()

def processar_importacao_df(df_import, preserve_status=False):
    try:
        # Backup antes de importar
        backup_database()
        
        st.write(f"Processando {len(df_import)} registros...")
        
        # Normalizar colunas para maiúsculo
        df_import.columns = [str(c).upper().strip() for c in df_import.columns]
        
        # Mapeamento de colunas (Excel Maiúsculo -> Banco) - EXPANDIDO
        map_cols = {
            # Matrícula
            'MATRÍCULA': 'matricula', 'MATRICULA': 'matricula', 'MATR': 'matricula', 'ID': 'matricula', 'RE': 'matricula', 'REGISTRO': 'matricula',
            # Nome
            'NOME': 'nome', 'COLABORADOR': 'nome', 'NOMES': 'nome', 'FUNCIONARIO': 'nome', 'BOLSISTA': 'nome',
            # CPF
            'CPF': 'cpf',
            # Diretoria
            'DIRETORIA': 'diretoria', 'AREA': 'diretoria', 'DEPARTAMENTO': 'diretoria', 'DEPTO': 'diretoria',
            # Código Local
            'COD. LOCAL': 'cod_local', 'COD LOCAL': 'cod_local', 'CODIGO LOCAL': 'cod_local', 'CÓDIGO LOCAL': 'cod_local', 
            'CENTRO DE CUSTO': 'cod_local', 'CC': 'cod_local', 'CR': 'cod_local', 'COD_LOCAL': 'cod_local',
            # Curso
            'CURSO': 'curso', 
            # Instituição (várias variantes)
            'INSTITUIÇÃO': 'instituicao', 'INSTITUICAO': 'instituicao', 'INSTITUIO': 'instituicao', 'FACULDADE': 'instituicao', 'UNIVERSIDADE': 'instituicao',
            # Tipo
            'TIPO': 'tipo', 'NIVEL': 'tipo',
            # Modalidade
            'MODALIDADE': 'modalidade',
            # Início do curso (várias variantes - SEM e COM acento)
            'INÍCIO CURSO': 'inicio_curso', 'INICIO CURSO': 'inicio_curso', 
            'INICIO DO CURSO': 'inicio_curso', 'INÍCIO DO CURSO': 'inicio_curso',
            'DATA INICIO': 'inicio_curso', 'DATA INÍCIO': 'inicio_curso', 'INICIO': 'inicio_curso',
            # Fim do curso (várias variantes)
            'FIM CURSO': 'fim_curso', 'TERMINO DO CURSO': 'fim_curso', 
            'TÉRMINO DO CURSO': 'fim_curso', 'FIM DO CURSO': 'fim_curso',
            'DATA FIM': 'fim_curso', 'DATA TERMINO': 'fim_curso', 'FIM': 'fim_curso',
            # Ano referência (várias variantes)
            'ANO PROGRAMA': 'ano_referencia', 'ANO': 'ano_referencia',
            'ANO REFERENCIA': 'ano_referencia', 'ANO REFERÊNCIA': 'ano_referencia', 'SAFRA': 'ano_referencia',
            # Mensalidade
            'MENSALIDADE': 'mensalidade', 'VALOR MENSALIDADE': 'mensalidade',
            'MENSALIDADE PREV CONTRATO': 'mensalidade',
            # Porcentagem
            '% BOLSA': 'porcentagem', 'PORCENTAGEM': 'porcentagem', '%BOLSA': 'porcentagem', '%': 'porcentagem',
            # Valor reembolso
            'VALOR REEMBOLSO': 'valor_reembolso', 'VALOR': 'valor_reembolso', 'REEMBOLSO': 'valor_reembolso',
            # Situação
            'SITUAÇÃO': 'situacao', 'SITUACAO': 'situacao', 'STATUS': 'situacao',
            # Checagem
            'CHECAGEM': 'checagem', 'CHECAGEM SITUACAO': 'checagem', 'CHECAGEM SITUAÇÃO': 'checagem'
        }
        
        stats = {'inseridos': 0, 'atualizados': 0, 'erros': 0}
        
        bar = st.progress(0)
        status_text = st.empty()
        
        for i, row in df_import.iterrows():
            # Preparar dados
            dados = {}
            for col_excel, col_db in map_cols.items():
                if col_excel in row:
                    val = row[col_excel]
                    # Tratamentos básicos
                    if pd.isna(val):
                        val = None
                    elif col_db in ['inicio_curso', 'fim_curso']:
                        try: val = pd.to_datetime(val).date()
                        except: val = None
                    elif col_db == 'porcentagem':
                        # Se vier como string "50%", converte. Se vier 0.5 mantem
                        if isinstance(val, str) and '%' in val:
                            try: val = float(val.strip('%').replace(',','.')) / 100
                            except: val = 0.5
                    elif col_db in ['mensalidade', 'valor_reembolso']:
                         if isinstance(val, str):
                            try: val = float(val.replace('R$','').replace('.','').replace(',','.').strip())
                            except: pass
                    
                    dados[col_db] = val
            
            # 1.5 Enriquecer com Organograma se tiver cod_local
            if 'cod_local' in dados and dados['cod_local'] and not df_org.empty:
                dir_org, _, _ = buscar_info_organograma(dados['cod_local'], df_org)
                if dir_org:
                    # Se não tiver diretoria ou se for N/A, usa a do organograma como prioridade
                    if not dados.get('diretoria') or dados.get('diretoria') in ['', 'N/A', 'None']:
                        dados['diretoria'] = dir_org

            # Cadastrar/Atualizar
            if 'matricula' in dados and dados['matricula']:
                dados['matricula'] = str(dados['matricula']).strip() # Garantir string
                if 'nome' in dados and dados['nome']:
                    dados['nome'] = str(dados['nome']).upper()
                    
                ok, msg = upsert_bolsista(dados, preserve_status=preserve_status)
                if ok: 
                    if "Atualizado" in msg: stats['atualizados'] += 1
                    else: stats['inseridos'] += 1
                else: 
                    stats['erros'] += 1
                    # st.error(f"Erro na linha {i}: {msg}")
            else:
                stats['erros'] += 1 # Sem matricula
            
            if i % 10 == 0:
                bar.progress((i + 1) / len(df_import))
                status_text.text(f"Processando {i+1}/{len(df_import)}...")
            
        bar.progress(100)
        status_text.empty()
        
        st.success(f"✅ Concluído! Inseridos: {stats['inseridos']} | Atualizados: {stats['atualizados']} | Erros/Ignorados: {stats['erros']}")
        
    except Exception as e:
        st.error(f"Erro ao processar dados: {e}")


def processar_importacao_historico(df, ano_padrao):
    conn = get_conn()
    # Forçar leitura fresca do organograma (sem cache)
    df_org = carregar_organograma()
    mapping = get_organograma_mapping(df_org)
    
    try:
        backup_database()
        
        # Padronizar nomes de colunas
        df.columns = [str(c).upper().strip() for c in df.columns]
        
        # Identificar colunas críticas
        col_data = next((c for c in ['DATA', 'PGTO', 'PAGTO', 'MÊS', 'MES'] if c in df.columns), None)
        col_mat = next((c for c in ['MATRICULA', 'MATRÍCULA', 'ID'] if c in df.columns), None)
        col_nome = next((c for c in ['NOMES', 'NOME', 'COLABORADOR'] if c in df.columns), None)
        col_valor = next((c for c in ['VALOR', 'VALOR LIQUIDO', 'LÍQUIDO', 'TOTAL'] if c in df.columns), None)
        col_cl = next((c for c in ['CODIGO LOCAL', 'CÓDIGO LOCAL', 'COD. LOCAL', 'COD LOCAL'] if c in df.columns), None)

        if not col_cl:
            st.warning("⚠️ Coluna 'CÓDIGO LOCAL' não encontrada no arquivo de pagamentos.")

        df_processado = []
        bar = st.progress(0)
        status = st.empty()
        
        total = len(df)
        for i, row in df.iterrows():
            try:
                # 1. Parsing da Data
                val_data = row[col_data] if col_data else None
                dt = None
                if isinstance(val_data, datetime): dt = val_data
                else:
                    try: dt = pd.to_datetime(val_data, dayfirst=True)
                    except: pass
                
                if not dt: continue
                
                mes, ano = dt.month, dt.year
                data_str = dt.strftime('%Y-%m-%d')
                
                # 2. Dados básicos
                matricula = str(row[col_mat]).split('.')[0].strip() if col_mat else "0"
                nome = str(row[col_nome]).upper().strip() if col_nome else "NÃO INFORMADO"
                
                # Valor
                v = row[col_valor] if col_valor else 0
                if isinstance(v, str):
                    try: v = float(v.replace('R$','').replace('.','').replace(',','.').strip())
                    except: v = 0.0
                valor = float(v) if pd.notna(v) else 0.0
                
                # Código Local
                cl = str(row[col_cl]).strip() if col_cl and pd.notna(row[col_cl]) else ""
                
                # Diretoria (Busca no organograma)
                diretoria, _, _ = buscar_info_organograma_fast(cl, mapping)
                
                # Ordem: matricula, nome, mes_referencia, data_pagamento, valor, ano, mes, cod_local, diretoria
                df_processado.append((
                    matricula, nome, f"{MESES[mes-1]}/{ano}", data_str, valor, ano, mes, cl, diretoria
                ))
            except:
                continue
            
            if i % 100 == 0:
                bar.progress((i+1)/total)
                status.text(f"Processando {i+1}/{total}...")

        if not df_processado:
            st.error("Nenhum dado válido processado.")
            return

        # 3. Salvar no Banco
        st.write(f"🧹 Limpando dados antigos e salvando {len(df_processado)} registros...")
        conn.execute("DELETE FROM historico_pagamentos")
        
        conn.executemany('''
            INSERT INTO historico_pagamentos (matricula, nome, mes_referencia, data_pagamento, valor, ano, mes, cod_local, diretoria)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', df_processado)
        
        conn.commit()
        st.cache_data.clear()
        st.success(f"✅ Sucesso! {len(df_processado)} registros vinculados ao Organograma.")
        st.balloons()
        st.rerun()
        
    except Exception as e:
        st.error(f"Erro crítico na importação: {e}")
    finally:
        conn.close()

def listar_bolsistas(situacao=None, diretoria=None, busca=None, ano_ref=None):
    conn = get_conn()
    query = "SELECT * FROM bolsistas WHERE 1=1"
    params = []
    if situacao and situacao != "Todos":
        query += " AND situacao = ?"
        params.append(situacao)
    if diretoria and diretoria != "Todas":
        query += " AND diretoria = ?"
        params.append(diretoria)
    if ano_ref and ano_ref != "Todos":
        query += " AND ano_referencia = ?"
        params.append(ano_ref)
    if busca:
        query += " AND (nome LIKE ? OR matricula LIKE ?)"
        params.extend([f'%{busca}%', f'%{busca}%'])
    query += " ORDER BY nome"
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df

def get_diretorias():
    """Busca as diretorias únicas do Organograma (Coluna C física)"""
    df_org = carregar_organograma()
    if df_org.empty:
        # Fallback para o banco se o organograma sumir
        conn = get_conn()
        df = pd.read_sql_query("SELECT DISTINCT diretoria FROM bolsistas WHERE diretoria IS NOT NULL AND diretoria != '' ORDER BY diretoria", conn)
        conn.close()
        return df['diretoria'].tolist()
    
    # Extrair Coluna C (índice 2)
    diretorias = df_org.iloc[:, 2].dropna().unique().tolist()
    # Adicionar as que já estão no banco de dados para segurança
    conn = get_conn()
    df_db = pd.read_sql_query("SELECT DISTINCT diretoria FROM bolsistas WHERE diretoria IS NOT NULL AND diretoria != ''", conn)
    conn.close()
    
    totais = sorted(list(set([str(d).upper().strip() for d in diretorias + df_db['diretoria'].tolist()])))
    return [d for d in totais if d not in ['NAN', 'NONE', 'N/A', '']]

def get_anos_referencia():
    """Busca anos de referência únicos no banco de forma robusta"""
    conn = get_conn()
    try:
        df = pd.read_sql_query("SELECT DISTINCT ano_referencia FROM bolsistas WHERE ano_referencia IS NOT NULL", conn)
        anos_brutos = df['ano_referencia'].unique().tolist()
        anos = []
        for a in anos_brutos:
            try:
                if not a: continue
                # Tratar strings de data (ex: 2024-01-01)
                if isinstance(a, str):
                    if '-' in a: a = a.split('-')[0]
                    elif '/' in a: a = a.split('/')[-1]
                    val = int(float(a))
                else:
                    val = int(a)
                anos.append(val)
            except:
                continue
        # Garantir 2026 e 2025
        res = set(anos)
        res.add(2025)
        res.add(2026)
        return sorted(list(res), reverse=True)
    except:
        return [2026, 2025]
    finally:
        conn.close()

def gerar_template_excel():
    """Gera um template Excel vazio com as colunas corretas para importação"""
    cols = ['MATRÍCULA', 'NOME', 'CPF', 'DIRETORIA', 'CURSO', 'INSTITUIÇÃO', 
            'TIPO', 'MODALIDADE', 'INÍCIO CURSO', 'FIM CURSO', 'ANO PROGRAMA', 
            'MENSALIDADE', '% BOLSA', 'VALOR REEMBOLSO', 'SITUAÇÃO', 'OBSERVAÇÃO']
    
    # Criar DataFrame vazio com uma linha de exemplo (opcional, ou vazio)
    df_template = pd.DataFrame(columns=cols)
    # Adicionar uma linha de exemplo para facilitar o entendimento
    df_template.loc[0] = ['123456', 'Exemplo Nome', '000.000.000-00', 'RH', 'Administração', 'Faculdade X', 
                          'GRADUAÇÃO', 'EAD', '01/01/2026', '31/12/2030', '2026', 
                          '1000.00', '50', '500.00', 'ATIVO', 'Cadastro via Importação']
    
    from io import BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_template.to_excel(writer, index=False, sheet_name='Modelo Importação')
        # Ajustar largura das colunas (opcional, mas bom pra UX)
        worksheet = writer.sheets['Modelo Importação']
        for idx, col in enumerate(df_template.columns):
            worksheet.column_dimensions[chr(65 + idx)].width = 15
            
    return output.getvalue()

def get_stats():
    conn = get_conn()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM bolsistas WHERE situacao = 'ATIVO'")
    ativos = cursor.fetchone()[0]
    
    # Média mensal da Safra 2025/2026 (Abr/25 a Mar/26)
    cursor.execute("""
        SELECT AVG(total_mensal)
        FROM (
            SELECT SUM(valor) as total_mensal
            FROM historico_pagamentos
            WHERE (ano = 2025 AND mes >= 4) OR (ano = 2026 AND mes <= 3)
            GROUP BY ano, mes
        )
    """)
    result = cursor.fetchone()
    inv = result[0] if result and result[0] else 0

    cursor.execute("SELECT COUNT(*) FROM bolsistas")
    total = cursor.fetchone()[0]
    conn.close()
    return {'ativos': ativos, 'investimento': inv, 'total': total}

def get_safra(ano, mes):
    """Retorna a safra no formato AAAA/AAAA baseado no mês"""
    if mes >= 4:  # Abril a Dezembro = primeira parte da safra
        return f"{ano}/{ano+1}"
    else:  # Janeiro a Março = segunda parte da safra
        return f"{ano-1}/{ano}"

def get_safras_disponiveis(df):
    """Retorna lista de safras disponíveis no dataframe"""
    safras = set()
    for _, row in df[['ano', 'mes']].drop_duplicates().iterrows():
        safras.add(get_safra(row['ano'], row['mes']))
    return sorted(list(safras), reverse=True)

DIRETORIAS = ["DIRETORIA AGRICOLA", "DIRETORIA INDUSTRIAL", "DIRETORIA ADMINISTRATIVA",
              "DIRETORIA GENTE E GESTAO", "DIRETORIA FINANCEIRA", "DIRETORIA CSC GRCI",
              "DIRETORIA COMERCIAL NOVOS PRODUTOS"]

# CSS MODERNO - AZUL E VERDE


def render_dashboard_geral():
    """Renderiza a página principal de Dashboard com indicadores completos."""
    st.markdown("### 📊 Dashboard Estratégico")
    
    # Carregar dados
    conn = get_conn()
    df_bolsistas = pd.read_sql_query("SELECT * FROM bolsistas WHERE situacao != 'INATIVO'", conn)
    
    # Busca TODO histórico para análises (sem travar em ano_atual-1, pois user quer safras)
    df_hist = pd.read_sql_query("SELECT mes, ano, valor FROM historico_pagamentos", conn)
    conn.close()
    
    # Tratamentos
    if not df_hist.empty:
        df_hist['data'] = df_hist.apply(lambda x: datetime(int(x['ano']), int(x['mes']), 1), axis=1)
        df_hist['safra'] = df_hist.apply(lambda x: get_safra(x['ano'], x['mes']), axis=1)
    
    # --- FILTRO SAFRA ---
    not_empty_hist = not df_hist.empty
    sel_val_label = "Todas"
    
    if not_empty_hist:
        col_radio, col_sel, _ = st.columns([1, 1, 2])
        
        with col_radio:
            tipo_filtro = st.radio("Filtrar por:", ["Safra", "Ano"], horizontal=True, label_visibility="collapsed")
            
        if tipo_filtro == "Safra":
            opts = sorted(df_hist['safra'].unique().tolist(), reverse=True)
            label = "📅 Selecione a Safra:"
            curr_col = 'safra'
        else:
            opts = sorted(df_hist['ano'].unique().tolist(), reverse=True)
            label = "📅 Selecione o Ano:"
            curr_col = 'ano'
            
        with col_sel:
            sel_val = st.selectbox(label, ["Todas"] + opts)
            sel_val_label = str(sel_val)
            
        if sel_val != "Todas":
            df_filtered_dash = df_hist[df_hist[curr_col] == sel_val]
        else:
            df_filtered_dash = df_hist
    else:
        df_filtered_dash = df_hist

        
    # KPI Cards Topo
    col1, col2, col3, col4 = st.columns(4)
    
    total_ativos = len(df_bolsistas[df_bolsistas['situacao'] == 'ATIVO'])
    
    # Nova Métrica: Total Pago no Período Selecionado (Baseado no histórico real)
    total_pago_periodo = df_filtered_dash['valor'].sum() if not df_filtered_dash.empty else 0
    
    avg_ticket = total_pago_periodo / len(df_filtered_dash) if len(df_filtered_dash) > 0 else 0
    if sel_val_label == "Todas" or not not_empty_hist:
         # Se for todas, avg_ticket fica estranho somado tudo. Melhor manter logica anterior ou apenas snapshot?
         # Vamos manter snapshot para ticket medio, mas usar total pago para o card 2
         total_invest_mensal_snap = df_bolsistas[df_bolsistas['situacao'] == 'ATIVO']['valor_reembolso'].sum()
         avg_ticket = total_invest_mensal_snap / total_ativos if total_ativos > 0 else 0
    
    # % em Graduação
    ativos_grad = len(df_bolsistas[(df_bolsistas['situacao'] == 'ATIVO') & (df_bolsistas['tipo'] == 'GRADUACAO')])
    pct_grad = (ativos_grad / total_ativos * 100) if total_ativos > 0 else 0
    
    with col1:
        render_modern_metric(
            "",
            "Bolsas Ativas (Hoje)",
            format_br_number(total_ativos),
            color="#30515F",  # Azul-cinza escuro
            bg_gradient=""
        )
        
    with col2:
        titulo_card = f"Total Pago ({sel_val_label})" if not_empty_hist else "Total Pago"
        render_modern_metric(
            "",
            titulo_card,
            format_br_currency(total_pago_periodo),
            color="#76B82A",  # Verde corporativo
            bg_gradient=""
        )
        
    with col3:
        render_modern_metric(
            "",
            "Ticket Médio",
            format_br_currency(avg_ticket),
            color="#EF7D00",  # Laranja corporativo
            bg_gradient=""
        )
        
    with col4:
        render_modern_metric(
            "",
            "% em Graduação",
            f"{pct_grad:.1f}%",
            color="#FFCC00",  # Amarelo corporativo
            bg_gradient=""
        )
    
    st.markdown("---")
    
    # ----------------------------------------------------
    # ROW 1: Charts Principais
    # ----------------------------------------------------
    c1, c2 = st.columns([2, 1])
    
    with c1:
        st.subheader("Investimento Estimado por Diretoria")
        if not df_bolsistas.empty and 'diretoria' in df_bolsistas.columns:
            # Filtrar apenas bolsistas ativos com diretoria válida
            df_ativos_dir = df_bolsistas[
                (df_bolsistas['situacao'] == 'ATIVO') & 
                (df_bolsistas['diretoria'].notna()) & 
                (df_bolsistas['diretoria'].str.strip() != '')
            ].copy()
            
            # Agrupar por diretoria
            if not df_ativos_dir.empty:
                df_dir = df_ativos_dir.groupby('diretoria').agg({
                    'valor_reembolso': 'sum',
                    'matricula': 'count'
                }).reset_index().sort_values('valor_reembolso', ascending=True)
            else:
                df_dir = pd.DataFrame()
            
            # Criar gráfico apenas se houver dados
            if not df_dir.empty:
                render_bar_chart(
                    df_dir, 
                    x_col='diretoria', 
                    y_col='valor_reembolso', 
                    title="", 
                    label_y="Valor Mensal"
                )
            else:
                st.info("📊 Sem dados de diretoria válidos para exibir.")
            
    with c2:
        st.subheader("Perfil dos Bolsistas")
        if not df_bolsistas.empty:
            tab_tipo, tab_mod = st.tabs(["Nível", "Modalidade"])
            
            with tab_tipo:
                if 'tipo' in df_bolsistas.columns:
                    df_ipo = df_bolsistas['tipo'].value_counts().reset_index()
                    df_ipo.columns = ['Tipo', 'Qtd']
                    # Degradê azul corporativo: do azul escuro ao claro
                    cores_azul = ['#1e3a5f', '#2c5282', '#3b6ba8', '#5a8ac7', '#7da9d9']
                    fig_pie1 = px.pie(df_ipo, values='Qtd', names='Tipo', hole=0.4, color_discrete_sequence=cores_azul)
                    fig_pie1.update_traces(textposition='inside', textinfo='percent+label')
                    fig_pie1.update_layout(showlegend=False, margin=dict(t=0, b=0, l=0, r=0), height=300, paper_bgcolor="white", plot_bgcolor="white")
                    st.plotly_chart(fig_pie1, use_container_width=True, config=PLOTLY_CONFIG)
            
            with tab_mod:
                if 'modalidade' in df_bolsistas.columns:
                    df_mod = df_bolsistas['modalidade'].value_counts().reset_index()
                    df_mod.columns = ['Modalidade', 'Qtd']
                    # Degradê azul corporativo: do azul escuro ao claro
                    cores_azul = ['#1e3a5f', '#2c5282', '#3b6ba8', '#5a8ac7', '#7da9d9']
                    fig_pie2 = px.pie(df_mod, values='Qtd', names='Modalidade', hole=0.4, color_discrete_sequence=cores_azul)
                    fig_pie2.update_traces(textposition='inside', textinfo='percent+label')
                    fig_pie2.update_layout(showlegend=False, margin=dict(t=0, b=0, l=0, r=0), height=300, paper_bgcolor="white", plot_bgcolor="white")
                    st.plotly_chart(fig_pie2, use_container_width=True, config=PLOTLY_CONFIG)

    # ----------------------------------------------------
    # ROW 2: Evolução e Top Cursos
    # ----------------------------------------------------
    c3, c4 = st.columns([2, 1])
    
    with c3:
        st.subheader("Evolução de Pagamentos (Realizado)")
        if not df_filtered_dash.empty:
            df_evo = df_filtered_dash.groupby('data')['valor'].sum().reset_index().sort_values('data')
            render_area_chart(df_evo, 'data', 'valor', "")
            # (Chart rendered inside function)
        else:
            st.info("Sem dados históricos de pagamento.")
            
    with c4:
        st.subheader("Top 5 Instituições")
        if not df_bolsistas.empty and 'instituicao' in df_bolsistas.columns:
            df_inst = df_bolsistas['instituicao'].value_counts().head(5).reset_index()
            df_inst.columns = ['Instituição', 'Alunos']
            # Sort para bar h
            df_inst = df_inst.sort_values('Alunos', ascending=True) 
            
            render_bar_chart(
                df_inst, 
                x_col='Instituição', 
                y_col='Alunos', 
                title="", 
                label_y="Alunos"
            )

    st.markdown("---")
    
    # ----------------------------------------------------
    # ROW 3: Top Cursos e Custos
    # ----------------------------------------------------
    st.subheader("Análise de Cursos (Top 10)")
    
    if not df_bolsistas.empty and 'curso' in df_bolsistas.columns:
        col_select, _ = st.columns([1,3])
        with col_select:
            metrica = st.selectbox("Métrica:", ["Quantidade de Alunos", "Custo Médio (R$)"])
        
        if metrica == "Quantidade de Alunos":
            df_curso = df_bolsistas['curso'].value_counts().head(10).reset_index()
            df_curso.columns = ['Curso', 'Valor']
            cor = 'Valor'
            fmt = '%{text}'
        else:
            df_curso = df_bolsistas.groupby('curso')['valor_reembolso'].mean().sort_values(ascending=False).head(10).reset_index()
            df_curso.columns = ['Curso', 'Valor']
            cor = 'Valor'
            fmt = 'R$ %{text:,.2f}'
            
        df_curso = df_curso.sort_values('Valor', ascending=True)
            
        render_bar_chart(
            df_curso, 
            x_col='Curso', 
            y_col='Valor', 
            title="", 
            label_y=metrica
        )


def criar_super_tabela(df, key="tabela"):
    """Cria tabela interativa com AgGrid, com espaçamento melhorado e edição salva no BD."""
    if AGGRID and len(df) > 0:
        # Preparar dados
        df_display = df.copy()
        
        # Renomear colunas
        col_map = {
            'matricula': 'Matrícula',
            'nome': 'Nome',
            'cpf': 'CPF',
            'diretoria': 'Diretoria',
            'cod_local': 'Cod. Local',
            'curso': 'Curso',
            'instituicao': 'Instituição',
            'tipo': 'Tipo',
            'modalidade': 'Modalidade',
            'inicio_curso': 'Início Curso',
            'fim_curso': 'Fim Curso',
            'ano_referencia': 'Ano Programa',
            'mensalidade': 'Mensalidade',
            'porcentagem': '% Bolsa',
            'valor_reembolso': 'Valor Reembolso',
            'situacao': 'Situação',
            'checagem': 'Checagem',
            'observacao': 'Observação'
        }
        
        # Selecionar colunas importantes (incluindo diretoria, cod_local e observacao)
        cols = ['matricula', 'nome', 'cpf', 'diretoria', 'cod_local', 'curso', 'instituicao', 'tipo', 'modalidade', 'inicio_curso', 'fim_curso', 'ano_referencia', 'mensalidade', 'porcentagem', 'valor_reembolso', 'situacao', 'checagem', 'observacao']
        cols = [c for c in cols if c in df_display.columns]
        df_display = df_display[cols].copy()
        df_display.columns = [col_map.get(c, c) for c in cols]
        
        # Formatar datas para padrão brasileiro DD/MM/YYYY
        if 'Início Curso' in df_display.columns:
            df_display['Início Curso'] = pd.to_datetime(df_display['Início Curso'], errors='coerce').dt.strftime('%d/%m/%Y')
            df_display['Início Curso'] = df_display['Início Curso'].fillna('')
        if 'Fim Curso' in df_display.columns:
            df_display['Fim Curso'] = pd.to_datetime(df_display['Fim Curso'], errors='coerce').dt.strftime('%d/%m/%Y')
            df_display['Fim Curso'] = df_display['Fim Curso'].fillna('')
        
        # Formatar valores monetários e percentuais
        if 'Mensalidade' in df_display.columns:
            df_display['Mensalidade'] = df_display['Mensalidade'].apply(lambda x: f"R$ {x:,.2f}" if pd.notna(x) else "")
        if 'Valor Reembolso' in df_display.columns:
            df_display['Valor Reembolso'] = df_display['Valor Reembolso'].apply(lambda x: f"R$ {x:,.2f}" if pd.notna(x) else "")
        if '% Bolsa' in df_display.columns:
            df_display['% Bolsa'] = df_display['% Bolsa'].apply(lambda x: f"{x*100:.0f}%" if pd.notna(x) else "")
        
        # Configurar AgGrid com layout aprimorado (SEM QUEBRA DE LINHA)
        gb = GridOptionsBuilder.from_dataframe(df_display)
        gb.configure_default_column(
            filterable=True,
            sortable=True,
            resizable=True,
            wrapText=False,
            autoHeight=False,
            minWidth=100,
            cellStyle={
                'padding': '4px 8px',
                'textOverflow': 'ellipsis',
                'overflow': 'hidden',
                'whiteSpace': 'nowrap'
            }
        )
        gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=50)
        gb.configure_selection('single', use_checkbox=False)
        
        # Larguras específicas e configuração de colunas
        gb.configure_column("Matrícula", width=130, pinned='left', editable=False)  # Matrícula não editável (é a chave)
        gb.configure_column("Nome", minWidth=250, pinned='left', editable=True)
        gb.configure_column("CPF", minWidth=130, editable=True)
        gb.configure_column("Diretoria", minWidth=150, editable=True)
        gb.configure_column("Cod. Local", minWidth=120, editable=True)
        gb.configure_column("Curso", minWidth=200, editable=True)
        gb.configure_column("Instituição", minWidth=200, editable=True)
        
        # Tipo com dropdown de opções
        gb.configure_column(
            "Tipo", 
            minWidth=150,
            editable=True,
            cellEditor='agSelectCellEditor',
            cellEditorParams={'values': ['', 'GRADUACAO', 'TECNICO', 'POS-GRADUACAO', 'MBA']}
        )
        
        # Modalidade com dropdown de opções
        gb.configure_column(
            "Modalidade", 
            minWidth=150,
            editable=True,
            cellEditor='agSelectCellEditor',
            cellEditorParams={'values': ['', 'PRESENCIAL', 'EAD', 'HÍBRIDO', 'SEMIPRESENCIAL']}
        )
        
        # Datas e ano editáveis
        gb.configure_column("Início Curso", minWidth=120, editable=True)
        gb.configure_column("Fim Curso", minWidth=120, editable=True)
        gb.configure_column("Ano Programa", minWidth=100, editable=True)
        
        # Valores monetários e percentuais editáveis
        gb.configure_column("Mensalidade", minWidth=130, editable=True)
        gb.configure_column("% Bolsa", minWidth=100, editable=True)
        gb.configure_column("Valor Reembolso", minWidth=140, editable=True)
        
        # Situação com dropdown
        gb.configure_column(
            "Situação",
            minWidth=120,
            editable=True,
            cellEditor='agSelectCellEditor',
            cellEditorParams={'values': ['ATIVO', 'INATIVO', 'CONCLUIDO', 'EM ANALISE']}
        )
        
        if AGGRID:
            try:
                # ---------------------------------------------------------
                # BLINDAGEM DE DADOS: Garantir que TUDO seja serializável
                # ---------------------------------------------------------
                # 1. Converter nomes de colunas para string
                df_display.columns = [str(c) for c in df_display.columns]
                
                # 2. Converter todos os valores de colunas tipo object/mixed para string
                for col in df_display.columns:
                    if not pd.api.types.is_numeric_dtype(df_display[col]):
                        df_display[col] = df_display[col].fillna('').astype(str)
                
                # Checagem com dropdown e estilo (JsCode)
                cell_style_jscode = JsCode("""
                function(params) {
                    if (!params.value) return {};
                    var v = params.value.toString().toUpperCase();
                    if (v === 'CANCELADO' || v === 'DESISTENCIA' || v === 'TRANCADO') {
                        return {'backgroundColor': '#fee2e2', 'color': '#991b1b', 'fontWeight': 'bold'};
                    } else if (v === 'IRREGULAR') {
                        return {'backgroundColor': '#fef3c7', 'color': '#92400e', 'fontWeight': 'bold'};
                    } else if (v === 'CONCLUIDO') {
                        return {'backgroundColor': '#dbeafe', 'color': '#1e40af', 'fontWeight': 'bold'};
                    } else if (v === 'REGULAR') {
                        return {'backgroundColor': '#dcfce7', 'color': '#166534', 'fontWeight': 'bold'};
                    }
                    return {};
                }
                """)

                gb.configure_column(
                    "Checagem",
                    minWidth=140,
                    editable=True,
                    cellEditor='agSelectCellEditor',
                    cellEditorParams={'values': ['REGULAR', 'IRREGULAR', 'CONCLUIDO', 'CANCELADO', 'DESISTENCIA', 'TRANCADO', 'DEMITIDO']},
                    cellStyle=cell_style_jscode
                )
                
                # Observação - campo de texto livre
                gb.configure_column("Observação", editable=True, cellEditor='agLargeTextCellEditor', minWidth=250)
                
                grid_options = gb.build()
                grid_options.update({"rowHeight": 36, "headerHeight": 36, "suppressRowHoverHighlight": True})

                # Renderizar AgGrid com parâmetros de segurança e Enums corretos
                grid_response = AgGrid(
                    df_display,
                    gridOptions=grid_options,
                    height=800,
                    theme='balham',  # Voltar para balham mas com CSS forçado
                    update_mode=GridUpdateMode.VALUE_CHANGED,
                    data_return_mode=DataReturnMode.AS_INPUT,
                    fit_columns_on_grid_load=False,
                    allow_unsafe_jscode=True,
                    key=key
                )
                df_edited = grid_response['data']
            except Exception as e:
                # Fallback limpo sem mensagem de erro assustadora
                df_edited = st.data_editor(df_display, key=f"editor_fallback_{key}", use_container_width=True, hide_index=True)
        else:
            st.info("ℹ️ Usando visualização básica de tabela.")
            df_edited = st.data_editor(df_display, key=f"editor_basic_{key}", use_container_width=True, hide_index=True)
        st.info("ℹ️ Nota: As alterações feitas aqui são salvas no **Banco de Dados do Sistema** (`bolsas.db`).")
        if st.button("💾 Salvar alterações", type="primary"):
            edited_df = df_edited.copy()
            
            # Reverter formatação para valores crus antes de gravar
            if 'Mensalidade' in edited_df.columns:
                edited_df['Mensalidade'] = edited_df['Mensalidade'].astype(str).str.replace(r"R\$ ", "", regex=True).str.replace(".", "", regex=True).str.replace(",", ".", regex=True)
                edited_df['Mensalidade'] = pd.to_numeric(edited_df['Mensalidade'], errors='coerce').fillna(0)
            if 'Valor Reembolso' in edited_df.columns:
                edited_df['Valor Reembolso'] = edited_df['Valor Reembolso'].astype(str).str.replace(r"R\$ ", "", regex=True).str.replace(".", "", regex=True).str.replace(",", ".", regex=True)
                edited_df['Valor Reembolso'] = pd.to_numeric(edited_df['Valor Reembolso'], errors='coerce').fillna(0)
            if '% Bolsa' in edited_df.columns:
                edited_df['% Bolsa'] = edited_df['% Bolsa'].astype(str).str.rstrip('%')
                edited_df['% Bolsa'] = pd.to_numeric(edited_df['% Bolsa'], errors='coerce').fillna(50) / 100
            
            # Atualizar banco de dados - MAPEAMENTO COMPLETO
            conn = get_conn()
            for _, row in edited_df.iterrows():
                update_fields = []
                params = []
                mapping = {
                    "Nome": "nome", "CPF": "cpf", "Diretoria": "diretoria", "Cod. Local": "cod_local",
                    "Curso": "curso", "Instituição": "instituicao", "Tipo": "tipo", "Modalidade": "modalidade",
                    "Início Curso": "inicio_curso", "Fim Curso": "fim_curso", "Ano Programa": "ano_referencia",
                    "Mensalidade": "mensalidade", "% Bolsa": "porcentagem", "Valor Reembolso": "valor_reembolso",
                    "Situação": "situacao", "Checagem": "checagem", "Observação": "observacao"
                }
                for col, db_col in mapping.items():
                    if col in row.index:
                        val = row[col]
                        if pd.isna(val) or val == '' or val == 'nan': val = None
                        update_fields.append(f"{db_col} = ?")
                        params.append(val)
                if update_fields:
                    sql = f"UPDATE bolsistas SET {', '.join(update_fields)} WHERE matricula = ?"
                    params.append(row['Matrícula'])
                    conn.execute(sql, tuple(params))
            conn.commit()
            conn.close()
            st.success("✅ Alterações salvas no banco de dados!")
            st.rerun()
    else:
        # Fallback para dataframe padrão
        st.dataframe(df, width='stretch', height=800, hide_index=True)

def main():
    # ---------------------------------------------------------------------------
    # Autenticação
    # ---------------------------------------------------------------------------
    if not check_authentication():
        login_page()
        return

    with st.sidebar:
        st.write(f"Usuário: **{st.session_state.get('username', 'gestao')}**")
        if st.button("🔒 Sair / Logout", use_container_width=True):
            logout()

    import plotly.graph_objects as go
    stats = get_stats()
    
    # Header com stats
    render_header(stats)
    
    # Menu horizontal moderno
    selected = option_menu(
        menu_title=None,
        options=["Tabela", "Conferência", "Histórico", "Perfil", "Pagamentos", "Cadastrar"],
        icons=['table', 'check-circle', 'clock-history', 'person', 'credit-card', 'plus-circle'],
        menu_icon="cast",
        default_index=0,
        orientation="horizontal",
        styles={
            "container": {"padding": "0!important", "background-color": "#ffffff", "border": "1px solid #e2e8f0", "border-radius": "8px"},
            "icon": {"color": "#475569", "font-size": "14px"}, 
            "nav-link": {"font-size": "14px", "text-align": "center", "margin": "0px", "--hover-color": "#f1f5f9", "color": "#475569"},
            "nav-link-selected": {"background-color": "#2563eb", "color": "#ffffff"},
        },
        key="main_menu"
    )
    
    # Compatibilidade com lógica existente
    # Mapear nome limpo de volta para o nome com emoji esperado pelo resto do código? 
    # Melhor: Alterar o resto do código para usar os nomes limpos.
    # Mas como são muitas seções, vou criar um mapa para manter o código abaixo funcionando sem mega refatoração agora.
    
    map_menus = {
        "Tabela": "📋 Tabela",
        "Conferência": "💰 Conferência",
        "Histórico": "📅 Histórico",
        "Perfil": "👤 Perfil",
        "Pagamentos": "💳 Pagamentos",
        "Cadastrar": "➕ Cadastrar"
    }
    
    # Se o menu for alterado pelo option_menu
    menu = map_menus.get(selected, "📋 Tabela")
    
    st.markdown("---")
    
    # =============================================
    # DASHBOARD GERAL
    # =============================================
    # =============================================
    # TABELA GERAL
    # =============================================
    if menu == "📋 Tabela":
        col_t, col_btn = st.columns([4, 1])
        with col_t:
            st.markdown("### 📋 Base de Bolsistas")
        with col_btn:
            # Opção de sobrescrever ou manter
            sobrescrever = st.checkbox("Sobrescrever Status/Obs?", value=False, help="Se marcado, o Excel substituirá os Status e Observações do sistema. Se desmarcado, mantém o que está no sistema atual.")
            
            if st.button("🔄 Atualizar Base", type="primary", use_container_width=True, help="Sincroniza com Google Sheets ou Excel Local"):
                try:
                    # Limpar cache antes de atualizar para garantir dados frescos
                    st.cache_data.clear()
                    
                    with st.spinner("Sincronizando dados..."):
                         df_local = get_dataset("BOLSAS")
                         
                         if not df_local.empty:
                            st.toast(f"Dados brutos carregados: {len(df_local)} linhas.", icon="📥")
                            st.write(f"Colunas encontradas: {list(df_local.columns)}") # Debug temporário
                            processar_importacao_df(df_local, preserve_status=not sobrescrever)
                            st.balloons()
                            st.rerun()
                         else:
                            st.error("Não foi possível carregar os dados. Verifique a conexão com o Google Sheets e se a planilha não está vazia.")
                            st.warning("Se estiver usando arquivo local, verifique se ele existe na pasta correta.")
                except Exception as e:
                    st.error(f"Erro ao sincronizar: {e}")
        
        # ---------------------------------------------------------
        # IMPORTAÇÃO DE NOVOS INSCRITOS (TEMPLATE + UPLOAD)
        # ---------------------------------------------------------
        with st.expander("📂 Importar Novos Inscritos / Baixar Modelo", expanded=False):
            st.info("Utilize esta área para cadastrar novos bolsistas em lote (ex: Novos de 2026).")
            
            c_mod, c_up = st.columns([1, 2])
            
            with c_mod:
                st.markdown("##### 1. Baixar Modelo")
                st.markdown("Baixe a planilha padrão para preenchimento.")
                excel_template = gerar_template_excel()
                st.download_button(
                    label="📥 Baixar Planilha Modelo",
                    data=excel_template,
                    file_name="Modelo_Cadastro_Bolsistas.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with c_up:
                st.markdown("##### 2. Importar Preenchida")
                st.markdown("Faça o upload da planilha preenchida para cadastro.")
                arquivo_novos = st.file_uploader("Selecione o arquivo Excel (.xlsx)", type=['xlsx'], key="upload_novos")
                if arquivo_novos is not None:
                    if st.button("📤 Processar Importação", type="primary"):
                        try:
                            df_novos = pd.read_excel(arquivo_novos)
                            processar_importacao_df(df_novos)
                            st.success("Importação concluída com sucesso!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erro ao processar arquivo: {e}")

        st.markdown("---")

        # Filtros em linha
        if 'filtros_version' not in st.session_state:
            st.session_state.filtros_version = 0
            
        st.markdown('<div class="filter-section">', unsafe_allow_html=True)
        col1, col2, col3, col4, col5, col_reset = st.columns([1, 1, 1, 1, 1.5, 0.5])
        with col1:
            situacao = st.selectbox("Situação", ["Todos", "ATIVO", "INATIVO", "CONCLUIDO", "EM ANÁLISE"], key=f"sit_{st.session_state.filtros_version}")
        with col2:
            anos_ref = get_anos_referencia()
            ano_selecionado = st.selectbox("Ano Safra", ["Todos"] + anos_ref, key=f"ano_{st.session_state.filtros_version}")
        with col3:
            diretorias = get_diretorias()
            diretoria = st.selectbox("Diretoria", ["Todas"] + diretorias, key=f"dir_{st.session_state.filtros_version}")
        with col4:
            ordem = st.selectbox("Ordenar por", ["Nome", "Matrícula", "Valor", "Diretoria", "Ano Ref."], key=f"ord_{st.session_state.filtros_version}")
        with col5:
            busca = st.text_input("🔍 Buscar", placeholder="Digite nome ou matrícula...", key=f"bus_{st.session_state.filtros_version}")
        with col_reset:
            st.write("") # alinhamento
            st.write("")
            if st.button("🔄 Limpar", help="Resetar todos os filtros", use_container_width=True):
                st.session_state.filtros_version += 1
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
        
        # 1. Buscar dados básicos (Filtros de Situação, Ano e Busca no Banco)
        df = listar_bolsistas(
            situacao=situacao if situacao != "Todos" else None,
            diretoria=None, # Filtrar diretoria depois do enriquecimento para maior precisão
            busca=busca if busca else None,
            ano_ref=ano_selecionado if ano_selecionado != "Todos" else None
        )
        
        # 2. Enriquecer com Diretoria e Gestores do ORGANOGRAMA via Cod. Local
        # Isso garante que mesmo quem está com "N/A" no cadastro seja encontrado pela diretoria certa
        df_org = carregar_organograma()
        df = enriquecer_com_organograma(df, df_org)
        
        # 3. Aplicar filtro de diretoria agora que temos os dados reais do Organograma
        if diretoria != "Todas":
            df = df[df['diretoria'].astype(str).str.upper() == diretoria.upper()]
        
        # 4. Ordenar
        if ordem == "Nome":
            df = df.sort_values('nome')
        elif ordem == "Matrícula":
            df = df.sort_values('matricula')
        elif ordem == "Valor":
            df = df.sort_values('valor_reembolso', ascending=False)
        elif ordem == "Diretoria":
            df = df.sort_values('diretoria')
        elif ordem == "Ano Ref.":
            df = df.sort_values('ano_referencia', ascending=False)
        
        # Estatísticas rápidas
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Encontrados", len(df))
        with col2:
            st.metric("Soma Reembolso", f"R$ {df['valor_reembolso'].sum():,.2f}")
        with col3:
            st.metric("Média Reembolso", f"R$ {df['valor_reembolso'].mean():,.2f}" if len(df) > 0 else "R$ 0")
        
        st.markdown("---")
        
        # SUPER TABELA
        if len(df) > 0:
            criar_super_tabela(df, "tabela_geral")
            
            # Botão de download Excel
            excel_data = df_to_excel(df)
            st.download_button("⬇️ Baixar Tabela", excel_data, "bolsistas.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
            # Seção de exclusão
            st.markdown("---")
            with st.expander("🗑️ Excluir Bolsista da Base", expanded=False):
                st.warning("⚠️ **Atenção:** A exclusão é permanente e não pode ser desfeita!")
                
                # Criar lista de opções para seleção
                opcoes_bolsistas = df[['matricula', 'nome']].apply(
                    lambda x: f"{x['matricula']} - {x['nome']}", axis=1
                ).tolist()
                
                bolsista_selecionado = st.selectbox(
                    "Selecione o bolsista para excluir:",
                    options=[""] + opcoes_bolsistas,
                    key="select_excluir"
                )
                
                if bolsista_selecionado:
                    matricula_excluir = bolsista_selecionado.split(" - ")[0]
                    nome_excluir = " - ".join(bolsista_selecionado.split(" - ")[1:])
                    
                    st.error(f"🚨 Você está prestes a excluir: **{nome_excluir}** (Matrícula: {matricula_excluir})")
                    
                    # Checkbox de confirmação
                    confirmar = st.checkbox(
                        f"Confirmo que desejo excluir permanentemente {nome_excluir}",
                        key="confirmar_exclusao"
                    )
                    
                    if confirmar:
                        if st.button("🗑️ Excluir Permanentemente", type="primary", key="btn_excluir"):
                            try:
                                conn = get_conn()
                                cursor = conn.cursor()
                                
                                # Buscar ID do bolsista
                                cursor.execute("SELECT id FROM bolsistas WHERE matricula = ?", (matricula_excluir,))
                                result = cursor.fetchone()
                                
                                if result:
                                    bolsista_id = result[0]
                                    
                                    # Excluir pagamentos relacionados
                                    cursor.execute("DELETE FROM pagamentos WHERE bolsista_id = ?", (bolsista_id,))
                                    
                                    # Excluir observações relacionadas
                                    cursor.execute("DELETE FROM observacoes WHERE bolsista_id = ?", (bolsista_id,))
                                    
                                    # Excluir bolsista
                                    cursor.execute("DELETE FROM bolsistas WHERE id = ?", (bolsista_id,))
                                    
                                    conn.commit()
                                    conn.close()
                                    
                                    st.success(f"✅ Bolsista **{nome_excluir}** excluído com sucesso!")
                                    st.balloons()
                                    st.rerun()
                                else:
                                    conn.close()
                                    st.error("Bolsista não encontrado!")
                            except Exception as e:
                                st.error(f"Erro ao excluir: {e}")
        else:
            st.info("Nenhum registro encontrado com os filtros selecionados.")
    
    # =============================================
    # CONFERÊNCIA
    # =============================================
    elif menu == "💰 Conferência":
        st.markdown("### 💰 Conferência Mensal")
        
        # FILTROS NO TOPO
        if 'filtros_version_conf' not in st.session_state:
            st.session_state.filtros_version_conf = 0
            
        col1, col2, col3, col4, col_reset = st.columns([1, 1, 1, 1.5, 0.5])
        with col1:
            mes = st.selectbox("📅 Mês", MESES, index=datetime.now().month-1, key=f"mes_conf_{st.session_state.filtros_version_conf}")
            mes_num = MESES.index(mes) + 1
        with col2:
            ano = st.selectbox("📅 Ano", [2026], index=0, key=f"ano_conf_{st.session_state.filtros_version_conf}")
        with col3:
            filtro_situacao = st.selectbox("📋 Situação", ["ATIVO", "CONCLUIDO", "INATIVO", "EM ANÁLISE", "Todos"], key=f"sit_conf_{st.session_state.filtros_version_conf}")
        with col4:
            busca_conf = st.text_input("🔍 Buscar", placeholder="Nome ou matrícula...", key=f"bus_conf_{st.session_state.filtros_version_conf}")
        with col_reset:
            st.write("") # alinhamento
            st.write("")
            if st.button("🔄 Limpar", help="Resetar todos os filtros da conferência", use_container_width=True, key="btn_reset_conf"):
                st.session_state.filtros_version_conf += 1
                st.rerun()
        
        # Buscar bolsistas com base no filtro de situação
        if filtro_situacao == "Todos":
            df_base = listar_bolsistas()
        else:
            df_base = listar_bolsistas(situacao=filtro_situacao)
        
        # Aplicar busca
        if busca_conf:
            df_base = df_base[
                df_base['nome'].str.contains(busca_conf, case=False, na=False) |
                df_base['matricula'].str.contains(busca_conf, na=False)
            ]
        
        if len(df_base) > 0:
            conn = get_conn()
            pagamentos = pd.read_sql_query(
                "SELECT bolsista_id, status, valor FROM pagamentos WHERE mes = ? AND ano = ?",
                conn, params=(int(mes_num), int(ano))
            )
            conn.close()
            
            pagos_ids = pagamentos[pagamentos['status'] == 'PAGO']['bolsista_id'].tolist()
            pend_ids = pagamentos[pagamentos['status'] == 'PENDENTE']['bolsista_id'].tolist()
            
            # Criar coluna de status para a tabela
            # FORÇAR CAST PARA INT E SET PARA EVITAR ERROS DE TIPO
            try:
                pagos_ids_set = set(pagamentos[pagamentos['status'] == 'PAGO']['bolsista_id'].astype(int).tolist())
                pend_ids_set = set(pagamentos[pagamentos['status'] == 'PENDENTE']['bolsista_id'].astype(int).tolist())
            except:
                pagos_ids_set = set()
                pend_ids_set = set()
            
            df_base['status_conf'] = df_base['id'].apply(
                lambda x: "✅ PAGO" if int(x) in pagos_ids_set else ("❌ PENDENTE" if int(x) in pend_ids_set else "⏳ AGUARDANDO")
            )
            
            total = len(df_base)
            pagos_count = len([x for x in df_base['id'] if int(x) in pagos_ids_set])
            pend_count = len([x for x in df_base['id'] if int(x) in pend_ids_set])
            aguard = total - pagos_count - pend_count
            
            # Métricas
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("📋 Total", total)
            with col2:
                st.metric("✅ Pagos", pagos_count)
            with col3:
                st.metric("❌ Pendentes", pend_count)
            with col4:
                st.metric("⏳ Conferir", aguard)
            
            prog = ((pagos_count + pend_count) / total * 100) if total > 0 else 0
            st.progress(prog / 100)
            st.caption(f"{prog:.0f}% conferido de {mes}/{ano}")
            
            st.markdown("---")
            
            tab1, tab2, tab3 = st.tabs(["📝 Conferir", "✅ Pagos", "📊 Relatório DP"])
            
            with tab1:
                # Filtro por status de conferência
                filtro_status = st.radio("Filtrar:", ["⏳ Aguardando", "Todos", "✅ Pagos", "❌ Pendentes"], horizontal=True)
                
                df_conf = df_base.copy()
                
                if filtro_status == "⏳ Aguardando":
                    df_conf = df_conf[df_conf['status_conf'] == '⏳ AGUARDANDO']
                elif filtro_status == "✅ Pagos":
                    df_conf = df_conf[df_conf['status_conf'] == '✅ PAGO']
                elif filtro_status == "❌ Pendentes":
                    df_conf = df_conf[df_conf['status_conf'] == '❌ PENDENTE']
                
                st.caption(f"{len(df_conf)} colaboradores")
                
                if len(df_conf) > 0:
                    # TABELA DE CONFERÊNCIA EDITÁVEL
                    st.markdown("#### 📋 Tabela de Conferência (edite Valor e Status)")
                    
                    # Preparar tabela para edição com Status editável
                    df_edit = df_conf[['id', 'matricula', 'nome', 'mensalidade', 'porcentagem', 'valor_reembolso', 'status_conf']].copy()
                    df_edit['porcentagem_display'] = df_edit['porcentagem'].apply(lambda x: f"{x*100:.0f}%" if pd.notna(x) else "50%")
                    # Simplificar status para dropdown
                    df_edit['status_edit'] = df_edit['status_conf'].apply(lambda x: 'PAGO' if 'PAGO' in x else ('PENDENTE' if 'PENDENTE' in x else 'AGUARDANDO'))
                    
                    # Configurar coluna de status como dropdown
                    column_config = {
                        'Status': st.column_config.SelectboxColumn(
                            'Status',
                            options=['AGUARDANDO', 'PAGO', 'PENDENTE'],
                            required=True
                        )
                    }
                    
                    if 'table_key_version' not in st.session_state:
                         st.session_state.table_key_version = 0

                    # Tabela editável
                    edited_df = st.data_editor(
                        df_edit[['matricula', 'nome', 'mensalidade', 'porcentagem_display', 'valor_reembolso', 'status_edit']].rename(columns={
                            'matricula': 'Matrícula',
                            'nome': 'Nome', 
                            'mensalidade': 'Mensalidade',
                            'porcentagem_display': '% Bolsa',
                            'valor_reembolso': 'Valor',
                            'status_edit': 'Status'
                        }),
                        disabled=['Matrícula', 'Nome', 'Mensalidade', '% Bolsa'],
                        column_config=column_config,
                        use_container_width=True,
                        height=700,
                        key=f"tabela_edit_{st.session_state.table_key_version}"
                    )
                    
                    # Botão para salvar alterações da tabela
                    if st.button("💾 SALVAR ALTERAÇÕES DA TABELA", type="primary", use_container_width=True):
                        conn = get_conn()
                        for i, (idx, row) in enumerate(df_conf.iterrows()):
                            valor = edited_df.iloc[i]['Valor']
                            status = edited_df.iloc[i]['Status']
                            if status in ['PAGO', 'PENDENTE']:
                                conn.execute('INSERT OR REPLACE INTO pagamentos (bolsista_id, mes, ano, valor, status) VALUES (?, ?, ?, ?, ?)',
                                            (row['id'], mes_num, ano, valor, status))
                        conn.commit()
                        conn.close()
                        st.success("✅ Alterações salvas!")
                        st.rerun()
                    
                    
                    # ========================================
                    # PREPARAÇÃO PARA CONFERÊNCIA INDIVIDUAL E HISTÓRICO
                    # ========================================
                    
                    # Inicializar índice no session_state (agora antes do histórico para filtrar)
                    if 'idx_colab' not in st.session_state:
                        st.session_state.idx_colab = 0
                    
                    # Lista de colaboradores
                    lista_colabs = df_conf['matricula'].astype(str).tolist()
                    total_colabs = len(lista_colabs)
                    
                    # Garantir que o índice está dentro dos limites
                    if st.session_state.idx_colab >= total_colabs:
                        st.session_state.idx_colab = total_colabs - 1
                    if st.session_state.idx_colab < 0:
                        st.session_state.idx_colab = 0
                    
                    # Identificar Matrícula Atual para Filtro
                    curr_mat = lista_colabs[st.session_state.idx_colab] if total_colabs > 0 else None
                    
                    st.markdown("---")
                    st.markdown("---")
                    label_hist = f"📊 Histórico de Pagamentos"
                    if curr_mat:
                        label_hist += f" - Matrícula: {curr_mat}"
                    
                    with st.expander(label_hist, expanded=True):
                        try:

                            df_pagos = get_dataset("PAGAMENTOS")
                            if not df_pagos.empty:
                                # Carregar sem cache para garantir que novos dados do Excel apareçam
                                df_pagos.columns = [str(c).upper().strip() for c in df_pagos.columns]
                                
                                # Limpeza de Matrícula (Excel)
                                df_pagos['MATRICULA'] = df_pagos['MATRICULA'].astype(str).str.split('.').str[0].str.strip()
                                df_pagos['DATA'] = pd.to_datetime(df_pagos['DATA'], dayfirst=True, errors='coerce')
                                
                                # Filtrar Colaborador Atual
                                if curr_mat:
                                    m_clean = str(curr_mat).strip().split('.')[0]
                                    df_hist = df_pagos[df_pagos['MATRICULA'] == m_clean].copy()
                                    
                                    if not df_hist.empty:
                                        # Ordenar por data (mais recente primeiro)
                                        df_hist = df_hist.sort_values('DATA', ascending=False)
                                        
                                        # Mostrar apenas os 5 mais recentes
                                        df_show = df_hist.head(5).copy()
                                        df_show['MES_ANO'] = df_show['DATA'].dt.strftime('%m/%Y')
                                        
                                        # Pivotar para colunas (meses)
                                        df_pivot = df_show.pivot_table(
                                            index=['MATRICULA', 'NOMES'],
                                            columns='MES_ANO',
                                            values='VALOR',
                                            aggfunc='sum'
                                        ).reset_index()
                                        
                                        df_pivot.columns.name = None
                                        df_pivot = df_pivot.rename(columns={'MATRICULA': 'Matrícula', 'NOMES': 'Nome'})
                                        
                                        # Formatar Moeda
                                        for col in df_pivot.columns:
                                            if col not in ['Matrícula', 'Nome']:
                                                df_pivot[col] = df_pivot[col].apply(lambda x: f"R$ {x:,.2f}" if pd.notna(x) else "-")
                                        
                                        st.dataframe(df_pivot, use_container_width=True, hide_index=True)
                                        
                                        # Totais
                                        t_pago = df_show['VALOR'].sum()
                                        st.markdown(f"**Total acumulado nos registros acima:** R$ {t_pago:,.2f}")
                                    else:
                                        st.info(f"ℹ️ Nenhuma informação de pagamento encontrada no Excel para a matrícula {m_clean}.")
                                else:
                                    st.info("ℹ️ Selecione um colaborador para ver o histórico.")
                            else:
                                st.warning("⚠️ Arquivo BASES.BOLSAS/BASE.PAGAMENTOS.xlsx não encontrado na pasta do sistema.")
                        except Exception as e:
                            st.error(f"Erro ao processar histórico: {e}")

                    st.markdown("#### 📝 Conferência Individual")
                    
                    # (Lógica de índice movida para cima para suportar histórico filtrado)
                    # Apenas renderização da navegação aqui
                    
                    # Navegação com botões (Usar callbacks para evitar conflito com selectbox)
                    col1, col2, col3 = st.columns([1, 4, 1])
                    with col1:
                        st.write("") # Spacer
                        st.write("") # Spacer
                        if st.button("⬅️ Anterior", use_container_width=True, key="btn_nav_anterior"):
                            if st.session_state.idx_colab > 0:
                                st.session_state.idx_colab -= 1
                            st.rerun()
                    with col2:
                        # Exibir apenas texto informativo para evitar conflito de estado com lista dinâmica
                        current_matricula = lista_colabs[st.session_state.idx_colab] if st.session_state.idx_colab < len(lista_colabs) else ""
                        html_collab = f'<div style="text-align: center; padding-top: 10px;"><strong>Colaborador {st.session_state.idx_colab + 1}</strong> de {total_colabs}<div style="font-size: 0.8rem; color: #64748b;">(Matrícula: {current_matricula})</div></div>'
                        st.markdown(html_collab, unsafe_allow_html=True)

                    with col3:
                        st.write("") # Spacer
                        st.write("") # Spacer
                        if st.button("Próximo ➡️", use_container_width=True, key="btn_nav_proximo"):
                            if st.session_state.idx_colab < total_colabs - 1:
                                st.session_state.idx_colab += 1
                            st.rerun()
                    
                    # Dados do colaborador atual
                    if st.session_state.idx_colab < len(df_conf):
                        row = df_conf.iloc[st.session_state.idx_colab]
                        
                        # Buscar últimos 3 pagamentos para contexto
                        conn_ctx = get_conn()
                        last_payments = pd.read_sql_query("SELECT mes, ano, valor, status FROM pagamentos WHERE bolsista_id = ? ORDER BY ano DESC, mes DESC LIMIT 3", conn_ctx, params=(row['id'],))
                        conn_ctx.close()
                        
                        hist_html = ""
                        if not last_payments.empty:
                            hist_items = []
                            for _, p in last_payments.iterrows():
                                m_name = MESES[p['mes']-1][:3]
                                hist_items.append(f"<span style='background:#e2e8f0; padding:2px 6px; border-radius:4px; font-size:0.8rem;'>{m_name}/{p['ano']}: <strong>R${p['valor']:.0f}</strong></span>")
                            hist_html = "<div style='margin-top:10px;'>" + " ".join(hist_items) + "</div>"
                        else:
                            hist_html = "<div style='margin-top:10px; font-size:0.8rem; color:#94a3b8;'>Sem histórico recente</div>"


                        # Botão de copiar matrícula
                        c_copy, _ = st.columns([1, 5])
                        with c_copy:
                            st.caption("📋 Copiar Matrícula")
                            st.code(row['matricula'], language=None)

                        # Layout em Card COM histórico (já exibido acima)
                        # Calcular valores para exibição (Fix para SyntaxError)
                        calc_mensalidade = row['mensalidade']
                        calc_porcentagem = row['porcentagem'] * 100
                        calc_reembolso = float(row['valor_reembolso'])
                        
                        # Pre-formatar strings para evitar erro de sintaxe no bloco HTML
                        str_mensalidade = f"{calc_mensalidade:,.2f}"
                        str_porcentagem = f"{calc_porcentagem:.0f}%"
                        str_reembolso = f"{calc_reembolso:,.2f}"

                        # Construir HTML do card
                        status_color = '#16a34a' if 'PAGO' in str(row['status_conf']) else ('#ca8a04' if 'AGUARDANDO' in str(row['status_conf']) else '#dc2626')
                        diretoria_display = row['diretoria'] or 'Sem diretoria'
                        
                        html_card = "".join([
                            f'<div style="background-color: #f8fafc; padding: 20px; border-radius: 12px; border: 1px solid #e2e8f0; margin-bottom: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.02);">',
                            f'<div style="display: flex; justify-content: space-between; align-items: start;">',
                            f'<div style="flex: 1;">',
                            f'<h3 style="margin: 0; color: #1e293b; font-size: 1.4rem;">{row["matricula"]}</h3>',
                            f'<p style="margin: 4px 0 0 0; color: #64748b; font-size: 0.9rem;">Nome: <strong>{row["nome"]}</strong> | {diretoria_display}</p>',
                            f'{hist_html}',
                            f'</div><div style="text-align: right; min-width: 120px;">',
                            f'<span style="font-size: 0.8rem; color: #64748b; text-transform: uppercase; letter-spacing: 0.5px;">Status</span><br>',
                            f'<span style="font-size: 1.1rem; font-weight: 700; color: {status_color};">{row["status_conf"]}</span></div></div>',
                            f'<hr style="margin: 15px 0; border: 0; border-top: 1px solid #e2e8f0;">',
                            f'<div style="display: flex; align-items: center; gap: 20px;"><div>',
                            f'<p style="margin: 0; font-size: 0.85rem; color: #64748b;">Cálculo Sugerido</p>',
                            f'<div style="font-size: 1.1rem; color: #334155; font-weight: 500;">R$ {str_mensalidade} <span style="color:#94a3b8">&times;</span> {str_porcentagem} <span style="color:#94a3b8">=</span> <strong>R$ {str_reembolso}</strong></div>',
                            f'</div><div>',
                            f'<p style="margin: 0; font-size: 0.85rem; color: #64748b;">Ação</p>',
                            f'<div style="font-size: 0.9rem; color: #64748b;">Clique em PAGO para salvar e avançar.</div>',
                            f'</div></div></div>'
                        ])
                        st.markdown(html_card, unsafe_allow_html=True)
                        
                        # --- STATUS FORA DO FORM PARA TER INTERATIVIDADE ---
                        # Novas opções solicitadas
                        status_opts = ["REGULAR", "IRREGULAR", "CONCLUIDO", "CANCELADO", "DESISTENCIA", "TRANCADO"]
                        
                        # Mapping de legado para novo sistema
                        raw_st = row['situacao']
                        if raw_st == "ATIVO": curr_status = "REGULAR"
                        elif raw_st == "INATIVO": curr_status = "CANCELADO"
                        elif raw_st == "EM ANÁLISE": curr_status = "IRREGULAR"
                        elif raw_st in status_opts: curr_status = raw_st
                        else: curr_status = "REGULAR"
                        
                        idx_status = status_opts.index(curr_status)
                        
                        c_stat_out, _ = st.columns([1, 2])
                        with c_stat_out:
                            novo_status = st.selectbox(
                                "📌 Situação / Checagem",
                                status_opts,
                                index=idx_status,
                                key=f"status_sel_{row['id']}",
                                help="REGULAR/IRREGULAR = Ativo | CANCELADO/DESISTENCIA = Inativo"
                            )

                        # Input e Botões dentro de um FORMULÁRIO para permitir Enter = Salvar
                        with st.form(key=f"form_pagto_{row['id']}"):
                            c_input, c_obs = st.columns([1, 1.5])
                            with c_input:
                                # Lógica de Valor: Se Cancelado/Desistencia/Concluido/Trancado, sugerir 0,00
                                if novo_status in ["CANCELADO", "DESISTENCIA", "TRANCADO", "CONCLUIDO"]:
                                    val_float = 0.0
                                else:
                                    val_float = float(row['mensalidade'])
                                    
                                val_inicial = f"{val_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                                
                                valor_pagar_str = st.text_input(
                                    "💰 Valor Boleto (100%):",
                                    value=val_inicial,
                                    key=f"valor_ind_{row['id']}_{novo_status}",
                                    help="Digite o valor CHEIO do boleto. O sistema calculará o reembolso na planilha."
                                )
                                
                                # Converter input texto para float (tratamento robusto BR)
                                try:
                                    clean_str = valor_pagar_str.replace("R$", "").strip()
                                    if "," in clean_str and "." in clean_str:
                                        # Formato 1.000,00
                                        clean_str = clean_str.replace(".", "").replace(",", ".")
                                    elif "," in clean_str:
                                        # Formato 1000,00
                                        clean_str = clean_str.replace(",", ".")
                                    valor_pagar = float(clean_str)
                                except ValueError:
                                    valor_pagar = 0.0
                                
                                # Se Status de Encerramento (Concluido, Cancelado, Desistencia), pedir Data
                                data_comprovante = None
                                if novo_status in ["CONCLUIDO", "CANCELADO", "DESISTENCIA", "TRANCADO"]:
                                    lbl_map = {
                                        "CONCLUIDO": "Data de Conclusão",
                                        "CANCELADO": "Data de Cancelamento",
                                        "DESISTENCIA": "Data da Desistência",
                                        "TRANCADO": "Data do Trancamento"
                                    }
                                    label_data = f"📅 {lbl_map.get(novo_status, 'Data')}"
                                    st.markdown(f"**{label_data}**")
                                    data_comprovante = st.date_input(
                                        "Selecione a data:",
                                        value=datetime.today(),
                                        format="DD/MM/YYYY",
                                        key=f"dt_comp_{row['id']}",
                                        label_visibility="collapsed"
                                    )
                            
                            with c_obs:
                                obs_texto = st.text_area("📝 Adicionar Obs. / Diário", height=105, key=f"obs_{row['id']}", placeholder="Digite uma observação para salvar no perfil do colaborador...")

                            col_btn1, col_btn2, col_btn3 = st.columns(3)
                            with col_btn1:
                                # Primeiro botão é o default do Enter
                                is_pago = st.form_submit_button("✅ PAGO", type="primary", use_container_width=True)
                            with col_btn2:
                                is_pendente = st.form_submit_button("❌ PENDENTE", use_container_width=True)
                            with col_btn3:
                                is_pular = st.form_submit_button("⏭️ PULAR", use_container_width=True)

                        # Logica de Processamento UNIFICADA
                        if is_pago or is_pendente:
                            try:
                                conn = get_conn()
                                
                                # 1. Salvar Observação se houver (com data extra se aplicável)
                                texto_final = obs_texto
                                
                                if (novo_status in ["CONCLUIDO", "CANCELADO", "DESISTENCIA", "TRANCADO"]) and data_comprovante:
                                    str_data = data_comprovante.strftime("%d/%m/%Y")
                                    prefixo = f"[{novo_status}] Data de Referência"
                                    obs_extra = f"{prefixo}: {str_data}"
                                    
                                    if texto_final:
                                        texto_final += f" | {obs_extra}"
                                    else:
                                        texto_final = obs_extra

                                if texto_final:
                                    data_hoje = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                    # Verificar se tabela tem colunas corretas, senão adaptar (Baseado no código do perfil)
                                    conn.execute("INSERT INTO observacoes (bolsista_id, data, texto) VALUES (?, ?, ?)", (int(row['id']), data_hoje, texto_final))
                                
                                # 2. Atualizar Status do Bolsista se Mudou
                                if novo_status != row['situacao']:
                                    conn.execute("UPDATE bolsistas SET situacao = ? WHERE id = ?", (novo_status, int(row['id'])))
                                    st.toast(f"🔄 Status alterado para {novo_status}!", icon="🔄")

                                # 3. Salvar Pagamento
                                status_pgto = 'PAGO' if is_pago else 'PENDENTE'
                                valor_pgto = float(valor_pagar) if is_pago else 0.0
                                
                                conn.execute('INSERT OR REPLACE INTO pagamentos (bolsista_id, mes, ano, valor, status) VALUES (?, ?, ?, ?, ?)',
                                            (int(row['id']), int(mes_num), int(ano), valor_pgto, status_pgto))
                                
                                conn.commit()
                                conn.close()
                                
                                
                                # Feedback Visual
                                if is_pago:
                                    st.success(f"✅ Salvo: {row['nome']} - R$ {valor_pgto:,.2f} (PAGO)")
                                else:
                                    st.warning(f"⚠️ Salvo: {row['nome']} (PENDENTE)")
                               
                                
                                if 'table_key_version' not in st.session_state: st.session_state.table_key_version = 0
                                st.session_state.table_key_version += 1
                                
                                import time
                                time.sleep(0.2)
                                
                                # Lógica de navegação inteligente
                                if is_pago:
                                    vai_sair_da_lista = (filtro_status == "⏳ Aguardando") or (filtro_status == "❌ Pendentes")
                                else: # Pendente
                                    vai_sair_da_lista = (filtro_status == "⏳ Aguardando") or (filtro_status == "✅ Pagos")
                                
                                if not vai_sair_da_lista:
                                    if st.session_state.idx_colab < total_colabs - 1:
                                        st.session_state.idx_colab += 1
                                else:
                                    if st.session_state.idx_colab >= total_colabs - 1 and total_colabs > 1:
                                            st.session_state.idx_colab -= 1
                                
                                st.rerun()
                            except Exception as e:
                                st.error(f"Erro ao salvar: {e}")

                        elif is_pular:
                            if st.session_state.idx_colab < total_colabs - 1:
                                st.session_state.idx_colab += 1
                            st.rerun()

                        # ==========================================
                        # PRÉVIA DO RELATÓRIO DP (Mirroring Tab 3)
                        # ==========================================
                        st.markdown("---")
                        st.markdown("##### 📄 Relatório DP Parcial (Todos os Conferidos)")
                        
                        c_prev = get_conn()
                        # Buscar totais
                        res_total = c_prev.execute('SELECT SUM(valor), COUNT(*) FROM pagamentos WHERE mes=? AND ano=? AND status=?', (int(mes_num), int(ano), 'PAGO')).fetchone()
                        total_pago_now = res_total[0] if res_total and res_total[0] else 0.0
                        count_pago_now = res_total[1] if res_total else 0
                        
                        # Buscar TODOS os pagos com ID para exclusão
                        df_prev = pd.read_sql_query('''
                            SELECT p.id, b.nome, b.matricula, p.valor, CAST(p.bolsista_id AS INTEGER) as bolsista_id_fix
                            FROM pagamentos p
                            LEFT JOIN bolsistas b ON CAST(p.bolsista_id AS INTEGER) = b.id
                            WHERE p.mes = ? AND p.ano = ? AND p.status = 'PAGO'
                            ORDER BY p.id DESC
                        ''', c_prev, params=(int(mes_num), int(ano)))
                        c_prev.close()
                        
                        if count_pago_now > 0:
                            st.caption(f"💰 Total Acumulado: **R$ {total_pago_now:,.2f}** ({count_pago_now} colaboradores)")
                            
                            # Ajustar nome 
                            def get_nome_display(row):
                                if row['nome'] and pd.notna(row['nome']):
                                    return row['nome']
                                return f"ID {row['bolsista_id_fix']}"
                            
                            df_prev['nome_final'] = df_prev.apply(get_nome_display, axis=1)
                            df_prev['Excluir'] = False # Coluna de Checkbox
                            df_prev['Competência'] = f"{mes}/{ano}" # Coluna de Competência
                            
                            # Formatar valor para BR string
                            df_prev['valor_fmt'] = df_prev['valor'].apply(lambda x: f"R$ {float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if pd.notna(x) else "R$ 0,00")
                            
                            # Tabela Interativa
                            edited_prev = st.data_editor(
                                df_prev[['Excluir', 'matricula', 'nome_final', 'Competência', 'valor_fmt']],
                                column_config={
                                    "Excluir": st.column_config.CheckboxColumn("🗑️", width="small", help="Selecione para excluir"),
                                    "matricula": st.column_config.TextColumn("Matrícula", width="medium", disabled=True),
                                    "nome_final": st.column_config.TextColumn("Nome", width="large", disabled=True),
                                    "Competência": st.column_config.TextColumn("Competência", width="medium", disabled=True),
                                    "valor_fmt": st.column_config.TextColumn("Valor", width="medium", disabled=True)
                                },
                                hide_index=True,
                                key=f"edit_prev_{st.session_state.get('table_key_version',0)}"
                            )
                            
                            # Botão de Exclusão (só aparece se houver seleção)
                            if edited_prev['Excluir'].any():
                                if st.button("🗑️ Apagar Selecionados", type="secondary"):
                                    # Pegar IDs reais baseados no index
                                    ids_to_del = df_prev.loc[edited_prev[edited_prev['Excluir']].index, 'id'].tolist()
                                    if ids_to_del:
                                        conn = get_conn()
                                        for pid in ids_to_del:
                                            conn.execute("DELETE FROM pagamentos WHERE id = ?", (pid,))
                                        conn.commit()
                                        conn.close()
                                        
                                        st.toast("✅ Registros excluídos com sucesso!")
                                        
                                        # Forçar refresh
                                        if 'table_key_version' not in st.session_state: st.session_state.table_key_version = 0
                                        st.session_state.table_key_version += 1
                                        import time; time.sleep(0.5)
                                        st.rerun()

                        else:
                            st.info("Nenhum pagamento confirmado para este mês ainda.")
            
            with tab2:
                if pagos_count > 0:
                    # Buscar valores reais pagos do banco
                    conn = get_conn()
                    df_pagos_db = pd.read_sql_query('''
                        SELECT p.bolsista_id, (p.valor * b.porcentagem) as valor_pago, b.matricula, b.nome, b.diretoria, b.cpf, b.curso
                        FROM pagamentos p
                        JOIN bolsistas b ON p.bolsista_id = b.id
                        WHERE p.mes = ? AND p.ano = ? AND p.status = 'PAGO'
                    ''', conn, params=(mes_num, ano))
                    conn.close()
                    
                    total_pago = df_pagos_db['valor_pago'].sum()
                    st.metric("💰 Total a Reembolsar", f"R$ {total_pago:,.2f}")
                    
                    st.dataframe(
                        df_pagos_db[['matricula', 'nome', 'diretoria', 'valor_pago']].rename(columns={
                            'matricula': 'Matrícula', 'nome': 'Nome', 'diretoria': 'Diretoria', 'valor_pago': 'Valor Pago'
                        }),
                        use_container_width=True, hide_index=True, height=400
                    )
                else:
                    st.info("Nenhum PAGO ainda.")
            
            with tab3:
                if pagos_count > 0:
                    # Buscar valores reais para relatório DP
                    conn = get_conn()
                    df_rel = pd.read_sql_query('''
                        SELECT b.matricula as MATRICULA, b.nome as NOME, b.cpf as CPF, 
                               b.diretoria as DIRETORIA, b.curso as CURSO, (p.valor * b.porcentagem) as VALOR,
                               p.valor as VALOR_CHEIO, b.porcentagem as PCT_BOLSA
                        FROM pagamentos p
                        JOIN bolsistas b ON p.bolsista_id = b.id
                        WHERE p.mes = ? AND p.ano = ? AND p.status = 'PAGO'
                        ORDER BY b.nome
                    ''', conn, params=(mes_num, ano))
                    conn.close()
                    
                    total_val = df_rel['VALOR'].sum()
                    
                    st.success(f"**{len(df_rel)} colaboradores** | **R$ {total_val:,.2f}**")
                    
                    # Versão de visualização com formatação
                    df_rel_display = df_rel.copy()
                    df_rel_display['VALOR'] = df_rel_display['VALOR'].apply(lambda x: f"R$ {float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if pd.notna(x) else "R$ 0,00")
                    
                    st.dataframe(df_rel_display, use_container_width=True, hide_index=True)
                    
                    excel_data = df_to_excel(df_rel)
                    st.download_button("⬇️ BAIXAR RELATÓRIO DP", excel_data, f"BOLSAS_{mes}_{ano}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
                else:
                    st.warning("Faça a conferência primeiro.")
    
    # =============================================
    # PERFIL DO COLABORADOR
    # =============================================
    elif menu == "👤 Perfil":
        st.markdown("### 👤 Perfil do Colaborador")
        
        conn = get_conn()
        
        # Buscar lista de colaboradores para seleção
        df_users = pd.read_sql_query("SELECT id, matricula, nome, diretoria, curso, instituicao, valor_reembolso, situacao, observacao FROM bolsistas ORDER BY nome", conn)
        
        if len(df_users) > 0:
            # Dropdown de busca
            col_search, _ = st.columns([1, 1])
            with col_search:
                # Mapa de nomes para otimizar busca (Nome + Matrícula)
                name_map = dict(zip(df_users['id'], df_users.apply(lambda x: f"{x['nome']} ({x['matricula']})", axis=1)))
                sel_id = st.selectbox("🔍 Buscar Colaborador:", 
                                    options=df_users['id'].tolist(), 
                                    format_func=lambda x: f"{name_map[x]}",
                                    index=0)
            
            if sel_id:
                user = df_users[df_users['id'] == sel_id].iloc[0]
                
                # Calcular cores do badge de status
                if user['situacao'] in ['ATIVO', 'REGULAR']:
                    bg_color = '#dcfce7'
                    text_color = '#166534'
                elif user['situacao'] == 'IRREGULAR':
                    bg_color = '#fef9c3'
                    text_color = '#854d0e'
                elif user['situacao'] == 'CONCLUIDO':
                    bg_color = '#dbeafe'
                    text_color = '#1e40af'
                else:
                    bg_color = '#fee2e2'
                    text_color = '#991b1b'
                
                # Formatar valor monetário
                valor_bolsa_fmt = f"R$ {user['valor_reembolso']:,.2f}"
                
                # Card Principal de Informações
                html_perfil = "".join([
                    f'<div style="background: white; padding: 25px; border-radius: 12px; border: 1px solid #e2e8f0; margin-bottom: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.02);">',
                    f'<div style="display: flex; justify-content: space-between; align-items: start;">',
                    f'<div>',
                    f'<h2 style="margin: 0; color: #1e293b; font-size: 1.8rem;">{user["nome"]}</h2>',
                    f'<p style="color: #64748b; margin: 5px 0 0 0; font-size: 1rem;">Matrícula: <strong>{user["matricula"]}</strong> | {user["diretoria"]}</p>',
                    f'</div>',
                    f'<div style="text-align: right;">',
                    f'<span style="background: {bg_color}; color: {text_color}; padding: 6px 12px; border-radius: 20px; font-weight: 600; font-size: 0.9rem;">',
                    f'{user["situacao"]}',
                    f'</span>',
                    f'</div>',
                    f'</div>',
                    f'<hr style="margin: 15px 0; border: 0; border-top: 1px solid #f1f5f9;">',
                    f'<div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px;">',
                    f'<div><strong style="color: #64748b; font-size: 0.85rem;">CURSO</strong><br><span style="color: #334155; font-size: 1.1rem; font-weight: 500;">{user["curso"]}</span></div>',
                    f'<div><strong style="color: #64748b; font-size: 0.85rem;">INSTITUIÇÃO</strong><br><span style="color: #334155; font-size: 1.1rem; font-weight: 500;">{user["instituicao"]}</span></div>',
                    f'<div><strong style="color: #64748b; font-size: 0.85rem;">VALOR BOLSA</strong><br><span style="color: #0d9488; font-size: 1.1rem; font-weight: 600;">{valor_bolsa_fmt}</span></div>',
                    f'</div>',
                    f'</div>'
                ])
                st.markdown(html_perfil, unsafe_allow_html=True)
                
                # Abas para Histórico e Informações
                tab_hist, tab_info = st.tabs(["💰 Histórico Financeiro", "📝 Informações & Obs"])
                
                with tab_hist:
                    # 1. BUSCAR DADOS DA CONFERÊNCIA ATUAL (Base de Dados)
                    # Isso garante que o que você está conferindo agora (ex: Janeiro/2026) apareça no perfil
                    hist_conf = pd.read_sql_query("SELECT p.mes, p.ano, (p.valor * b.porcentagem) as valor, p.status, p.observacao FROM pagamentos p JOIN bolsistas b ON p.bolsista_id = b.id WHERE p.bolsista_id = ?", conn, params=(int(sel_id),))
                    
                    # 2. BUSCAR DADOS DO EXCEL (Histórico Consolidado)
                    hist_excel = pd.DataFrame()
                    if os.path.exists("VALORES.PAGOS.xlsx"):
                        try:
                            df_ex = pd.read_excel("VALORES.PAGOS.xlsx")
                            df_ex.columns = [str(c).upper().strip() for c in df_ex.columns]
                            m_clean = str(user['matricula']).strip().split('.')[0]
                            # Limpeza da matrícula para o match
                            df_ex['MATRICULA'] = df_ex['MATRICULA'].astype(str).str.split('.').str[0].str.strip()
                            df_colab = df_ex[df_ex['MATRICULA'] == m_clean].copy()
                            
                            if not df_colab.empty:
                                df_colab['DATA'] = pd.to_datetime(df_colab['DATA'], dayfirst=True, errors='coerce')
                                hist_excel = pd.DataFrame({
                                    'mes': df_colab['DATA'].dt.month,
                                    'ano': df_colab['DATA'].dt.year,
                                    'valor': df_colab['VALOR'],
                                    'status': 'PAGO',
                                    'observacao': 'Relatório Pagos'
                                })
                        except:
                            pass
                        
                    # 3. BUSCAR HISTÓRICO LEGADO (Backup do Sistema)
                    try:
                        hist_legacy = pd.read_sql_query("SELECT mes, ano, valor, 'PAGO' as status, 'Importado' as observacao FROM historico_pagamentos WHERE matricula = ?", conn, params=(str(user['matricula']),))
                    except:
                        hist_legacy = pd.DataFrame()
                        
                    # 4. UNIFICAÇÃO COM PRIORIDADE (Conferência > Excel > Legado)
                    hist = pd.concat([hist_conf, hist_excel, hist_legacy], ignore_index=True)
                    
                    # Remover duplicatas de Mês/Ano, mantendo a primeira ocorrência (que é a de maior prioridade)
                    if not hist.empty:
                        # Criar uma chave temporária para drop_duplicates
                        hist['temp_key'] = hist.apply(lambda x: f"{int(x['mes'])}/{int(x['ano'])}", axis=1)
                        hist = hist.drop_duplicates(subset=['temp_key'], keep='first').drop(columns=['temp_key'])
                        
                        # Ordenar por Ano/Mês (mais recente primeiro)
                        hist = hist.sort_values(by=['ano', 'mes'], ascending=False)
                    else:
                        hist = pd.DataFrame(columns=['mes', 'ano', 'valor', 'status', 'observacao'])
                    
                    # Calcular métricas
                    if len(hist) > 0:
                        total_pago = hist[hist['status'] == 'PAGO']['valor'].sum()
                        pendente = hist[hist['status'] == 'PENDENTE']['valor'].sum()
                    else:
                        total_pago = 0
                        pendente = 0
                        
                    # Cards de estatísticas com estilo
                    last_pagto = f"{MESES[hist.iloc[0]['mes']-1]}/{hist.iloc[0]['ano']}" if len(hist) > 0 else "-"
                    st.markdown("".join([
                        f'<div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 20px;">',
                        f'<div style="background: #eff6ff; padding: 15px; border-radius: 10px; border: 1px solid #bfdbfe; text-align: center;">',
                        f'<span style="color: #64748b; font-size: 0.9rem;">Total Pago</span><br>',
                        f'<strong style="color: #2563eb; font-size: 1.4rem;">R$ {total_pago:,.2f}</strong>',
                        f'</div>',
                        f'<div style="background: #fff7ed; padding: 15px; border-radius: 10px; border: 1px solid #fed7aa; text-align: center;">',
                        f'<span style="color: #64748b; font-size: 0.9rem;">Pendente</span><br>',
                        f'<strong style="color: #ea580c; font-size: 1.4rem;">R$ {pendente:,.2f}</strong>',
                        f'</div>',
                        f'<div style="background: #f8fafc; padding: 15px; border-radius: 10px; border: 1px solid #e2e8f0; text-align: center;">',
                        f'<span style="color: #64748b; font-size: 0.9rem;">Último Pagamento</span><br>',
                        f'<strong style="color: #334155; font-size: 1.4rem;">{last_pagto}</strong>',
                        f'</div></div>'
                    ]), unsafe_allow_html=True)
                    
                    c_chart, c_table = st.columns([1, 1])
                    
                    if len(hist) > 0:
                        # Gráfico de Evolução
                        with c_chart:
                            st.markdown("##### 📈 Evolução dos Pagamentos")
                            df_chart = hist[hist['status']=='PAGO'].copy()
                            # Criar data para ordenação correta no gráfico
                            df_chart['Data'] = df_chart.apply(lambda x: pd.Timestamp(year=x['ano'], month=x['mes'], day=1), axis=1)
                            df_chart = df_chart.sort_values('Data')
                            df_chart['Mês/Ano'] = df_chart.apply(lambda x: f"{x['mes']:02d}/{str(x['ano'])[2:]}", axis=1)
                            
                            fig = px.bar(df_chart, x='Mês/Ano', y='valor', text_auto='.2s', color_discrete_sequence=['#3b82f6'])
                            fig.update_layout(
                                plot_bgcolor="white",
                                paper_bgcolor="white",
                                font=dict(color="#374151"),
                                xaxis_title=None, 
                                yaxis_title=None, 
                                height=300, 
                                margin=dict(l=0, r=0, t=10, b=0)
                            )
                            st.plotly_chart(fig, use_container_width=True)

                        with c_table:
                            st.markdown("##### 🧾 Histórico Detalhado")
                            hist_display = hist.copy()
                            hist_display['Competência'] = hist_display.apply(lambda x: f"{MESES[x['mes']-1]}/{x['ano']}", axis=1)
                            # Traduzir nomes de fontes para ficar mais amigável
                            font_map = {
                                'Relatório Pagos': '📁 Excel',
                                'Importado': '📦 Legado',
                            }
                            hist_display['Fonte'] = hist_display['observacao'].apply(lambda x: font_map.get(x, '✅ Conferência'))
                            
                            hist_display = hist_display[['Competência', 'valor', 'status', 'Fonte']].rename(columns={
                                'valor': 'Valor', 'status': 'Status'
                            })
                            
                            if AGGRID:
                                try:
                                    # Limpeza de dados para o Histórico
                                    hist_display.columns = [str(c) for c in hist_display.columns]
                                    for col in hist_display.columns:
                                        if not pd.api.types.is_numeric_dtype(hist_display[col]):
                                            hist_display[col] = hist_display[col].fillna('').astype(str)

                                    # Usar AgGrid para tabela também
                                    gb_h = GridOptionsBuilder.from_dataframe(hist_display)
                                    gb_h.configure_column("Valor", type=["numericColumn", "numberColumnFilter", "customNumericFormat"], precision=2, valueFormatter=JsCode("params.value ? params.value.toLocaleString('pt-BR', {style: 'currency', currency: 'BRL'}) : ''"))
                                    gb_h.configure_column("Status", cellStyle=JsCode("function(params) { return {'fontWeight': 'bold'}; }"))
                                    gb_h.configure_pagination(paginationAutoPageSize=False, paginationPageSize=20)
                                    gb_h.configure_default_column(resizable=False)
                                    
                                    AgGrid(
                                        hist_display,
                                        gridOptions=gb_h.build(),
                                        height=500,
                                        theme='balham',
                                        update_mode=GridUpdateMode.VALUE_CHANGED,
                                        data_return_mode=DataReturnMode.AS_INPUT,
                                        allow_unsafe_jscode=True,
                                        key=f"hist_grid_{sel_id}"
                                    )
                                except Exception as e:
                                    st.dataframe(hist_display, use_container_width=True, hide_index=True)
                            else:
                                st.dataframe(hist_display, use_container_width=True, hide_index=True)
                    else:
                        st.info("Nenhum registro de pagamento encontrado para este colaborador.")

                        
                with tab_info:
                    st.markdown("#### 📝 Diário de Bordo e Observações")
                    
                    # Layout: Coluna da esquerda (Histórico) e form abaixo
                    
                    # Buscar observações registradas
                    try:
                        obs_hist = pd.read_sql_query("SELECT id, data, texto, nome_anexo, anexo_blob FROM observacoes WHERE bolsista_id = ? ORDER BY data DESC", conn, params=(int(sel_id),))
                    except:
                        try:
                            obs_hist = pd.read_sql_query("SELECT data, texto FROM observacoes WHERE bolsista_id = ? ORDER BY data DESC", conn, params=(int(sel_id),))
                        except:
                            obs_hist = pd.DataFrame()

                    # Migrar observação antiga se existir e não tiver no histórico novo (apenas visualização inicial)
                    if obs_hist.empty and user['observacao']:
                        st.info("⚠️ Observação antiga migrada para o novo histórico.")
                        try:
                            c_temp = get_conn()
                            c_temp.execute("INSERT INTO observacoes (bolsista_id, texto) VALUES (?, ?)", (int(sel_id), user['observacao']))
                            c_temp.commit()
                            c_temp.close()
                            st.rerun()
                        except: pass
                    
                    # Exibir histórico Estilo Timeline
                    if not obs_hist.empty:
                        st.markdown('<div style="max-height: 500px; overflow-y: auto; padding-right: 10px;">', unsafe_allow_html=True)
                        for idx, row in obs_hist.iterrows():
                            # Converter string de data
                            try:
                                data_obj = pd.to_datetime(row['data'])
                                data_fmt = data_obj.strftime("%d/%m/%Y às %H:%M")
                            except:
                                data_fmt = str(row['data'])
                            
                            st.markdown("".join([
                                f'<div style="background: #f8fafc; border-left: 4px solid #3b82f6; padding: 15px; border-radius: 4px; margin-bottom: 15px;">',
                                f'<div style="display: flex; justify-content: space-between;">',
                                f'<div style="font-size: 0.8rem; color: #64748b; margin-bottom: 5px;">📅 {data_fmt}</div>',
                                f'</div>',
                                f'<div style="color: #334155; white-space: pre-wrap; font-size: 0.95rem;">{row["texto"]}</div>'
                            ]), unsafe_allow_html=True)
                            
                            # Mostrar anexo se houver
                            if 'anexo_blob' in row and row['anexo_blob']:
                                try:
                                    if isinstance(row['nome_anexo'], str) and row['nome_anexo'].lower().endswith(('.png', '.jpg', '.jpeg')):
                                        with st.expander(f"🖼️ Ver Imagem: {row['nome_anexo']}", expanded=False):
                                            st.image(row['anexo_blob'], width=400)
                                    else:
                                        st.download_button(
                                            label=f"📎 Baixar Anexo ({row['nome_anexo']})",
                                            data=row['anexo_blob'],
                                            file_name=row['nome_anexo'] or "anexo",
                                            key=f"dl_{row['id']}"
                                        )
                                except:
                                    st.error("Erro ao exibir anexo")
                            
                            st.markdown("</div>", unsafe_allow_html=True)
                            
                            # Botão de excluir fora do markdown container, logo abaixo
                            col_del_btn, _ = st.columns([1, 10])
                            with col_del_btn:
                                if st.button("🗑️ Excluir", key=f"del_obs_{row['id']}", type="secondary", use_container_width=True):
                                    c_del = get_conn()
                                    c_del.execute("DELETE FROM observacoes WHERE id = ?", (row['id'],))
                                    c_del.commit()
                                    c_del.close()
                                    st.warning("Registro excluído com sucesso!")
                                    st.rerun()
                                    
                        st.markdown('</div>', unsafe_allow_html=True)
                    else:
                        st.info("Nenhuma anotação registrada ainda.")
                        
                    st.markdown("---")
                    
                    # Formulário de Nova Anotação
                    with st.form(key=f"form_obs_{sel_id}"):
                        c_date, c_file = st.columns([1, 2])
                        with c_date:
                            data_registro = st.date_input("Data do Registro:", datetime.today())
                        with c_file:
                            uploaded_file = st.file_uploader("📎 Anexar Foto/Arquivo:", type=['png', 'jpg', 'jpeg', 'pdf', 'docx', 'txt'])
                            
                        novo_texto = st.text_area("✍️ Descrição / Anotação:", height=100, placeholder="Digite aqui os detalhes...")
                        
                        col_btn, _ = st.columns([1, 4])
                        submit = col_btn.form_submit_button("💾 Registar Nota com Data", type="primary", use_container_width=True)
                        
                        if submit and novo_texto:
                            blob_data = None
                            filename = None
                            if uploaded_file is not None:
                                blob_data = uploaded_file.read()
                                filename = uploaded_file.name
                            
                            # Usar data selecionada + hora atual para ordenação correta
                            data_final = datetime.combine(data_registro, datetime.now().time())
                            
                            c_obs = get_conn()
                            c_obs.execute("INSERT INTO observacoes (bolsista_id, texto, data, anexo_blob, nome_anexo) VALUES (?, ?, ?, ?, ?)", (int(sel_id), novo_texto, data_final, blob_data, filename))
                            
                            # Atualizar 'observacao' apenas com texto para compatibilidade
                            c_obs.execute("UPDATE bolsistas SET observacao = ? WHERE id = ?", (novo_texto, int(sel_id)))
                            
                            c_obs.commit()
                            c_obs.close()
                            st.success("Registro adicionado com sucesso!")
                            st.rerun()

        
        else:
            st.warning("Nenhum colaborador cadastrado ainda.")
            
        conn.close()
    
    # =============================================
    # HISTÓRICO DE PAGAMENTOS (PARA ENVIO POR E-MAIL)
    # =============================================
    elif menu == "📅 Histórico":
        st.markdown("### 📅 Histórico de Pagamentos para Envio")
        st.caption("Relatório dos últimos 3 meses com Reporte 1 e Reporte 2 para envio aos gestores")
        
        try:
            

            # Carregar dados de pagamentos
            @st.cache_data(ttl=300)
            def carregar_pagamentos_completo():
                if os.path.exists("BASES.BOLSAS/BASE.PAGAMENTOS.xlsx"):
                    df = pd.read_excel("BASES.BOLSAS/BASE.PAGAMENTOS.xlsx")
                    df.columns = [str(c).upper().strip() for c in df.columns]
                    df['MATRICULA'] = df['MATRICULA'].astype(str).str.strip()
                    df['DATA'] = pd.to_datetime(df['DATA'], errors='coerce')
                    return df
                return pd.DataFrame()
            
            # Carregar dados do organograma (para Gestor N3 e Gestor N4)
            @st.cache_data(ttl=300)
            def carregar_organograma_reportes():
                if os.path.exists("BASES.BOLSAS/ORGANOGRAMA.xlsx"):
                    df = pd.read_excel("BASES.BOLSAS/ORGANOGRAMA.xlsx")
                    df.columns = [str(c).strip() for c in df.columns]
                    df['Cod. Local'] = df['Cod. Local'].astype(str).str.strip()
                    df = df.sort_values('Cod. Local', key=lambda s: s.str.len(), ascending=False)
                    return df
                return pd.DataFrame()
            
            # Carregar TODOS os bolsistas do banco (para mostrar mesmo sem pagamento)
            @st.cache_data(ttl=300)
            def carregar_todos_bolsistas():
                conn = get_conn()
                df = pd.read_sql("SELECT matricula, nome, situacao, diretoria, cod_local FROM bolsistas", conn)
                df['matricula'] = df['matricula'].astype(str).str.strip()
                conn.close()
                return df
            
            df_pagos = carregar_pagamentos_completo()
            df_organograma = carregar_organograma_reportes()
            df_bolsistas = carregar_todos_bolsistas()
            
            # Começar com TODOS os bolsistas (não apenas os que têm pagamento)
            if len(df_bolsistas) > 0:
                # Filtros
                col1, col2, col3, col4 = st.columns([1, 1, 1, 1])
                with col1:
                    # Opção para escolher período
                    periodo = st.selectbox("📅 Período", ["Últimos 3 Meses", "Últimos 6 Meses", "Todo o Histórico"])
                with col2:
                    # Filtro por diretoria - pegar do banco de bolsistas
                    diretorias_disp = df_bolsistas['diretoria'].dropna().unique().tolist()
                    diretoria_filtro = st.selectbox("🏢 Diretoria", ["Todas"] + sorted([d for d in diretorias_disp if d]))
                with col3:
                    # Filtro por situação
                    situacao_filtro = st.selectbox("📋 Situação", ["Todas", "REGULAR", "IRREGULAR", "CONCLUIDO", "CANCELADO", "DESISTENCIA", "TRANCADO"])
                with col4:
                    search_term = st.text_input("🔍 Buscar", placeholder="Nome ou Matrícula...")
                
                # Preparar base de bolsistas
                df_base = df_bolsistas.copy()
                df_base.columns = ['MATRICULA', 'NOME', 'SITUACAO', 'DIRETORIA', 'COD_LOCAL']
                
                # Enriquecer com organograma (Gestor N3 e Gestor N4 via Cod. Local)
                if len(df_organograma) > 0 and 'COD_LOCAL' in df_base.columns:
                    gestor_n3_list = []
                    gestor_n4_list = []
                    diretoria_org_list = []
                    for _, row in df_base.iterrows():
                        cod = str(row.get('COD_LOCAL', '')).strip() if pd.notna(row.get('COD_LOCAL')) else ''
                        found_n3 = None
                        found_n4 = None
                        found_dir = None
                        if cod:
                            for _, org_row in df_organograma.iterrows():
                                org_cod = str(org_row['Cod. Local']).strip()
                                if cod.startswith(org_cod):
                                    found_dir = org_row.get('Diretoria')
                                    found_n3 = org_row.get('Gestor N3')
                                    found_n4 = org_row.get('Gestor N4')
                                    break
                        gestor_n3_list.append(found_n3 if found_n3 and str(found_n3) != 'nan' else '')
                        gestor_n4_list.append(found_n4 if found_n4 and str(found_n4) != 'nan' else '')
                        diretoria_org_list.append(found_dir)
                    
                    df_base['GESTOR_N3'] = gestor_n3_list
                    df_base['GESTOR_N4'] = gestor_n4_list
                    # Atualizar diretoria do organograma se disponível
                    df_base['DIRETORIA_ORG'] = diretoria_org_list
                    df_base['DIRETORIA'] = df_base['DIRETORIA_ORG'].fillna(df_base['DIRETORIA'])
                    df_base = df_base.drop(columns=['DIRETORIA_ORG'], errors='ignore')
                else:
                    df_base['GESTOR_N3'] = ''
                    df_base['GESTOR_N4'] = ''
                
                # Criar coluna Gestor Responsável (Gestor N4, se vazio usa Gestor N3)
                df_base['GESTOR_N3'] = df_base['GESTOR_N3'].fillna('')
                df_base['GESTOR_N4'] = df_base['GESTOR_N4'].fillna('')
                df_base['GESTOR'] = df_base.apply(
                    lambda row: row['GESTOR_N4'] if row['GESTOR_N4'] != '' else row['GESTOR_N3'],
                    axis=1
                )
                df_base['GESTOR'] = df_base['GESTOR'].replace('', 'SEM GESTOR')
                
                # Aplicar filtros
                if diretoria_filtro != "Todas":
                    df_base = df_base[df_base['DIRETORIA'] == diretoria_filtro]
                
                if situacao_filtro != "Todas":
                    df_base = df_base[df_base['SITUACAO'] == situacao_filtro]
                
                if search_term:
                    t = search_term.lower()
                    df_base = df_base[
                        df_base['NOME'].astype(str).str.lower().str.contains(t, na=False) | 
                        df_base['MATRICULA'].astype(str).str.contains(t, na=False)
                    ]
                
                # Processar pagamentos se existirem
                if len(df_pagos) > 0:
                    # Filtrar por período
                    if periodo == "Últimos 3 Meses":
                        # Aumentar para 120 dias para garantir pegar o início dos meses anteriores
                        data_limite = datetime.now() - timedelta(days=120)
                    elif periodo == "Últimos 6 Meses":
                        # Aumentar para 210 dias
                        data_limite = datetime.now() - timedelta(days=210)
                    else:
                        data_limite = pd.Timestamp.min
                    
                    df_filtrado = df_pagos[df_pagos['DATA'] >= data_limite].copy()
                    
                    if len(df_filtrado) > 0:
                        # Criar coluna Mês/Ano
                        df_filtrado['MES_ANO'] = df_filtrado['DATA'].dt.strftime('%m/%Y')
                        
                        # Agregar valores por colaborador e mês
                        df_pivot = df_filtrado.pivot_table(
                            index=['MATRICULA'],
                            columns='MES_ANO',
                            values='VALOR',
                            aggfunc='sum'
                        ).reset_index()
                        
                        df_pivot.columns.name = None
                        
                        # LEFT JOIN - bolsistas com pagamentos
                        df_final = df_base.merge(df_pivot, on='MATRICULA', how='left')
                    else:
                        df_final = df_base.copy()
                else:
                    df_final = df_base.copy()
                
                # Tratar nome vazio/NaN como INATIVO (pessoas desligadas)
                df_final['NOME'] = df_final['NOME'].fillna('DESLIGADO')
                df_final.loc[df_final['NOME'].isin(['', 'nan', 'NaN', 'None']), 'NOME'] = 'DESLIGADO'
                df_final.loc[df_final['NOME'] == 'DESLIGADO', 'SITUACAO'] = 'INATIVO'
                
                # Reorganizar colunas - REMOVER GESTOR_N3, GESTOR_N4 e COD_LOCAL, manter só GESTOR
                cols_base = ['MATRICULA', 'NOME', 'DIRETORIA', 'SITUACAO', 'GESTOR']
                cols_valores = [c for c in df_final.columns if c not in cols_base + ['GESTOR_N3', 'GESTOR_N4', 'COD_LOCAL']]
                df_final = df_final[cols_base + sorted(cols_valores)]
                
                # Renomear colunas para exibição
                df_display = df_final.rename(columns={
                    'MATRICULA': 'Matrícula',
                    'NOME': 'Nome',
                    'DIRETORIA': 'Diretoria',
                    'SITUACAO': 'Situação',
                    'GESTOR': 'Gestor Responsável'
                })
                
                # Métricas
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("👥 Colaboradores", len(df_display))
                with col2:
                    # Total pago no período - calcular dinamicamente
                    cols_meses = [c for c in df_display.columns if '/' in str(c)]
                    if cols_meses:
                        total = df_display[cols_meses].sum().sum()
                    else:
                        total = 0
                    st.metric("💰 Total Pago", f"R$ {total:,.2f}")
                with col3:
                    # Média por colaborador
                    media = total / len(df_display) if len(df_display) > 0 else 0
                    st.metric("📊 Média/Colaborador", f"R$ {media:,.2f}")
                
                st.markdown("---")
                
                # Exibir tabela
                st.dataframe(
                    df_display,
                    use_container_width=True,
                    height=500,
                    hide_index=True
                )
                
                # Botões de download
                st.markdown("### 📥 Download para Envio por E-mail")
                
                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    # Download Excel completo
                    excel_data = df_to_excel(df_display)
                    st.download_button(
                        "📥 Baixar Relatório Completo (Excel)",
                        excel_data,
                        f"historico_pagamentos_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary"
                    )
                
                with col_btn2:
                    # Lista de gestores
                    gestores_list = df_display['Gestor Responsável'].dropna().unique().tolist()
                    gestores_list = [g for g in gestores_list if g and g != 'SEM GESTOR']
                
                # ========================================
                # RELATÓRIO POR GESTOR PARA ENVIO
                # ========================================
                st.markdown("---")
                st.markdown("### 📧 Relatório por Gestor (Para Envio)")
                
                col_g1, col_g2 = st.columns([2, 2])
                with col_g1:
                    # Selecionar gestor específico
                    gestores_opcoes = sorted([g for g in df_display['Gestor Responsável'].unique() if g and str(g) != 'nan'])
                    gestor_selecionado = st.selectbox("👤 Selecione o Gestor", ["Todos"] + gestores_opcoes, key="gestor_sel")
                
                # Filtrar por gestor selecionado
                if gestor_selecionado != "Todos":
                    df_gestor = df_display[df_display['Gestor Responsável'] == gestor_selecionado].copy()
                else:
                    df_gestor = df_display.copy()
                
                if len(df_gestor) > 0:
                    # Métricas do gestor
                    st.markdown(f"**{len(df_gestor)} colaboradores** sob responsabilidade")
                    
                    # Mostrar tabela do gestor
                    st.dataframe(df_gestor, use_container_width=True, height=300, hide_index=True)
                    
                    # Botão de download do relatório do gestor
                    excel_gestor = df_to_excel(df_gestor)
                    nome_arquivo = f"relatorio_{gestor_selecionado.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.xlsx" if gestor_selecionado != "Todos" else f"relatorio_todos_{datetime.now().strftime('%Y%m%d')}.xlsx"
                    
                    st.download_button(
                        f"📥 Baixar Relatório para {gestor_selecionado}",
                        excel_gestor,
                        nome_arquivo,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary",
                        key="download_gestor"
                    )
                
                # Estatísticas por Gestor
                with st.expander("📊 Resumo Geral por Gestor", expanded=False):
                    if 'Gestor Responsável' in df_display.columns:
                        resumo_gestor = df_display.groupby('Gestor Responsável').agg({
                            'Matrícula': 'count'
                        }).reset_index()
                        resumo_gestor.columns = ['Gestor', 'Qtd Colaboradores']
                        
                        # Ordenar por quantidade
                        resumo_gestor = resumo_gestor.sort_values('Qtd Colaboradores', ascending=False)
                        
                        st.dataframe(resumo_gestor, use_container_width=True, hide_index=True)
                        
                        # Download do resumo
                        excel_resumo = df_to_excel(resumo_gestor)
                        st.download_button(
                            "📥 Baixar Resumo por Gestor",
                            excel_resumo,
                            f"resumo_gestores_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            key="download_resumo_gestor"
                        )
            else:
                st.warning("⚠️ Arquivo BASES.BOLSAS/BASE.PAGAMENTOS.xlsx não encontrado ou vazio.")
                
        except Exception as e:
            st.error(f"Erro ao carregar histórico: {e}")
    
    # =============================================
    # PAGAMENTOS (HISTÓRICO COMPLETO)
    # =============================================

    elif menu == "💳 Pagamentos":
        col_p, col_btn_p = st.columns([4, 1])
        with col_p:
            st.markdown("### 💳 Histórico de Pagamentos")
        with col_btn_p:
            if st.button("🔄 Atualizar Pagamentos", type="primary", use_container_width=True, help="Atualiza o histórico usando o arquivo: BASES.BOLSAS/BASE.PAGAMENTOS.xlsx", key="btn_update_pag"):
                try:
                    import glob
                    import shutil

                    
                    with st.spinner("Lendo arquivo de pagamentos..."):
                        # Tenta nome específico primeiro
                        arquivo_pag = "BASES.BOLSAS/BASE.PAGAMENTOS.xlsx"
                        
                        # Se não existir exato, tenta achar por padrão
                        if not os.path.exists(arquivo_pag):
                            procura = glob.glob("BASES.BOLSAS/BASE.PAGAMENTOS*.xlsx")
                            if procura:
                                arquivo_pag = procura[0]
                            else:
                                arquivo_pag = None
                        
                        if arquivo_pag and os.path.exists(arquivo_pag):
                            # Mostrar timestamp do arquivo
                            mod_time = os.path.getmtime(arquivo_pag)
                            dt_mod = datetime.fromtimestamp(mod_time).strftime('%d/%m/%Y %H:%M:%S')
                            st.info(f"📁 Processando: `{arquivo_pag}`\n\n🕒 Última modificação do arquivo: **{dt_mod}**")
                            
                            # Limpar cache do Streamlit para garantir
                            st.cache_data.clear()
                            
                            # Copiar para temp para evitar erro de arquivo aberto
                            temp_pag = "temp_pagamentos_sync.xlsx"
                            try:
                                shutil.copy2(arquivo_pag, temp_pag)
                                
                                # Leitura inteligente de Abas (PAGAMENTOS > Sheet1)
                                xl_file = pd.ExcelFile(temp_pag)
                                sheet_name = 0
                                if 'PAGAMENTOS' in [s.upper() for s in xl_file.sheet_names]:
                                    for s in xl_file.sheet_names:
                                        if s.upper() == 'PAGAMENTOS':
                                            sheet_name = s
                                            break
                                            
                                st.info(f"📄 Lendo aba: `{sheet_name}`")
                                df_hist = pd.read_excel(temp_pag, sheet_name=sheet_name)
                                
                                processar_importacao_historico(df_hist, datetime.now().year)
                            except PermissionError:
                                st.error(f"⚠️ O arquivo `{arquivo_pag}` parece estar aberto. Feche-o e tente novamente.")
                            except Exception as e:
                                st.error(f"Erro ao ler arquivo: {e}")
                            finally:
                                # Sempre remover arquivo temporário
                                if os.path.exists(temp_pag):
                                    try:
                                        os.remove(temp_pag)
                                    except:
                                        pass
                        else:
                            st.warning("⚠️ Arquivo 'BASE.PAGAMENTOS.xlsx' não encontrado na pasta BASES.BOLSAS.")
                            
                        # Atualiza também a base cadastral para garantir integridade (mas sem estardalhaço)
                        if os.path.exists("BASES.BOLSAS/BASE.BOLSAS.2025.xlsx"):
                            df_base = pd.read_excel("BASES.BOLSAS/BASE.BOLSAS.2025.xlsx")
                            # processar_importacao_df -> comentado para não poluir, ou chamamos silenciosamente?
                            # Melhor focar no que o usuário pediu: Pagamentos.
                            
                    st.balloons()
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao sincronizar: {e}")
        
        # ABAS PRINCIPAIS PARA ORGANIZAÇÃO
        tab_consulta, tab_dashboard, tab_ranking = st.tabs([
            "📋 Consulta de Pagamentos",
            "📊 Dashboard / Análises", 
            "🏆 Ranking Colaboradores"
        ])
        
        # ===================
        # ABA 1: CONSULTA
        # ===================
        with tab_consulta:
            st.markdown("#### 🔍 Filtrar Pagamentos")
            
            # Filtros
            col1, col2, col3 = st.columns([1, 1, 2])
            
            # Buscar Anos disponíveis no banco
            conn = get_conn()
            try:
                df_anos = pd.read_sql_query("SELECT DISTINCT ano FROM historico_pagamentos ORDER BY ano DESC", conn)
                lista_anos = df_anos['ano'].dropna().unique().tolist()
                ano_atual = datetime.now().year
                if ano_atual not in lista_anos: lista_anos.append(ano_atual)
                if (ano_atual + 1) not in lista_anos: lista_anos.append(ano_atual + 1)
                lista_anos = sorted(list(set(lista_anos)), reverse=True)
            except:
                lista_anos = [2026, 2025, 2024, 2023, 2022]
            finally:
                conn.close()

            with col1:
                ano_filter = st.selectbox("Ano:", lista_anos + ["Todos"], key="ano_consulta")
            with col2:
                mes_filter = st.selectbox("Mês:", ["Todos"] + MESES, key="mes_consulta")
            with col3:
                busca_pag = st.text_input("🔍 Buscar:", placeholder="Matrícula ou nome...", key="busca_consulta")
            
            # Buscar histórico
            conn = get_conn()
            query = "SELECT * FROM historico_pagamentos WHERE 1=1"
            params = []
            
            if ano_filter != "Todos":
                query += " AND ano = ?"
                params.append(ano_filter)
            if mes_filter != "Todos":
                mes_num = MESES.index(mes_filter) + 1
                query += " AND mes = ?"
                params.append(mes_num)
            if busca_pag:
                query += " AND (matricula LIKE ? OR nome LIKE ?)"
                params.extend([f'%{busca_pag}%', f'%{busca_pag}%'])
            
            query += " ORDER BY ano DESC, mes DESC, nome"
            
            df_hist = pd.read_sql_query(query, conn, params=params)
            conn.close()
            
            if len(df_hist) > 0:
                # Estatísticas
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Registros", len(df_hist))
                with col2:
                    st.metric("Colaboradores", df_hist['matricula'].nunique())
                with col3:
                    st.metric("Total Pago", f"R$ {df_hist['valor'].sum():,.2f}")
                
                st.markdown("---")
                
                # Tabela
                df_show = df_hist[['matricula', 'nome', 'mes_referencia', 'data_pagamento', 'valor']].copy()
                df_show.columns = ['Matrícula', 'Nome', 'Mês Ref.', 'Data Pagto', 'Valor']
                df_show['Valor'] = df_show['Valor'].apply(lambda x: f"R$ {x:,.2f}")
                
                st.dataframe(df_show, use_container_width=True, hide_index=True, height=500)
                
                # Download Excel
                excel_data = df_to_excel(df_hist)
                st.download_button("⬇️ Baixar Histórico", excel_data, "historico_pagamentos.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_consulta")
            else:
                st.info("Nenhum pagamento encontrado com os filtros selecionados.")
        
        # ===================
        # ABA 2: DASHBOARD
        # ===================
        with tab_dashboard:
            # =============================================
            # LINHA DO TEMPO - VISÃO GERAL DE TODOS OS PAGAMENTOS
            # =============================================
            st.markdown("#### Evolução dos Pagamentos")
        
            # Buscar TODOS os pagamentos para a linha do tempo (incluindo cod_local e diretoria do histórico)
            conn = get_conn()
            df_timeline = pd.read_sql_query("SELECT mes, ano, mes_referencia, matricula, nome, valor, data_pagamento, cod_local, diretoria FROM historico_pagamentos ORDER BY ano, mes", conn)
            conn.close()
            
            if len(df_timeline) > 0:
                # Adicionar coluna de safra
                df_timeline['safra'] = df_timeline.apply(lambda row: get_safra(row['ano'], row['mes']), axis=1)
                
                # FILTROS DINÂMICOS
                anos_disponiveis = sorted(df_timeline['ano'].unique().tolist(), reverse=True)
                safras_disponiveis = get_safras_disponiveis(df_timeline)
                
                col_filter1, col_filter2, col_filter3 = st.columns([1, 1, 2])
                with col_filter1:
                    tipo_filtro = st.radio("Filtrar por:", ["📅 Ano", "🌾 Safra"], horizontal=True, key="tipo_filtro_timeline")
            
                with col_filter2:
                    if tipo_filtro == "📅 Ano":
                        filtro_periodo = st.selectbox(
                            "Selecione o Ano:",
                            ["Todos"] + anos_disponiveis,
                            key="filtro_ano_timeline"
                        )
                    else:
                        filtro_periodo = st.selectbox(
                            "Selecione a Safra:",
                            ["Todas"] + safras_disponiveis,
                            key="filtro_safra_timeline"
                        )
            
                with col_filter3:
                    if tipo_filtro == "📅 Ano" and filtro_periodo != "Todos":
                        meses_do_ano = df_timeline[df_timeline['ano'] == filtro_periodo]['mes'].unique()
                        meses_nomes = ["Todos"] + [MESES[m-1] for m in sorted(meses_do_ano)]
                        filtro_mes = st.selectbox("📆 Mês:", meses_nomes, key="filtro_mes_timeline")
                    elif tipo_filtro == "🌾 Safra" and filtro_periodo != "Todas":
                        # Mostrar meses na ordem da safra
                        meses_nomes = ["Todos"] + MESES_SAFRA
                        filtro_mes = st.selectbox("📆 Mês:", meses_nomes, key="filtro_mes_safra")
                    else:
                        filtro_mes = "Todos"
            
                # Aplicar filtro
                df_filtered = df_timeline.copy()
            
                if tipo_filtro == "📅 Ano":
                    if filtro_periodo != "Todos":
                        df_filtered = df_filtered[df_filtered['ano'] == filtro_periodo]
                        if filtro_mes != "Todos":
                            mes_num = MESES.index(filtro_mes) + 1
                            df_filtered = df_filtered[df_filtered['mes'] == mes_num]
                else:  # Safra
                    if filtro_periodo != "Todas":
                        df_filtered = df_filtered[df_filtered['safra'] == filtro_periodo]
                        if filtro_mes != "Todos":
                            mes_num = MESES.index(filtro_mes) + 1
                            df_filtered = df_filtered[df_filtered['mes'] == mes_num]
            
                # Criar coluna de período para ordenação (considerando safra)
                def get_periodo_safra(row):
                    """Ordena por safra: Abr=1, Mai=2... Mar=12"""
                    mes_ordem = MESES_SAFRA_NUM.index(row['mes']) + 1 if row['mes'] in MESES_SAFRA_NUM else row['mes']
                    ano_safra = row['ano'] if row['mes'] >= 4 else row['ano'] - 1
                    return f"{ano_safra}-{mes_ordem:02d}"
            
                df_filtered['periodo'] = df_filtered.apply(get_periodo_safra, axis=1)
                df_filtered['periodo_label'] = df_filtered['mes_referencia'] + '/' + df_filtered['ano'].astype(str)
            
                # Agregação por período
                df_agg = df_filtered.groupby(['ano', 'mes', 'periodo', 'periodo_label']).agg({
                    'valor': 'sum',
                    'matricula': ['nunique', 'count']
                }).reset_index()
                df_agg.columns = ['Ano', 'Mês', 'Periodo', 'Periodo_Label', 'Valor_Total', 'Qtd_Colaboradores', 'Qtd_Pagamentos']
                df_agg = df_agg.sort_values('Periodo')
            
                # Métricas do período selecionado
                st.markdown("---")
                col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                
                total_periodo = df_filtered['valor'].sum()
                qtd_colab = df_filtered['matricula'].nunique()
                qtd_pagtos = len(df_filtered)
                media_pag = df_filtered['valor'].mean() if len(df_filtered) > 0 else 0
                
                with col_m1:
                    st.metric("💰 Total do Período", format_br_currency(total_periodo))
                with col_m2:
                    st.metric("👥 Colaboradores", format_br_number(qtd_colab))
                with col_m3:
                    st.metric("📝 Pagamentos", format_br_number(qtd_pagtos))
                with col_m4:
                    st.metric("📊 Média/Pagamento", format_br_currency(media_pag))
            
                # Tabs para diferentes visualizações
                tab_graf1, tab_graf_ano, tab_graf2, tab_dir, tab_tabela, tab_top = st.tabs([
                    "📈 Evolução Mensal", 
                    "📅 Evolução Anual",
                    "👥 Evolução Colaboradores", 
                    "🏢 Por Diretoria",
                    "📋 Tabela Resumo & Comparativo",
                    "🏆 Todos Colaboradores"
                ])
            
                with tab_graf1:
                    render_area_chart(df_agg, 'Periodo_Label', 'Valor_Total', "Evolução do Valor Total Pago por Mês", label_y="Valor Total (R$)")

                with tab_graf_ano:
                    # Agregação por Ano para o gráfico anual
                    df_agg_ano_chart = df_filtered.groupby('ano')['valor'].sum().reset_index()
                    df_agg_ano_chart.columns = ['Ano', 'Valor Total']
                    # Converter Ano para string para ficar categórico no eixo X
                    df_agg_ano_chart['Ano'] = df_agg_ano_chart['Ano'].astype(str)
                    
                    render_area_chart(df_agg_ano_chart, 'Ano', 'Valor Total', "Evolução do Valor Total Pago por Ano", label_y="Valor Total (R$)")
            
                with tab_graf2:
                    render_bar_chart(df_agg, 'Periodo_Label', 'Qtd_Colaboradores', "Quantidade de Colaboradores Pagos por Mês", label_y="Nº Colaboradores")

                with tab_dir:
                    st.markdown("#### Análise por Diretoria")
                    
                    # Para esta aba, utilizaremos EXCLUSIVAMENTE os dados de Pagamentos + Organograma,
                    # conforme solicitado. Ignoramos a tabela de bolsistas (cadastro) aqui.
                    df_merged_dir = df_filtered.copy()
                    
                    # 1. Limpar campos para garantir que o organograma tenha chance de preencher
                    for col in ['diretoria', 'cod_local']:
                        if col in df_merged_dir.columns:
                            df_merged_dir[col] = df_merged_dir[col].astype(str).replace(['N/A', 'NAN', 'NONE', 'nan', ''], None)
                    
                    # 2. Conectar com Organograma para preencher Diretorias via Cod. Local
                    df_org_dash = carregar_organograma()
                    if not df_org_dash.empty:
                        # O enriquecer já prioriza o organograma sobre o N/A
                        df_merged_dir = enriquecer_com_organograma(df_merged_dir, df_org_dash)
                    
                    # Garantir que a diretoria seja normalizada antes do agrupamento
                    df_merged_dir['diretoria'] = df_merged_dir['diretoria'].fillna('N/A').astype(str).str.upper().str.strip()
                    df_merged_dir.loc[df_merged_dir['diretoria'] == 'NAN', 'diretoria'] = 'N/A'
                    df_merged_dir.loc[df_merged_dir['diretoria'] == 'NONE', 'diretoria'] = 'N/A'
                    df_merged_dir.loc[df_merged_dir['diretoria'] == '', 'diretoria'] = 'N/A'
                    
                    col_d1, col_d2 = st.columns([1, 1])
                    
                    with col_d1:
                        # Gráfico Total por Diretoria (Barras Horizontais)
                        st.markdown("##### Total Investido por Diretoria (Período Selecionado)")
                        df_total_dir = df_merged_dir.groupby('diretoria')['valor'].sum().reset_index().sort_values('valor', ascending=True)
                        
                        if not df_total_dir.empty:
                            # Usar render_bar_chart padrão (que já é horizontal e verde)
                            render_bar_chart(
                                df_total_dir, 
                                x_col='diretoria', 
                                y_col='valor', 
                                title="", 
                                label_y="Valor Total",
                                currency=True
                            )
                        else:
                            st.info("Sem dados para exibir.")
                            
                    with col_d2:
                        # Gráfico de Pizza ou Tabela? Vamos de Tabela para detalhes
                        st.markdown("##### Detalhes do Período")
                        df_total_dir_show = df_total_dir.sort_values('valor', ascending=False).copy()
                        df_total_dir_show.columns = ['Diretoria', 'Valor Total']
                        df_total_dir_show['Valor Total'] = df_total_dir_show['Valor Total'].apply(lambda x: format_br_currency(x))
                        st.dataframe(df_total_dir_show, use_container_width=True, hide_index=True)

                    # --- NOVO: Detalhamento de N/A para auxílio ao usuário ---
                    if 'N/A' in df_merged_dir['diretoria'].unique():
                        with st.expander("🕵️ Ver quem são os colaboradores em 'N/A' (Sem Diretoria)", expanded=False):
                            df_na_details = df_merged_dir[df_merged_dir['diretoria'] == 'N/A'].copy()
                            # Agrupar por matrícula e nome para não repetir
                            df_na_grouped = df_na_details.groupby(['matricula', 'nome']).agg({
                                'valor': 'sum',
                                'cod_local': 'first'
                            }).reset_index().sort_values('valor', ascending=False)
                            
                            df_na_grouped.columns = ['Matrícula', 'Nome', 'Valor Total no Período', 'Código Local']
                            
                            st.write("Estes colaboradores estão sem diretoria mapeada. Verifique se a matrícula existe no cadastro ou se o Código Local está correto no Organograma.")
                            st.dataframe(
                                df_na_grouped.style.format({'Valor Total no Período': 'R$ {:,.2f}'}),
                                use_container_width=True,
                                hide_index=True
                            )

                    st.markdown("---")
                    st.markdown("##### 📊 Orçamento vs Realizado (Acompanhamento)")
                    
                    # Definição do Teto Orçamentário (Budget Fixo Anual)
                    BUDGET_ANUAL_TOTAL = 724000.00
                    BUDGET_MENSAL_TOTAL = BUDGET_ANUAL_TOTAL / 12
                    
                    # Definição dos Limites Orçamentários (% do total) por Diretoria
                    METAS_BUDGET = {
                        "AGRICOLA": 0.25,
                        "ADMINISTRATIVA": 0.15,
                        "FINANCEIRA": 0.10,
                        "CSC GRCI": 0.10,
                        "INDUSTRIAL": 0.15,
                        "GENTE E GESTAO": 0.15,
                        "COMERCIAL NOVOS PRODUTOS": 0.10
                    }
                    
                    def normalizar_nome_diretoria(nome):
                        import unicodedata
                        if not nome: return ""
                        n = unicodedata.normalize('NFD', str(nome)).encode('ascii', 'ignore').decode('utf-8').upper()
                        n = n.replace("DIRETORIA", "").replace("/", " ").replace("-", " ").strip()
                        return n

                    # Cálculo do gasto no período selecionado
                    total_geral_periodo = df_merged_dir['valor'].sum()
                    num_meses_selecionados = len(df_merged_dir['periodo_label'].unique()) if 'periodo_label' in df_merged_dir.columns else 1
                    if num_meses_selecionados == 0: num_meses_selecionados = 1
                    
                    budget_teto_periodo = BUDGET_MENSAL_TOTAL * num_meses_selecionados

                    # Cards de Resumo de Budget
                    c_bud1, c_bud2, c_bud3, c_bud4 = st.columns(4)
                    with c_bud1:
                        st.metric("🎯 Teto Anual (Alvo)", format_br_currency(BUDGET_ANUAL_TOTAL))
                    with c_bud2:
                        st.metric("📅 Meta do Período", format_br_currency(budget_teto_periodo), 
                                  help="Valor proporcional aos meses selecionados (R$ 60.333/mês)")
                    with c_bud3:
                        percent_consumo_ano = (total_geral_periodo / BUDGET_ANUAL_TOTAL) * 100
                        st.metric("💸 Gasto Realizado", format_br_currency(total_geral_periodo),
                                  f"{percent_consumo_ano:.1f}% do ano")
                    with c_bud4:
                        saldo = budget_teto_periodo - total_geral_periodo
                        st.metric("⚖️ Saldo do Período", format_br_currency(saldo), 
                                  delta="No Plano" if saldo >= 0 else "Excedido",
                                  delta_color="normal")

                    if total_geral_periodo > 0:
                        df_budget = df_merged_dir.groupby('diretoria')['valor'].sum().reset_index()
                        df_budget.columns = ['Diretoria', 'Gasto Atual']
                        metas_norm = {normalizar_nome_diretoria(k): v for k, v in METAS_BUDGET.items()}
                        
                        def vincular_meta(nome_dir):
                            norm = normalizar_nome_diretoria(nome_dir)
                            return metas_norm.get(norm, 0)

                        df_budget['Meta %'] = df_budget['Diretoria'].apply(vincular_meta)
                        # O limite agora é baseado no BUDGET FIXO, não no gasto total
                        df_budget['Limite Orçamentário'] = budget_teto_periodo * df_budget['Meta %']
                        
                        df_budget['Diferença'] = df_budget['Limite Orçamentário'] - df_budget['Gasto Atual']
                        df_budget['Status'] = df_budget['Diferença'].apply(lambda x: "✅ No Limite" if x >= 0 else "🚨 Acima do Limite")
                        df_budget['% do Budget'] = (df_budget['Gasto Atual'] / df_budget['Limite Orçamentário']) * 100
                        
                        df_budget_plot = df_budget[df_budget['Meta %'] > 0].copy()
                        
                        if not df_budget_plot.empty:
                            import plotly.graph_objects as go
                            fig_budget = go.Figure()
                            
                            fig_budget.add_trace(go.Bar(
                                x=df_budget_plot['Diretoria'],
                                y=df_budget_plot['Gasto Atual'],
                                name="Gasto Atual",
                                marker_color=df_budget_plot['Diferença'].apply(lambda x: '#22c55e' if x >= 0 else '#ef4444'),
                                text=df_budget_plot['Gasto Atual'].apply(lambda x: format_br_currency(x)),
                                textposition='auto',
                            ))
                            
                            fig_budget.add_trace(go.Scatter(
                                x=df_budget_plot['Diretoria'],
                                y=df_budget_plot['Limite Orçamentário'],
                                name="Meta da Diretoria",
                                mode='markers+text',
                                text=df_budget_plot['Limite Orçamentário'].apply(lambda x: f"{x/1000:.1f}k"),
                                textposition='top center',
                                marker=dict(color='#1f2937', size=15, symbol='line-ns-open', line=dict(width=3)),
                                hoverinfo='y'
                            ))
                            
                            fig_budget.update_layout(
                                title=f"Gasto Atual vs Budget Fixo por Diretoria (Base: R$ 724k/ano)",
                                plot_bgcolor="white",
                                paper_bgcolor="white",
                                yaxis=dict(title="Valor Total (R$)", gridcolor="#f3f4f6"),
                                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                            )
                            st.plotly_chart(fig_budget, use_container_width=True)
                            
                            st.markdown("###### Detalhamento da Meta vs Realizado")
                            df_tbl_budget = df_budget_plot.copy()
                            df_tbl_budget['Meta % Tabela'] = df_tbl_budget['Meta %'].apply(lambda x: f"{x*100:.1f}%")
                            df_tbl_budget['Utilização %'] = df_tbl_budget['% do Budget'].apply(lambda x: f"{x:.1f}%")
                            
                            for col in ['Gasto Atual', 'Limite Orçamentário', 'Diferença']:
                                df_tbl_budget[col] = df_tbl_budget[col].apply(lambda x: format_br_currency(x))
                            
                            # Reordenar colunas para incluir a Meta da Diretoria
                            st.dataframe(
                                df_tbl_budget[['Diretoria', 'Meta % Tabela', 'Gasto Atual', 'Limite Orçamentário', 'Utilização %', 'Status']] \
                                .rename(columns={'Meta % Tabela': 'Sua Fatia do Budget (%)'}), 
                                use_container_width=True, 
                                hide_index=True
                            )
                        else:
                            st.warning("Nenhuma diretoria com meta definida foi encontrada nos dados atuais.")
                    else:
                        st.info("Aguardando dados para análise orçamentária.")

                    st.markdown("---")
                    st.markdown("##### 📈 Evolução por Diretoria")
                    # Evolução Temporal por Diretoria (Linhas)
                    # Agrupar por Periodo_Label e Diretoria
                    df_evo_dir_tab = df_merged_dir.groupby(['periodo_label', 'periodo', 'diretoria'])['valor'].sum().reset_index()
                    df_evo_dir_tab = df_evo_dir_tab.sort_values('periodo')
                    
                    if not df_evo_dir_tab.empty:
                        # Criar rótulos formatados (ex: 38.8k)
                        df_evo_dir_tab['text_label'] = df_evo_dir_tab['valor'].apply(lambda x: f"{x/1000:.1f}k" if x >= 1000 else f"{x:.0f}")
                        
                        fig_evo_dir = px.line(
                            df_evo_dir_tab, 
                            x="periodo_label", 
                            y="valor", 
                            color="diretoria",
                            markers=True,
                            text="text_label",
                            title="Evolução Mensal do Investimento por Diretoria",
                            color_discrete_sequence=[APP_COLORS['secondary'], APP_COLORS['primary'], '#1f2937', '#9ca3af', '#d1d5db']
                        )
                        fig_evo_dir.update_traces(textposition="top center")
                        fig_evo_dir.update_layout(
                            plot_bgcolor="white",
                            paper_bgcolor="white",
                            font=dict(color="#374151", size=11),
                            xaxis=dict(showgrid=False, title=None, tickangle=-45),
                            yaxis=dict(showgrid=True, gridcolor="#f3f4f6", title="Valor (R$)"),
                            hovermode="x unified"
                        )
                        st.plotly_chart(fig_evo_dir, use_container_width=True, config=PLOTLY_CONFIG)
                    else:
                        st.info("Sem dados temporais suficientes.")

                    st.markdown("---")
                    st.markdown("##### ⚔️ Comparativo Direto (Side-by-Side)")
                    
                    if not df_merged_dir.empty:
                        lista_dirs = sorted(df_merged_dir['diretoria'].unique().tolist())
                        
                        c_sel1, c_sel2 = st.columns(2)
                        with c_sel1:
                            dir_A = st.selectbox("Diretoria A", lista_dirs, index=0, key="sel_dir_A")
                        with c_sel2:
                            # Tentar selecionar o segundo item por padrão
                            idx_B = 1 if len(lista_dirs) > 1 else 0
                            dir_B = st.selectbox("Diretoria B", lista_dirs, index=idx_B, key="sel_dir_B")
                            
                        if dir_A and dir_B:
                            # Filtrar dados das duas diretorias
                            df_A = df_merged_dir[df_merged_dir['diretoria'] == dir_A]
                            df_B = df_merged_dir[df_merged_dir['diretoria'] == dir_B]
                            
                            val_A = df_A['valor'].sum()
                            val_B = df_B['valor'].sum()
                            
                            # Métricas de Comparação
                            col_comp_m1, col_comp_m2 = st.columns(2)
                            
                            with col_comp_m1:
                                delta_val = val_A - val_B
                                # Se A > B, o delta é positivo, mas em gastos isso é "pior" (cor inversa)
                                st.metric(
                                    f"Total {dir_A}", 
                                    format_br_currency(val_A), 
                                    f"{format_br_currency(delta_val)} vs {dir_B}",
                                    delta_color="inverse" 
                                )
                                
                            with col_comp_m2:
                                delta_val_B = val_B - val_A
                                st.metric(
                                    f"Total {dir_B}", 
                                    format_br_currency(val_B), 
                                    f"{format_br_currency(delta_val_B)} vs {dir_A}",
                                    delta_color="inverse"
                                )
                            
                            # Gráfico Comparativo Mensal (Agrupado)
                            df_comp_chart = pd.concat([
                                df_A.assign(Grupo=dir_A),
                                df_B.assign(Grupo=dir_B)
                            ])
                            
                            # Agrupar por Mês/Periodo
                            if 'periodo_label' in df_comp_chart.columns:
                                df_comp_agg = df_comp_chart.groupby(['periodo_label', 'periodo', 'Grupo'])['valor'].sum().reset_index().sort_values('periodo')
                                
                                # Criar rótulo formatado
                                df_comp_agg['text_label'] = df_comp_agg['valor'].apply(lambda x: f"{x/1000:.1f}k" if x >= 1000 else f"{x:.0f}")

                                fig_comp = px.bar(
                                    df_comp_agg,
                                    x="periodo_label",
                                    y="valor",
                                    color="Grupo",
                                    barmode="group",
                                    text="text_label",
                                    title=f"Comparativo Mensal: {dir_A} vs {dir_B}",
                                    color_discrete_map={dir_A: APP_COLORS['primary'], dir_B: APP_COLORS['secondary']}
                                )
                                
                                fig_comp.update_traces(textposition='outside')
                                
                                fig_comp.update_layout(
                                    plot_bgcolor="white",
                                    paper_bgcolor="white",
                                    font=dict(color="#374151", size=11),
                                    xaxis=dict(showgrid=False, title=None, tickangle=-45),
                                    yaxis=dict(showgrid=True, gridcolor="#f3f4f6", title="Valor (R$)"),
                                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                                )
                                st.plotly_chart(fig_comp, use_container_width=True, config=PLOTLY_CONFIG)

            
                with tab_tabela:
                    st.markdown("#### 📋 Resumo por Ano e Mês")
                    
                    # Seletor de visualização (Valor ou Quantidade)
                    tipo_visao = st.radio(
                        "Visualizar por:", 
                        ["💰 Valor Investido (R$)", "🔢 Quantidade de Pagamentos"],
                        horizontal=True,
                        key="radio_visao_tabela"
                    )
                
                    # Pivot table por SAFRA (usando dados filtrados)
                    if "Valor" in tipo_visao:
                        val_col = 'valor'
                        agg_func = 'sum'
                        y_chart = 'Valor Total'
                        color_scale = 'Blues'
                    else:
                        val_col = 'matricula' # Contar IDs
                        agg_func = 'count' # Contagem
                        y_chart = 'Quantidade'
                        color_scale = 'Greens'
                        
                    # Criar pivot por SAFRA x MÊS
                    df_pivot = df_filtered.pivot_table(
                        values=val_col,
                        index='safra',
                        columns='mes',
                        aggfunc=agg_func,
                        fill_value=0
                    )
                
                    # Ordenar colunas conforme ordem da safra (Abr a Mar)
                    cols_ordenadas = []
                    for mes_num in MESES_SAFRA_NUM:
                        if mes_num in df_pivot.columns:
                            cols_ordenadas.append(mes_num)
                    
                    df_pivot = df_pivot[cols_ordenadas] if cols_ordenadas else df_pivot
                    
                    # Renomear colunas para nomes dos meses (abreviados)
                    df_pivot.columns = [MESES[m-1][:3] for m in df_pivot.columns]
                    df_pivot['TOTAL'] = df_pivot.sum(axis=1)
                
                    # Formatar valores para exibição - FORMATO BRASILEIRO
                    df_pivot_fmt = df_pivot.copy()
                    for col in df_pivot_fmt.columns:
                        if "Valor" in tipo_visao:
                            df_pivot_fmt[col] = df_pivot_fmt[col].apply(lambda x: format_br_currency(x) if x > 0 else "-")
                        else:
                            df_pivot_fmt[col] = df_pivot_fmt[col].apply(lambda x: format_br_number(x) if x > 0 else "-")
                    
                    # Renomear index para mostrar "Safra"
                    df_pivot_fmt.index.name = 'Safra'
                
                    # Botão de Download Excel
                    import io
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df_pivot.to_excel(writer, sheet_name='Resumo')
                        
                    st.download_button(
                        label="📥 Baixar em Excel",
                        data=buffer,
                        file_name="resumo_pagamentos_safra.xlsx",
                        mime="application/vnd.ms-excel"
                    )

                    st.dataframe(df_pivot_fmt, use_container_width=True)
                
                    # Gráfico Dual Axis (Valor + Quantidade)
                    import plotly.graph_objects as go
                    from plotly.subplots import make_subplots

                    df_ano = df_filtered.groupby('ano').agg({
                        'valor': 'sum',
                        'matricula': 'count' 
                    }).reset_index()
                    df_ano.columns = ['Ano', 'Valor Total', 'Quantidade']
                    
                    st.markdown("---")
                    st.markdown("#### Análise Comparativa: Valor vs Quantidade")
                    
                    # Layout com Tabs - Visão Mensal como Padrão
                    tab_mensal, tab_anual = st.tabs(["📆 Visão Mensal", "📅 Visão Anual"])
                    
                    with tab_mensal:
                        # Gráfico Mês a Mês
                        fig_month = make_subplots(specs=[[{"secondary_y": True}]])
                        
                        # Barra - Valor (GRADIENTE VERDE)
                        fig_month.add_trace(
                            go.Bar(
                                x=df_agg['Periodo_Label'], y=df_agg['Valor_Total'],
                                name="Valor Investido",
                                marker=dict(
                                    color=APP_COLORS['primary'], # Cor Sólida
                                    showscale=False
                                ), 
                                text=df_agg['Valor_Total'],
                                texttemplate='R$ %{text:,.2f}',
                                textposition='outside',
                                textfont=dict(color='black', family="Arial", size=11, weight='bold')
                            ),
                            secondary_y=False
                        )
                        
                        # Linha - Quantidade (DARK SLATE)
                        fig_month.add_trace(
                            go.Scatter(
                                x=df_agg['Periodo_Label'], y=df_agg['Qtd_Pagamentos'],
                                name="Qtd Pagamentos",
                                mode='lines+markers+text',
                                marker=dict(color=APP_COLORS['secondary'], size=8), 
                                line=dict(width=3, color=APP_COLORS['secondary']),
                                text=df_agg['Qtd_Pagamentos'],
                                texttemplate='<b>%{text}</b>',
                                textposition='top center',
                                textfont=dict(
                                    color=APP_COLORS['secondary'],
                                    size=11,
                                    family="Arial Black"
                                )
                            ),
                            secondary_y=True
                        )
                        
                        fig_month.update_layout(
                            title="Valor vs Quantidade (Mês a Mês)",
                            height=500,
                            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                            yaxis=dict(showgrid=True, gridcolor='rgba(0,0,0,0.1)'),
                            xaxis=dict(type='category', tickangle=-45),
                            plot_bgcolor="white",
                            paper_bgcolor="white",
                        )
                        
                        # Ajuste de escalas 
                        max_val_m = df_agg['Valor_Total'].max() * 1.3 if len(df_agg) > 0 else 100
                        max_qtd_m = df_agg['Qtd_Pagamentos'].max() * 2.5 if len(df_agg) > 0 else 10

                        fig_month.update_yaxes(title_text="Valor Investido (R$)", range=[0, max_val_m], secondary_y=False)
                        fig_month.update_yaxes(title_text="Quantidade de Pagamentos", range=[0, max_qtd_m], secondary_y=True, showgrid=False)
                        
                        st.plotly_chart(fig_month, use_container_width=True, config=PLOTLY_CONFIG)

                    with tab_anual:
                        fig_combo = make_subplots(specs=[[{"secondary_y": True}]])
                        
                        # Barra - Valor (GRADIENTE VERDE)
                        fig_combo.add_trace(
                            go.Bar(
                                x=df_ano['Ano'], y=df_ano['Valor Total'],
                                name="Valor Investido",
                                marker=dict(
                                    color=APP_COLORS['primary'],
                                    showscale=False
                                ), 
                                text=df_ano['Valor Total'],
                                texttemplate='R$ %{text:,.2f}',
                                textposition='outside',
                                textfont=dict(color='black', family="Arial", size=11, weight='bold')
                            ),
                            secondary_y=False
                        )
                        
                        # Linha - Quantidade (DARK SLATE)
                        fig_combo.add_trace(
                            go.Scatter(
                                x=df_ano['Ano'], y=df_ano['Quantidade'],
                                name="Qtd Pagamentos",
                                mode='lines+markers+text',
                                marker=dict(color=APP_COLORS['secondary'], size=10),
                                line=dict(width=3, color=APP_COLORS['secondary']),
                                text=df_ano['Quantidade'],
                                texttemplate='<b>%{text}</b>',
                                textposition='top center',
                                textfont=dict(
                                    color=APP_COLORS['secondary'],
                                    size=11,
                                    family="Arial Black"
                                )
                            ),
                            secondary_y=True
                        )
                        
                        fig_combo.update_layout(
                            title="Valor vs Quantidade (Por Ano)",
                            height=500,
                            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                            yaxis=dict(showgrid=True, gridcolor='rgba(0,0,0,0.1)'),
                            xaxis=dict(type='category'), # Garantir que anos sejam mostrados como categorias
                            plot_bgcolor="white",
                            paper_bgcolor="white",
                        )
                        
                        # Ajuste de escalas 
                        max_val = df_ano['Valor Total'].max() * 1.3 
                        max_qtd = df_ano['Quantidade'].max() * 2.5

                        fig_combo.update_yaxes(title_text="Valor Investido (R$)", range=[0, max_val], secondary_y=False)
                        fig_combo.update_yaxes(title_text="Quantidade de Pagamentos", range=[0, max_qtd], secondary_y=True, showgrid=False)
                        
                        st.plotly_chart(fig_combo, use_container_width=True, config=PLOTLY_CONFIG)

                    st.markdown("---")
                    st.markdown("#### Análise por Ano Safra")
                    
                    # Criar tabela detalhada mês a mês por safra
                    df_safra_detalhada = df_filtered.copy()
                    
                    # Criar pivot table: Safra x Mês
                    df_pivot_safra = df_safra_detalhada.pivot_table(
                        values='valor',
                        index='safra',
                        columns='mes',
                        aggfunc='sum',
                        fill_value=0
                    )
                    
                    # Renomear colunas para ordenar corretamente na sequência da safra (Abr a Mar)
                    # MESES_SAFRA_NUM = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]
                    cols_ordenadas = []
                    for mes_num in MESES_SAFRA_NUM:
                        if mes_num in df_pivot_safra.columns:
                            cols_ordenadas.append(mes_num)
                    
                    # Reordenar colunas conforme ordem da safra
                    df_pivot_safra = df_pivot_safra[cols_ordenadas] if cols_ordenadas else df_pivot_safra
                    
                    # Renomear para nomes dos meses
                    df_pivot_safra.columns = [MESES[m-1] for m in df_pivot_safra.columns]
                    
                    # Adicionar coluna de Total
                    df_pivot_safra['Total'] = df_pivot_safra.sum(axis=1)
                    
                    # Adicionar linha de Quantidade de bolsistas por total da safra
                    df_qtd_safra = df_safra_detalhada.groupby('safra').size()
                    df_pivot_safra['Quantidade'] = df_qtd_safra
                    
                    # Formatar valores para exibição - FORMATO BRASILEIRO
                    df_safra_show = df_pivot_safra.copy()
                    
                    for col in df_safra_show.columns:
                        if col == 'Quantidade':
                            df_safra_show[col] = df_safra_show[col].apply(lambda x: format_br_number(x) if x > 0 else "-")
                        else:
                            df_safra_show[col] = df_safra_show[col].apply(lambda x: format_br_currency(x) if x > 0 else "-")

                    
                    # Resetar index para mostrar Safra como coluna
                    df_safra_show = df_safra_show.reset_index()
                    df_safra_show = df_safra_show.rename(columns={'safra': 'Safra'})
                    
                    # Preparar dados numéricos para o gráfico (antes da formatação)
                    df_safra_grafico = df_pivot_safra.copy()
                    df_safra_grafico = df_safra_grafico.reset_index()
                    df_safra_grafico['Valor Total'] = df_safra_grafico['Total']
                    df_safra_grafico['Quantidade_num'] = df_safra_grafico['Quantidade']
                    df_safra_grafico = df_safra_grafico.rename(columns={'safra': 'Safra'})
                    
                    c_safra_tab, c_safra_chart = st.columns([1, 2])
                    
                    with c_safra_tab:
                         st.markdown("##### Tabela Resumo")
                         # Mostrar apenas Safra, Total e Quantidade
                         df_resumo_compacto = df_safra_show[['Safra', 'Total', 'Quantidade']].copy()
                         df_resumo_compacto.columns = ['Safra', 'Valor Total', 'Quantidade']
                         st.dataframe(df_resumo_compacto, use_container_width=True, hide_index=True)
                         
                    with c_safra_chart:
                        # Gráfico Safra (Dual Axis)
                        fig_safra = make_subplots(specs=[[{"secondary_y": True}]])
                        
                        # Barra - Valor (GRADIENTE VERDE)
                        fig_safra.add_trace(
                            go.Bar(
                                x=df_safra_grafico['Safra'], y=df_safra_grafico['Valor Total'],
                                name="Valor Investido",
                                marker=dict(
                                    color=APP_COLORS['primary'],
                                    showscale=False
                                ), 
                                text=df_safra_grafico['Valor Total'],
                                texttemplate='R$ %{text:,.2f}',
                                textposition='outside',
                                textfont=dict(color='black', family="Arial", size=11, weight='bold')
                            ),
                            secondary_y=False
                        )
                        
                        # Linha - Quantidade (DARK SLATE)
                        fig_safra.add_trace(
                            go.Scatter(
                                x=df_safra_grafico['Safra'], y=df_safra_grafico['Quantidade_num'],
                                name="Qtd Pagamentos",
                                mode='lines+markers+text',
                                marker=dict(color=APP_COLORS['secondary'], size=10),
                                line=dict(width=3, color=APP_COLORS['secondary']),
                                text=df_safra_grafico['Quantidade_num'],
                                texttemplate='<b>%{text}</b>',
                                textposition='top center',
                                textfont=dict(color=APP_COLORS['secondary'], size=11, family="Arial Black")
                            ),
                            secondary_y=True
                        )
                        
                        # Ajuste de escalas
                        max_val_s = df_safra_grafico['Valor Total'].max() * 1.3
                        max_qtd_s = df_safra_grafico['Quantidade_num'].max() * 2.5

                        fig_safra.update_layout(
                            title="Safra: Valor vs Quantidade",
                            height=400,
                            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                            yaxis=dict(showgrid=True, gridcolor='rgba(0,0,0,0.1)'),
                            xaxis=dict(type='category'),
                            separators=',.',  # Formato brasileiro para Plotly
                            plot_bgcolor="white",
                            paper_bgcolor="white",
                        )
                        fig_safra.update_yaxes(title_text="Valor (R$)", range=[0, max_val_s], secondary_y=False)
                        fig_safra.update_yaxes(title_text="Qtd", range=[0, max_qtd_s], secondary_y=True, showgrid=False)
                        
                        st.plotly_chart(fig_safra, use_container_width=True, config=PLOTLY_CONFIG)
            
                with tab_top:
                    st.markdown("#### Todos os Colaboradores - Ranking por Valor Recebido")
                
                    # Busca
                    busca_colab = st.text_input("🔍 Buscar colaborador:", placeholder="Nome ou matrícula...", key="busca_colab_top")
                
                    # Todos colaboradores por valor total (usando dados filtrados)
                    df_top = df_filtered.groupby(['matricula', 'nome']).agg({
                        'valor': ['sum', 'count'],
                        'ano': ['min', 'max']
                    }).reset_index()
                    df_top.columns = ['Matrícula', 'Nome', 'Valor Total', 'Qtd Pagamentos', 'Primeiro Ano', 'Último Ano']
                    df_top = df_top.sort_values('Valor Total', ascending=False)
                
                    # Aplicar busca
                    if busca_colab:
                        df_top = df_top[
                            df_top['Nome'].str.contains(busca_colab, case=False, na=False) |
                            df_top['Matrícula'].astype(str).str.contains(busca_colab, case=False, na=False)
                        ]
                
                    # Métricas
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("👥 Total Colaboradores", len(df_top))
                    with col2:
                        st.metric("💰 Valor Total", f"R$ {df_top['Valor Total'].sum():,.2f}")
                    with col3:
                        st.metric("📊 Média/Colaborador", f"R$ {df_top['Valor Total'].mean():,.2f}" if len(df_top) > 0 else "R$ 0")
                
                    # Gráfico dos Top 20 (se não houver busca) ou todos da busca
                    df_chart = df_top.head(20) if not busca_colab else df_top
                
                    if len(df_chart) > 0:
                        n_items = len(df_chart)
                        
                        fig_top = px.bar(
                            df_chart,
                            y='Nome',
                            x='Valor Total',
                            orientation='h',
                            text='Valor Total',
                            color_discrete_sequence=[APP_COLORS['primary']]
                        )
                        fig_top.update_traces(
                            texttemplate='R$ %{text:,.2f}',
                            textposition='outside',
                            textfont=dict(color='black', family="Arial", size=11, weight='bold')
                        )
                        fig_top.update_layout(
                            coloraxis_showscale=False,
                            height=max(400, n_items * 35),
                            yaxis={'categoryorder': 'total ascending'},
                            showlegend=False,
                            margin=dict(l=20, r=120, t=40, b=40),
                            plot_bgcolor="white",
                            paper_bgcolor="white",
                        )
                        st.plotly_chart(fig_top, use_container_width=True, config=PLOTLY_CONFIG)
                
                    # Tabela completa com todos
                    st.markdown("##### 📋 Lista Completa")
                    df_top_show = df_top.copy()
                    df_top_show['Valor Total Fmt'] = df_top_show['Valor Total'].apply(lambda x: f"R$ {x:,.2f}")
                    df_top_show['Período'] = df_top_show['Primeiro Ano'].astype(str) + ' - ' + df_top_show['Último Ano'].astype(str)
                    df_top_show['Ranking'] = range(1, len(df_top_show) + 1)
                
                    st.dataframe(
                        df_top_show[['Ranking', 'Matrícula', 'Nome', 'Valor Total Fmt', 'Qtd Pagamentos', 'Período']].rename(
                            columns={'Valor Total Fmt': 'Valor Total'}
                        ),
                        use_container_width=True,
                        hide_index=True,
                        height=600
                    )
                
                    # Download Excel
                    excel_top = df_to_excel(df_top)
                    st.download_button("⬇️ Baixar Lista Completa", excel_top, "ranking_colaboradores.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_ranking")
            else:
                st.info("Nenhum dado de pagamento disponível para gerar a linha do tempo.")

    # =============================================
    # CADASTRAR
    # =============================================
    elif menu == "➕ Cadastrar":
        st.markdown("### ➕ Novo Bolsista")
        
        tab1, tab2 = st.tabs(["📝 Cadastro Manual", "📂 Importar Excel"])
        
        with tab1:
            # Função para buscar colaborador na base de gestores
            @st.cache_data(ttl=300)
            def carregar_base_gestores():
                """Carrega a base de gestores do Excel"""

                # Carregar Gestores do Organograma (Substitui o antigo gestores.xlsx)
                try:
                    df_gestores_file = get_dataset("ORGANOGRAMA")
                    if not df_gestores_file.empty:
                        df_gestores_file.columns = [str(c).upper().strip() for c in df_gestores_file.columns]
                        return df_gestores_file
                except Exception as e:
                    logger.error(f"Erro ao carregar ORGANOGRAMA para gestores: {e}")
                return pd.DataFrame()
            
            def buscar_colaborador(matricula_busca):
                """Busca um colaborador na base de gestores pela matricula"""
                df_gestores = carregar_base_gestores()
                if df_gestores.empty:
                    return None
                
                # Normalizar matrícula para busca
                matricula_busca = str(matricula_busca).strip()
                
                # Buscar na coluna MATRICULA
                if 'MATRICULA' in df_gestores.columns:
                    df_gestores['MATRICULA'] = df_gestores['MATRICULA'].astype(str).str.strip()
                    resultado = df_gestores[df_gestores['MATRICULA'] == matricula_busca]
                    if not resultado.empty:
                        return resultado.iloc[0].to_dict()
                return None
            
            # Campo de matrícula FORA do form para permitir busca dinâmica
            st.markdown("#### 🔍 Buscar Colaborador")
            col_mat, col_btn = st.columns([3, 1])
            with col_mat:
                matricula = st.text_input("Matrícula *", key="matricula_cadastro", placeholder="Digite a matrícula e clique em Buscar")
            with col_btn:
                st.write("")  # Spacer
                buscar = st.button("🔍 Buscar", type="primary", use_container_width=True)
            
            # Inicializar session_state para os dados do colaborador
            if 'dados_gestor' not in st.session_state:
                st.session_state.dados_gestor = {}
            
            # Buscar dados quando clicar no botão
            if buscar and matricula:
                colab = buscar_colaborador(matricula)
                if colab:
                    st.session_state.dados_gestor = colab
                    st.success(f"✅ Colaborador encontrado: **{colab.get('COLABORADOR', 'N/A')}**")
                else:
                    st.session_state.dados_gestor = {}
                    st.warning("⚠️ Matrícula não encontrada na base de gestores. Preencha manualmente.")
            
            # Pegar dados do session_state
            dados_g = st.session_state.get('dados_gestor', {})
            
            st.markdown("---")
            st.markdown("#### 📝 Dados do Bolsista")
            
            with st.form("cadastro"):
                col1, col2 = st.columns(2)
                with col1:
                    # Matrícula já preenchida
                    st.text_input("Matrícula", value=matricula, disabled=True, key="mat_display")
                    nome = st.text_input("Nome *", value=dados_g.get('COLABORADOR', ''))
                    cpf = st.text_input("CPF", value=dados_g.get('CPF FORMATADO', dados_g.get('CPF', '')))
                    
                    # Diretoria - tentar pegar do gestor
                    diretoria_gestor = dados_g.get('DIRETORIA', '')
                    if diretoria_gestor and diretoria_gestor in DIRETORIAS:
                        idx_dir = DIRETORIAS.index(diretoria_gestor) + 1
                    else:
                        idx_dir = 0
                    diretoria = st.selectbox("Diretoria", [""] + DIRETORIAS, index=idx_dir)
                    
                    ano_ref = st.number_input("Ano Referência", min_value=2000, max_value=2100, value=datetime.now().year)
                
                with col2:
                    # Tentar pegar curso da base gestores se existir
                    curso_gestor = dados_g.get('BASE BOLSAS.CURSO', '')
                    curso = st.text_input("Curso", value=curso_gestor if curso_gestor else '')
                    instituicao = st.text_input("Instituição")
                    
                    c_tipo, c_mod = st.columns(2)
                    with c_tipo:
                        tipo = st.selectbox("Tipo", ["", "Técnico", "Graduação", "Pós-Graduação", "MBA", "Mestrado", "Outros"])
                    with c_mod:
                        modalidade = st.selectbox("Modalidade", ["", "Presencial", "EAD", "Híbrido", "Semipresencial"])
                    
                    inicio = st.date_input("Início Curso", value=datetime.today(), format="DD/MM/YYYY")
                    fim = st.date_input("Fim Curso", value=datetime.today(), format="DD/MM/YYYY")
                    situacao = st.selectbox("Situação", ["ATIVO", "INATIVO"])
                    checagem = st.selectbox("Checagem", ["REGULAR", "IRREGULAR", "CONCLUIDO", "CANCELADO", "DESISTENCIA", "TRANCADO", "DEMITIDO"])
                
                st.markdown("#### 💰 Valores")
                c1, c2, c3 = st.columns(3)
                with c1:
                    mensalidade = st.number_input("Mensalidade", min_value=0.0, step=10.0)
                with c2:
                    pct = st.slider("% Bolsa", 0, 100, 50) / 100
                with c3:
                    valor = mensalidade * pct
                    st.metric("Reembolso", f"R$ {valor:,.2f}")
                
                obs = st.text_area("Observações")
                
                if st.form_submit_button("💾 Cadastrar", type="primary", use_container_width=True):
                    if not matricula or not nome:
                        st.error("Matrícula e Nome obrigatórios!")
                    else:
                        dados = {
                            'matricula': matricula.strip(), 'nome': nome.upper(), 'cpf': cpf,
                            'diretoria': diretoria, 'curso': curso, 'instituicao': instituicao,
                            'tipo': tipo, 'modalidade': modalidade,
                            'inicio_curso': inicio, 'fim_curso': fim, 'ano_referencia': int(ano_ref),
                            'mensalidade': mensalidade, 'porcentagem': pct, 
                            'valor_reembolso': valor, 'situacao': situacao, 
                            'checagem': checagem, 'observacao': obs
                        }
                        ok, msg = cadastrar_bolsista(dados)
                        if ok:
                            st.success(f"✅ {msg}")
                            # Limpar dados do gestor após cadastro
                            st.session_state.dados_gestor = {}
                            st.balloons()
                        else:
                            st.error(msg)

                            
        with tab2:
            st.markdown("### 🔄 Sincronização e Importação")
            st.info("Utilize esta aba para cadastrar novos ou atualizar dados existentes a partir do Excel.")
            
            # Novo botão de download do template
            st.markdown("#### 1. Baixar Modelo de Importação")
            st.markdown("Caso não tenha o arquivo, baixe o modelo padrão abaixo:")
            excel_template = gerar_template_excel()
            st.download_button(
                label="📥 Baixar Planilha Modelo (Vazia)",
                data=excel_template,
                file_name="Modelo_Cadastro_Bolsistas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=False
            )
            st.markdown("---")
            
            st.markdown("#### 2. Importar Dados")
            col_local, col_upload = st.columns(2)
            
            with col_local:
                st.markdown("#### 📂 Arquivo Local")
                # Template / Arquivo padrão / Google Sheets
                st.caption(f"Fonte: Google Sheets (Se conectado) ou `BASES.BOLSAS/BASE.BOLSAS.2025.xlsx`")
                
                sobrescrever_imp = st.checkbox("Sobrescrever Status/Obs?", value=False, key="check_sobrescrever_imp")
                
                if st.button(f"🔄 Sincronizar Agora", type="primary", use_container_width=True):
                    try:
                        df_local = get_dataset("BOLSAS")
                        if not df_local.empty:
                            st.info(f"Dados carregados! {len(df_local)} registros.")
                            processar_importacao_df(df_local, preserve_status=not sobrescrever_imp)
                        else:
                            st.error(f"Não foi possível carregar dados da fonte (Sheets ou Local).")
                    except Exception as e:
                        st.error(f"Erro ao ler dados: {e}")
            
            with col_upload:
                 st.markdown("#### ⬆️ Upload de Arquivo")
                 uploaded_file = st.file_uploader("Arraste seu Excel aqui", type=['xlsx', 'xlsm'])
                 if uploaded_file:
                    if st.button("🚀 Processar Upload", use_container_width=True):
                         try:
                             df_up = pd.read_excel(uploaded_file)
                             if df_up is not None:
                                 processar_importacao_df(df_up, preserve_status=not sobrescrever_imp)
                         except Exception as e:
                             st.error(f"Erro ao ler upload: {e}")

    # =============================================
    # RODAPÉ DISCRETO COM AÇÕES DO SISTEMA
    # =============================================
    st.markdown("---")
    with st.expander("⚙️ Configurações do Sistema", expanded=False):
        col_sys1, col_sys2, col_sys_spacer = st.columns([1, 1, 4])
        with col_sys1:
            if st.button("🗑️ Limpar Cache", help="Força o sistema a recarregar todos os dados", use_container_width=True, key="btn_limpar_cache_footer"):
                st.cache_data.clear()
                st.success("Cache limpo com sucesso!")
                st.rerun()
        with col_sys2:
            if st.button("🔄 Atualizar Dados", help="Reprocessa o cruzamento com o Organograma", use_container_width=True, key="btn_atualizar_footer"):
                st.cache_data.clear()
                st.rerun()

if __name__ == "__main__":
    main()
