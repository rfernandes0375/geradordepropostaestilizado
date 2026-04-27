import streamlit as st
import pandas as pd
import io
import os
import tempfile
from datetime import datetime
import time
import base64
from pathlib import Path
import sys
import re
import zipfile
import subprocess
import tempfile
import base64
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from streamlit_gsheets import GSheetsConnection

def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

def get_img_with_href(local_img_path):
    img_format = local_img_path.split(".")[-1]
    bin_str = get_base64_of_bin_file(local_img_path)
    return f"data:image/{img_format};base64,{bin_str}"

logo_base64 = ""
if os.path.exists("logo_jardim.png"):
    logo_base64 = get_img_with_href("logo_jardim.png")
else:
    logo_base64 = "https://i.postimg.cc/qqvQS9S9/jardim.png?text=AJCE+BRASIL" # Fallback

temp_dir = tempfile.gettempdir()

from motor_pdf import *
# --- Configuração da Página Streamlit ---
st.set_page_config(
    page_title="Gerador de Propostas Jardim Equipamentos",
    page_icon=logo_base64 if logo_base64 else "📊",
    layout="wide",
    initial_sidebar_state="collapsed" # Começa com sidebar recolhida
)

# --- Estilos CSS Customizados (Mantidos) ---
def load_css():
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&display=swap');

        /* Esconder a marca do Streamlit */
        #MainMenu {visibility: hidden;}
        header {visibility: hidden;}
        footer {visibility: hidden;}

        /* Fundo principal Estilo Claro Moderno */
        [data-testid="stAppViewContainer"] {
            background-color: #f8fafc;
            background-image: radial-gradient(circle at 10% 20%, rgba(59, 130, 246, 0.05), transparent 20%),
                              radial-gradient(circle at 90% 80%, rgba(16, 185, 129, 0.05), transparent 20%);
            color: #1e293b;
        }

        .main .block-container {
             padding-top: 2rem;
             padding-bottom: 3rem;
             max-width: 1050px;
        }

        /* Títulos e Textos */
        h1, h2, h3, h4, h5, h6 {
            color: #0f172a !important;
            font-family: 'Inter', sans-serif !important;
            font-weight: 800 !important;
            letter-spacing: -0.025em !important;
        }
        
        p, label, .stMarkdown {
            color: #334155 !important;
            font-family: 'Inter', sans-serif !important;
        }

        /* Cabeçalho com logo — layout horizontal */
        .header-container {
            display: flex;
            flex-direction: row;
            align-items: center;
            justify-content: center;
            gap: 1.5rem;
            margin-bottom: 1.5rem;
            margin-top: 0.5rem;
            padding: 1.25rem 2rem;
            background: linear-gradient(135deg, rgba(37,99,235,0.05) 0%, rgba(16,185,129,0.05) 100%);
            border-radius: 20px;
            border: 1px solid rgba(37, 99, 235, 0.1);
        }
        .header-identity { display: flex; flex-direction: column; }
        .logo-img {
            height: 70px;
            filter: drop-shadow(0px 4px 10px rgba(0,0,0,0.08));
            flex-shrink: 0;
        }
        .header-title {
            text-align: left;
            background: linear-gradient(135deg, #2563eb 0%, #1e40af 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            font-size: 2.2rem !important;
            font-weight: 800 !important;
            margin: 0 !important;
        }
        .header-subtitle {
            color: #64748b;
            font-size: 0.875rem;
            margin: 0.2rem 0 0;
            font-weight: 500;
        }

        /* Soft Glassmorphism nos Containers (Tema Claro) */
        [data-testid="stVerticalBlock"] > [style*="border: 1px solid rgba(49, 51, 63, 0.2)"],
        [data-testid="stVerticalBlock"] > [style*="border: 1px solid rgba(250, 250, 250, 0.2)"],
        div[data-testid="stVerticalBlockBorderWrapper"] > div {
            background: rgba(255, 255, 255, 0.75) !important;
            backdrop-filter: blur(12px) !important;
            -webkit-backdrop-filter: blur(12px) !important;
            border: 1px solid rgba(226, 232, 240, 0.8) !important;
            border-radius: 16px !important;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.08), 0 2px 4px -1px rgba(0, 0, 0, 0.04) !important;
            padding: 1.5rem !important;
            margin-bottom: 1.25rem !important;
        }

        /* Estilos das Abas (Tabs) Claras */
        .stTabs [data-baseweb="tab-list"] {
            gap: 12px;
            background-color: #f1f5f9 !important;
            border-radius: 40px;
            padding: 6px !important;
            border: 1px solid #e2e8f0;
            margin-bottom: 2rem;
            width: fit-content;
            margin-left: auto;
            margin-right: auto;
        }
        .stTabs [data-baseweb="tab"] {
            background-color: transparent !important;
            border: none !important;
            border-radius: 35px !important;
            padding: 8px 24px !important;
            color: #64748b !important;
            font-weight: 600 !important;
            transition: all 0.2s ease !important;
        }
        .stTabs [aria-selected="true"] {
             background-color: white !important;
             color: #2563eb !important;
             box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05) !important;
        }

        /* Botões Primários (Gradiente Elegante) */
        .stButton>button[kind="primary"] {
            background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
            color: white !important;
            border: none !important;
            border-radius: 12px !important;
            font-weight: 600 !important;
            padding: 0.8rem 1.5rem !important;
            box-shadow: 0 4px 12px rgba(16, 185, 129, 0.2) !important;
            width: 100%;
            transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
        }
        .stButton>button[kind="primary"]:hover {
             transform: translateY(-2px);
             box-shadow: 0 8px 16px rgba(16, 185, 129, 0.3) !important;
        }

        /* Inputs e Caixa de Arquivo (Tema Claro) */
        input[type="text"], input[type="number"], .stNumberInput > div > div > input {
            background-color: white !important;
            border: 1px solid #e2e8f0 !important;
            color: #1e293b !important;
            border-radius: 10px !important;
            padding: 0.6rem !important;
        }
        
        .stSelectbox > div > div > div {
            background-color: white !important;
            border: 1px solid #e2e8f0 !important;
            color: #1e293b !important;
            border-radius: 10px !important;
        }

        /* Inputs Desabilitados (Passo 3) */
        input:disabled {
            background-color: #f1f5f9 !important;
            color: #475569 !important;
            opacity: 1 !important;
            -webkit-text-fill-color: #475569 !important;
        }

        /* Caixa de Upload do Streamlit */
        [data-testid="stFileUploadDropzone"] {
            background: #ffffff !important;
            border: 2px dashed #cbd5e1 !important;
            border-radius: 16px !important;
        }
        [data-testid="stFileUploadDropzone"]:hover {
            border-color: #3b82f6 !important;
            background: #f8fafc !important;
        }

        /* Rodapé */
        .footer {
            margin-top: 2rem;
            padding: 1.5rem 2rem;
            border-top: 1px solid #e2e8f0;
            text-align: center;
            color: #94a3b8;
            font-size: 0.875rem;
            background: linear-gradient(135deg, rgba(37,99,235,0.03) 0%, rgba(16,185,129,0.03) 100%);
            border-radius: 16px;
        }
        .footer strong { color: #475569; }
        .footer a { color: #2563eb; text-decoration: none; }
        .footer a:hover { text-decoration: underline; }

        /* Barra de Progresso por Etapas */
        .progress-stepper {
            display: flex;
            align-items: flex-start;
            justify-content: center;
            max-width: 480px;
            margin: 0 auto 1.5rem;
        }
        .step-item {
            display: flex;
            flex-direction: column;
            align-items: center;
            flex: 1;
        }
        .step-circle {
            width: 34px; height: 34px;
            border-radius: 50%;
            background: #e2e8f0;
            border: 2px solid #cbd5e1;
            display: flex; align-items: center; justify-content: center;
            font-weight: 700; font-size: 0.85rem;
            color: #94a3b8;
            transition: all 0.3s ease;
        }
        .step-circle.done { background: #10b981; border-color: #10b981; color: white; }
        .step-circle.active { background: #2563eb; border-color: #2563eb; color: white; box-shadow: 0 0 0 4px rgba(37,99,235,0.15); }
        .step-label {
            font-size: 0.68rem; color: #94a3b8;
            margin-top: 0.35rem; text-align: center;
            font-weight: 600; text-transform: uppercase; letter-spacing: 0.04em;
        }
        .step-label.active { color: #2563eb; }
        .step-label.done { color: #10b981; }
        .step-connector {
            flex: 1; height: 2px;
            background: #e2e8f0;
            margin-top: 16px;
        }
        .step-connector.done { background: #10b981; }

        /* Card de Resumo (Passo 3) */
        .summary-card {
            display: flex;
            flex-wrap: wrap;
            gap: 1rem;
            background: linear-gradient(135deg, rgba(37,99,235,0.05) 0%, rgba(16,185,129,0.05) 100%);
            border: 1px solid rgba(37, 99, 235, 0.15);
            border-radius: 14px;
            padding: 1.25rem 1.5rem;
            margin-bottom: 1.25rem;
        }
        .summary-item { flex: 1; min-width: 140px; }
        .summary-item .s-label {
            font-size: 0.7rem; font-weight: 700;
            text-transform: uppercase; letter-spacing: 0.05em;
            color: #64748b; margin-bottom: 0.2rem;
        }
        .summary-item .s-value {
            font-size: 1rem; font-weight: 700; color: #0f172a;
        }
        .summary-item .s-value.money { color: #059669; }

        /* Responsividade mobile */
        @media (max-width: 768px) {
            .header-container { flex-direction: column; gap: 0.75rem; padding: 1rem; }
            .header-title { font-size: 1.6rem !important; text-align: center; }
            .header-subtitle { text-align: center; }
            .main .block-container { padding-left: 1rem !important; padding-right: 1rem !important; }
            .summary-card { flex-direction: column; }
        }
        @media (max-width: 480px) {
            .header-title { font-size: 1.3rem !important; }
            .logo-img { height: 50px; }
        }
    </style>
    """, unsafe_allow_html=True)

load_css()

# --- Cabeçalho com Logo ---
def render_header():
    st.markdown(f"""
    <div class="header-container">
        <img class="logo-img" src="{logo_base64}" alt="Logo Jardim Equipamentos">
        <div class="header-identity">
            <h1 class="header-title">Gerador de Propostas</h1>
            <span class="header-subtitle">Jardim Equipamentos — Propostas Comerciais</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

# --- Barra de Progresso por Etapas ---
def render_progress_bar():
    step = {"Upload": 1, "Seleção": 2, "Geração": 3}.get(st.session_state.get('current_tab', 'Upload'), 1)
    steps = [("1", "Upload"), ("2", "Seleção"), ("3", "Gerar PDF")]
    html = '<div class="progress-stepper">'
    for i, (num, label) in enumerate(steps, 1):
        if i < step:
            cc, lc, cn = 'done', 'done', '✓'
        elif i == step:
            cc, lc, cn = 'active', 'active', num
        else:
            cc, lc, cn = '', '', num
        html += f'<div class="step-item"><div class="step-circle {cc}">{cn}</div><span class="step-label {lc}">{label}</span></div>'
        if i < len(steps):
            html += f'<div class="step-connector {"done" if i < step else ""}" ></div>'
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)

render_header()

# --- Inicialização do Estado da Sessão ---
if 'current_tab' not in st.session_state:
    st.session_state['current_tab'] = "Upload"
if 'planilha_data' not in st.session_state:
    st.session_state['planilha_data'] = None 
if 'planilha_nome' not in st.session_state:
    st.session_state['planilha_nome'] = None
if 'modelos_info' not in st.session_state:
    st.session_state['modelos_info'] = {} 
if 'dados_linha_selecionada' not in st.session_state:
    st.session_state['dados_linha_selecionada'] = None
if 'modelo_selecionado_nome' not in st.session_state:
    st.session_state['modelo_selecionado_nome'] = None

render_progress_bar()

# --- Criação das Abas ---
tab_upload, tab_selecao, tab_geracao = st.tabs([
    "📤 1. Upload de Arquivos",
    "📊 2. Seleção de Dados",
    "🖨️ 3. Gerar Proposta"
])


# --- Aba 1: Upload de Arquivos ---
with tab_upload:
    st.header("Passo 1: Fonte dos Modelos e Dados")
    st.markdown("---")

    col_meta1, col_meta2 = st.columns(2)
    
    with col_meta1:
        with st.container(border=True):
            status_planilha = "✅" if st.session_state.get('planilha_data') is not None else "⏳"
            st.subheader(f"📊 Banco de Dados {status_planilha}")
            origem_dados = st.radio("Origem da Planilha:", ["Upload Manual", "Google Sheets (Nuvem)"], horizontal=True, key="origem_dados", index=1)
            
            if origem_dados == "Upload Manual":
                arquivo_planilha = st.file_uploader(
                    "Upload da Planilha (.ods, .xlsx)",
                    type=["ods", "xlsx", "xls"],
                    key="planilha_upload_widget",
                    label_visibility="collapsed"
                )
                if arquivo_planilha:
                    try:
                        planilha_bytes = arquivo_planilha.getvalue()
                        engine = 'odf' if arquivo_planilha.name.endswith('.ods') else None
                        df = pd.read_excel(io.BytesIO(planilha_bytes), engine=engine)
                        st.session_state['planilha_data'] = df
                        st.session_state['planilha_nome'] = arquivo_planilha.name
                        
                        # Lógica para setar a última linha preenchida como padrão
                        if 'Cliente' in df.columns:
                            last_filled_index = df[df['Cliente'].notna() & (df['Cliente'].astype(str).str.strip() != "")].index
                            if not last_filled_index.empty:
                                st.session_state['last_selected_line'] = int(last_filled_index[-1]) + 2
                            else: st.session_state['last_selected_line'] = 2
                        
                        st.success(f"✅ Planilha '{arquivo_planilha.name}' carregada.")
                    except Exception as e:
                        st.error(f"❌ Erro ao ler planilha: {e}")
            else:
                st.info("🔗 Conectando ao Google Sheets...")
                try:
                    # Tenta conectar via Streamlit GSheets Connection
                    conn = st.connection("gsheets", type=GSheetsConnection)
                    # O usuário precisará fornecer a URL ou configurar no secrets
                    # Endereço padrão da planilha fornecido pelo usuário
                    url_sheets_default = "https://docs.google.com/spreadsheets/d/1CkllnC9xWKzZEnB_bKFRuNY5dsu1FoGRF5phdHirjbs/edit#gid=465824771"
                    
                    with st.expander("🔗 Configurar Link da Planilha", expanded=False):
                        url_sheets = st.text_input("Link da Planilha Google:", value=url_sheets_default, placeholder="https://docs.google.com/spreadsheets/d/...")
                        
                        col_refresh, _ = st.columns([1, 2])
                        with col_refresh:
                            if st.button("🔄 Atualizar Dados", use_container_width=True):
                                st.cache_data.clear() # Limpa qualquer cache residual
                                st.rerun()

                    if url_sheets:
                        # ttl=0 desativa o cache de 10 minutos para atualizações instantâneas
                        df = conn.read(spreadsheet=url_sheets, ttl=0)
                        st.session_state['planilha_data'] = df
                        st.session_state['planilha_nome'] = "Google Sheets"

                        # Identifica a última linha preenchida na coluna 'Cliente'
                        if 'Cliente' in df.columns:
                            # Filtra linhas onde 'Cliente' não é nulo nem vazio
                            linhas_preenchidas = df[df['Cliente'].notna() & (df['Cliente'].astype(str).str.strip() != "")].index
                            if not linhas_preenchidas.empty:
                                # Define a última linha como padrão (Índice + 2 para bater com o número da linha no Excel/Sheets)
                                st.session_state['last_selected_line'] = int(linhas_preenchidas[-1]) + 2
                            else:
                                st.session_state['last_selected_line'] = 2
                except Exception as e:
                    if "not found" in str(e).lower():
                        st.error("❌ Planilha não encontrada. Verifique se o link está correto e se você compartilhou com o e-mail da Service Account.")
                    elif "credential" in str(e).lower():
                        st.error("⚠️ Credenciais do Google não configuradas corretamente no secrets.toml.")
                    else:
                        st.error(f"❌ Erro de Conexão: {e}")

    with col_meta2:
        with st.container(border=True):
            status_modelos = "✅" if (st.session_state.get('modelos_info') or st.session_state.get('modelos_drive_info')) else "⏳"
            st.subheader(f"📄 Modelos de Proposta {status_modelos}")
            origem_modelos = st.radio("Origem dos Modelos:", ["Upload Manual", "Google Drive"], horizontal=True, key="origem_modelos", index=1) # Google Drive como padrão
            
            if origem_modelos == "Upload Manual":
                arquivos_modelo = st.file_uploader(
                    "Upload de Modelos ODT",
                    type=["odt"],
                    accept_multiple_files=True,
                    key="modelos_upload_widget",
                    label_visibility="collapsed"
                )
                if arquivos_modelo:
                     st.session_state['modelos_info'] = {modelo.name: modelo.getvalue() for modelo in arquivos_modelo}
                     st.success(f"✅ {len(arquivos_modelo)} modelos carregados.")
            else:
                st.info("📂 Buscando modelos no Google Drive...")
                # ID padrão da pasta fornecido pelo usuário
                folder_id_default = "1daF_KyzA1te7cMFBuKJaAzvYxX7D7gKu"
                
                with st.expander("📂 Configurar Pasta de Modelos", expanded=False):
                    folder_id = st.text_input("ID da Pasta no Drive:", value=folder_id_default, placeholder="1A2B3... (final da URL da pasta)")
                
                if folder_id:
                    modelos_drive = listar_modelos_google_drive(folder_id)
                    if modelos_drive:
                        st.success(f"✅ {len(modelos_drive)} modelos encontrados no Drive!")
                        # Armazenamos os IDs, Links e MIME types
                        st.session_state['modelos_drive_info'] = {m['name']: {'id': m['id'], 'mimeType': m['mimeType']} for m in modelos_drive}
                        st.session_state['modelos_drive_links'] = {m['name']: m['webViewLink'] for m in modelos_drive}
                        # Para compatibilidade com o resto do código, inicializamos o dicionário de bytes
                        st.session_state['modelos_info'] = {m['name']: None for m in modelos_drive} 
                    else:
                        st.error("❌ Nenhum modelo .odt encontrado ou acesso negado.")

    st.divider()

    if st.session_state.get('planilha_data') is not None and (st.session_state.get('modelos_info') or st.session_state.get('modelos_drive_info')):
        if st.button("Avançar para Seleção de Dados →", type="primary", key="goto_selecao"):
            st.session_state['current_tab'] = "Seleção"
            st.rerun() 
    else:
         st.info("ℹ️ Configure a planilha e os modelos para continuar.")


# --- Aba 2: Seleção de Dados ---
with tab_selecao:
    st.header("Passo 2: Selecione os Dados para a Proposta")
    st.markdown("---")

    if st.session_state['planilha_data'] is None or not st.session_state['modelos_info']:
        st.warning("⚠️ Volte ao Passo 1 e faça o upload da planilha e dos modelos ODT.")
        if st.button("← Voltar para Upload", key="back_to_upload_selecao"):
            st.session_state['current_tab'] = "Upload"
            st.rerun()
    else:
        df = st.session_state['planilha_data']

        with st.expander("👁️ Visualizar Planilha Carregada", expanded=False):
             st.dataframe(df, use_container_width=True, height=300) 

        st.divider()

        # Novo Layout Lado a Lado: Seleção (Esquerda) e Gerenciamento (Direita)
        col_selecao, col_gerenciamento = st.columns([0.35, 0.65], gap="large")

        with col_selecao:
            with st.container(border=True):
                st.subheader("Selecione a Linha da Planilha")
                st.caption("Escolha a linha que contém os dados para esta proposta específica.")

                linha_selecionada_usuario = st.number_input(
                     f"Número da linha (de 2 a {len(df) + 1}):", 
                     min_value=2,
                     max_value=len(df) + 1,
                     value=st.session_state.get('last_selected_line', 2), 
                     step=1,
                     key="linha_input_selecao"
                 )

                linha_indice_zero = linha_selecionada_usuario - 2

                if 0 <= linha_indice_zero < len(df):
                     st.session_state['dados_linha_selecionada'] = df.iloc[linha_indice_zero].fillna('').to_dict()
                     st.session_state['last_selected_line'] = linha_selecionada_usuario

                     with st.expander("🔍 Pré-visualizar Dados da Linha Selecionada", expanded=True):
                          cliente = st.session_state['dados_linha_selecionada'].get('Cliente', '—')
                          modelo = st.session_state['dados_linha_selecionada'].get('Modelo', '—')
                          v_romp = st.session_state['dados_linha_selecionada'].get('Valor Rompedor', '—')
                          v_kit = st.session_state['dados_linha_selecionada'].get('Valor Kit', '—')
                          
                          st.markdown(f"""
                          <div style="background: rgba(37,99,235,0.05); padding: 1rem; border-radius: 12px; border: 1px solid rgba(37,99,235,0.1);">
                              <div style="font-weight: 800; color: #1e293b; font-size: 1.1rem; margin-bottom: 0.5rem;">👤 {cliente}</div>
                              <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 0.5rem; font-size: 0.9rem;">
                                  <div><strong style="color: #64748b;">Modelo:</strong> <br>{modelo}</div>
                                  <div>
                                      <strong style="color: #64748b;">Valores:</strong> <br>
                                      <span style="color: #059669; font-weight: 600;">Rompedor: {v_romp}</span><br>
                                      <span style="color: #059669; font-weight: 600;">Kit: {v_kit}</span>
                                  </div>
                              </div>
                          </div>
                          """, unsafe_allow_html=True)

                else:
                     st.error(f"❌ Linha {linha_selecionada_usuario} inválida. Selecione um valor entre 2 e {len(df) + 1}.")
                     st.session_state['dados_linha_selecionada'] = None 

        with col_gerenciamento:
            # --- GERENCIAMENTO DE DADOS GOOGLE SHEETS ---
            if origem_dados == "Google Sheets (Nuvem)" and st.session_state['planilha_data'] is not None:
                with st.container(border=True):
                    st.subheader("📝 Gerenciar Banco de Dados")
                    tab_edit, tab_add = st.tabs(["✏️ Editar Selecionada", "➕ Adicionar Nova"])
                    
                    # Todas as colunas da planilha carregada
                    todas_colunas = list(st.session_state['planilha_data'].columns)
                    
                    with tab_edit:
                        if st.session_state['dados_linha_selecionada']:
                            linha_atual = st.session_state['last_selected_line']
                            st.write(f"Editando dados da linha **{linha_atual}**:")
                            with st.form("form_edit_sheets"):
                                novos_dados = {}
                                # Renderiza em pares de 2 colunas
                                for i in range(0, len(todas_colunas), 2):
                                    cols_f = st.columns(2)
                                    for j, campo in enumerate(todas_colunas[i:i+2]):
                                        valor_atual = str(st.session_state['dados_linha_selecionada'].get(campo, ""))
                                        if valor_atual in ("nan", "None", "NaT"):
                                            valor_atual = ""
                                        with cols_f[j]:
                                            novos_dados[campo] = st.text_input(
                                                campo,
                                                value=valor_atual,
                                                key=f"edit_{campo}_{linha_atual}"
                                            )
                                
                                if st.form_submit_button("💾 Salvar Alterações na Nuvem", use_container_width=True):
                                    try:
                                        conn = st.connection("gsheets", type=GSheetsConnection)
                                        df_atualizado = st.session_state['planilha_data'].copy()
                                        idx = st.session_state['last_selected_line'] - 2
                                        
                                        # Automação do campo NOME DO ARQUIVO no Edit
                                        if 'Número' in novos_dados and 'Cliente' in novos_dados:
                                            num_f = novos_dados['Número']
                                            cli_f = novos_dados['Cliente']
                                            if num_f and cli_f:
                                                 if 'NOME DO ARQUIVO' in df_atualizado.columns:
                                                     novos_dados['NOME DO ARQUIVO'] = f"JE {num_f} - {cli_f}"

                                        for key, val in novos_dados.items():
                                            df_atualizado.at[idx, key] = val
                                        conn.update(spreadsheet=url_sheets, data=df_atualizado)
                                        st.session_state['planilha_data'] = df_atualizado
                                        st.success("✅ Planilha atualizada!")
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"Erro ao salvar: {e}")
                        else:
                            st.info("Selecione uma linha à esquerda para editar.")

                    with tab_add:
                        st.write(f"Novo cliente:")
                        with st.form("form_add_sheets"):
                            dados_novos_row = {}
                            for i in range(0, len(todas_colunas), 2):
                                cols_f = st.columns(2)
                                for j, campo in enumerate(todas_colunas[i:i+2]):
                                    with cols_f[j]:
                                        if campo == "Data":
                                            hoje = datetime.today()
                                            meses = {1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL", 5: "MAIO", 6: "JUNHO", 7: "JULHO", 8: "AGOSTO", 9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"}
                                            valor_padrao = f"Goiânia, {hoje.day} de {meses[hoje.month]} de {hoje.year}"
                                        elif campo == "Número":
                                            valor_padrao = obter_proximo_numero(st.session_state['planilha_data'])
                                        else:
                                            valor_padrao = ""
                                            
                                        dados_novos_row[campo] = st.text_input(campo, value=valor_padrao, key=f"add_{campo}")
                            
                            if st.form_submit_button("➕ Adicionar à Planilha", use_container_width=True):
                                try:
                                    conn = st.connection("gsheets", type=GSheetsConnection)
                                    df_atualizado = st.session_state['planilha_data'].copy()
                                    nova_linha = {col: "" for col in df_atualizado.columns}
                                    nova_linha.update(dados_novos_row)
                                    
                                    if 'Número' in nova_linha and 'Cliente' in nova_linha:
                                        num_f = nova_linha['Número']
                                        cli_f = nova_linha['Cliente']
                                        if num_f and cli_f and 'NOME DO ARQUIVO' in df_atualizado.columns:
                                            nova_linha['NOME DO ARQUIVO'] = f"JE {num_f} - {cli_f}"
                                    
                                    df_atualizado = pd.concat([df_atualizado, pd.DataFrame([nova_linha])], ignore_index=True)
                                    conn.update(spreadsheet=url_sheets, data=df_atualizado)
                                    st.session_state['planilha_data'] = df_atualizado
                                    st.success(f"✅ Proposta adicionada!")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Erro ao adicionar: {e}")
            else:
                 st.info("💡 Carregue dados do Google Sheets no Passo 1 para gerenciar o banco aqui.")

        st.divider()
        
        # Seleção de Modelo ODT
        with st.container(border=True):
             st.subheader("📄 Selecione o Modelo ODT")
             st.caption("Escolha qual modelo ODT será usado para esta proposta.")
             nomes_modelos = list(st.session_state['modelos_info'].keys())

             if nomes_modelos:
                  col_m1, col_m2 = st.columns([0.7, 0.3])
                  with col_m1:
                       modelo_selecionado = st.selectbox(
                           "Modelos Disponíveis:",
                           options=nomes_modelos,
                           index=nomes_modelos.index(st.session_state.get('modelo_selecionado_nome', nomes_modelos[0])) if st.session_state.get('modelo_selecionado_nome') in nomes_modelos else 0,
                           key="modelo_select_widget"
                       )
                       st.session_state['modelo_selecionado_nome'] = modelo_selecionado
                  
                  with col_m2:
                       st.write("") # alinhamento
                       st.write("") 
                       # Link para visualizar/editar modelo no Google Drive
                       if 'modelos_drive_links' in st.session_state and modelo_selecionado in st.session_state['modelos_drive_links']:
                            base_url = st.session_state['modelos_drive_links'][modelo_selecionado]
                            file_info = st.session_state['modelos_drive_info'].get(modelo_selecionado, {})
                            is_gdoc = file_info.get('mimeType') == 'application/vnd.google-apps.document'
                            
                            # Se for Doc Google, o link já é o editor. Se for ODT, tentamos forçar a visualização que permite abrir com Doc Google.
                            edit_url = base_url if is_gdoc else base_url.replace("/view", "/edit")
                            
                            label_btn = "✏️ Abrir Documento Google" if is_gdoc else "✏️ Editar Modelo (Drive)"
                            
                            st.markdown(f"""
                                <a href='{edit_url}' target='_blank' style='text-decoration: none;'>
                                    <button style='background-color: white; border: 1px solid #2563eb; color: #2563eb; border-radius: 12px; padding: 10px 15px; cursor: pointer; font-size: 0.9rem; width: 100%; font-weight: 600;'>
                                        {label_btn}
                                    </button>
                                </a>
                            """, unsafe_allow_html=True)
                 
                  st.info(f"📄 Modelo selecionado: **{modelo_selecionado}**")
             else:
                  st.error("Nenhum modelo ODT encontrado. Volte ao Passo 1.")
                  st.session_state['modelo_selecionado_nome'] = None

        st.divider()

        col_btn_back, col_btn_next = st.columns(2)
        with col_btn_back:
            if st.button("← Voltar para Upload", key="back_to_upload_selecao_2", use_container_width=True):
                st.session_state['current_tab'] = "Upload"
                st.rerun()

        with col_btn_next:
             if st.session_state['dados_linha_selecionada'] is not None and st.session_state['modelo_selecionado_nome'] is not None:
                 if st.button("Avançar para Gerar Proposta →", type="primary", key="goto_geracao", use_container_width=True):
                     st.session_state['current_tab'] = "Geração"
                     st.rerun()
             else:
                  st.button("Avançar para Gerar Proposta →", type="primary", key="goto_geracao_disabled", use_container_width=True, disabled=True)


# --- Aba 3: Gerar Proposta ---
with tab_geracao:
    st.header("Passo 3: Revise e Gere a Proposta em PDF")
    st.markdown("---")

    if st.session_state.get('dados_linha_selecionada') is None or st.session_state.get('modelo_selecionado_nome') is None:
        st.warning("⚠️ Por favor, complete os Passos 1 e 2 primeiro (selecione uma linha válida e um modelo).")
        if st.button("← Voltar para Seleção", key="back_to_selecao_geracao"):
            st.session_state['current_tab'] = "Seleção"
            st.rerun()
    else:
        dados_linha = st.session_state['dados_linha_selecionada']
        nome_modelo_selecionado = st.session_state['modelo_selecionado_nome']
        
        # Lógica para buscar modelo - ou do Upload ou do Drive
        modelo_bytes = st.session_state.get('modelos_info', {}).get(nome_modelo_selecionado)
        
        # Se não tiver em memória (bytes), tenta baixar do Drive se for o caso
        if modelo_bytes is None and 'modelos_drive_info' in st.session_state:
            file_info = st.session_state['modelos_drive_info'].get(nome_modelo_selecionado)
            if file_info:
                file_id = file_info['id']
                mime_type = file_info.get('mimeType')
                with st.spinner(f"📥 Baixando '{nome_modelo_selecionado}' do Google Drive..."):
                    modelo_bytes = baixar_arquivo_drive(file_id, mime_type)
                    if modelo_bytes:
                        st.session_state['modelos_info'][nome_modelo_selecionado] = modelo_bytes

        if not modelo_bytes:
             st.error(f"❌ Erro: Modelo ODT '{nome_modelo_selecionado}' não encontrado. Verifique a conexão com o Drive.")
        else:
            with st.container(border=True):
                st.subheader("Revisão das Informações")
                substituicoes = criar_substituicoes(dados_linha)

                col_rev1, col_rev2 = st.columns(2)
                with col_rev1:
                     st.markdown("**Cliente e Contato:**")
                     rev_cliente = st.text_input("Cliente:", value=substituicoes.get("<Cliente>", ""), key=f"rev_cliente_{st.session_state.get('last_selected_line', 2)}")
                     rev_nome = st.text_input("Contato (Nome):", value=substituicoes.get("<Nome>", ""), key=f"rev_nome_{st.session_state.get('last_selected_line', 2)}")
                     
                     # Split logic for city/state if edited
                     local_valor = f"{substituicoes.get('<Cidade>', '')}/{substituicoes.get('<Estado>', '')}"
                     rev_local = st.text_input("Local (Cidade/Estado):", value=local_valor, key=f"rev_local_{st.session_state.get('last_selected_line', 2)}")

                with col_rev2:
                     st.markdown("**Produto e Valores:**")
                     rev_modelo_p = st.text_input("Modelo Proposta:", value=substituicoes.get("<Modelo>", ""), key=f"rev_modelo_prod_{st.session_state.get('last_selected_line', 2)}")
                     rev_val_romp = st.text_input("Valor Rompedor:", value=substituicoes.get("<Valor Rompedor>", ""), key=f"rev_val_romp_{st.session_state.get('last_selected_line', 2)}")
                     rev_val_kit = st.text_input("Valor Kit:", value=substituicoes.get("<Valor Kit>", ""), key=f"rev_val_kit_{st.session_state.get('last_selected_line', 2)}")

                # Campo de Observações (Área de texto para parágrafos maiores)
                st.markdown("**Observações / Descrição Adicional:**")
                rev_obs = st.text_area("Este texto será inserido no campo <Observações> do documento:", value=substituicoes.get("<Observações>", ""), height=100, key=f"rev_obs_{st.session_state.get('last_selected_line', 2)}")

                # Atualiza dicionário de substituições com os valores da revisão manual
                substituicoes["<Cliente>"] = rev_cliente
                substituicoes["<Nome>"] = rev_nome
                if "/" in rev_local:
                    cidade, estado = rev_local.split("/", 1)
                    substituicoes["<Cidade>"] = cidade.strip()
                    substituicoes["<Estado>"] = estado.strip()
                else:
                    substituicoes["<Cidade>"] = rev_local
                    substituicoes["<Estado>"] = ""
                
                substituicoes["<Modelo>"] = rev_modelo_p
                substituicoes["<Valor Rompedor>"] = rev_val_romp
                substituicoes["<Valor Kit>"] = rev_val_kit
                substituicoes["<Observações>"] = rev_obs

                with st.expander("Ver todas as substituições que serão feitas no documento"):
                     substituicoes_df = pd.DataFrame({
                         'Placeholder no Documento': list(substituicoes.keys()),
                         'Valor a ser Inserido': [str(v) for v in substituicoes.values()] 
                     })
                     st.dataframe(substituicoes_df, hide_index=True, use_container_width=True)

            st.divider()

            # Card de resumo antes do botão gerar
            st.markdown(f"""
            <div class="summary-card">
                <div class="summary-item">
                    <div class="s-label">👤 Cliente</div>
                    <div class="s-value">{substituicoes.get('<Cliente>', '—')}</div>
                </div>
                <div class="summary-item">
                    <div class="s-label">🔢 Nº Proposta</div>
                    <div class="s-value">{substituicoes.get('<Número>', '—')}</div>
                </div>
                <div class="summary-item">
                    <div class="s-label">🏗️ Modelo</div>
                    <div class="s-value">{substituicoes.get('<Modelo>', '—')}</div>
                </div>
                <div class="summary-item">
                    <div class="s-label">💰 Valor Rompedor</div>
                    <div class="s-value money">{substituicoes.get('<Valor Rompedor>', '—')}</div>
                </div>
                <div class="summary-item">
                    <div class="s-label">📦 Valor Kit</div>
                    <div class="s-value money">{substituicoes.get('<Valor Kit>', '—')}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            if st.button("🚀 Gerar Documento PDF Agora", type="primary", key="generate_pdf_final", use_container_width=True):
                 pdf_bytes_result = None  
                 pdf_filename_result = None 

                 with st.status("⚙️ Iniciando geração da proposta...", expanded=True) as status:
                    try:
                        status.update(label="1/4 - Extraindo conteúdo do modelo ODT...")
                        content_xml = extrair_conteudo_odt(modelo_bytes)
                        if not content_xml: raise ValueError("Falha ao extrair 'content.xml' do modelo ODT.")

                        status.update(label="2/4 - Aplicando substituições nos dados...")
                        content_xml_modificado, num_substituicoes = substituir_no_xml(content_xml, substituicoes)

                        status.update(label="3/4 - Recriando arquivo ODT modificado...")
                        documento_odt_modificado = criar_odt_modificado(modelo_bytes, content_xml_modificado)
                        if not documento_odt_modificado: raise ValueError("Falha ao recriar o arquivo ODT modificado.")

                        # DEFINIÇÃO DO NOME DO ARQUIVO: JE [Número] - [Cliente]
                        try:
                            # Pega os valores reais para compor o nome do arquivo
                            num_doc = substituicoes.get("<Número>", "S-N")
                            cliente_doc = substituicoes.get("<Cliente>", "Proposta")
                            # Limpa caracteres que o Windows/Linux não aceitam em nomes de arquivos
                            cliente_limpo = str(cliente_doc).replace('/', '-').replace('\\', '-').strip()
                            nome_base_desejado = f"JE {num_doc} - {cliente_limpo}"
                        except Exception:
                            # Fallback de segurança
                            nome_base_desejado = f"Proposta_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                        
                        nome_arquivo_pdf = f"{nome_base_desejado}.pdf"

                        status.update(label=f"4/4 - Convertendo para PDF ('{nome_arquivo_pdf}')... (pode levar alguns segundos)")
                        pdf_bytes = converter_para_pdf(documento_odt_modificado, nome_base_desejado)
                        if not pdf_bytes: raise ValueError("Falha ao converter o documento ODT para PDF usando LibreOffice.")

                        pdf_bytes_result = pdf_bytes
                        pdf_filename_result = nome_arquivo_pdf

                        status.update(label="🎉 Proposta gerada com sucesso!", state="complete", expanded=False)

                    except (ValueError, Exception) as e:
                        status.update(label=f"❌ Erro ao gerar proposta: {str(e)}", state="error", expanded=True)

                 if pdf_bytes_result and pdf_filename_result:
                      st.success(f"✅ Documento '{pdf_filename_result}' pronto!") 
                      st.download_button(
                           label=f"📥 Baixar {pdf_filename_result}",
                           data=pdf_bytes_result,
                           file_name=pdf_filename_result,
                           mime="application/pdf",
                           key="download_pdf_final_btn", 
                           use_container_width=True,
                           type="primary" 
                      )

            st.divider()

            col_btn_back_geracao, col_btn_new_geracao = st.columns(2)
            with col_btn_back_geracao:
                 if st.button("← Voltar para Seleção", key="back_to_selecao_geracao_2", use_container_width=True):
                      st.session_state['current_tab'] = "Seleção"
                      st.rerun()
            with col_btn_new_geracao:
                 if st.button("✨ Iniciar Nova Proposta (Voltar ao Início)", key="new_proposal_geracao", use_container_width=True):
                      st.session_state['current_tab'] = "Upload"
                      st.session_state['planilha_data'] = None
                      st.session_state['planilha_nome'] = None
                      st.session_state['dados_linha_selecionada'] = None
                      st.session_state['modelo_selecionado_nome'] = None
                      if 'last_selected_line' in st.session_state: del st.session_state['last_selected_line'] 
                      st.rerun()

st.markdown(f"""
<div class="footer">
    <img src="{logo_base64}" alt="Jardim" style="height:36px; margin-bottom:0.75rem; opacity:0.7;">
    <p><strong>Jardim Equipamentos</strong> — Gerador de Propostas Comerciais</p>
    <p style="margin-top:0.35rem;">© 2026 · Todos os direitos reservados · Desenvolvido por <strong>Rodrigo Ferreira</strong></p>
</div>
""", unsafe_allow_html=True)

tab_map = {"Upload": 0, "Seleção": 1, "Geração": 2}
current_tab_index = tab_map.get(st.session_state['current_tab'], 0) 

if st.session_state['current_tab'] != "Upload": 
    js = f"""
    <script>
        function selectTab() {{
            const tabIndex = {current_tab_index};
            const tabs = parent.document.querySelectorAll('button[data-baseweb="tab"]');
            if (tabs && tabs.length > tabIndex) {{
                tabs[tabIndex].click();
            }} else {{
                 console.warn('Streamlit Tabs:', 'Tab index', tabIndex, 'not found or tabs not rendered yet. Available tabs:', tabs ? tabs.length : 0);
             }}
        }}
         if (window.parent) {{ 
            setTimeout(selectTab, 150);
         }}
    </script>
    """
    st.components.v1.html(js, height=0, width=0)
