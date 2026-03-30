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

# --- INTEGRAÇÃO GOOGLE CLOUD ---

def get_google_drive_service():
    """Autentica e retorna o serviço do Google Drive"""
    try:
        if "google_cloud" in st.secrets:
            creds_dict = dict(st.secrets["google_cloud"])
            creds = service_account.Credentials.from_service_account_info(creds_dict)
            return build('drive', 'v3', credentials=creds)
        else:
            st.error("⚠️ Seção [google_cloud] não encontrada no secrets.toml.")
    except Exception as e:
        st.error(f"❌ Erro de Autenticação Google: {e}")
        return None
    return None

def listar_modelos_google_drive(folder_id):
    """Lista arquivos ODT em uma pasta específica do Drive"""
    service = get_google_drive_service()
    if not service:
        return []
    
    # Busca por ODT com múltiplos MIME types possíveis (arquivos ODT podem ter tipos diferentes no Drive)
    query = (
        f"'{folder_id}' in parents and trashed=false and ("
        f"mimeType='application/vnd.oasis.opendocument.text' or "
        f"mimeType='application/octet-stream' or "
        f"name contains '.odt'"
        f")"
    )
    results = service.files().list(q=query, fields="files(id, name)").execute()
    # Filtra apenas arquivos .odt pelo nome
    todos = results.get('files', [])
    return [f for f in todos if f['name'].lower().endswith('.odt')]

def baixar_arquivo_drive(file_id):
    """Baixa um arquivo do Drive em memória"""
    service = get_google_drive_service()
    if not service:
        return None
    
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    return fh.getvalue()

# --- FUNÇÕES AUXILIARES (Mantidas exatamente como no original) ---

def extrair_conteudo_odt(arquivo_bytes):
    """Extrai o conteúdo de um arquivo ODT"""
    with tempfile.NamedTemporaryFile(suffix='.odt', delete=False) as temp_file:
        temp_file.write(arquivo_bytes)
        temp_path = temp_file.name

    try:
        with zipfile.ZipFile(temp_path, 'r') as zip_ref:
            content_xml = zip_ref.read('content.xml').decode('utf-8')
        os.unlink(temp_path)
        return content_xml
    except Exception as e:
        st.error(f"Erro ao extrair conteúdo do arquivo ODT: {str(e)}")
        if os.path.exists(temp_path):
             os.unlink(temp_path)
        return None
    finally:
        # Garante que o arquivo temporário seja removido mesmo se ocorrer um erro inesperado antes do unlink
        if 'temp_path' in locals() and os.path.exists(temp_path):
            try:
                os.unlink(temp_path)
            except OSError:
                pass # Ignora erros se o arquivo já foi removido

def substituir_no_xml(content_xml, substituicoes):
    """Substitui texto no conteúdo XML do arquivo ODT"""
    texto_modificado = content_xml
    substituicoes_feitas = 0

    # Mapeamento dos nomes das colunas para os placeholders
    mapeamento_colunas = {
        "Cliente": "<Cliente>", "Cidade": "<Cidade>", "Estado": "<Estado>",
        "Número": "<Número>", "Nome": "<Nome>", "Telefone": "<Telefone>",
        "Email": "<Email>", "Modelo": "<Modelo>", "TIPO DE MÁQUINA": "<TIPO DE MÁQUINA>",
        "MODELO DE MÁQUINA": "<MODELO DE MÁQUINA>", "Valor Rompedor": "<Valor Rompedor>",
        "Valor Kit": "<Valor Kit>", "Condição de pagamento": "<Condição de pagamento>",
        "FRETE": "<FRETE>", "Data": "<Data>"
    }

    # Primeiro, substituir os placeholders no formato de database-display
    for coluna, placeholder in mapeamento_colunas.items():
        if placeholder in substituicoes:
            padrao = f'<text:database-display[^>]*text:column-name="{re.escape(coluna)}"[^>]*>([^<]*)</text:database-display>'
            # Usamos uma função lambda para preservar a estrutura original da tag, apenas mudando o conteúdo
            texto_modificado, num_subs = re.subn(
                padrao,
                lambda m: f'<text:database-display text:column-name="{coluna}" text:table-name="Planilha1" text:table-type="table" text:database-name="Formulário propostas Rompedor1">{substituicoes[placeholder]}</text:database-display>',
                texto_modificado
            )
            substituicoes_feitas += num_subs

    # Depois, substituir os placeholders como texto simples (se existirem)
    for placeholder, valor in substituicoes.items():
        padrao_simples = re.escape(placeholder)
        texto_modificado, num_subs_simples = re.subn(padrao_simples, str(valor), texto_modificado)
        substituicoes_feitas += num_subs_simples

    return texto_modificado, substituicoes_feitas


def criar_odt_modificado(arquivo_original_bytes, content_xml_modificado):
    """Cria um novo arquivo ODT com o conteúdo modificado"""
    temp_original_path = None
    temp_modificado_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix='.odt', delete=False) as temp_original:
            temp_original.write(arquivo_original_bytes)
            temp_original_path = temp_original.name

        with tempfile.NamedTemporaryFile(suffix='.odt', delete=False) as temp_modificado:
            temp_modificado_path = temp_modificado.name

        with zipfile.ZipFile(temp_original_path, 'r') as zip_original:
            with zipfile.ZipFile(temp_modificado_path, 'w', zipfile.ZIP_DEFLATED) as zip_modificado: # Usar compressão
                for item in zip_original.infolist():
                    if item.filename == 'content.xml':
                        zip_modificado.writestr('content.xml', content_xml_modificado.encode('utf-8')) # Garantir encoding utf-8
                    else:
                        zip_modificado.writestr(item, zip_original.read(item.filename))

        with open(temp_modificado_path, 'rb') as f:
            conteudo_modificado = f.read()

        return conteudo_modificado

    except Exception as e:
        st.error(f"Erro ao criar arquivo ODT modificado: {str(e)}")
        return None
    finally:
        # Limpeza robusta dos arquivos temporários
        if temp_original_path and os.path.exists(temp_original_path):
            os.unlink(temp_original_path)
        if temp_modificado_path and os.path.exists(temp_modificado_path):
            os.unlink(temp_modificado_path)

def converter_para_pdf(odt_bytes, nome_arquivo_base):
    """Converte ODT para PDF usando LibreOffice"""
    libreoffice_path = None
    paths_to_try = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/libreoffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice"
    ]

    for path in paths_to_try:
        if os.path.exists(path):
            libreoffice_path = path
            break

    if not libreoffice_path:
        st.error("⚠️ **LibreOffice não encontrado.** Verifique a instalação ou o caminho no código.")
        return None

    temp_odt_path = None
    temp_pdf_dir = None
    pdf_path = None # Inicializa pdf_path

    try:
        with tempfile.NamedTemporaryFile(suffix='.odt', delete=False) as temp_odt:
            temp_odt.write(odt_bytes)
            temp_odt_path = temp_odt.name

        temp_pdf_dir = tempfile.mkdtemp()

        comando = [
            libreoffice_path,
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', temp_pdf_dir,
            temp_odt_path
        ]

        # Usar Popen para melhor controle, especialmente no Windows
        process = subprocess.Popen(comando, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=(os.name == 'nt'))
        stdout, stderr = process.communicate(timeout=120) # Timeout aumentado

        if process.returncode != 0:
            error_message = stderr.decode('utf-8', errors='ignore')
            # Tentar extrair mensagem mais útil do erro do LibreOffice
            if "Error: source file could not be loaded" in error_message:
                 raise Exception("Erro do LibreOffice: O arquivo ODT de origem não pôde ser carregado (pode estar corrompido ou ter permissões incorretas).")
            elif "error while loading shared libraries" in error_message:
                 raise Exception(f"Erro do LibreOffice: Falta de bibliotecas compartilhadas. Detalhes: {error_message}")
            else:
                 raise Exception(f"Erro na conversão (código {process.returncode}): {error_message}")


        # O nome do arquivo PDF gerado pelo LibreOffice será o mesmo do ODT, mas com extensão .pdf
        pdf_filename = os.path.basename(temp_odt_path).replace('.odt', '.pdf')
        pdf_path = os.path.join(temp_pdf_dir, pdf_filename)


        if not os.path.exists(pdf_path):
             # Adicionar verificação do stdout para pistas
             output_message = stdout.decode('utf-8', errors='ignore')
             raise Exception(f"Arquivo PDF não foi gerado em '{temp_pdf_dir}'. Output: {output_message}")


        with open(pdf_path, 'rb') as f:
            pdf_bytes = f.read()

        return pdf_bytes

    except subprocess.TimeoutExpired:
        st.error("⏳ A conversão para PDF demorou muito (timeout). Tente novamente ou verifique o arquivo ODT.")
        return None
    except Exception as e:
        st.error(f"Falha na conversão para PDF: {str(e)}")
        # Adicionar log extra para depuração
        st.error(f"Comando executado: {' '.join(comando)}")
        if 'stderr' in locals() and stderr: st.error(f"Saída de erro do processo: {stderr.decode('utf-8', errors='ignore')}")
        return None
    finally:
        # Limpeza final
        if temp_odt_path and os.path.exists(temp_odt_path):
            os.unlink(temp_odt_path)
        if pdf_path and os.path.exists(pdf_path):
             os.unlink(pdf_path)
        if temp_pdf_dir and os.path.exists(temp_pdf_dir):
             try:
                  os.rmdir(temp_pdf_dir)
             except OSError:
                  # Pode falhar se o LibreOffice ainda tiver algum lock, mas tentamos
                  st.warning(f"Não foi possível remover o diretório temporário {temp_pdf_dir}. Pode ser necessário remover manualmente.")
                  pass


def formatar_valor_monetario(valor):
    """Formata um valor como moeda brasileira (R$)"""
    try:
        # Tenta converter para float, tratando vírgula como separador decimal se necessário
        if isinstance(valor, str):
            valor = valor.replace('.', '').replace(',', '.')
        valor_float = float(valor)
        # Formatação padrão brasileira
        return f"R$ {valor_float:,.2f}".replace(',', 'v').replace('.', ',').replace('v', '.')
    except (ValueError, TypeError):
        return "R$ 0,00" # Retorna R$ 0,00 se a conversão falhar


def criar_substituicoes(dados):
    """Prepara dicionário de substituições a partir de uma linha (dict) do DataFrame"""
    substituicoes = {}
    data_hoje = datetime.today().strftime("%d/%m/%Y")

    # Mapeamento dos placeholders para as colunas (considerando nomes exatos)
    mapeamento_placeholders = {
        "<Cliente>": "Cliente", "<Cidade>": "Cidade", "<Estado>": "Estado",
        "<Número>": "Número", "<Nome>": "Nome", "<Telefone>": "Telefone",
        "<Email>": "Email", "<Modelo>": "Modelo", "<TIPO DE MÁQUINA>": "TIPO DE MÁQUINA",
        "<MODELO DE MÁQUINA>": "MODELO DE MÁQUINA", "<Valor Rompedor>": "Valor Rompedor",
        "<Valor Kit>": "Valor Kit", "<Condição de pagamento>": "Condição de pagamento",
        "<FRETE>": "FRETE", "<Data>": "Data"
    }

    for placeholder, coluna in mapeamento_placeholders.items():
        valor = dados.get(coluna, "") # Pega o valor da coluna correspondente

        # Tratamento especial para valores monetários
        if coluna in ["Valor Rompedor", "Valor Kit"]:
            valor_formatado = formatar_valor_monetario(valor)
            substituicoes[placeholder] = valor_formatado
        # Tratamento especial para Data
        elif coluna == "Data":
             if pd.isna(valor) or valor == "":
                  substituicoes[placeholder] = data_hoje
             elif isinstance(valor, datetime):
                  substituicoes[placeholder] = valor.strftime("%d/%m/%Y")
             else:
                  # Tenta converter string para data, se falhar usa o valor como está ou data de hoje
                  try:
                       data_obj = pd.to_datetime(valor, errors='coerce')
                       if pd.isna(data_obj):
                            substituicoes[placeholder] = str(valor) if valor else data_hoje
                       else:
                            substituicoes[placeholder] = data_obj.strftime("%d/%m/%Y")
                  except Exception:
                       substituicoes[placeholder] = str(valor) if valor else data_hoje
        # Para outros campos, apenas converte para string
        else:
            substituicoes[placeholder] = str(valor)

    return substituicoes

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

        /* Cabeçalho com logo */
        .header-container {
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            margin-bottom: 3rem;
            margin-top: 1rem;
        }
        .logo-img {
             height: 100px;
             margin-bottom: 1.5rem;
             filter: drop-shadow(0px 4px 10px rgba(0,0,0,0.08));
        }
        .header-title {
            text-align: center;
            background: linear-gradient(135deg, #2563eb 0%, #1e40af 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            font-size: 2.8rem !important;
            font-weight: 800 !important;
            margin: 0 !important;
        }

        /* Soft Glassmorphism nos Containers (Tema Claro) */
        [data-testid="stVerticalBlock"] > [style*="border: 1px solid rgba(49, 51, 63, 0.2)"],
        [data-testid="stVerticalBlock"] > [style*="border: 1px solid rgba(250, 250, 250, 0.2)"] {
            background: rgba(255, 255, 255, 0.7) !important;
            backdrop-filter: blur(12px) !important;
            -webkit-backdrop-filter: blur(12px) !important;
            border: 1px solid rgba(255, 255, 255, 0.8) !important;
            border-radius: 20px !important;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06) !important;
            padding: 2.5rem !important;
            margin-bottom: 2rem !important;
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

        /* Rodapé Fixo */
        .footer {
            margin-top: 4rem;
            padding: 2rem 0;
            border-top: 1px solid #e2e8f0;
            text-align: center;
            color: #94a3b8;
            font-size: 0.9rem;
        }
    </style>
    """, unsafe_allow_html=True)

load_css()

# --- Cabeçalho com Logo ---
def render_header():
    st.markdown(f"""
    <div class="header-container">
        <img class="logo-img" src="{logo_base64}" alt="Logo Jardim Equipamentos">
        <h1 class="header-title">Gerador de Propostas</h1>
    </div>
    """, unsafe_allow_html=True)

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
            st.subheader("📊 Banco de Dados (Planilha)")
            origem_dados = st.radio("Origem da Planilha:", ["Upload Manual", "Google Sheets (Nuvem)"], horizontal=True, key="origem_dados")
            
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
                    url_sheets = st.text_input("Cole o link da sua Planilha Google:", placeholder="https://docs.google.com/spreadsheets/d/...")
                    
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
                        st.success("✅ Dados carregados da nuvem!")
                except Exception as e:
                    if "not found" in str(e).lower():
                        st.error("❌ Planilha não encontrada. Verifique se o link está correto e se você compartilhou com o e-mail da Service Account.")
                    elif "credential" in str(e).lower():
                        st.error("⚠️ Credenciais do Google não configuradas corretamente no secrets.toml.")
                    else:
                        st.error(f"❌ Erro de Conexão: {e}")

    with col_meta2:
        with st.container(border=True):
            st.subheader("📄 Modelos de Proposta")
            origem_modelos = st.radio("Origem dos Modelos:", ["Upload Manual", "Google Drive"], horizontal=True, key="origem_modelos")
            
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
                folder_id = st.text_input("ID da Pasta no Drive:", placeholder="1A2B3... (final da URL da pasta)")
                if folder_id:
                    modelos_drive = listar_modelos_google_drive(folder_id)
                    if modelos_drive:
                        st.success(f"✅ {len(modelos_drive)} modelos encontrados no Drive!")
                        # Armazenamos apenas os IDs, baixaremos sob demanda no Passo 3
                        st.session_state['modelos_drive_info'] = {m['name']: m['id'] for m in modelos_drive}
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

        col_linha, col_modelo = st.columns(2)

        with col_linha:
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
                          preview_data = {k: v for k, v in st.session_state['dados_linha_selecionada'].items() if k in ['Cliente', 'Modelo', 'Valor Rompedor', 'Valor Kit', 'Data']}
                          st.dataframe(pd.Series(preview_data).astype(str), use_container_width=True)

                else:
                     st.error(f"❌ Linha {linha_selecionada_usuario} inválida. Selecione um valor entre 2 e {len(df) + 1}.")
                     st.session_state['dados_linha_selecionada'] = None 

            # --- GERENCIAMENTO DE DADOS GOOGLE SHEETS ---
            if origem_dados == "Google Sheets (Nuvem)" and st.session_state['planilha_data'] is not None:
                st.markdown("---")
                with st.expander("📝 Gerenciar Banco de Dados (Nuvem)", expanded=False):
                    tab_edit, tab_add = st.tabs(["✏️ Editar Selecionada", "➕ Adicionar Nova"])
                    
                    # Todas as colunas da planilha carregada
                    todas_colunas = list(st.session_state['planilha_data'].columns)
                    
                    with tab_edit:
                        if st.session_state['dados_linha_selecionada']:
                            linha_atual = st.session_state['last_selected_line']
                            st.write(f"Editando dados da linha **{linha_atual}** — {len(todas_colunas)} campos:")
                            with st.form("form_edit_sheets"):
                                novos_dados = {}
                                # Renderiza em pares de 2 colunas
                                # A linha_atual entra no key para forçar recriação ao mudar de linha
                                for i in range(0, len(todas_colunas), 2):
                                    cols = st.columns(2)
                                    for j, campo in enumerate(todas_colunas[i:i+2]):
                                        valor_atual = str(st.session_state['dados_linha_selecionada'].get(campo, ""))
                                        if valor_atual in ("nan", "None", "NaT"):
                                            valor_atual = ""
                                        with cols[j]:
                                            novos_dados[campo] = st.text_input(
                                                campo,
                                                value=valor_atual,
                                                key=f"edit_{campo}_{linha_atual}"  # key muda com a linha!
                                            )
                                
                                if st.form_submit_button("💾 Salvar Alterações na Nuvem", use_container_width=True):
                                    try:
                                        conn = st.connection("gsheets", type=GSheetsConnection)
                                        df_atualizado = st.session_state['planilha_data'].copy()
                                        idx = st.session_state['last_selected_line'] - 2
                                        for key, val in novos_dados.items():
                                            df_atualizado.at[idx, key] = val
                                        conn.update(spreadsheet=url_sheets, data=df_atualizado)
                                        st.session_state['planilha_data'] = df_atualizado
                                        st.success("✅ Planilha atualizada no Google Sheets!")
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"Erro ao salvar: {e}")
                        else:
                            st.info("Selecione uma linha válida para editar.")

                    with tab_add:
                        st.write(f"Insira os dados para o novo cliente — {len(todas_colunas)} campos:")
                        with st.form("form_add_sheets"):
                            dados_novos_row = {}
                            for i in range(0, len(todas_colunas), 2):
                                cols = st.columns(2)
                                for j, campo in enumerate(todas_colunas[i:i+2]):
                                    with cols[j]:
                                        # Preenche Data com hoje por padrão
                                        valor_padrao = datetime.today().strftime("%d/%m/%Y") if campo == "Data" else ""
                                        dados_novos_row[campo] = st.text_input(campo, value=valor_padrao, key=f"add_{campo}")
                            
                            if st.form_submit_button("➕ Adicionar à Planilha", use_container_width=True):
                                try:
                                    conn = st.connection("gsheets", type=GSheetsConnection)
                                    df_atualizado = st.session_state['planilha_data'].copy()
                                    nova_linha = {col: "" for col in df_atualizado.columns}
                                    nova_linha.update(dados_novos_row)
                                    df_atualizado = pd.concat([df_atualizado, pd.DataFrame([nova_linha])], ignore_index=True)
                                    conn.update(spreadsheet=url_sheets, data=df_atualizado)
                                    st.session_state['planilha_data'] = df_atualizado
                                    st.success("✅ Novo cliente adicionado com sucesso!")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Erro ao adicionar: {e}")


        with col_modelo:
             with st.container(border=True):
                 st.subheader("Selecione o Modelo ODT")
                 st.caption("Escolha qual modelo ODT será usado para esta proposta.")
                 nomes_modelos = list(st.session_state['modelos_info'].keys())

                 if nomes_modelos:
                      modelo_selecionado = st.selectbox(
                           "Modelos Disponíveis:",
                           options=nomes_modelos,
                           index=nomes_modelos.index(st.session_state.get('modelo_selecionado_nome', nomes_modelos[0])) if st.session_state.get('modelo_selecionado_nome') in nomes_modelos else 0,
                           key="modelo_select_widget"
                      )
                      st.session_state['modelo_selecionado_nome'] = modelo_selecionado
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
            file_id = st.session_state['modelos_drive_info'].get(nome_modelo_selecionado)
            if file_id:
                with st.spinner(f"📥 Baixando '{nome_modelo_selecionado}' do Google Drive..."):
                    modelo_bytes = baixar_arquivo_drive(file_id)
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
                     st.text_input("Cliente:", value=substituicoes.get("<Cliente>", ""), disabled=True, key=f"rev_cliente_{st.session_state.get('last_selected_line', 2)}")
                     st.text_input("Contato (Nome):", value=substituicoes.get("<Nome>", ""), disabled=True, key=f"rev_nome_{st.session_state.get('last_selected_line', 2)}")
                     st.text_input("Local:", value=f"{substituicoes.get('<Cidade>', '')}/{substituicoes.get('<Estado>', '')}", disabled=True, key=f"rev_local_{st.session_state.get('last_selected_line', 2)}")

                with col_rev2:
                     st.markdown("**Produto e Valores:**")
                     st.text_input("Modelo Proposta:", value=substituicoes.get("<Modelo>", ""), disabled=True, key=f"rev_modelo_prod_{st.session_state.get('last_selected_line', 2)}")
                     st.text_input("Valor Rompedor:", value=substituicoes.get("<Valor Rompedor>", ""), disabled=True, key=f"rev_val_romp_{st.session_state.get('last_selected_line', 2)}")
                     st.text_input("Valor Kit:", value=substituicoes.get("<Valor Kit>", ""), disabled=True, key=f"rev_val_kit_{st.session_state.get('last_selected_line', 2)}")

                with st.expander("Ver todas as substituições que serão feitas no documento"):
                     substituicoes_df = pd.DataFrame({
                         'Placeholder no Documento': list(substituicoes.keys()),
                         'Valor a ser Inserido': [str(v) for v in substituicoes.values()] 
                     })
                     st.dataframe(substituicoes_df, hide_index=True, use_container_width=True)

            st.divider()

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

                        # AQUI ESTÁ A IMPLEMENTAÇÃO DO NOME DA ÚLTIMA COLUNA
                        try:
                            # Pega todos os nomes das colunas da planilha
                            colunas = list(st.session_state['planilha_data'].columns)
                            ultima_coluna = colunas[-1]
                            # Pega o valor da ultima coluna
                            nome_base_desejado = dados_linha.get(ultima_coluna, "")
                            if not nome_base_desejado or pd.isna(nome_base_desejado):
                                nome_cliente = str(dados_linha.get('Cliente', 'Proposta')).replace(' ', '_').replace('/','-')
                                nome_base_desejado = f"Proposta_{nome_cliente}_{datetime.now().strftime('%Y%m%d')}"
                        except Exception:
                            # Fallback original
                            nome_base_desejado = dados_linha.get("NOME DO ARQUIVO", "Proposta_Gerada")
                        
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

st.markdown("---") 
st.markdown("""
<div class="footer">
    <p>Jardim Equipamentos - Gerador de Propostas Comerciais</p>
    <p>© 2025 - Todos os direitos reservados - Desenvolvido por Rodrigo Ferreira</p>
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
