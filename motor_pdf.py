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
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# --- INTEGRAÇÃO GOOGLE CLOUD ---

def get_google_drive_service():
    """Autentica via conta de serviço (apenas para leitura de templates)"""
    try:
        if "google_cloud" in st.secrets:
            creds_dict = dict(st.secrets["google_cloud"])
            scopes = ['https://www.googleapis.com/auth/drive']
            creds = service_account.Credentials.from_service_account_info(
                creds_dict, scopes=scopes
            )
            return build('drive', 'v3', credentials=creds)
        else:
            st.error("⚠️ Seção [google_cloud] não encontrada no secrets.toml.")
    except Exception as e:
        st.error(f"❌ Erro de Autenticação Google: {e}")
        return None
    return None

def get_user_drive_service():
    """Autentica via OAuth2 do usuário (usa quota do usuário - 5TB)"""
    try:
        if "oauth2_drive" not in st.secrets:
            return None
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request as GRequest
        o = st.secrets["oauth2_drive"]
        creds = Credentials(
            token=None,
            refresh_token=o["refresh_token"],
            token_uri="https://oauth2.googleapis.com/token",
            client_id=o["client_id"],
            client_secret=o["client_secret"],
            scopes=["https://www.googleapis.com/auth/drive"]
        )
        creds.refresh(GRequest())
        return build('drive', 'v3', credentials=creds)
    except Exception as e:
        st.warning(f"⚠️ OAuth2 falhou: {e}")
        return None


def listar_modelos_google_drive(folder_id):
    """Lista arquivos ODT e Documentos Google em uma pasta específica do Drive"""
    service = get_google_drive_service()
    if not service:
        return []
    
    # Busca por ODT e Documentos Google
    query = (
        f"'{folder_id}' in parents and trashed=false and ("
        f"mimeType='application/vnd.oasis.opendocument.text' or "
        f"mimeType='application/vnd.google-apps.document' or "
        f"mimeType='application/octet-stream' or "
        f"name contains '.odt'"
        f")"
    )
    # Incluímos mimeType nos campos retornados
    results = service.files().list(q=query, fields="files(id, name, webViewLink, mimeType)").execute()
    
    todos = results.get('files', [])
    # Filtra apenas arquivos .odt ou Documentos Google nativos
    return [f for f in todos if f['name'].lower().endswith('.odt') or f['mimeType'] == 'application/vnd.google-apps.document']

def baixar_arquivo_drive(file_id, mime_type=None):
    """Baixa um arquivo do Drive em memória (ou exporta se for Doc Google)新闻网"""
    service = get_google_drive_service()
    if not service:
        return None
    
    try:
        if mime_type == 'application/vnd.google-apps.document':
            # Se for Doc Google, exportamos para ODT e retornamos os bytes diretamente
            return service.files().export(fileId=file_id, mimeType='application/vnd.oasis.opendocument.text').execute()
        else:
            # Se for binário (ODT), baixamos usando MediaIoBaseDownload
            request = service.files().get_media(fileId=file_id)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
            return fh.getvalue()
    except Exception as e:
        st.error(f"❌ Erro ao baixar/exportar arquivo do Drive: {e}")
        return None

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
        if 'temp_path' in locals() and os.path.exists(temp_path):
             os.unlink(temp_path)
        return None

def obter_proximo_numero(df):
    """Calcula o próximo número de proposta (ex: 852-26 -> 853-26)"""
    if 'Número' not in df.columns or df.empty:
        return ""
    
    # Pega o último valor não nulo da coluna Número
    ultimos_nums = df['Número'].dropna().astype(str).tolist()
    if not ultimos_nums:
        return ""
    
    ultimo_val = ultimos_nums[-1].strip()
    
    # Tenta quebrar pelo hífen (formato 852-26)
    if '-' in ultimo_val:
        partes = ultimo_val.split('-')
        try:
            # Pega apenas os dígitos da primeira parte (ex: 852)
            num_base_str = re.sub(r'\D', '', partes[0])
            if num_base_str:
                num_base = int(num_base_str)
                sufixo = partes[1]
                return f"{num_base + 1}-{sufixo}"
        except:
            pass
            
    return ultimo_val # Fallback se não conseguir processar

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
    import html as _html_mod

    for coluna, placeholder in mapeamento_colunas.items():
        if placeholder in substituicoes:
            valor_xml = _html_mod.escape(str(substituicoes[placeholder]))
            padrao = f'<text:database-display[^>]*text:column-name="{re.escape(coluna)}"[^>]*>([^<]*)</text:database-display>'
            texto_modificado, num_subs = re.subn(
                padrao,
                lambda m, v=valor_xml, c=coluna: f'<text:database-display text:column-name="{c}" text:table-name="Planilha1" text:table-type="table" text:database-name="Formulário propostas Rompedor1">{v}</text:database-display>',
                texto_modificado
            )
            substituicoes_feitas += num_subs

    # Depois, substituir os placeholders como texto simples (se existirem)
    for placeholder, valor in substituicoes.items():
        padrao_simples = re.escape(placeholder)
        valor_xml = _html_mod.escape(str(valor))
        texto_modificado, num_subs_simples = re.subn(padrao_simples, valor_xml, texto_modificado)
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

def _limpar_drive_conta_servico(service):
    """
    Deleta permanentemente TODOS os arquivos de propriedade da conta de serviço.
    A conta de serviço não deve conter arquivos legítimos — é apenas uma conta técnica.
    Isso libera a quota de armazenamento (15GB) antes de fazer o upload de conversão.
    """
    try:
        page_token = None
        total_deletados = 0
        while True:
            params = {
                "q": "'me' in owners and trashed=false",
                "fields": "nextPageToken, files(id, name)",
                "pageSize": 100
            }
            if page_token:
                params["pageToken"] = page_token

            results = service.files().list(**params).execute()
            arquivos = results.get('files', [])

            for arq in arquivos:
                try:
                    service.files().delete(fileId=arq['id']).execute()
                    total_deletados += 1
                except Exception:
                    pass

            page_token = results.get('nextPageToken')
            if not page_token:
                break

        # Esvazia a lixeira para liberar quota imediatamente
        try:
            service.files().emptyTrash().execute()
        except Exception:
            pass

    except Exception:
        pass


def converter_para_pdf_drive(odt_bytes, nome_arquivo_base):
    """
    Converte ODT para PDF usando a API do Google Drive.
    """
    from googleapiclient.http import MediaIoBaseUpload
    import uuid as _uuid

    if not odt_bytes or len(odt_bytes) == 0:
        st.error("⚠️ ODT vazio (0 bytes).")
        return None

    service = get_user_drive_service()
    if not service:
        service = get_google_drive_service()
    if not service:
        st.error("❌ [Drive] Falha na autenticação.")
        return None

    temp_file_id = None
    try:
        nome_temp = f"_temp_{_uuid.uuid4().hex}"
        file_metadata = {
            'name': nome_temp,
            'mimeType': 'application/vnd.google-apps.document'
        }
        media = MediaIoBaseUpload(
            io.BytesIO(odt_bytes),
            mimetype='application/vnd.oasis.opendocument.text',
            resumable=False
        )
        uploaded = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()
        temp_file_id = uploaded.get('id')

        pdf_bytes = service.files().export(
            fileId=temp_file_id,
            mimeType='application/pdf'
        ).execute()
        return pdf_bytes

    except Exception as e:
        st.error(f"❌ [Drive] ERRO: {str(e)}")
        return None
    finally:
        if temp_file_id:
            try:
                service.files().delete(fileId=temp_file_id).execute()
            except Exception:
                pass


def converter_para_pdf_python(odt_bytes):
    """
    Converte ODT → HTML customizado → PDF (via weasyprint).
    Usa parser XML direto para capturar text:p e draw:text-box.
    """
    import weasyprint
    from xml.etree import ElementTree as ET
    import html as html_lib

    NS_TEXT   = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'
    NS_DRAW   = 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0'
    NS_TABLE  = 'urn:oasis:names:tc:opendocument:xmlns:table:1.0'
    NS_OFFICE = 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'

    with zipfile.ZipFile(io.BytesIO(odt_bytes)) as z:
        content_xml = z.read('content.xml')

    try:
        root = ET.fromstring(content_xml)
    except ET.ParseError:
        # Remove caracteres de controle inválidos em XML (exceto tab, LF, CR)
        import re as _re
        content_str = content_xml.decode('utf-8', errors='replace')
        content_str = _re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', content_str)
        root = ET.fromstring(content_str.encode('utf-8'))
    parts = []

    def get_text(el):
        texts = []
        if el.text:
            texts.append(el.text)
        for child in el:
            texts.append(get_text(child))
            if child.tail:
                texts.append(child.tail)
        return ''.join(texts)

    def process(el):
        tag = el.tag.split('}')[1] if '}' in el.tag else el.tag
        ns  = el.tag.split('}')[0][1:] if '}' in el.tag else ''

        if ns == NS_TEXT and tag == 'p':
            text = get_text(el).strip()
            parts.append(f'<p>{html_lib.escape(text) if text else "&nbsp;"}</p>\n')
        elif ns == NS_TEXT and tag == 'h':
            lvl  = el.get(f'{{{NS_TEXT}}}outline-level', '2')
            text = get_text(el).strip()
            if text:
                parts.append(f'<h{lvl}>{html_lib.escape(text)}</h{lvl}>\n')
        elif ns == NS_DRAW and tag in ('text-box', 'frame'):
            for child in el:
                process(child)
        elif ns == NS_TABLE and tag == 'table':
            parts.append('<table>\n')
            for row in el.iter(f'{{{NS_TABLE}}}table-row'):
                parts.append('<tr>')
                for cell in row:
                    cell_tag = cell.tag.split('}')[1] if '}' in cell.tag else ''
                    if 'table-cell' in cell_tag:
                        parts.append('<td>')
                        for p in cell:
                            process(p)
                        parts.append('</td>')
                parts.append('</tr>\n')
            parts.append('</table>\n')
        else:
            for child in el:
                process(child)

    # O corpo do documento ODT é office:text (não text:text)
    body = root.find(f'.//{{{NS_OFFICE}}}text')
    if body is not None:
        for child in body:
            process(child)

    css = """<style>
    @page { size: A4; margin: 2cm; }
    body { font-family: Arial, sans-serif; font-size: 11pt; line-height: 1.6; color: #000; }
    p { margin: 0.3em 0; }
    h1,h2,h3 { color: #111; margin-top: 1em; }
    table { width: 100%; border-collapse: collapse; margin: 0.8em 0; }
    td { padding: 5px 10px; border: 1px solid #ccc; vertical-align: top; }
    </style>"""

    html = f'<!DOCTYPE html><html><head><meta charset="utf-8">{css}</head><body>{"".join(parts)}</body></html>'
    pdf_bytes = weasyprint.HTML(string=html).write_pdf()
    return pdf_bytes


def converter_para_pdf(odt_bytes, nome_arquivo_base):
    """
    Converte ODT para PDF com três estratégias em cascata:
    1. Google Drive API — qualidade idêntica ao template (conta nova com quota zerada)
    2. Python nativo (weasyprint) — sem layout mas funciona
    3. LibreOffice local — fallback desenvolvimento local
    """
    # --- Tentativa 1: Google Drive API (qualidade perfeita) ---
    if "google_cloud" in st.secrets or "oauth2_drive" in st.secrets:
        resultado = converter_para_pdf_drive(odt_bytes, nome_arquivo_base)
        if resultado:
            return resultado
        st.warning("⚠️ Falha na conversão via Google Drive. Tentando conversão Python...")

    # --- Tentativa 2: Python nativo (sem layout perfeito, mas funciona) ---
    try:
        resultado = converter_para_pdf_python(odt_bytes)
        if resultado:
            return resultado
    except Exception as e:
        st.warning(f"⚠️ Conversão Python nativa falhou ({e}). Tentando LibreOffice local...")

    # --- Tentativa 3: LibreOffice local (dev local) ---
    libreoffice_path = None
    for path in [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/libreoffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice"
    ]:
        if os.path.exists(path):
            libreoffice_path = path
            break

    if not libreoffice_path:
        st.error("⚠️ Nenhum método de conversão disponível. Instale o LibreOffice ou configure o Google Drive.")
        return None

    temp_odt_path = None
    pdf_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix='.odt', delete=False) as f:
            f.write(odt_bytes)
            temp_odt_path = f.name

        temp_pdf_dir = tempfile.gettempdir()
        profile_dir = os.path.join(temp_pdf_dir, f"lo_profile_{os.getpid()}")
        comando = [
            libreoffice_path,
            f'-env:UserInstallation=file://{profile_dir}',
            '--headless', '--norestore', '--nofirststartwizard',
            '--convert-to', 'pdf',
            '--outdir', temp_pdf_dir,
            temp_odt_path
        ]
        env_config = os.environ.copy()
        env_config['HOME'] = temp_pdf_dir

        process = subprocess.Popen(comando, stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=env_config, shell=(os.name == 'nt'))
        stdout, stderr = process.communicate(timeout=120)

        pdf_filename = os.path.basename(temp_odt_path).replace('.odt', '.pdf')
        pdf_path = os.path.join(temp_pdf_dir, pdf_filename)

        if not os.path.exists(pdf_path):
            raise Exception(f"PDF não gerado. Stderr: {stderr.decode('utf-8', errors='ignore')}")

        with open(pdf_path, 'rb') as f:
            return f.read()

    except Exception as e:
        st.error(f"Falha na conversão local via LibreOffice: {str(e)}")
        return None
    finally:
        if temp_odt_path and os.path.exists(temp_odt_path):
            os.unlink(temp_odt_path)
        if pdf_path and os.path.exists(pdf_path):
            os.unlink(pdf_path)


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


def formatar_data_extenso(data_obj=None):
    """Formata a data no estilo: Goiânia, 27 de ABRIL de 2026"""
    if not data_obj:
        data_obj = datetime.today()
    meses = {1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL", 5: "MAIO", 6: "JUNHO", 7: "JULHO", 8: "AGOSTO", 9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"}
    return f"Goiânia, {data_obj.day} de {meses.get(data_obj.month, '')} de {data_obj.year}"


def criar_substituicoes(dados):
    """Prepara dicionário de substituições a partir de uma linha (dict) do DataFrame"""
    substituicoes = {}
    data_hoje = formatar_data_extenso()

    # Mapeamento dos placeholders para as colunas (considerando nomes exatos)
    mapeamento_placeholders = {
        "<Cliente>": "Cliente", "<Cidade>": "Cidade", "<Estado>": "Estado",
        "<Número>": "Número", "<Nome>": "Nome", "<Telefone>": "Telefone",
        "<Email>": "Email", "<Modelo>": "Modelo", "<TIPO DE MÁQUINA>": "TIPO DE MÁQUINA",
        "<MODELO DE MÁQUINA>": "MODELO DE MÁQUINA", "<Valor Rompedor>": "Valor Rompedor",
        "<Valor Kit>": "Valor Kit", "<Condição de pagamento>": "Condição de pagamento",
        "<FRETE>": "FRETE", "<Data>": "Data", "<Observações>": "Observações"
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
                  substituicoes[placeholder] = formatar_data_extenso(valor)
             else:
                  # Tenta converter string para data, se falhar usa o valor como está ou data de hoje
                  try:
                       data_obj = pd.to_datetime(valor, errors='coerce')
                       if pd.isna(data_obj):
                            substituicoes[placeholder] = str(valor) if valor else data_hoje
                       else:
                            substituicoes[placeholder] = formatar_data_extenso(data_obj)
                  except Exception:
                       substituicoes[placeholder] = str(valor) if valor else data_hoje
        # Para outros campos, apenas converte para string
        else:
            substituicoes[placeholder] = str(valor)

    return substituicoes

