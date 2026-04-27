import os
import json
import toml
import google.generativeai as genai
import time

# 1. Carregar a chave de API do arquivo secrets.toml
api_key = None
try:
    secrets_path = os.path.join(os.path.dirname(__file__), ".streamlit", "secrets.toml")
    if os.path.exists(secrets_path):
        with open(secrets_path, "r", encoding="utf-8") as f:
            dados = toml.load(f)
            api_key = dados.get("gemini", {}).get("api_key")
except Exception as e:
    print(f"Erro ao ler secrets: {e}")

if not api_key:
    print("ERRO CRÍTICO: Chave do Gemini (api_key) não encontrada.")
else:
    genai.configure(api_key=api_key)

def extrair_dados_proposta(texto_ou_audio_path, tipo="texto", prompt_personalizado=None):
    """
    Versão com Retry Automático para evitar Erro 429 (Limite de Cota).
    """
    if not api_key:
        return {"erro": "Chave de API não configurada."}

    prompt_base = """
Você é um assistente comercial da Jardim Equipamentos. Extraia os dados e gere um JSON.
Campos: Transcricao, Cliente, Cidade, Estado, Nome, Email, Telefone, Modelo, TIPO DE MÁQUINA, MODELO DE MÁQUINA, Valor Rompedor, Condição de pagamento, FRETE.
Use MAIÚSCULAS. No FRETE inclua o local (ex: FOB RECIFE). No VALOR use apenas números e vírgula.
"""

    # Modelos que apareceram como disponíveis na sua lista oficial
    modelos_para_tentar = ["gemini-flash-latest", "gemini-2.0-flash", "gemini-2.0-flash-lite"]
    
    response = None
    erro_final = ""

    for nome_modelo in modelos_para_tentar:
        tentativas = 2
        while tentativas > 0:
            try:
                print(f"IA: Tentando {nome_modelo} (Restam {tentativas} tentativas)...")
                # Usa prompt_base como system_instruction
                model = genai.GenerativeModel(
                    model_name=nome_modelo, 
                    generation_config={"response_mime_type": "application/json", "temperature": 0.1}, 
                    system_instruction=prompt_base
                )

                if tipo == "texto":
                    # Se tiver prompt personalizado (correção), envia junto
                    msg = [texto_ou_audio_path]
                    if prompt_personalizado:
                        msg = [f"JSON ATUAL: {prompt_personalizado}. CORREÇÃO: {texto_ou_audio_path}"]
                    response = model.generate_content(msg)
                elif tipo == "audio":
                    print("IA: Upload do áudio...")
                    arquivo = genai.upload_file(path=texto_ou_audio_path, mime_type="audio/ogg")
                    while arquivo.state.name == "PROCESSING":
                        time.sleep(2)
                        arquivo = genai.get_file(arquivo.name)
                    
                    # Se for correção via áudio, manda o JSON atual junto com o arquivo de áudio
                    msg_audio = [arquivo]
                    if prompt_personalizado:
                        msg_audio.append(f"JSON ATUAL PARA CORRIGIR: {prompt_personalizado}")
                    
                    response = model.generate_content(msg_audio)
                    genai.delete_file(arquivo.name)
                
                if response and response.text:
                    print(f"IA: Sucesso com {nome_modelo}!")
                    res = json.loads(response.text.replace("```json", "").replace("```", "").strip())
                    return res[0] if isinstance(res, list) and len(res) > 0 else res

            except Exception as e:
                erro_str = str(e)
                print(f"⚠️ IA: Erro em {nome_modelo}: {erro_str[:100]}")
                
                if "429" in erro_str:
                    print("⏳ Limite de cota atingido. Esperando 10 segundos para tentar novamente...")
                    time.sleep(10)
                    tentativas -= 1
                    continue
                else:
                    erro_final = erro_str
                    break # Se não for 429, tenta o próximo modelo
            
            tentativas = 0 # Sai do while se não houve erro ou se já tentou o retry

    return {"erro": f"Limite de cota do Google excedido. Tente novamente em 1 minuto. ({erro_final[:50]})"}
