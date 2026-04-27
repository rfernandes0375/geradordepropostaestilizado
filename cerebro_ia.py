import os
import json
import toml
import google.generativeai as genai
import time
from groq import Groq

# 1. Carregar as chaves de API do arquivo secrets.toml
api_key_gemini = None
api_key_groq = None
try:
    secrets_path = os.path.join(os.path.dirname(__file__), ".streamlit", "secrets.toml")
    if os.path.exists(secrets_path):
        with open(secrets_path, "r", encoding="utf-8") as f:
            dados = toml.load(f)
            api_key_gemini = dados.get("gemini", {}).get("api_key")
            api_key_groq = dados.get("groq", {}).get("api_key")
except Exception as e:
    print(f"Erro ao ler secrets: {e}")

# Configura o Gemini se a chave existir
if api_key_gemini:
    genai.configure(api_key=api_key_gemini)

def extrair_dados_proposta_groq(texto_ou_audio_path, tipo="texto", prompt_personalizado=None, status_callback=None):
    """Fallback usando Groq (Llama 3 + Whisper)"""
    if not api_key_groq:
        return None
    
    def log(msg):
        print(f"IA (Groq): {msg}")
        if status_callback: status_callback(f"☁️ IA (Groq): {msg}")

    try:
        client = Groq(api_key=api_key_groq)
        # Prompt mais detalhado para o Llama
        prompt_base = """
Você é um especialista em extração de dados comerciais da Jardim Equipamentos.
Sua missão é ler um texto ou transcrição e retornar um JSON com estes campos EXATOS:
- Transcricao: O texto completo ouvido ou lido.
- Cliente: Nome da empresa ou pessoa.
- Cidade: Nome da cidade.
- Estado: SIGLA do estado (ex: GO, SP, RJ).
- Nome: Nome do contato/vendedor.
- Email: E-mail em letras minúsculas.
- Telefone: Apenas números e traços.
- Modelo: O modelo do equipamento (ex: AT 810M).
- TIPO DE MÁQUINA: Ex: ESCAVADEIRA, RETROESCAVADEIRA.
- MODELO DE MÁQUINA: Ex: KOMATSU PC200, CAT 320.
- Valor Rompedor: Apenas números e vírgula (ex: 110.000,00).
- Condição de pagamento: Detalhes do parcelamento.
- FRETE: CIF ou FOB + Local (ex: FOB RECIFE, CIF GOIÂNIA).

REGRAS CRÍTICAS:
1. Retorne APENAS o JSON.
2. Se não encontrar um dado, use "---".
3. Use letras MAIÚSCULAS para tudo (exceto e-mail).
"""
        texto_final = ""
        if tipo == "audio":
            log("Transcrevendo áudio...")
            with open(texto_ou_audio_path, "rb") as file:
                transcription = client.audio.transcriptions.create(
                    file=(texto_ou_audio_path, file.read()),
                    model="whisper-large-v3",
                    response_format="text"
                )
                texto_final = transcription
        else:
            texto_final = texto_ou_audio_path

        msg_sistema = {"role": "system", "content": prompt_base}
        msg_usuario = {"role": "user", "content": f"Extraia os dados deste texto: {texto_final}"}
        
        if prompt_personalizado:
            msg_usuario["content"] = f"JSON ATUAL: {prompt_personalizado}\n\nCORREÇÃO DO USUÁRIO: {texto_final}\n\nRetorne o JSON com as correções aplicadas."
            
        log("Extraindo JSON detalhado...")
        chat_completion = client.chat.completions.create(
            messages=[msg_sistema, msg_usuario],
            model="llama-3.3-70b-versatile",
            response_format={"type": "json_object"},
            temperature=0.1
        )
        
        res = json.loads(chat_completion.choices[0].message.content)
        if "Transcricao" not in res or res["Transcricao"] == "---":
            res["Transcricao"] = texto_final
        return res
    except Exception as e:
        log(f"Erro: {e}")
        return None

def extrair_dados_proposta(texto_ou_audio_path, tipo="texto", prompt_personalizado=None, status_callback=None):
    """Tenta Gemini, se der erro de cota, tenta Groq"""
    
    def log(msg):
        print(f"IA: {msg}")
        if status_callback: status_callback(f"🧠 IA: {msg}")

    prompt_base = """
Você é um assistente comercial da Jardim Equipamentos. Extraia os dados e gere um JSON.
Campos: Transcricao, Cliente, Cidade, Estado, Nome, Email, Telefone, Modelo, TIPO DE MÁQUINA, MODELO DE MÁQUINA, Valor Rompedor, Condição de pagamento, FRETE.
Use MAIÚSCULAS. No FRETE inclua o local (ex: FOB RECIFE). No VALOR use apenas números e vírgula.
"""

    if api_key_gemini:
        modelos_para_tentar = ["gemini-flash-latest", "gemini-2.0-flash", "gemini-2.0-flash-lite"]
        erro_final = ""

        for nome_modelo in modelos_para_tentar:
            tentativas = 2
            while tentativas > 0:
                try:
                    log(f"Tentando {nome_modelo}...")
                    model = genai.GenerativeModel(
                        model_name=nome_modelo, 
                        generation_config={"response_mime_type": "application/json", "temperature": 0.1}, 
                        system_instruction=prompt_base
                    )

                    if tipo == "texto":
                        msg = [texto_ou_audio_path]
                        if prompt_personalizado:
                            msg = [f"JSON ATUAL: {prompt_personalizado}. CORREÇÃO: {texto_ou_audio_path}"]
                        response = model.generate_content(msg)
                    elif tipo == "audio":
                        log("Processando áudio...")
                        arquivo = genai.upload_file(path=texto_ou_audio_path, mime_type="audio/ogg")
                        while arquivo.state.name == "PROCESSING":
                            time.sleep(2)
                            arquivo = genai.get_file(arquivo.name)
                        msg_audio = [arquivo]
                        if prompt_personalizado:
                            msg_audio.append(f"JSON ATUAL PARA CORRIGIR: {prompt_personalizado}")
                        response = model.generate_content(msg_audio)
                        genai.delete_file(arquivo.name)
                    
                    if response and response.text:
                        res = json.loads(response.text.replace("```json", "").replace("```", "").strip())
                        return res[0] if isinstance(res, list) and len(res) > 0 else res

                except Exception as e:
                    erro_str = str(e)
                    if "429" in erro_str:
                        log(f"Limite atingido em {nome_modelo}...")
                        tentativas -= 1
                        if tentativas > 0: time.sleep(5)
                        continue
                    else:
                        erro_final = erro_str
                        break
                tentativas = 0

    if status_callback: status_callback("🔄 Usando Backup (Groq)...")
    res_groq = extrair_dados_proposta_groq(texto_ou_audio_path, tipo, prompt_personalizado, status_callback)
    if res_groq:
        return res_groq

    return {"erro": "Limite de cota excedido. Tente novamente em 1 minuto."}
