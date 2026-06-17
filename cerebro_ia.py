import os
import json
import toml
import time
import requests

# google.generativeai mantido para suporte a áudio (multimodal)
try:
    import google.generativeai as genai
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False

# --- Carrega chaves de API ---
api_key_gemini = None
api_key_openrouter = None
try:
    secrets_path = os.path.join(os.path.dirname(__file__), ".streamlit", "secrets.toml")
    if os.path.exists(secrets_path):
        with open(secrets_path, "r", encoding="utf-8") as f:
            dados = toml.load(f)
            api_key_gemini    = dados.get("gemini",    {}).get("api_key")
            api_key_openrouter = dados.get("openrouter", {}).get("api_key")
except Exception as e:
    print(f"Erro ao ler secrets: {e}")

if api_key_gemini and GEMINI_AVAILABLE:
    genai.configure(api_key=api_key_gemini)

# --- Prompt compartilhado entre todos os motores ---
PROMPT_BASE = """
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
- Valor Kit: Apenas números e vírgula (ex: 15.500,00).
- Condição de pagamento: Detalhes do parcelamento.
- FRETE: CIF ou FOB + Local (ex: FOB RECIFE, CIF GOIÂNIA).

REGRAS CRÍTICAS:
1. Retorne APENAS o JSON.
2. Se não encontrar um dado, use "---".
3. Use letras MAIÚSCULAS para tudo (exceto e-mail).
4. Se houver um JSON ATUAL, aplique as correções do usuário sobre ele.
"""


def extrair_dados_proposta_openrouter(texto, prompt_personalizado=None, status_callback=None):
    """
    Motor Principal: OpenRouter (agregador de modelos gratuitos).
    Suporte apenas a TEXTO. Tenta múltiplos modelos em cascata.
    """
    if not api_key_openrouter:
        return None

    def log(msg):
        print(f"IA (OpenRouter): {msg}")
        if status_callback: status_callback(f"🌐 IA (OpenRouter): {msg}")

    if prompt_personalizado:
        user_content = (
            f"JSON ATUAL: {prompt_personalizado}\n\n"
            f"CORREÇÃO A APLICAR: {texto}\n\n"
            f"Retorne o JSON atualizado com as mudanças."
        )
    else:
        user_content = f"Texto do usuário:\n{texto}"

    messages = [
        {"role": "system", "content": PROMPT_BASE},
        {"role": "user",   "content": user_content},
    ]

    # Modelos gratuitos disponíveis no OpenRouter (em ordem de preferência)
    # Ordem baseada em testes reais de disponibilidade (17/06/2026)
    modelos = [
        "nvidia/nemotron-3-ultra-550b-a55b:free",   # confirmado funcionando
        "meta-llama/llama-3.3-70b-instruct:free",   # fallback — pode ter rate limit
        "google/gemma-4-31b-it:free",               # fallback — pode ter rate limit
        "qwen/qwen3-next-80b-a3b-instruct:free",    # fallback — pode ter rate limit
        "openai/gpt-oss-120b:free",                 # fallback adicional
    ]

    for modelo in modelos:
        try:
            log(f"Tentando {modelo}...")
            resp = requests.post(
                "https://openrouter.ai/api/v1/chat/completions",
                headers={
                    "Authorization": f"Bearer {api_key_openrouter}",
                    "Content-Type": "application/json",
                    "HTTP-Referer": "https://jardimequipamentos.com.br",
                    "X-Title": "Jardim Equipamentos - Gerador de Propostas",
                },
                json={
                    "model": modelo,
                    "messages": messages,
                    "response_format": {"type": "json_object"},
                    "temperature": 0.1,
                },
                timeout=60,
            )
            if resp.status_code == 200:
                content = resp.json()["choices"][0]["message"]["content"]
                res = json.loads(content)
                log("Dados extraídos com sucesso!")
                return res
            else:
                log(f"Modelo {modelo} retornou {resp.status_code}. Próximo...")
        except Exception as e:
            log(f"Erro em {modelo}: {e}")
            continue

    return None


def extrair_dados_proposta_gemini(texto_ou_audio_path, tipo="texto", prompt_personalizado=None, status_callback=None):
    """
    Fallback / Motor de Áudio: Gemini (suporte multimodal nativo para áudio).
    Para texto funciona como fallback quando o OpenRouter falha.
    """
    if not api_key_gemini or not GEMINI_AVAILABLE:
        return None

    def log(msg):
        print(f"IA (Gemini): {msg}")
        if status_callback: status_callback(f"🧠 IA (Gemini): {msg}")

    modelos = ["gemini-2.0-flash", "gemini-2.0-flash-lite", "gemini-1.5-flash"]

    for nome_modelo in modelos:
        try:
            log(f"Tentando {nome_modelo}...")
            model = genai.GenerativeModel(
                model_name=nome_modelo,
                generation_config={"response_mime_type": "application/json", "temperature": 0.1},
                system_instruction=PROMPT_BASE,
            )

            if tipo == "texto":
                msg = [texto_ou_audio_path]
                if prompt_personalizado:
                    msg = [f"JSON ATUAL: {prompt_personalizado}. CORREÇÃO: {texto_ou_audio_path}"]
                response = model.generate_content(msg)

            else:  # audio
                log("Processando áudio...")
                arquivo = genai.upload_file(path=texto_ou_audio_path, mime_type="audio/ogg")
                while arquivo.state.name == "PROCESSING":
                    time.sleep(2)
                    arquivo = genai.get_file(arquivo.name)
                msg_audio = [arquivo]
                if prompt_personalizado:
                    msg_audio.append(f"JSON ATUAL PARA CORRIGIR: {prompt_personalizado}")
                response = model.generate_content(msg_audio)
                try:
                    genai.delete_file(arquivo.name)
                except Exception:
                    pass

            if response and response.text:
                texto_limpo = response.text.replace("```json", "").replace("```", "").strip()
                res = json.loads(texto_limpo)
                return res[0] if isinstance(res, list) and res else res

        except Exception as e:
            if "429" in str(e):
                log(f"Limite atingido em {nome_modelo}...")
                continue
            log(f"Erro em {nome_modelo}: {e}")
            break

    return None


def extrair_dados_proposta(texto_ou_audio_path, tipo="texto", prompt_personalizado=None, status_callback=None):
    """
    Ponto de entrada principal. Cascata de IAs:

    TEXTO → OpenRouter (principal, gratuito, multi-modelo)
           → Gemini (fallback se OpenRouter falhar)

    ÁUDIO → Gemini (único motor com suporte multimodal nativo)
    """
    if tipo == "audio":
        # Áudio: apenas Gemini suporta nativamente
        res = extrair_dados_proposta_gemini(texto_ou_audio_path, tipo, prompt_personalizado, status_callback)
        if res and "erro" not in res:
            return res
        return {"erro": "IA indisponível para áudio. Envie os dados por texto."}

    # Texto: OpenRouter primeiro
    res = extrair_dados_proposta_openrouter(texto_ou_audio_path, prompt_personalizado, status_callback)
    if res and "erro" not in res:
        return res

    # Fallback: Gemini
    if status_callback:
        status_callback("🔄 OpenRouter indisponível, tentando Gemini...")
    res = extrair_dados_proposta_gemini(texto_ou_audio_path, tipo, prompt_personalizado, status_callback)
    if res and "erro" not in res:
        return res

    return {"erro": "Todas as IAs estão fora do ar ou sem cota."}
