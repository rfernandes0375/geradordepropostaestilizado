import sys
import json
import os
import toml
import tempfile
import re
from datetime import datetime

# --- MOCK DO STREAMLIT PARA O BOT RODAR SEM ERRO ---
class DummyStreamlit:
    secrets = {}
    @staticmethod
    def error(msg): print(f"ERRO (Streamlit): {msg}")
    @staticmethod
    def warning(msg): print(f"AVISO (Streamlit): {msg}")

try:
    sec_path = os.path.join(os.path.dirname(__file__), '.streamlit', 'secrets.toml')
    if os.path.exists(sec_path):
        with open(sec_path, 'r', encoding='utf-8') as f:
            DummyStreamlit.secrets = toml.load(f)
except Exception as e:
    print(f"Aviso ao carregar secrets local: {e}")

sys.modules['streamlit'] = DummyStreamlit
# ----------------------------------------------------

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, CallbackQueryHandler
import gspread

from cerebro_ia import extrair_dados_proposta
from motor_pdf import (
    listar_modelos_google_drive, 
    baixar_arquivo_drive, 
    criar_substituicoes,
    extrair_conteudo_odt,
    substituir_no_xml,
    criar_odt_modificado,
    converter_para_pdf
)

FOLDER_ID_MODELOS = "1daF_KyzA1te7cMFBuKJaAzvYxX7D7gKu"

def normalizar(texto):
    return re.sub(r'[^a-z0-9]', '', str(texto).lower())

def normalizar_uf(estado_texto):
    """Converte 'Goiás' ou 'GOIÁS' para 'GO'"""
    ufs = {
        "ACRE": "AC", "ALAGOAS": "AL", "AMAPA": "AP", "AMAZONAS": "AM", "BAHIA": "BA",
        "CEARA": "CE", "DISTRITO FEDERAL": "DF", "ESPIRITO SANTO": "ES", "GOIAS": "GO",
        "MARANHAO": "MA", "MATO GROSSO": "MT", "MATO GROSSO DO SUL": "MS", "MINAS GERAIS": "MG",
        "PARA": "PA", "PARAIBA": "PB", "PARANA": "PR", "PERNAMBUCO": "PE", "PIAUI": "PI",
        "RIO DE JANEIRO": "RJ", "RIO GRANDE DO NORTE": "RN", "RIO GRANDE DO SUL": "RS",
        "RONDONIA": "RO", "RORAIMA": "RR", "SANTA CATARINA": "SC", "SAO PAULO": "SP",
        "SERGIPE": "SE", "TOCANTINS": "TO"
    }
    texto = str(estado_texto).upper().strip()
    # Remove acentos básicos para comparação
    texto_limpo = "".join(c for c in texto if c.isalnum() or c == " ")
    texto_sem_acento = texto_limpo.replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U").replace("Ã", "A").replace("Õ", "O").replace("Ç", "C")
    
    if texto_sem_acento in ufs:
        return ufs[texto_sem_acento]
    return texto[:2] # Retorna as 2 primeiras letras se não achar

def is_authorized(user_id):
    authorized = DummyStreamlit.secrets.get("telegram", {}).get("authorized_users", [])
    if not authorized: return True
    return user_id in authorized

def salvar_na_planilha_google(dados_proposta):
    try:
        if "google_cloud" not in DummyStreamlit.secrets: return False
        gc = gspread.service_account_from_dict(DummyStreamlit.secrets["google_cloud"])
        sh = gc.open_by_url("https://docs.google.com/spreadsheets/d/1CkllnC9xWKzZEnB_bKFRuNY5dsu1FoGRF5phdHirjbs/edit?gid=465824771#gid=465824771")
        worksheet = sh.get_worksheet_by_id(465824771)
        headers = worksheet.row_values(1)
        row_to_add = []
        
        if "Número" in headers:
            idx_num = headers.index("Número") + 1
            col_nums = worksheet.col_values(idx_num)
            ultimo = col_nums[-1] if len(col_nums) > 1 else ""
            if '-' in ultimo:
                partes = ultimo.split('-')
                try:
                    num = int(''.join(filter(str.isdigit, partes[0]))) + 1
                    dados_proposta["Número"] = f"{num}-{partes[1]}"
                except:
                    dados_proposta["Número"] = datetime.now().strftime("%H%M")
            else:
                dados_proposta["Número"] = datetime.now().strftime("%H%M")
        else:
            dados_proposta["Número"] = datetime.now().strftime("%H%M")

        for h in headers:
            if h == "Data" and not dados_proposta.get("Data"):
                 hoje = datetime.today()
                 meses = {1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL", 5: "MAIO", 6: "JUNHO", 7: "JULHO", 8: "AGOSTO", 9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"}
                 row_to_add.append(f"Goiânia, {hoje.day} de {meses.get(hoje.month, '')} de {hoje.year}")
            elif h == "Número":
                 row_to_add.append(dados_proposta.get("Número", ""))
            elif h == "NOME DO ARQUIVO":
                 row_to_add.append(f"JE {dados_proposta.get('Número', '')} - {dados_proposta.get('Cliente', '')}")
            else:
                 row_to_add.append(str(dados_proposta.get(h, "")))
                 
        worksheet.append_row(row_to_add)
        return True
    except Exception as e:
        print(f"Erro planilha: {e}")
        return False

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("👋 Robô Comercial Pronto!")

async def exibir_resumo_edicao(update: Update, context: ContextTypes.DEFAULT_TYPE):
    dados = context.user_data.get('dados_temp')
    if not dados: return

    resumo = (
        "📝 **Conferência de Proposta:**\n\n"
        f"🏢 **Cliente:** {dados.get('Cliente', '---')}\n"
        f"📍 **Cidade/UF:** {dados.get('Cidade', '---')} / {dados.get('Estado', '---')}\n"
        f"👤 **Contato:** {dados.get('Nome', '---')}\n"
        f"📞 **Fone:** {dados.get('Telefone', '---')}\n"
        f"📧 **E-mail:** {dados.get('Email', '---')}\n"
        f"🛠️ **Equipamento:** {dados.get('Modelo', '---')}\n"
        f"🏗️ **Tipo Máq:** {dados.get('TIPO DE MÁQUINA', '---')}\n"
        f"🚜 **Mod. Máq:** {dados.get('MODELO DE MÁQUINA', '---')}\n"
        f"💰 **Valor Romp:** R$ {dados.get('Valor Rompedor', '---')}\n"
        f"💳 **Pagamento:** {dados.get('Condição de pagamento', '---')}\n"
        f"🚚 **Frete:** {dados.get('FRETE', '---')}\n\n"
        "💡 *Dica: Digite a correção direto no chat ou use os botões:* "
    )

    keyboard = [
        [InlineKeyboardButton("🏢 Cliente", callback_data="edit_Cliente"), InlineKeyboardButton("📍 Cidade/UF", callback_data="edit_Cidade")],
        [InlineKeyboardButton("👤 Contato", callback_data="edit_Nome"), InlineKeyboardButton("📞 Fone", callback_data="edit_Telefone")],
        [InlineKeyboardButton("📧 E-mail", callback_data="edit_Email"), InlineKeyboardButton("🛠️ Equipamento", callback_data="edit_Modelo")],
        [InlineKeyboardButton("🏗️ Tipo Máq", callback_data="edit_TIPO DE MÁQUINA"), InlineKeyboardButton("🚜 Mod. Máq", callback_data="edit_MODELO DE MÁQUINA")],
        [InlineKeyboardButton("💰 Valor", callback_data="edit_Valor Rompedor"), InlineKeyboardButton("💳 Pagamento", callback_data="edit_Condição de pagamento")],
        [InlineKeyboardButton("🚚 Frete", callback_data="edit_FRETE")],
        [InlineKeyboardButton("🚀 CONFIRMAR TUDO", callback_data="confirmar_tudo")],
        [InlineKeyboardButton("❌ Limpar Tudo", callback_data="cancelar")]
    ]
    
    reply_markup = InlineKeyboardMarkup(keyboard)
    if update.callback_query:
        await update.callback_query.edit_message_text(resumo, reply_markup=reply_markup, parse_mode="Markdown")
    else:
        await update.message.reply_text(resumo, reply_markup=reply_markup, parse_mode="Markdown")

async def exibir_selecao_modelo(query, context: ContextTypes.DEFAULT_TYPE):
    dados = context.user_data.get('dados_temp')
    await query.edit_message_text("⏳ Buscando arquivos ODT no Drive...")
    modelos = listar_modelos_google_drive(FOLDER_ID_MODELOS)
    
    palavras_busca = []
    for t in [dados.get("Modelo", ""), dados.get("MODELO DE MÁQUINA", ""), dados.get("TIPO DE MÁQUINA", "")]:
        palavras = re.findall(r'[A-Z0-9]+', str(t).upper())
        for p in palavras:
            if len(p) >= 2 or p.isdigit():
                if p not in palavras_busca: palavras_busca.append(p)

    ranking = []
    for m in modelos:
        nome_norm = normalizar(m['name'])
        score = sum(5 if p.isdigit() and normalizar(p) in nome_norm else (1 if normalizar(p) in nome_norm else 0) for p in palavras_busca)
        ranking.append((score, m))
    
    ranking.sort(key=lambda x: x[0], reverse=True)
    keyboard = [[InlineKeyboardButton(f"{'✅ ' if score > 0 else ''}{m['name']}", callback_data=f"file_{m['id']}")] for score, m in ranking[:6]]
    keyboard.append([InlineKeyboardButton("⬅️ Voltar para Edição", callback_data="voltar_edicao")])
    await query.edit_message_text("🎯 **Selecione o Modelo de Proposta:**", reply_markup=InlineKeyboardMarkup(keyboard), parse_mode="Markdown")

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if not is_authorized(user.id): return

    texto_usuario = update.message.text
    waiting_field = context.user_data.get('waiting_for')
    
    if waiting_field:
        if waiting_field == "Valor Rompedor": texto_usuario = re.sub(r'[^0-9,]', '', texto_usuario)
        val = texto_usuario.lower() if waiting_field == "Email" else texto_usuario.upper()
        context.user_data['dados_temp'][waiting_field] = val
        context.user_data['waiting_for'] = None
        await exibir_resumo_edicao(update, context)
        return

    if context.user_data.get('dados_temp'):
        msg_wait = await update.message.reply_text("🔄 Atualizando dados...")
        contexto_atual = json.dumps(context.user_data['dados_temp'])
        prompt_correcao = f"Com base no JSON atual: {contexto_atual}. Aplique esta correção do usuário: {texto_usuario}. Retorne o JSON atualizado."
        dados_novos = extrair_dados_proposta(prompt_correcao, tipo="texto")
        for k, v in dados_novos.items():
            if v and v != "---" and k != "Transcricao":
                if k == "Estado": val = normalizar_uf(v)
                elif k == "Email": val = str(v).lower()
                else: val = v
                context.user_data['dados_temp'][k] = val

        await msg_wait.delete()
        await exibir_resumo_edicao(update, context)
        return

    msg_wait = await update.message.reply_text("🧠 Extraindo dados...")
    dados = extrair_dados_proposta(texto_usuario, tipo="texto")
    if "Estado" in dados: dados["Estado"] = normalizar_uf(dados["Estado"])
    if "Email" in dados: dados["Email"] = str(dados["Email"]).lower()
    await msg_wait.delete()
    context.user_data['dados_temp'] = dados
    await exibir_resumo_edicao(update, context)

async def handle_voice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if not is_authorized(user.id): return
    msg_wait = await update.message.reply_text("🎤 Analisando áudio...")
    voice_file = await context.bot.get_file(update.message.voice.file_id)
    with tempfile.NamedTemporaryFile(suffix=".ogg", delete=False) as tmp:
        await voice_file.download_to_drive(tmp.name)
        tmp_path = tmp.name
    
    if context.user_data.get('dados_temp'):
        contexto_atual = json.dumps(context.user_data['dados_temp'])
        dados_novos = extrair_dados_proposta(f"Audio de correção para este JSON: {contexto_atual}. Audio path: {tmp_path}", tipo="audio")
        for k, v in dados_novos.items():
            if v and v != "---" and k != "Transcricao":
                if k == "Estado": val = normalizar_uf(v)
                elif k == "Email": val = v.lower()
                else: val = v
                context.user_data['dados_temp'][k] = val
    else:
        dados = extrair_dados_proposta(tmp_path, tipo="audio")
        if "Estado" in dados: dados["Estado"] = normalizar_uf(dados["Estado"])
        if "Email" in dados: dados["Email"] = str(dados["Email"]).lower()
        context.user_data['dados_temp'] = dados

    os.unlink(tmp_path)
    await msg_wait.delete()
    await exibir_resumo_edicao(update, context)

async def on_button_click(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "cancelar":
        context.user_data.clear()
        await query.edit_message_text("🗑️ Dados limpos. Pode começar uma nova proposta.")
        return
    if query.data == "voltar_edicao":
        await exibir_resumo_edicao(update, context)
        return
    if query.data.startswith("edit_"):
        campo = query.data.replace("edit_", "")
        context.user_data['waiting_for'] = campo
        await query.edit_message_text(f"📝 Digite o novo valor para **{campo}**:")
        return
    if query.data == "confirmar_tudo":
        await exibir_selecao_modelo(query, context)
        return
    if query.data.startswith("file_"):
        file_id = query.data.replace("file_", "")
        dados = context.user_data.get('dados_temp')
        await query.edit_message_text("⏳ Gerando proposta final...")
        salvar_na_planilha_google(dados)
        modelos = listar_modelos_google_drive(FOLDER_ID_MODELOS)
        modelo_escolhido = next((m for m in modelos if m['id'] == file_id), None)
        modelo_bytes = baixar_arquivo_drive(modelo_escolhido['id'], modelo_escolhido['mimeType'])
        substituicoes = criar_substituicoes(dados)
        content_xml = extrair_conteudo_odt(modelo_bytes)
        content_xml_modificado, _ = substituir_no_xml(content_xml, substituicoes)
        odt_modificado = criar_odt_modificado(modelo_bytes, content_xml_modificado)
        nome_arquivo = f"JE {dados.get('Número')} - {dados.get('Cliente', 'Proposta')}"
        pdf_bytes = converter_para_pdf(odt_modificado, nome_arquivo)

        if pdf_bytes:
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                tmp.write(pdf_bytes)
                tmp_path = tmp.name
            with open(tmp_path, "rb") as f:
                await context.bot.send_document(chat_id=query.message.chat_id, document=f, filename=f"{nome_arquivo}.pdf", write_timeout=60)
            os.unlink(tmp_path)
            await query.delete_message()
        else:
            await query.edit_message_text("❌ Falha no PDF.")

def main():
    token = DummyStreamlit.secrets.get("telegram", {}).get("bot_token")
    if not token: return
    app = Application.builder().token(token).read_timeout(60).connect_timeout(60).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.add_handler(MessageHandler(filters.VOICE, handle_voice))
    app.add_handler(CallbackQueryHandler(on_button_click))
    print("🚀 Robô Online!")
    app.run_polling()

if __name__ == "__main__":
    main()
