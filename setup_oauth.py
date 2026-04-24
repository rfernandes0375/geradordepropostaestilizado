"""
Script de setup OAuth2 - Execute UMA VEZ localmente para obter o refresh token.

Pré-requisitos:
  pip install google-auth-oauthlib

Instruções:
  1. Acesse: console.cloud.google.com
  2. APIs & Services > Credentials > "+ Create Credentials" > "OAuth 2.0 Client IDs"
  3. Application type: "Desktop app", nome: "Gerador Propostas"
  4. Baixe o JSON e salve como "client_secret.json" na pasta do projeto
  5. Execute este script: python setup_oauth.py
  6. Faça login no navegador que abrir
  7. Copie o output e cole nos secrets do Streamlit Cloud
"""

from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ['https://www.googleapis.com/auth/drive']

flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
creds = flow.run_local_server(port=0)

print("\n" + "="*60)
print("COPIE O TRECHO ABAIXO NOS SECRETS DO STREAMLIT CLOUD:")
print("="*60)
print(f"""
[oauth2_drive]
client_id = "{creds.client_id}"
client_secret = "{creds.client_secret}"
refresh_token = "{creds.refresh_token}"
""")
print("="*60)
