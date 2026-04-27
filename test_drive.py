import os, toml
from google.oauth2 import service_account
from googleapiclient.discovery import build

sec_path = os.path.join('.streamlit', 'secrets.toml')
with open(sec_path, 'r', encoding='utf-8') as f: secrets = toml.load(f)

creds = service_account.Credentials.from_service_account_info(secrets['google_cloud'], scopes=['https://www.googleapis.com/auth/drive'])
service = build('drive', 'v3', credentials=creds)

query = "'1daF_KyzA1te7cMFBuKJaAzvYxX7D7gKu' in parents and trashed=false"
for item in service.files().list(q=query, fields='files(id, name, mimeType)').execute().get('files', []):
    print(item['name'])
