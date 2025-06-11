import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import gspread
import os
from pathlib import Path
from dotenv import load_dotenv

# Carregar o .env corretamente
env_path = Path(__file__).parent / '.env'
load_dotenv(dotenv_path=env_path)

# Configurações da API Google
GOOGLE_DRIVE_SCOPE = 'https://www.googleapis.com/auth/drive'
GOOGLE_SHEETS_SCOPE = 'https://www.googleapis.com/auth/spreadsheets'

def autenticar_google():
    """Autentica com as APIs do Google e retorna os serviços"""
    try:
        cred_path = os.getenv('GOOGLE_CREDS_JSON_PATH')
        if cred_path is None:
            raise ValueError("A variável de ambiente 'GOOGLE_CREDS_JSON_PATH' não foi carregada do .env!")

        creds = Credentials.from_service_account_file(
            cred_path,
            scopes=[GOOGLE_DRIVE_SCOPE, GOOGLE_SHEETS_SCOPE]
        )
        service_drive = build('drive', 'v3', credentials=creds)
        service_sheets = build('sheets', 'v4', credentials=creds)
        gs_client = gspread.authorize(creds)
        return service_drive, service_sheets, gs_client
    except Exception as e:
        print(f"Erro na autenticação: {e}")
        raise

def processar_dados_incubadas():
    """Cria um DataFrame com os dados das incubadas"""
    dados = {
        'Nome': [
            'Meu Troco', 'Prime', 'Canteiro', 'João de Barro', 'Telsonn', 'Bem vivido', 'Skeed', 'School King', 'AMSA',
            'Energy Greed', 'Guiar', 'Healpp tech', 'MertricDev', 'Higia', 'Printese', 'SmartGears', 'Carreira Hub',
            'Inclusão na Prática', 'Innovacci Alimentos Inteligentes', 'Lis Saúde', 'MAKE A VISION', 'Ponto a Ponto',
            'Reversa com Sistema Solar Offgrid', 'Shopping Cidadão', 'Sistema de Gerenciamento de Fichas Online - SisGeFiO',
            'Smart Solutions IoT', 'Vexa Lab', 'Alimentos Upcycled', 'Cata cata', 'Chemall', 'Coco Dog', 'Ecocamping Lumiar',
            'Ookami', 'Universo das plantas', 'BioPoliTech', 'ForCE Metabolomics', "K'auy bebidas", 'Mãe do Mato Foodtech',
            'Qualileite Neonatal', 'Sitio Mangara', 'Smart Chef'
        ],
        'Ano': [
            2020, 2020, 2020, 2020, 2020, 2020, 2020, 2020, 2020,
            2021, 2021, 2021, 2021, 2021, 2021, 2021, 2021,
            2022, 2022, 2022, 2022, 2022, 2022, 2022, 2022, 2022, 2022,
            2023, 2023, 2023, 2023, 2023, 2023, 2023,
            2024, 2024, 2024, 2024, 2024, 2024, 2024
        ]
    }
    return pd.DataFrame(dados)


