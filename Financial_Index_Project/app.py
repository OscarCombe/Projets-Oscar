import streamlit as st
import pandas as pd
import os
import seaborn as sns
import matplotlib.pyplot as plt
import numpy as np
import smtplib
from io import BytesIO
import zipfile
import urllib.parse
import shutil
import requests

# Titre de l'application
st.title("📊 Analyse et Création Interactive d'Indices Financiers")

# Explication introductive
st.markdown("""
Bienvenue dans cette application interactive dédiée à l'analyse et à la création d'indices financiers.  
Vous pourrez explorer, filtrer, et construire des indices basés sur des entreprises américaines (**SPX**) et européennes (**SXXP**) par **secteurs** et **sous-secteurs**.

### Objectifs :
1. **Analyse sectorielle** : Identifiez les entreprises pertinentes dans le secteur de votre choix.
2. **Création d'indices** : Construisez et visualisez des indices sectoriels adaptés à vos critères.
3. **Comparaison avec benchmarks** : Évaluez les performances des indices en les comparant à des benchmarks globaux comme SPX et SXXP.

Grâce à cette plateforme, vous pourrez également explorer des indices basés sur des styles d'investissement spécifiques (Momentum, Solidité Financière) pour mieux comprendre les dynamiques de marché.

**👉 Commencez dès maintenant en sélectionnant un secteur à analyser via le panneau latéral.**
""")

# Chargement des données avec mise en cache

@st.cache_data
def charger_donnees():
    # url = "https://www.dropbox.com/scl/fi/aazc2gnzofqjee5fsc9sm/Data-projet-indices-python.xlsx?rlkey=6vyz3mbazfqx4c665ud6mnesj&st=9vzw1bf8&dl=1"
    local_path = "Data projet indices python.xlsx"

    #if not os.path.exists(local_path):
       #r = requests.get(url)
       # with open(local_path, 'wb') as f:
           # f.write(r.content)

    index_data = pd.read_excel(local_path, sheet_name='Index', engine='openpyxl')
    forex_data = pd.read_excel(local_path, sheet_name="Forex", engine='openpyxl')
    members_data = pd.read_excel(local_path, sheet_name='Members', engine='openpyxl')
    spx_prices = pd.read_excel(local_path, sheet_name='SPX_PX_LAST', engine='openpyxl')
    sxxp_prices = pd.read_excel(local_path, sheet_name='SXXP_PX_LAST', engine='openpyxl')
    qualitativ_2018 = pd.read_excel(local_path, sheet_name="Qualitativ_2018", engine='openpyxl')
    qualitativ_2019 = pd.read_excel(local_path, sheet_name="Qualitativ_2019", engine='openpyxl')
    qualitativ_2020 = pd.read_excel(local_path, sheet_name="Qualitativ_2020", engine='openpyxl')

    return {
        'index_data': index_data,
        'forex_data': forex_data,
        'members_data': members_data,
        'spx_prices': spx_prices,
        'sxxp_prices': sxxp_prices,
        'qualitativ_2018': qualitativ_2018,
        'qualitativ_2019': qualitativ_2019,
        'qualitativ_2020': qualitativ_2020,
    }

donnees = charger_donnees()
