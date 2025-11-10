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
import openpyxl


#pip install streamlit pandas matplotlib seaborn numpy
#Il faut avoir le fichier de données dans le même dossier que le fichier python

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

def charger_donnees():
    # Chemin du fichier Excel dans le même dossier que le script
    chemin = os.path.join(os.path.dirname(__file__), "data_projet_indices_python.xlsx")
    
    # Vérification de la présence du fichier
    if not os.path.exists(chemin):
        st.error(f"Fichier non trouvé : {chemin}")
        return None

    try:
        index_data = pd.read_excel(chemin, sheet_name='Index', engine='openpyxl')
        forex_data = pd.read_excel(chemin, sheet_name='Forex', engine='openpyxl')
        members_data = pd.read_excel(chemin, sheet_name='Members', engine='openpyxl')
        spx_prices = pd.read_excel(chemin, sheet_name='SPX_PX_LAST', engine='openpyxl')
        sxxp_prices = pd.read_excel(chemin, sheet_name='SXXP_PX_LAST', engine='openpyxl')
        qualitativ_2018 = pd.read_excel(chemin, sheet_name='Qualitativ_2018', engine='openpyxl')
        qualitativ_2019 = pd.read_excel(chemin, sheet_name='Qualitativ_2019', engine='openpyxl')
        qualitativ_2020 = pd.read_excel(chemin, sheet_name='Qualitativ_2020', engine='openpyxl')
    except Exception as e:
        st.error("Erreur lors du chargement du fichier Excel : " + str(e))
        return None

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

# Charger les données
donnees = charger_donnees()

def dataframe_to_image(df, filename, decimals=2):
    """
    Convertit un DataFrame en image PNG et l'enregistre avec le nom spécifié.

    Args:
        df (pd.DataFrame): Le DataFrame à convertir.
        filename (str): Le chemin du fichier PNG de sortie.
        decimals (int): Nombre de décimales pour arrondir les valeurs numériques.
    """
    # Arrondir les valeurs numériques
    df_rounded = df.round(decimals)

    # Création de l'image
    fig, ax = plt.subplots(figsize=(min(15, 5 + 0.5 * len(df_rounded.columns)), 0.5 * len(df_rounded) + 1))
    ax.axis('off')  # Pas d'axes
    ax.axis('tight')
    table = ax.table(cellText=df_rounded.values, colLabels=df_rounded.columns, cellLoc='center', loc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.auto_set_column_width(col=list(range(len(df_rounded.columns))))  # Ajuste la largeur des colonnes

    plt.savefig(filename, format='png', bbox_inches='tight')
    plt.close(fig)


# Fonction pour sauvegarder un graphique en PNG
def save_figure(fig, filename):
    """
    Enregistre un graphique en PNG avec le nom spécifié.
    """
    temp_dir = "temp_reports"
    os.makedirs(temp_dir, exist_ok=True)
    filepath = os.path.join(temp_dir, filename)
    fig.savefig(filepath, format="png", bbox_inches="tight")
    plt.close(fig)

# Fonction pour sauvegarder un fichier ZIP avec tous les résultats
def create_zip():
    """
    Crée un fichier ZIP contenant tous les fichiers enregistrés dans le répertoire temporaire.
    """
    temp_dir = "temp_reports"
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(temp_dir):
            for file in files:
                zf.write(os.path.join(root, file), arcname=file)
    return zip_buffer

# Nettoyer le dossier temporaire au démarrage
def clear_temp_folder(temp_dir="temp_reports"):
    """
    Supprime le contenu du dossier temporaire s'il existe.
    """
    if os.path.exists(temp_dir):
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                try:
                    os.unlink(os.path.join(root, file))  # Supprime les fichiers
                except PermissionError:
                    print(f"Impossible de supprimer le fichier : {file}. Il est en cours d'utilisation.")
        shutil.rmtree(temp_dir, ignore_errors=True)  # Supprime le dossier
    os.makedirs(temp_dir, exist_ok=True)  # Recrée un dossier vide

# Appel de la fonction au démarrage
clear_temp_folder()

# Section Indice Sectoriel
st.title("📈 Création d'un Indice Sectoriel")


