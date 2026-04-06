import sys
import os
import re
import io
import tempfile
import pandas as pd
import numpy as np
from sklearn.metrics import (
    mean_absolute_error, mean_squared_error, r2_score,
    accuracy_score, confusion_matrix, roc_auc_score, roc_curve
)
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image
from lightgbm import LGBMRegressor, LGBMClassifier
from pypfopt import EfficientFrontier, EfficientCVaR, risk_models, expected_returns
import streamlit as st
import datetime
from scipy.optimize import minimize 

DEFAULT_WORKBOOK = r"C:\Users\oscar\Documents\Files\Backtest and Interactive_Tool.xlsm"
SHEET_NAME = "Bloomberg Data"

st.set_page_config(layout="wide")
st.title("Outil Interactif Streamlit")

# champ texte pré-rempli avec le default
workbook_path = st.text_input("Chemin du fichier Excel source (modifiable)", value=DEFAULT_WORKBOOK, key="workbook_path")

# bouton pour (re)charger le fichier
if st.button("Relire le fichier"):
    st.session_state['workbook_path_chosen'] = workbook_path
    st.cache_data.clear()

# cache avec invalidation si le fichier change (mtime)
@st.cache_data(show_spinner=False)
def read_sheet_cached(path, sheet_name, mtime=None):
    return pd.read_excel(path, sheet_name=sheet_name, header=None, engine="openpyxl")

# Utilise la valeur sélectionnée en session_state si elle existe, sinon la valeur actuelle du champ
DEFAULT_WORKBOOK = st.session_state.get('workbook_path_chosen', workbook_path)

if not DEFAULT_WORKBOOK:
    st.info("Renseigne un chemin puis clique sur 'Relire le fichier'.")
else:
    if not os.path.exists(DEFAULT_WORKBOOK):
        st.error(f"Fichier introuvable depuis le serveur : {DEFAULT_WORKBOOK}")
        st.info("Veille à ce que le chemin soit accessible depuis le serveur Streamlit.")
    else:
        try:
            mtime = os.path.getmtime(DEFAULT_WORKBOOK)
            df = read_sheet_cached(DEFAULT_WORKBOOK, sheet_name=SHEET_NAME, mtime=mtime)
            st.success(f"Feuille '{SHEET_NAME}' lue depuis : {DEFAULT_WORKBOOK}")
            st.dataframe(df.head())
        except Exception as e:
            st.error(f"Erreur lors de la lecture : {e}")

EXCEL_HEADER_ROW = 0
DATA_START_ROW = 2
DATE_COL_INDEX = 0
HORIZON = 1
TEST_SIZE_RATIO = 0.2
MODEL_RANDOM_STATE = 42
N_LAGS = 10
MA_WINDOWS = [5, 10, 20]
VOL_WINDOWS = [5, 10]

ROLLING_MAE_WINDOW = 20

st.title("Simulations de portefeuille rapide ")

st.markdown("""
**Note :**
- Le rendement annualisé est calculé comme le rendement composé hebdo annualisé sur la période sélectionnée.
- La volatilité annualisée est calculée comme l'écart-type des rendements annualisées.
- La VaR et la CVaR mensuelle sont calculées sur les rendements mensuels composés du portefeuille.
- La VaR et la CVaR annualisée sont calculées sur les rendements hebdos annualisés.
""")

df = pd.read_excel(DEFAULT_WORKBOOK, sheet_name=SHEET_NAME, header=None, engine="openpyxl")
title_names = df.iloc[0, 1:].tolist()
data = df.iloc[2:, :].reset_index(drop=True)
data.columns = ['date'] + title_names
data['date'] = pd.to_datetime(data['date'], errors='coerce')
data = data.dropna(subset=['date'])
title_cols = data.columns[1:]

# --- Sélection de la période ---
min_date = pd.to_datetime(data['date'].min()).to_pydatetime()
max_date = pd.to_datetime(data['date'].max()).to_pydatetime()
start_date, end_date = st.slider(
    "Sélectionne la période",
    min_value=min_date,
    max_value=max_date,
    value=(min_date, max_date),
    format="YYYY-MM-DD",
    key="période"
)
df_period = data[(data['date'] >= pd.Timestamp(start_date)) & (data['date'] <= pd.Timestamp(end_date))].copy()

# --- Sélection des titres et des poids ---
st.sidebar.header("Titres et poids du portefeuille")
selected_titles = []
weights = []
for col in title_cols:
    include = st.sidebar.checkbox(f"Inclure {col}", value=True)
    if include:
        selected_titles.append(col)
        w = st.sidebar.number_input(f"Poids de {col} (%)", min_value=0.0, max_value=100.0, value=100.0/len(title_cols), step=1.0, key=f"w_{col}")
        weights.append(w)
if not selected_titles:
    st.warning("Aucun titre sélectionné.")
    st.stop()
weights = np.array(weights)
total_weight = weights.sum()

# --- Calcul des rendements hebdos des titres ---
prices = df_period[selected_titles]
returns = prices.pct_change().dropna()
returns.index = pd.to_datetime(df_period['date'].iloc[1:])  # Aligne les dates

# --- Calcul de la valeur base 100 du portefeuille semaine après semaine ---
port_cum = (returns @ (weights/100) + 1).cumprod() * 100
port_cum.index = returns.index

# --- Calcul des rendements hebdos du portefeuille ---
port_ret_hebdo = port_cum.pct_change().dropna()

# --- Choix du niveau de VaR/CVaR ---
alpha_var = st.selectbox("Niveau de VaR/CVaR (%)", [95, 99], index=0) / 100

freq = 52  # hebdo

# --- Calcul VaR/CVaR ---
def var_cvar(series, alpha_var):
    if len(series) == 0:
        return np.nan, np.nan
    var = np.quantile(series, 1 - alpha_var)
    cvar = series[series <= var].mean()
    return var, cvar

var_h, cvar_h = var_cvar(port_ret_hebdo, alpha_var)
var_h_annual = var_h * np.sqrt(freq)
cvar_h_annual = cvar_h * np.sqrt(freq)

# --- Calcul des rendements mensuels composés ---
returns.index = pd.to_datetime(returns.index)
returns_monthly = returns.resample('M').apply(lambda x: (x + 1).prod() - 1)

# VaR/CVaR mensuelle sur le portefeuille
port_ret_monthly = port_ret_hebdo.resample('M').apply(lambda x: (x + 1).prod() - 1)
var_m, cvar_m = var_cvar(port_ret_monthly, alpha_var)

# --- Affichage du tableau interactif ---
mean_returns = returns.mean()
vols = returns.std() * np.sqrt(freq)
rendement_ann_actifs = ((mean_returns + 1) ** freq - 1) * 100

table = pd.DataFrame({
    "Rendement annualisé (%)": rendement_ann_actifs,
    "Volatilité annualisée (%)": vols * 100,
    "Poids (%)": weights
}, index=selected_titles)

st.subheader("Tableau Simulation Rapide")
st.dataframe(table.style.format("{:.2f}"))

# --- Affichage du total des poids ---
if abs(total_weight - 100) > 1e-6:
    st.markdown(f"<span style='color:red'><b>Total des poids : {total_weight:.2f} %</b></span>", unsafe_allow_html=True)
else:
    st.markdown(f"**Total des poids** : {total_weight:.2f} %")

# --- Statistiques portefeuille ---
port_return_ann = (port_ret_hebdo.mean() + 1) ** freq - 1
port_vol_ann = port_ret_hebdo.std() * np.sqrt(freq)

st.markdown(f"""
**Rendement portefeuille annualisé** : {port_return_ann*100:.2f} %  
**Volatilité portefeuille annualisée** : {port_vol_ann*100:.2f} %  

**VaR {int(alpha_var*100)}% mensuelle** : {var_m*100:.2f} %  
**CVaR {int(alpha_var*100)}% mensuelle** : {cvar_m*100:.2f} %  
**VaR {int(alpha_var*100)}% annualisée ** : {var_h_annual*100:.2f} %  
**CVaR {int(alpha_var*100)}% annualisée ** : {cvar_h_annual*100:.2f} %  
""")

# --- Tableau des contributions ---
port_std_hebdo = port_ret_hebdo.std()
cov_matrix = returns.cov()
contrib_vol = []
for i, col in enumerate(selected_titles):
    cov = np.dot(cov_matrix.loc[col, selected_titles], weights/100)
    contrib = weights[i]/100 * cov / port_std_hebdo if port_std_hebdo > 0 else np.nan
    contrib_vol.append((contrib / port_std_hebdo) * 100 if port_std_hebdo > 0 else np.nan)
contrib_vol = np.array(contrib_vol)
if np.isfinite(contrib_vol).all() and np.nansum(contrib_vol) != 0:
    contrib_vol = contrib_vol * (100 / np.nansum(contrib_vol))

returns_annual = returns.resample('Y').apply(lambda x: (x + 1).prod() - 1)
cvar_ann = []
cvar_mens = []
for col in selected_titles:
    cvar_ann.append(var_cvar(returns_annual[col], alpha_var)[1] * 100)
    cvar_mens.append(var_cvar(returns_monthly[col], alpha_var)[1] * 100)

contrib_table = pd.DataFrame({
    "Contribution à la volatilité annualisée (%)": contrib_vol,
    f"CVaR {int(alpha_var*100)}% annuelle (%)": cvar_ann,
    f"CVaR {int(alpha_var*100)}% mensuelle (%)": cvar_mens,
    "Poids (%)": weights
}, index=selected_titles)

st.subheader("Contribution de chaque actif au risque du portefeuille (annuelle)")
st.dataframe(contrib_table.style.format("{:.2f}"))

if st.button("Exporter le tableau des contributions en Excel"):
    excel_buffer = io.BytesIO()
    contrib_table.to_excel(excel_buffer, sheet_name="Contributions_Portefeuille")
    st.download_button(
        label="Télécharger le tableau des contributions Excel",
        data=excel_buffer.getvalue(),
        file_name="contributions_portefeuille.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.subheader("Valeur du portefeuille (base 100)")
st.line_chart(port_cum)

def sanitize_sheet_name(name):
    name = str(name)
    name = re.sub(r'[:\\/*?\[\]]', '_', name)
    return name[:31]

def read_bloomberg_sheet_with_dates(path, sheet_name=SHEET_NAME):
    xl = pd.read_excel(path, sheet_name=sheet_name, header=None, engine="openpyxl")
    header = xl.iloc[EXCEL_HEADER_ROW].tolist()
    data = xl.iloc[DATA_START_ROW:].copy().reset_index(drop=True)
    data.columns = [str(c) if not pd.isnull(c) else f"col_{i}" for i, c in enumerate(header)]
    return data

def get_risk_free_rate(excel_path):
    try:
        val = pd.read_excel(excel_path, sheet_name="Eonia", header=None, engine="openpyxl").iloc[8, 7]
        if pd.isnull(val):
            return 0.0
        return float(val)
    except Exception as e:
        st.warning(f"Erreur lecture taux sans risque : {e}")
        return 0.0

def detect_frequency(dates):
    if len(dates) < 3: return 252, "quotidienne"
    diffs = pd.Series(dates).sort_values().diff().dropna().dt.days
    median_diff = diffs.median()
    if median_diff <= 2: return 252, "quotidienne"
    elif median_diff <= 8: return 52, "hebdomadaire"
    elif median_diff <= 32: return 12, "mensuelle"
    else: return 252, "quotidienne"

def markowitz_optimisation(prices_df, mode="sharpe", freq=252, max_vol_annual=None, risk_free_rate=0.0):
    returns = prices_df.pct_change().dropna()
    mu = returns.mean()
    S = returns.cov()
    ef = EfficientFrontier(mu, S)
    if mode == "sharpe":
        weights = ef.max_sharpe(risk_free_rate=risk_free_rate)
    elif mode == "vol":
        weights = ef.min_volatility()
    elif mode == "ret":
        target = mu.max() * 0.99
        weights = ef.efficient_return(target_return=target)
    elif mode == "risk":
        if max_vol_annual is None: raise ValueError("max_vol_annual requis")
        max_vol_period = max_vol_annual / np.sqrt(freq)
        weights = ef.efficient_risk(target_volatility=max_vol_period)
    else:
        weights = ef.max_sharpe(risk_free_rate=risk_free_rate)
    cleaned_weights = ef.clean_weights()
    perf = ef.portfolio_performance(verbose=False, risk_free_rate=risk_free_rate)
    return cleaned_weights, (perf[0]*freq, perf[1]*np.sqrt(freq), perf[2]*np.sqrt(freq)), mu, S


st.title("Optimisations sous contraintes")

# --- Sélection de la période pour optimisations ---
df_raw = read_bloomberg_sheet_with_dates(workbook_path, sheet_name=SHEET_NAME)
date_col_name = df_raw.columns[DATE_COL_INDEX]
df_raw[date_col_name] = pd.to_datetime(df_raw[date_col_name], errors='coerce')
df_raw = df_raw.dropna(subset=[date_col_name])

min_date = df_raw[date_col_name].min()
max_date = df_raw[date_col_name].max()
start_date, end_date = st.slider(
    "Sélectionne la période pour l'optimisation",
    min_value=min_date.to_pydatetime(),
    max_value=max_date.to_pydatetime(),
    value=(min_date.to_pydatetime(), max_date.to_pydatetime()),
    format="YYYY-MM-DD",
    key="slider_periode_optim"
)
df_raw = df_raw[(df_raw[date_col_name] >= pd.Timestamp(start_date)) & (df_raw[date_col_name] <= pd.Timestamp(end_date))].copy()

try:
    freq, freq_label = detect_frequency(df_raw[date_col_name])
    st.info(f"Fréquence détectée : {freq_label} ({freq} périodes/an)")
except:
    freq, freq_label = 252, "quotidienne"

all_titles = [c for i, c in enumerate(df_raw.columns) if i != DATE_COL_INDEX]

# --- Section : Optimisation Tracking Error (TE) - CORRIGÉE & ÉLÉGANTE ---
st.header("Optimisation de la Tracking Error (TE) par rapport à un benchmark")
benchmark_title = st.selectbox("Choisir un titre comme benchmark", all_titles)
other_titles = [t for t in all_titles if t != benchmark_title]
te_candidates = st.multiselect("Univers d'investissement pour répliquer le benchmark", other_titles, default=other_titles)

st.markdown("**Note :** Les poids sont optimisés pour minimiser l'écart (MSE) au benchmark.")

if st.button("Lancer optimisation TE"):
    if not te_candidates:
        st.error("Sélectionnez au moins un actif dans l'univers d'investissement.")
    else:
        # Préparation des données
        prices_bench = df_raw[[benchmark_title]].set_index(df_raw[date_col_name])
        prices_others = df_raw[te_candidates].set_index(df_raw[date_col_name])
        
        # Conversion numérique (sécurité)
        prices_bench = prices_bench.apply(pd.to_numeric, errors='coerce')
        prices_others = prices_others.apply(pd.to_numeric, errors='coerce')
        
        returns_bench = prices_bench.pct_change().dropna()
        returns_others = prices_others.pct_change().dropna()
        
        # Alignement des dates
        common_idx = returns_bench.index.intersection(returns_others.index)
        returns_bench = returns_bench.loc[common_idx]
        returns_others = returns_others.loc[common_idx]
        
        X = returns_others.values
        y = returns_bench.values.flatten()
        
        # --- Fonction d'optimisation robuste ---
        def optimize_te(X, y):
            n = X.shape[1]
            # Fonction objectif : MSE * 100000 pour éviter les gradients plats
            def objective(w):
                return np.mean((X @ w - y)**2) * 100000
            
            cons = ({'type': 'eq', 'fun': lambda w: np.sum(w) - 1})
            bounds = tuple((0.0, 1.0) for _ in range(n))
            w0 = np.ones(n) / n
            
            res = minimize(objective, w0, method='SLSQP', bounds=bounds, constraints=cons, options={'ftol': 1e-12, 'maxiter': 2000})
            return res.x

        w_opt = optimize_te(X, y)
        
        # --- Affichage des Poids Optimaux ---
        res_df = pd.DataFrame({
            "Actif": te_candidates,
            "Poids Optimal (%)": w_opt * 100
        }).set_index("Actif")
        
        # On trie par poids décroissant et on colore
        st.subheader("Allocation Optimale")
        st.dataframe(res_df.sort_values("Poids Optimal (%)", ascending=False).style.format("{:.2f}").background_gradient(cmap="Greens"))
        
        # --- Graphique Élégant (Matplotlib) ---
        st.subheader("Comparaison : Benchmark vs Portefeuille Optimisé (Base 100)")
        
        # Calcul des performances cumulées
        port_returns = (returns_others @ w_opt)
        perf_bench = (1 + returns_bench[benchmark_title]).cumprod() * 100
        perf_port = (1 + pd.Series(port_returns, index=returns_bench.index)).cumprod() * 100
        
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.plot(perf_bench.index, perf_bench, label="Benchmark", color="black", linestyle="--", linewidth=1.5)
        ax.plot(perf_port.index, perf_port, label="Portefeuille Optimisé", color="green", linewidth=2)
        
        ax.set_title("Évolution comparée (Base 100)")
        ax.set_ylabel("Valeur")
        ax.legend()
        ax.grid(True, linestyle=":", alpha=0.6)
        
        # Format des dates pour être "élégant"
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
        fig.autofmt_xdate()
        
        st.pyplot(fig)
        plt.close(fig)
        
        # Metrics TE
        te_val = np.std(port_returns - y) * np.sqrt(freq)
        st.info(f"Tracking Error Annualisée atteinte : {te_val*100:.2f}%")


# --- Section : Optimisation Markowitz rendement sous contrainte de volatilité annualisée ---
st.header("Optimisation rendement sous contrainte de volatilité annualisée")
selected_titles_vol = st.multiselect("Titres à inclure (volatilité)", all_titles, default=all_titles, key="vol_titles")
max_vol_annual = st.slider("Volatilité annualisée maximale autorisée (%)", min_value=1.0, max_value=50.0, value=15.0, step=0.1)
if st.button("Lancer optimisation rendement sous contrainte de volatilité annualisée"):
    symbol_cols = selected_titles_vol
    prices_df = df_raw[symbol_cols].copy()
    prices_df.index = pd.to_datetime(df_raw[date_col_name])
    prices_df = prices_df.apply(pd.to_numeric, errors="coerce")
    prices_df = prices_df.dropna(how="any")
    risk_free_rate_annual = get_risk_free_rate(workbook_path)
    risk_free_rate_period = risk_free_rate_annual / freq

    weights, perf, mu, S = markowitz_optimisation(
        prices_df,
        mode="risk",
        freq=freq,
        max_vol_annual=max_vol_annual/100,
        risk_free_rate=risk_free_rate_period
    )

    st.subheader("Poids optimaux (Max rendement sous contrainte de volatilité annualisée)")
    st.dataframe(pd.DataFrame([weights]).T)
    st.bar_chart(pd.DataFrame([weights]).T)
    perf_labels = ["Rendement annualisé", "Volatilité annualisée", "Sharpe ratio"]
    perf_units = ["%", "%", ""]
    perf_values = [perf[0]*100, perf[1]*100, perf[2]]
    col1, col2, col3 = st.columns(3)
    col1.metric(perf_labels[0], f"{perf_values[0]:.2f} {perf_units[0]}")
    col2.metric(perf_labels[1], f"{perf_values[1]:.2f} {perf_units[1]}")
    col3.metric(perf_labels[2], f"{perf_values[2]:.2f}")

# --- Section : Optimisation Markowitz Sharpe ---
st.header("Optimisation Sharpe Ratio")
if st.button("Lancer optimisation Sharpe"):
    date_col_name = df_raw.columns[0]
    symbol_cols = [c for c in df_raw.columns if c != date_col_name]
    prices_df = df_raw[symbol_cols].copy()
    prices_df = prices_df.apply(pd.to_numeric, errors="coerce")
    prices_df = prices_df.dropna(how="any")
    risk_free_rate_annual = get_risk_free_rate(workbook_path)
    risk_free_rate_period = risk_free_rate_annual / freq
    weights, perf, mu, S = markowitz_optimisation(prices_df, mode="sharpe", freq=freq, risk_free_rate=risk_free_rate_period)
    st.subheader("Poids optimaux (Sharpe)")
    st.dataframe(pd.DataFrame([weights]).T)
    st.bar_chart(pd.DataFrame([weights]).T)
    perf_labels = ["Rendement annualisé", "Volatilité annualisée", "Sharpe ratio"]
    perf_units = ["%", "%", ""]
    perf_values = [perf[0]*100, perf[1]*100, perf[2]]
    col1, col2, col3 = st.columns(3)
    col1.metric(perf_labels[0], f"{perf_values[0]:.2f} {perf_units[0]}")
    col2.metric(perf_labels[1], f"{perf_values[1]:.2f} {perf_units[1]}")
    col3.metric(perf_labels[2], f"{perf_values[2]:.2f}")

from pypfopt.expected_returns import mean_historical_return

# --- Section CVaR ---
st.header("Maximisation sous contrainte de CVaR annualisé, calculé sur toute la période")
cvar_alpha = st.slider("Niveau de confiance CVaR (%)", min_value=90, max_value=99, value=95, step=1) / 100
max_cvar = st.slider("CVaR max autorisée (%)", min_value=1.0, max_value=50.0, value=10.0, step=0.1)
max_cvar_annual = max_cvar / 100
max_cvar_period = max_cvar_annual / np.sqrt(freq)

if st.button("Lancer optimisation rendement sous contrainte de CVaR"):
    date_col_name = df_raw.columns[0]
    symbol_cols = [c for c in df_raw.columns if c != date_col_name]
    prices_df = df_raw[symbol_cols].copy()
    prices_df = prices_df.apply(pd.to_numeric, errors="coerce")
    prices_df = prices_df.dropna(axis=1, how="all")
    prices_df = prices_df.dropna(how="any")

    returns1 = prices_df.pct_change().dropna()
    valid_cols = [c for c in returns1.columns if str(c).strip() not in ["0", "col_0", "nan", "NaN"] and not pd.isnull(c)]
    returns1 = returns1.loc[:, valid_cols].copy()
    returns1 = returns1.apply(pd.to_numeric, errors="coerce")
    returns1 = returns1.dropna(axis=1, how="all")
    returns1 = returns1.reset_index(drop=True)

    if not isinstance(returns1, pd.DataFrame):
        st.error("Erreur : returns1 n'est pas un DataFrame. Vérifie la sélection des colonnes.")
    elif returns1.shape[1] < 2:
        st.error(f"La CVaR nécessite au moins 2 colonnes valides.")
    else:
        try:
            mu = mean_historical_return(prices_df, frequency=freq)
            # Pas de risk_free_rate ici pour éviter l'erreur
            ef_cvar = EfficientCVaR(mu, returns1, beta=cvar_alpha)
            weights = ef_cvar.efficient_risk(target_cvar=max_cvar_period)
            cleaned_weights = ef_cvar.clean_weights()
            perf = ef_cvar.portfolio_performance(verbose=False)
            rendement_annualise = perf[0]
            cvar_annualise = perf[1] * np.sqrt(freq)
            st.subheader("Poids optimaux (Max rendement sous contrainte de CVaR)")
            st.dataframe(pd.DataFrame([cleaned_weights]).T)
            st.bar_chart(pd.DataFrame([cleaned_weights]).T)
            perf_labels = ["Rendement annualisé", "CVaR annualisée sur tout la période"]
            perf_units = ["%", "%"]
            perf_values = [rendement_annualise*100, - cvar_annualise*100]
            col1, col2 = st.columns(2)
            col1.metric(perf_labels[0], f"{perf_values[0]:.2f} {perf_units[0]}")
            col2.metric(perf_labels[1], f"{perf_values[1]:.2f} {perf_units[1]}")
        except Exception as e:
            st.error(f"Erreur lors de l'optimisation CVaR : {e}")

    if 'cleaned_weights' in locals():
        port_returns = returns1 @ pd.Series(cleaned_weights)
        window = freq 
        alpha = cvar_alpha

        def rolling_cvar(series, window, alpha):
            return series.rolling(window).apply(lambda x: x[x <= np.quantile(x, 1-alpha)].mean() if len(x) > 0 else np.nan)

        cvar_glissante = rolling_cvar(port_returns, window, alpha)

        st.subheader("CVaR glissante sur 1 an")
        st.line_chart(cvar_glissante * 100)
        st.write(f"CVaR glissante max sur 1 an ({int(alpha*100)}%) : {cvar_glissante.max()*100:.2f} %")
        st.write(f"CVaR glissante min sur 1 an ({int(alpha*100)}%) : {cvar_glissante.min()*100:.2f} %")

# --- Section Reporting Avancé (Original + Fix invisible) ---
st.title("Reporting Avancé")

PORTFOLIO_SHEET_PREFIX = "Portefeuille"
START_ROW = 14
COL_NAMES = ['date', 'valeur', 'perf', 'var99', 'ecart_type', 'cvar99', 'drawdown', 'recovery']

xls = pd.ExcelFile(DEFAULT_WORKBOOK)
portfolio_sheets = [s for s in xls.sheet_names if s.startswith(PORTFOLIO_SHEET_PREFIX)]

data_dict = {var: {} for var in COL_NAMES[1:]}
for sheet in portfolio_sheets:
    df = pd.read_excel(DEFAULT_WORKBOOK, sheet_name=sheet, header=None, usecols="A:H", engine="openpyxl")
    df = df.iloc[START_ROW:, :].reset_index(drop=True)
    df.columns = COL_NAMES
    
    # Fix invisible
    df['date'] = pd.to_datetime(df['date'], errors='coerce')
    for c in COL_NAMES[1:]:
        df[c] = pd.to_numeric(df[c], errors='coerce')

    for var in COL_NAMES[1:]:
        if var in ['var99', 'ecart_type', 'cvar99']:
            first_valid = df[var].first_valid_index()
            if first_valid is not None:
                sub_df = df.loc[first_valid:, ['date', var]].dropna()
            else:
                sub_df = pd.DataFrame(columns=['date', var])
        else:
            sub_df = df[['date', var]].dropna()
        if not sub_df.empty:
            data_dict[var][sheet] = pd.Series(sub_df[var].values, index=pd.to_datetime(sub_df['date']))
        else:
            data_dict[var][sheet] = pd.Series(dtype=float)

color_map = plt.get_cmap('tab10')
default_colors = {pf: '#%02x%02x%02x' % tuple(int(255*x) for x in color_map(idx % 10)[:3]) for idx, pf in enumerate(portfolio_sheets)}

st.sidebar.header("Options graphiques avancées (optionnelles)")
rename_dict_global = {}
color_dict_global = {}
for idx, pf in enumerate(portfolio_sheets):
    rename_dict_global[pf] = st.sidebar.text_input(f"Nom affiché global pour {pf}", value=pf, key=f"global_name_{pf}")
    color_dict_global[pf] = st.sidebar.color_picker(
        f"Couleur pour {pf} (optionnel)", 
        value=default_colors[pf], 
        key=f"global_color_{pf}"
    )

for var in COL_NAMES[1:]:
    with st.expander(f"Options avancées pour {var}", expanded=False):
        rename_dict = {}
        color_dict = {}
        for idx, pf in enumerate(portfolio_sheets):
            rename_dict[pf] = st.text_input(f"Nom affiché pour {pf} ({var})", value=rename_dict_global[pf], key=f"name_{var}_{pf}")
            color_dict[pf] = st.color_picker(f"Couleur pour {pf} ({var})", value=color_dict_global[pf], key=f"color_{var}_{pf}")
        custom_title = st.text_input(f"Titre du graphique {var}", value=f"Comparaison {var}", key=f"title_{var}")
        custom_ylabel = st.text_input(f"Label axe Y {var}", value=var, key=f"ylabel_{var}")
        yscale = st.selectbox(f"Echelle axe Y {var}", ["linear", "log"], key=f"yscale_{var}")
        date_format = st.selectbox(
            f"Format de l'axe X (dates) {var}",
            [("Année", "%Y"), ("Mois", "%Y-%m"), ("Semaine", "%Y-%W"), ("Jour", "%Y-%m-%d")],
            format_func=lambda x: x[0],
            key=f"datefmt_{var}"
        )
        unite = st.text_input(f"Unité affichée (ex: %, base 100)", value="", key=f"unite_{var}")
        ymin = st.number_input(f"Borne Y min {var}", value=float("nan"), format="%.2f", key=f"ymin_{var}")
        ymax = st.number_input(f"Borne Y max {var}", value=float("nan"), format="%.2f", key=f"ymax_{var}")

    st.subheader(f"Comparaison {var}")
    fig, ax = plt.subplots(figsize=(10, 5))
    for idx, (sheet, serie) in enumerate(data_dict[var].items()):
        if not serie.empty:
            user_color = color_dict.get(sheet, None)
            color = user_color if user_color and user_color != default_colors.get(sheet, None) else None
            label = rename_dict.get(sheet, sheet)
            ax.plot(serie.index, serie.values, label=label, color=color)
    ax.set_title(custom_title)
    ax.set_xlabel("Date")
    ax.set_ylabel((custom_ylabel + (f" ({unite})" if unite else "")) if custom_ylabel else var)
    ax.set_yscale(yscale)
    fmt = date_format[1]
    if fmt == "%Y":
        ax.xaxis.set_major_locator(mdates.YearLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    elif fmt == "%Y-%m":
        ax.xaxis.set_major_locator(mdates.MonthLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
    elif fmt == "%Y-%W":
        ax.xaxis.set_major_locator(mdates.WeekdayLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%W"))
    else:
        ax.xaxis.set_major_locator(mdates.AutoDateLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m-%d"))
    if not pd.isna(ymin): ax.set_ylim(bottom=ymin)
    if not pd.isna(ymax): ax.set_ylim(top=ymax)
    ax.legend()
    ax.grid(True)
    fig.autofmt_xdate()
    st.pyplot(fig)
    plt.close(fig)

with st.expander("Options avancées pour graphique combiné valeur base 100 + écart type", expanded=False):
    rename_dict_comb = {}
    color_dict_comb = {}
    for idx, pf in enumerate(portfolio_sheets):
        rename_dict_comb[pf] = st.text_input(f"Nom affiché pour {pf} (combiné)", value=rename_dict_global[pf], key=f"name_comb_{pf}")
        color_dict_comb[pf] = st.color_picker(f"Couleur pour {pf} (combiné)", value=color_dict_global[pf], key=f"color_comb_{pf}")
    custom_title_comb = st.text_input("Titre du graphique combiné", value="Valeur base 100 (courbes) et écart type (aires)", key="title_comb")
    yscale_comb = st.selectbox("Echelle axe Y (valeur)", ["linear", "log"], key="yscale_comb")
    y2scale_comb = st.selectbox("Echelle axe Y (écart type)", ["linear", "log"], key="y2scale_comb")
    date_format_comb = st.selectbox(
        "Format de l'axe X (dates) (combiné)",
        [("Année", "%Y"), ("Mois", "%Y-%m"), ("Semaine", "%Y-%W"), ("Jour", "%Y-%m-%d")],
        format_func=lambda x: x[0],
        key="datefmt_comb"
    )
    unite_comb = st.text_input("Unité affichée (valeur)", value="", key="unite_comb")
    ymin_comb = st.number_input("Borne Y min (valeur)", value=float("nan"), format="%.2f", key="ymin_comb")
    ymax_comb = st.number_input("Borne Y max (valeur)", value=float("nan"), format="%.2f", key="ymax_comb")
    y2min_comb = st.number_input("Borne Y min (écart type)", value=float("nan"), format="%.2f", key="y2min_comb")
    y2max_comb = st.number_input("Borne Y max (écart type)", value=float("nan"), format="%.2f", key="y2max_comb")

st.subheader("Comparaison valeur base 100 (courbes) et écart type (aires, axe secondaire)")
fig, ax1 = plt.subplots(figsize=(12, 6))
ax2 = ax1.twinx()
for idx, pf in enumerate(portfolio_sheets):
    user_color = color_dict_comb[pf]
    color = user_color if user_color and user_color != default_colors[pf] else default_colors[pf]
    val = data_dict['valeur'][pf]
    et = data_dict['ecart_type'][pf]
    name = rename_dict_comb[pf] if rename_dict_comb[pf] != pf else pf
    if not val.empty:
        ax1.plot(val.index, val.values, label=f"{name} (valeur)", color=color, linewidth=2)
    if not et.empty:
        et_clean = et.dropna()
        if not et_clean.empty:
            ax2.fill_between(et_clean.index, 0, et_clean.values.astype(float), label=f"{name} (écart type)", color=color, alpha=0.25)
ax1.set_xlabel("Date")
ax1.set_ylabel("Valeur base 100" + (f" ({unite_comb})" if unite_comb else ""))
ax2.set_ylabel("Ecart type")
ax1.set_title(custom_title_comb)
ax1.set_yscale(yscale_comb)
ax2.set_yscale(y2scale_comb)
if date_format_comb[1] == "%Y":
    ax1.xaxis.set_major_locator(mdates.YearLocator())
    ax1.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
elif date_format_comb[1] == "%Y-%m":
    ax1.xaxis.set_major_locator(mdates.MonthLocator())
    ax1.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
elif date_format_comb[1] == "%Y-%W":
    ax1.xaxis.set_major_locator(mdates.WeekdayLocator())
    ax1.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%W"))
elif date_format_comb[1] == "%Y-%m-%d":
    ax1.xaxis.set_major_locator(mdates.AutoDateLocator())
    ax1.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m-%d"))
if not pd.isna(ymin_comb): ax1.set_ylim(bottom=ymin_comb)
if not pd.isna(ymax_comb): ax1.set_ylim(top=ymax_comb)
if not pd.isna(y2min_comb): ax2.set_ylim(bottom=y2min_comb)
if not pd.isna(y2max_comb): ax2.set_ylim(top=y2max_comb)
ax1.grid(True)
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper left")
fig.autofmt_xdate()
st.pyplot(fig)
plt.close(fig)

with st.expander("Options avancées pour graphique combiné valeur base 100 + VaR 99%", expanded=False):
    rename_dict_var = {}
    color_dict_var = {}
    for idx, pf in enumerate(portfolio_sheets):
        rename_dict_var[pf] = st.text_input(f"Nom affiché pour {pf} (valeur+VaR)", value=rename_dict_global[pf], key=f"name_var_{pf}")
        color_dict_var[pf] = st.color_picker(f"Couleur pour {pf} (valeur+VaR)", value=color_dict_global[pf], key=f"color_var_{pf}")
    custom_title_var = st.text_input("Titre du graphique combiné", value="Valeur base 100 (courbes) et VaR 99% (aires)", key="title_var")
    yscale_var = st.selectbox("Echelle axe Y (valeur)", ["linear", "log"], key="yscale_var")
    y2scale_var = st.selectbox("Echelle axe Y (VaR)", ["linear", "log"], key="y2scale_var")
    date_format_var = st.selectbox(
        "Format de l'axe X (dates) (valeur+VaR)",
        [("Année", "%Y"), ("Mois", "%Y-%m"), ("Semaine", "%Y-%W"), ("Jour", "%Y-%m-%d")],
        format_func=lambda x: x[0],
        key="datefmt_var"
    )
    unite_var = st.text_input("Unité affichée (valeur)", value="", key="unite_var")
    ymin_var = st.number_input("Borne Y min (valeur)", value=float("nan"), format="%.2f", key="ymin_var")
    ymax_var = st.number_input("Borne Y max (valeur)", value=float("nan"), format="%.2f", key="ymax_var")
    y2min_var = st.number_input("Borne Y min (VaR)", value=float("nan"), format="%.2f", key="y2min_var")
    y2max_var = st.number_input("Borne Y max (VaR)", value=float("nan"), format="%.2f", key="y2max_var")

st.subheader("Comparaison valeur base 100 (courbes) et VaR 99% (aires, axe secondaire)")
fig, ax1 = plt.subplots(figsize=(12, 6))
ax2 = ax1.twinx()
export_df_var = pd.DataFrame()
for idx, pf in enumerate(portfolio_sheets):
    user_color = color_dict_var[pf]
    color = user_color if user_color and user_color != default_colors[pf] else default_colors[pf]
    val = data_dict['valeur'][pf]
    var99 = data_dict['var99'][pf]
    name = rename_dict_var[pf] if rename_dict_var[pf] != pf else pf
    if not val.empty:
        ax1.plot(val.index, val.values, label=f"{name} (valeur)", color=color, linewidth=2)
        export_df_var[f"{name} (valeur)"] = val
    if not var99.empty:
        var_clean = var99.dropna()
        if not var_clean.empty:
            ax2.fill_between(var_clean.index, 0, var_clean.values.astype(float), label=f"{name} (VaR 99%)", color=color, alpha=0.25)
            export_df_var[f"{name} (VaR 99%)"] = var_clean
ax1.set_xlabel("Date")
ax1.set_ylabel("Valeur base 100" + (f" ({unite_var})" if unite_var else ""))
ax2.set_ylabel("VaR 99% (négatif)")
ax1.set_title(custom_title_var)
ax1.set_yscale(yscale_var)
ax2.set_yscale(y2scale_var)
if date_format_var[1] == "%Y":
    ax1.xaxis.set_major_locator(mdates.YearLocator())
    ax1.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
elif date_format_var[1] == "%Y-%m":
    ax1.xaxis.set_major_locator(mdates.MonthLocator())
    ax1.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
elif date_format_var[1] == "%Y-%W":
    ax1.xaxis.set_major_locator(mdates.WeekdayLocator())
    ax1.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%W"))
elif date_format_var[1] == "%Y-%m-%d":
    ax1.xaxis.set_major_locator(mdates.AutoDateLocator())
    ax1.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m-%d"))
if not pd.isna(ymin_var): ax1.set_ylim(bottom=ymin_var)
if not pd.isna(ymax_var): ax1.set_ylim(top=ymax_var)
if not pd.isna(y2min_var): ax2.set_ylim(bottom=y2min_var)
if not pd.isna(y2max_var): ax2.set_ylim(top=y2max_var)
ax2.invert_yaxis()
ax1.grid(True)
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper left")
fig.autofmt_xdate()
st.pyplot(fig)
plt.close(fig)

with st.expander("Options avancées pour graphique combiné valeur base 100 + CVaR 99%", expanded=False):
    rename_dict_cvar = {}
    color_dict_cvar = {}
    for idx, pf in enumerate(portfolio_sheets):
        rename_dict_cvar[pf] = st.text_input(f"Nom affiché pour {pf} (valeur+CVaR)", value=rename_dict_global[pf], key=f"name_cvar_{pf}")
        color_dict_cvar[pf] = st.color_picker(f"Couleur pour {pf} (valeur+CVaR)", value=color_dict_global[pf], key=f"color_cvar_{pf}")
    custom_title_cvar = st.text_input("Titre du graphique combiné", value="Valeur base 100 (courbes) et CVaR 99% (aires)", key="title_cvar")
    yscale_cvar = st.selectbox("Echelle axe Y (valeur)", ["linear", "log"], key="yscale_cvar")
    y2scale_cvar = st.selectbox("Echelle axe Y (CVaR)", ["linear", "log"], key="y2scale_cvar")
    date_format_cvar = st.selectbox(
        "Format de l'axe X (dates) (valeur+CVaR)",
        [("Année", "%Y"), ("Mois", "%Y-%m"), ("Semaine", "%Y-%W"), ("Jour", "%Y-%m-%d")],
        format_func=lambda x: x[0],
        key="datefmt_cvar"
    )
    unite_cvar = st.text_input("Unité affichée (valeur)", value="", key="unite_cvar")
    ymin_cvar = st.number_input("Borne Y min (valeur)", value=float("nan"), format="%.2f", key="ymin_cvar")
    ymax_cvar = st.number_input("Borne Y max (valeur)", value=float("nan"), format="%.2f", key="ymax_cvar")
    y2min_cvar = st.number_input("Borne Y min (CVaR)", value=float("nan"), format="%.2f", key="y2min_cvar")
    y2max_cvar = st.number_input("Borne Y max (CVaR)", value=float("nan"), format="%.2f", key="y2max_cvar")

st.subheader("Comparaison valeur base 100 (courbes) et CVaR 99% (aires, axe secondaire)")
fig, ax1 = plt.subplots(figsize=(12, 6))
ax2 = ax1.twinx()
export_df_cvar = pd.DataFrame()
for idx, pf in enumerate(portfolio_sheets):
    user_color = color_dict_cvar[pf]
    color = user_color if user_color and user_color != default_colors[pf] else default_colors[pf]
    val = data_dict['valeur'][pf]
    cvar99 = data_dict['cvar99'][pf]
    name = rename_dict_cvar[pf] if rename_dict_cvar[pf] != pf else pf
    if not val.empty:
        ax1.plot(val.index, val.values, label=f"{name} (valeur)", color=color, linewidth=2)
        export_df_cvar[f"{name} (valeur)"] = val
    if not cvar99.empty:
        cvar_clean = cvar99.dropna()
        if not cvar_clean.empty:
            ax2.fill_between(cvar_clean.index, 0, cvar_clean.values.astype(float), label=f"{name} (CVaR 99%)", color=color, alpha=0.25)
            export_df_cvar[f"{name} (CVaR 99%)"] = cvar_clean
ax1.set_xlabel("Date")
ax1.set_ylabel("Valeur base 100" + (f" ({unite_cvar})" if unite_cvar else ""))
ax2.set_ylabel("CVaR 99% (négatif)")
ax1.set_title(custom_title_cvar)
ax1.set_yscale(yscale_cvar)
ax2.set_yscale(y2scale_cvar)
if date_format_cvar[1] == "%Y":
    ax1.xaxis.set_major_locator(mdates.YearLocator())
    ax1.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
elif date_format_cvar[1] == "%Y-%m":
    ax1.xaxis.set_major_locator(mdates.MonthLocator())
    ax1.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
elif date_format_cvar[1] == "%Y-%W":
    ax1.xaxis.set_major_locator(mdates.WeekdayLocator())
    ax1.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%W"))
elif date_format_cvar[1] == "%Y-%m-%d":
    ax1.xaxis.set_major_locator(mdates.AutoDateLocator())
    ax1.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m-%d"))
if not pd.isna(ymin_cvar): ax1.set_ylim(bottom=ymin_cvar)
if not pd.isna(ymax_cvar): ax1.set_ylim(top=ymax_cvar)
if not pd.isna(y2min_cvar): ax2.set_ylim(bottom=y2min_cvar)
if not pd.isna(y2max_cvar): ax2.set_ylim(top=y2max_cvar)
ax2.invert_yaxis()
ax1.grid(True)
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper left")
fig.autofmt_xdate()
st.pyplot(fig)
plt.close(fig)