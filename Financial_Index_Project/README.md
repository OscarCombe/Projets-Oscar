
# Interactive Financial Index Builder 📊

## Description
Interactive Financial Index Builder est une application interactive développée avec **Python** et **Streamlit**. Elle permet d'analyser et de créer des indices financiers personnalisés basés sur des données sectorielles et fondamentales. L'objectif est d'explorer les dynamiques de marché et de tester des stratégies d'investissement.

Cliquer ici : 
https://financial-index-builder-ocombe.streamlit.app/
---

## Fonctionnalités principales
- **Analyse sectorielle** : Filtrage des entreprises par secteurs et sous-secteurs (SPX et SXXP).  
- **Création d'indices** : Génération d'indices sectoriels, Momentum, ou basés sur la solidité financière.  
- **Benchmarking** : Comparaison des indices personnalisés avec des benchmarks globaux comme SPX et SXXP.  
- **Visualisation interactive** : Graphiques et analyses statistiques pour suivre les performances des indices.  
- **Exportation de résultats** : Téléchargement des analyses et graphiques sous forme de fichiers.

---

## Installation
### Prérequis
- Python 3.7 ou plus récent
- Librairies nécessaires :
  ```bash
  pip install streamlit pandas matplotlib seaborn numpy
  ```

### Lancement de l'application
1. Les données nécessaires au projet sont disponibles ici :  
   [Télécharger les données](https://docs.google.com/uc?id=1aTZh6I3Xe2kzgXlkS9520-ThuYDBVYFi&export=download)  

2. Placez le fichier **Data projet indices python.xlsx** dans le même répertoire que le fichier Python.  

3. Lancez l'application avec la commande suivante :  
   ```bash
   streamlit run Interactive_Financial_Index_Builder.py
   ```

## Utilisation
1. **Analyse sectorielle** : Sélectionnez un secteur et des sous-secteurs via le panneau latéral.  
2. **Création d'indices** :
   - Indices sectoriels : Construction basée sur les entreprises filtrées.  
   - Indices Momentum : Identifiez les entreprises à forte dynamique.  
   - Indices basés sur la solidité financière : Analyse des fondamentaux (Price-to-Book, PER, etc.).
3. **Visualisation et téléchargement** : Comparez les indices aux benchmarks et exportez les résultats.

---

## Fichiers importants
- `Data projet indices python.xlsx` : Fichier contenant les données nécessaires (indices, forex, membres, etc.).
- `temp_reports/` : Répertoire temporaire où sont stockés les graphiques et analyses.

---

## Contribution
Les contributions sont les bienvenues ! Veuillez soumettre une pull request ou ouvrir une issue pour toute suggestion ou correction.

---

## Licence
Ce projet est sous licence **MIT**. Vous êtes libre de l'utiliser, de le modifier et de le redistribuer.
