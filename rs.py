import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime, timedelta

# Configuration de la page
st.set_page_config(
    page_title="Détection des doublons",
    page_icon="📊",
    layout="wide"
)

# Styles CSS simples
st.markdown("""
<style>
    .header { font-size: 26px; font-weight: bold; margin-bottom: 15px; }
    .subheader { font-size: 20px; font-weight: bold; margin: 10px 0; }
</style>
""", unsafe_allow_html=True)

# Fonction pour créer des données d'exemple
def create_example_data():
    """Crée un jeu de données d'exemple."""
    # Valeurs cibles pour le filtrage
    valeurs_cibles = [
        'IGD HORS IDF 1 REP.', 'IGD HORS IDF 2 REP.', 'IGD IDF 1 REP.', 
        'IGD IDF 2 REP.', 'IPD Ticket restaurant'
    ]
    
    # Créer un DataFrame d'exemple
    data = {
        'MATRICULE': ['MAT001', 'MAT002', 'MAT003', 'MAT001', 'MAT004', 'MAT005'],
        'DATE DEBUT': ['2025-01-15', '2025-01-20', '2025-02-05', '2025-01-15', '2025-02-10', '2025-02-15'],
        'NOM': ['Dupont', 'Martin', 'Durand', 'Dupont', 'Bernard', 'Petit'],
        'PRENOM': ['Jean', 'Marie', 'Pierre', 'Jean', 'Sophie', 'Thomas'],
        'ACTIVITE': ['Production', 'Commercial', 'Technique', 'Production', 'Administratif', 'Production'],
        'REFERENCE': ['IGD IDF 1 REP.', 'IGD HORS IDF 2 REP.', 'IPD Ticket restaurant', 
                      'IGD IDF 1 REP.', 'IGD HORS IDF 1 REP.', 'IGD IDF 2 REP.'],
        'CUMUL': [25, 30, 0, 25, 42, 18]
    }
    
    df = pd.DataFrame(data)
    df['DATE DEBUT'] = pd.to_datetime(df['DATE DEBUT'])
    return df

# Étape 1 : Télécharger le fichier et choisir la feuille
def load_excel():
    """Charge un fichier Excel avec prévisualisation des données."""
    st.markdown('<div class="header">1. Téléchargement du fichier</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        uploaded_file = st.file_uploader("Téléchargez votre fichier Excel", type=["xlsx", "xls"])
    
    with col2:
        st.write(" ")  # Espacement
        if st.button("Données d'exemple"):
            return create_example_data()
    
    if uploaded_file:
        try:
            xls = pd.ExcelFile(uploaded_file)
            sheet_name = st.selectbox("Sélectionnez une feuille", xls.sheet_names)
            
            df = pd.read_excel(xls, sheet_name=sheet_name)
            
            # Prévisualisation des données
            st.subheader("Aperçu des données")
            st.dataframe(df.head())
            
            # Informations de base
            col1, col2 = st.columns(2)
            col1.metric("Nombre de lignes", df.shape[0])
            col2.metric("Nombre de colonnes", df.shape[1])
            
            return df
            
        except Exception as e:
            st.error(f"Erreur lors du chargement du fichier: {str(e)}")
            return None
    return None

# Étape 2 : Filtrer les données
def filter_data(df):
    """Filtre les données selon les critères définis."""
    st.markdown('<div class="header">2. Filtrage des données</div>', unsafe_allow_html=True)
    
    # Valeurs cibles par défaut
    valeurs_cibles = [
        'IGD HORS IDF 1 REP.', 'IGD HORS IDF 2 REP.', 'IGD HORS IDF LOG. + 1 REP.',
        'IGD HORS IDF LOG. + 2 REP.', 'IGD IDF 1 REP.', 'IGD IDF 2 REP.',
        'IGD IDF LOG. + 1 REP.', 'IGD IDF LOG. + 2 REP.', 'IPD Repas hors locaux (TX)',
        'Repas pris restaurant', 'IPD Ticket restaurant', 'Panier Sedentaire (TX)'
    ]
    
    # Mise en page en deux colonnes
    col1, col2 = st.columns(2)
    
    with col1:
        # Sélection de la colonne de référence
        colonne_reference = st.selectbox("Sélectionnez la colonne de filtrage", df.columns)
        
        # Affichage des valeurs uniques disponibles
        unique_values = df[colonne_reference].unique().tolist()
        
        # Option pour utiliser les valeurs par défaut ou personnalisées
        use_default = st.checkbox("Utiliser les valeurs cibles par défaut", value=True)
        
        if use_default:
            # Filtrer les valeurs qui existent dans les données
            valeurs_filtrage = [val for val in valeurs_cibles if val in unique_values]
            if not valeurs_filtrage:
                st.warning("Aucune valeur cible par défaut n'existe dans les données.")
                valeurs_filtrage = unique_values[:5]  # Prendre les 5 premières valeurs
        else:
            # Sélection manuelle des valeurs
            valeurs_filtrage = st.multiselect(
                "Sélectionnez les valeurs à inclure",
                options=unique_values
            )
    
    with col2:
        # Filtre CUMUL
        filter_cumul = st.checkbox("Filtrer sur CUMUL (supprimer les valeurs 0)", value=True)
        
        # Filtre de dates
        filter_dates = st.checkbox("Filtrer par dates", value=True)
        
        if filter_dates and 'DATE DEBUT' in df.columns:
            # Convertir la colonne DATE DEBUT en format datetime si nécessaire
            if not pd.api.types.is_datetime64_any_dtype(df['DATE DEBUT']):
                df['DATE DEBUT'] = pd.to_datetime(df['DATE DEBUT'], errors='coerce')
            
            # Trouver la plage de dates disponibles
            min_date = df['DATE DEBUT'].min().date()
            max_date = df['DATE DEBUT'].max().date()
            
            # Sélection de la période
            date_range = st.date_input(
                "Sélectionnez une période",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date
            )
            
            if len(date_range) == 2:
                start_date, end_date = date_range
            else:
                start_date, end_date = min_date, max_date
    
    # Application des filtres
    # 1. Filtre sur les valeurs cibles
    if valeurs_filtrage:
        df_filtered = df[df[colonne_reference].isin(valeurs_filtrage)]
    else:
        df_filtered = df.copy()
        st.warning("Aucune valeur sélectionnée pour le filtrage.")
    
    # 2. Filtre CUMUL
    if filter_cumul and 'CUMUL' in df.columns:
        df_filtered = df_filtered[df_filtered['CUMUL'] != 0]
    
    # 3. Filtre dates
    if filter_dates and 'DATE DEBUT' in df.columns:
        df_filtered = df_filtered[
            (df_filtered['DATE DEBUT'] >= pd.to_datetime(start_date)) & 
            (df_filtered['DATE DEBUT'] <= pd.to_datetime(end_date))
        ]
    
    # Statistiques simples
    st.metric("Nombre d'enregistrements après filtrage", df_filtered.shape[0])
    
    # Afficher les données filtrées
    st.subheader("Données filtrées")
    st.dataframe(df_filtered)
    
    return df_filtered

# Étape 3 : Détection des doublons
def detect_duplicates(df):
    """Détecte les doublons de matricules pour une même date."""
    st.markdown('<div class="header">3. Détection des doublons</div>', unsafe_allow_html=True)
    
    # Sélection des colonnes pour la détection
    col1, col2 = st.columns(2)
    
    with col1:
        col_matricule = st.selectbox(
            "Sélectionnez la colonne des matricules", 
            df.columns,
            index=list(df.columns).index('MATRICULE') if 'MATRICULE' in df.columns else 0
        )
        
        col_date = st.selectbox(
            "Sélectionnez la colonne des dates", 
            df.columns,
            index=list(df.columns).index('DATE DEBUT') if 'DATE DEBUT' in df.columns else 0
        )
    
    with col2:
        col_nom = st.selectbox(
            "Sélectionnez la colonne du nom", 
            df.columns,
            index=list(df.columns).index('NOM') if 'NOM' in df.columns else 0
        )
        
        col_prenom = st.selectbox(
            "Sélectionnez la colonne du prénom", 
            df.columns,
            index=list(df.columns).index('PRENOM') if 'PRENOM' in df.columns else 0
        )
    
    # Convertir la colonne des dates en format datetime si nécessaire
    if not pd.api.types.is_datetime64_any_dtype(df[col_date]):
        df[col_date] = pd.to_datetime(df[col_date], errors='coerce')
    
    # Formater la colonne des dates pour affichage
    df['date_affichage'] = df[col_date].dt.strftime('%d/%m/%Y')
    
    # Détecter les doublons
    duplicate_df = df[df.duplicated(subset=[col_matricule, col_date], keep=False)]
    
    if not duplicate_df.empty:
        # Afficher les doublons
        st.warning(f"🚨 {len(duplicate_df)} enregistrements en double détectés!")
        
        # Colonnes à afficher
        display_cols = ['date_affichage', col_matricule, col_nom, col_prenom]
        
        # Ajouter la colonne ACTIVITE si elle existe
        if 'ACTIVITE' in df.columns:
            display_cols.append('ACTIVITE')
        
        # Afficher le tableau des doublons
        st.subheader("Enregistrements en double")
        st.dataframe(duplicate_df[display_cols])
        
        # Visualisation simple des doublons
        if len(duplicate_df) > 1:
            st.subheader("Visualisation des doublons")
            
            # Graphique par mois si suffisamment de données
            duplicate_df['month'] = duplicate_df[col_date].dt.strftime('%Y-%m')
            date_counts = duplicate_df['month'].value_counts().reset_index()
            date_counts.columns = ['Mois', 'Nombre']
            
            fig = px.bar(
                date_counts,
                x='Mois',
                y='Nombre',
                title="Doublons par mois",
                height=400
            )
            
            st.plotly_chart(fig, use_container_width=True)
        
        # Export des données
        output = io.BytesIO()
        duplicate_df[display_cols].to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        
        st.download_button(
            label="Télécharger les doublons (Excel)",
            data=output,
            file_name="doublons_detectes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.success("✅ Aucun doublon trouvé!")

# Interface principale de l'application
def main():
    st.title("Détection des doublons de matricules")
    
    # Étape 1: Chargement des données
    df = load_excel()
    
    if df is not None:
        # Étape 2: Filtrage des données
        df_filtered = filter_data(df)
        
        # Étape 3: Détection des doublons
        detect_duplicates(df_filtered)

if __name__ == "__main__":
    main()
