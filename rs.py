import streamlit as st
import pandas as pd
import io

# Étape 1 : Télécharger le fichier et choisir la feuille
def load_excel():
    uploaded_file = st.file_uploader("Téléchargez votre fichier Excel", type=["xlsx", "xls"])
    if uploaded_file:
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = st.selectbox("Sélectionnez une feuille", xls.sheet_names)
        df = pd.read_excel(xls, sheet_name=sheet_name)
        return df
    return None

# Étape 2 : Filtrer les données sur les valeurs spécifiques
def filter_data(df):
    # Liste des valeurs à filtrer
    valeurs_cibles = [
        'IGD HORS IDF 1 REP.', 'IGD HORS IDF 2 REP.', 'IGD HORS IDF LOG. + 1 REP.',
        'IGD HORS IDF LOG. + 2 REP.', 'IGD IDF 1 REP.', 'IGD IDF 2 REP.',
        'IGD IDF LOG. + 1 REP.', 'IGD IDF LOG. + 2 REP.', 'IPD Repas hors locaux (TX)',
        'Repas pris restaurant', 'IPD Ticket restaurant', 'Panier Sedentaire (TX)'
    ]

    # Liste des valeurs interdites dans la colonne 'CODE CRA'
    valeurs_interdites_code_cra = [
        'j_B0534_Paie', 'j_B0670_Paie', 'j_BDI09_Paie', 'j_BDI13_Pai3',
        'j_BDI19_Paie', 'j_BNU24_Paie', 'j_BNU28_Paie', 'j_BNU37_Paie',
        'j_BNU38_Paie', 'j_BNU40_Paie', 'j_BSA21_Paie', 'j_BTICK_paie',
        'j_WIRRE_paie'
    ]

    colonne_reference = st.selectbox("Sélectionnez la colonne de filtrage", df.columns)
    df_filtered = df[df[colonne_reference].isin(valeurs_cibles)]
    
    # Filtrer la colonne 'CUMUL' pour ne pas prendre les lignes où CUMUL == '0'
    if 'CUMUL' in df.columns:
        df_filtered = df_filtered[~df_filtered['CUMUL'].isin([0, '0'])]

    # Supprimer les lignes où 'CODE CRA' contient une valeur interdite
    if 'CODE CRA' in df.columns:
        df_filtered = df_filtered[~df_filtered['CODE CRA'].isin(valeurs_interdites_code_cra)]

    # Filtre de dates : Ajout du filtre basé sur la colonne 'DATE DEBUT'
    if 'DATE DEBUT' in df.columns:
        df['DATE DEBUT'] = pd.to_datetime(df['DATE DEBUT'], errors='coerce')
        min_date = df['DATE DEBUT'].min()
        max_date = df['DATE DEBUT'].max()
        
        start_date, end_date = st.date_input(
            "Sélectionnez une période", 
            value=(min_date, max_date), 
            min_value=min_date, 
            max_value=max_date
        )

        df_filtered = df_filtered[(df['DATE DEBUT'] >= pd.to_datetime(start_date)) & (df['DATE DEBUT'] <= pd.to_datetime(end_date))]

    return df_filtered

# Étape 3 : Détection des doublons de matricules pour une même date
def detect_duplicates(df):
    col_matricule = st.selectbox("Sélectionnez la colonne des matricules", df.columns)
    col_date = st.selectbox("Sélectionnez la colonne des dates", df.columns)
    col_nom = st.selectbox("Sélectionnez la colonne du nom", df.columns)
    col_prenom = st.selectbox("Sélectionnez la colonne du prénom", df.columns)
    
    df[col_date] = pd.to_datetime(df[col_date], errors='coerce')
    df[col_date] = df[col_date].dt.strftime('%d/%m/%Y')  # Format : 'DD/MM/YYYY'
    
    duplicate_df = df[df.duplicated(subset=[col_matricule, col_date], keep=False)]
    
    if not duplicate_df.empty:
        st.write("### Matricules en double pour la même date", unsafe_allow_html=True)
        st.dataframe(duplicate_df[[col_date, col_matricule, col_nom, col_prenom, 'ACTIVITE']])
        
        output = io.BytesIO()
        duplicate_df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.download_button(
            label="Exporter les doublons en Excel",
            data=output,
            file_name="doublons_detectes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.success("Aucun doublon trouvé.", icon="✅")

# Interface principale de l'application
def main():
    st.markdown("<h1 style='text-align: center; color: #0066cc;'>Détection des doublons de matricules</h1>", unsafe_allow_html=True)
    
    df = load_excel()
    if df is not None:
        st.subheader("📊 Filtrer les données")
        
        # Filtrer les données selon les critères donnés
        df_filtered = filter_data(df)

        # Vérifier si la colonne 'ACTIVITE' existe
        if 'ACTIVITE' in df.columns:
            df_filtered['ACTIVITE'] = df['ACTIVITE']
        
        # Affichage du DataFrame filtré avec la colonne 'ACTIVITE'
        st.write("### Résultats filtrés", unsafe_allow_html=True)
        st.dataframe(df_filtered)
        
        # Détecter les doublons dans le DataFrame filtré
        st.write("### Détection des doublons", unsafe_allow_html=True)
        detect_duplicates(df_filtered)

if __name__ == "__main__":
    main()
