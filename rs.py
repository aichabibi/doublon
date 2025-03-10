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
    valeurs_cibles = [
        'IGD HORS IDF 1 REP.', 'IGD HORS IDF 2 REP.', 'IGD HORS IDF LOG. + 1 REP.',
        'IGD HORS IDF LOG. + 2 REP.', 'IGD IDF 1 REP.', 'IGD IDF 2 REP.',
        'IGD IDF LOG. + 1 REP.', 'IGD IDF LOG. + 2 REP.', 'IPD Repas hors locaux (TX)',
        'Repas pris restaurant', 'IPD Ticket restaurant', 'Panier Sedentaire (TX)'
    ]
    colonne_reference = st.selectbox("Sélectionnez la colonne de filtrage", df.columns)
    df_filtered = df[df[colonne_reference].isin(valeurs_cibles)]
    
    # Filtre de dates : Ajout du filtre basé sur la colonne 'DATE DEBUT'
    if 'DATE DEBUT' in df.columns:  # Vérifiez si la colonne DATE DEBUT existe dans votre fichier
        # Convertir la colonne 'DATE DEBUT' en format datetime si ce n'est pas déjà fait
        df['DATE DEBUT'] = pd.to_datetime(df['DATE DEBUT'], errors='coerce')
        
        # Trouver la date minimale et maximale pour la plage de dates
        min_date = df['DATE DEBUT'].min()
        max_date = df['DATE DEBUT'].max()

        # Ajout du filtre de plage de dates avec un calendrier interactif
        start_date, end_date = st.date_input(
            "Sélectionnez une période", 
            value=(min_date, max_date), 
            min_value=min_date, 
            max_value=max_date
        )

        # Filtrer le DataFrame en fonction de la période sélectionnée
        df_filtered = df_filtered[(df['DATE DEBUT'] >= pd.to_datetime(start_date)) & (df['DATE DEBUT'] <= pd.to_datetime(end_date))]

    # Retourner les données filtrées sans modifier la structure initiale
    return df_filtered

# Étape 3 : Détection des doublons de matricules pour une même date
def detect_duplicates(df):
    col_matricule = st.selectbox("Sélectionnez la colonne des matricules", df.columns)
    col_date = st.selectbox("Sélectionnez la colonne des dates", df.columns)
    col_nom = st.selectbox("Sélectionnez la colonne du nom", df.columns)
    col_prenom = st.selectbox("Sélectionnez la colonne du prénom", df.columns)
    
    # Convertir la colonne des dates en format datetime
    df[col_date] = pd.to_datetime(df[col_date], errors='coerce')  # Convertir en format date
    
    # Formater la colonne des dates pour afficher au format français (jj/mm/aaaa)
    df[col_date] = df[col_date].dt.strftime('%d/%m/%Y')  # Format : 'DD/MM/YYYY'
    
    # Détecter les doublons en fonction du matricule et de la date
    duplicate_df = df[df.duplicated(subset=[col_matricule, col_date], keep=False)]
    
    if not duplicate_df.empty:
        # Afficher le tableau des doublons avec les colonnes souhaitées, y compris 'ACTIVITE'
        st.write("### Matricules en double pour la même date")
        st.dataframe(duplicate_df[[col_date, col_matricule, col_nom, col_prenom, 'ACTIVITE']])
        
        # Créer un fichier Excel à partir du DataFrame des doublons
        output = io.BytesIO()
        duplicate_df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        # Bouton de téléchargement pour le fichier Excel
        st.download_button(
            label="Exporter les doublons en Excel",
            data=output,
            file_name="doublons_detectes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.success("Aucun doublon trouvé.")

# Interface principale de l'application
def main():
    st.title("Détection des doublons de matricules")
    df = load_excel()
    if df is not None:
        # Filtrer les données selon les critères donnés
        df_filtered = filter_data(df)
        
        # Vérifier si la colonne 'ACTIVITE' existe, et l'ajouter à l'affichage sans duplication
        if 'ACTIVITE' in df.columns:
            # Ajouter 'ACTIVITE' à l'affichage des données filtrées
            df_filtered['ACTIVITE'] = df['ACTIVITE']
        
        # Affichage du DataFrame filtré avec la colonne 'ACTIVITE'
        st.dataframe(df_filtered)
        
        # Détecter les doublons dans le DataFrame filtré, avec 'ACTIVITE' incluse
        detect_duplicates(df_filtered)

if __name__ == "__main__":
    main()
