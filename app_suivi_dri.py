import streamlit as st
import pandas as pd
import geopandas as gpd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from datetime import datetime
import shutil
import zipfile
import os
import tempfile
from io import BytesIO
import requests
import base64
import time
from openpyxl.styles import PatternFill, Font
import plotly.express as px


# --- Fonction pour push sur GitHub ---
def push_to_github(token, repo, path_in_repo, local_filepath, commit_message="Mise à jour fichier DRI"):
    api_url = f"https://api.github.com/repos/{repo}/contents/{path_in_repo}"
    headers = {"Authorization": f"token {token}"}

    response = requests.get(api_url, headers=headers)
    if response.status_code == 200:
        sha = response.json()["sha"]
    else:
        sha = None

    with open(local_filepath, "rb") as f:
        content = f.read()
        encoded_content = base64.b64encode(content).decode("utf-8")

    data = {
        "message": commit_message,
        "content": encoded_content,
        "branch": "main"
    }

    if sha:
        data["sha"] = sha

    put_response = requests.put(api_url, headers=headers, json=data)
    return put_response.status_code, put_response.json()

# --- Fonction d'appel simplifiée ---
def try_push_to_github():
    try:
        token = st.secrets["GITHUB_TOKEN"]
        repo = "Ali-geomarket/fiche-dri-app"
        path_in_repo = "Suivi_demandes_AUTOMATISATION.xlsx"
        status, result = push_to_github(token, repo, path_in_repo, suivi_file_path)
        if status in [200, 201]:
            st.success("Fichier synchronisé avec GitHub")
        else:
            st.warning(f"Push GitHub échoué : {result}")
    except Exception as e:
        st.warning(f"Erreur lors du push GitHub : {e}")

# --- Authentification ---
USERS = {"sg": "dri", "ps": "dri"}

st.set_page_config(page_title="Fiche DRI & TCD MA", layout="wide")

# --- Session State ---
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "main_page" not in st.session_state:
    st.session_state["main_page"] = "login"
if "ligne_temporaire" not in st.session_state:
    st.session_state["ligne_temporaire"] = None

# --- Constantes ---
COLONNES = [
    "DATE RECEPTION", "RESEAU", "RESPONSABLE PROD", "COMMERCIAL", "PROJET", "TYPE DE DEMANDE",
    "COUT EXTENSION", "COUT GLOBAL PROJET", "OPERATEUR", "GAIN DRI", "ROI", "NB CLIENTS AMORTISSEMENT",
    "COMMANDE", "DATE TRAITEMENT", "DELAI TRAITEMENT", "ETAT GEOMARKETING", "RESP GEOMARKET",
    "CONCLUSION", "COMMENTAIRE"
]

suivi_file_path = "Suivi_demandes_AUTOMATISATION.xlsx"
temp_export_path = "Suivi_demandes_EXPORT.xlsx"

# -------------------------
# PAGE : CONNEXION
# -------------------------
if not st.session_state["authenticated"]:
    st.title("Connexion requise")
    with st.form("login_form"):
        username = st.text_input("Nom d'utilisateur")
        password = st.text_input("Mot de passe", type="password")
        if st.form_submit_button("Se connecter"):
            if username in USERS and USERS[username] == password:
                st.session_state["authenticated"] = True
                st.session_state["user"] = username
                st.session_state["main_page"] = "home"
                st.session_state["current_section"] = "visualisation_modif" 
                st.rerun()
            else:
                st.error("Identifiants incorrects")

# -------------------------
# INTERFACE PRINCIPALE APRÈS CONNEXION
# -------------------------
elif st.session_state["authenticated"]:
    st.title("Application Fiche DRI & TCD MA")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        if st.button("Visualisation globale & Modification"):
            st.session_state["current_section"] = "visualisation_modif"
    with col2:
        if st.button("Ajouter une nouvelle fiche DRI"):
            st.session_state["current_section"] = "ajout"
            st.session_state["ligne_temporaire"] = None
    with col3:
        if st.button("TCD MA"):
            st.session_state["current_section"] = "tcd_ma"
    with col4:
        if st.button("Analyse des données"):
            st.session_state["current_section"] = "analyse_donnees"

    # Ensuite on affiche la section choisie :
    section = st.session_state.get("current_section")

    if section == "visualisation_modif":
        st.subheader("Visualisation & Modification du fichier Excel")
    
        if "show_full_df" not in st.session_state:
            st.session_state["show_full_df"] = False
    
        try:
            df = pd.read_excel(suivi_file_path, engine="openpyxl")
            df.columns = [col.strip().upper() for col in df.columns]
            df = df.reindex(columns=COLONNES)
    
            df["DATE RECEPTION"] = pd.to_datetime(df["DATE RECEPTION"], errors="coerce", dayfirst=True)
            df["DATE TRAITEMENT"] = pd.to_datetime(df["DATE TRAITEMENT"], errors="coerce", dayfirst=True)
    
            toggle_label = "Afficher tout le fichier" if not st.session_state["show_full_df"] else "Afficher un extrait"
            toggle_clicked = st.button(toggle_label)
            if toggle_clicked:
                st.session_state["show_full_df"] = not st.session_state["show_full_df"]
                st.rerun()
    
            if st.session_state["show_full_df"]:
                display_df = df.copy()
            else:
                display_df = df.tail(15).copy()
    
            edited_df = st.data_editor(display_df, use_container_width=True, num_rows="dynamic")
    
            if not st.session_state["show_full_df"]:
                df_update = df.copy()
                df_tail_idx = df.tail(15).index
                for i, idx in enumerate(df_tail_idx):
                    df_update.loc[idx] = edited_df.iloc[i]
            else:
                df_update = edited_df.copy()
    
            if st.button("Enregistrer les modifications"):
                try:
                    full_df = pd.read_excel(suivi_file_path, engine="openpyxl")
                    full_df.columns = [col.strip().upper() for col in full_df.columns]
                    full_df = full_df.reindex(columns=COLONNES)
    
                    if not st.session_state["show_full_df"]:
                        tail_indices = full_df.tail(15).index
                        for i, idx in enumerate(tail_indices):
                            full_df.loc[idx] = edited_df.iloc[i]
                    else:
                        full_df = edited_df.copy()
    
                    wb = load_workbook(suivi_file_path)
                    ws = wb.active
                    ws.delete_rows(2, ws.max_row)
    
                    for _, row in full_df.iterrows():
                        cleaned_row = [None if pd.isna(cell) else cell for cell in row]
                        ws.append(cleaned_row)
    
                    headers = [cell.value for cell in ws[1]]
                    if "ETAT GEOMARKETING" in headers:
                        etat_col_index = headers.index("ETAT GEOMARKETING")
                        for i, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
                            cell = row[etat_col_index]
                            if cell.value == "OK":
                                cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                                cell.font = Font(color="006100")
    
                    wb.save(suivi_file_path)
                    try_push_to_github()
                    st.success("Fichier mis à jour avec succès")
    
                except Exception as e:
                    st.error(f"Erreur lors de la sauvegarde : {e}")
    
            shutil.copy(suivi_file_path, temp_export_path)
            st.download_button("Télécharger le fichier Excel", data=open(temp_export_path, "rb"),
                               file_name="Suivi_demandes_EXPORT.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            with st.expander("Ajouter une date de traitement"):
                commande_input = st.text_input("Entrez la commande (ex: CMD_X_00003320)")
                date_traitement_input = st.date_input("Date de traitement", format="DD/MM/YYYY")
            
                if st.button("Mettre à jour la date de traitement"):
                    try:
                        df = pd.read_excel(suivi_file_path, engine="openpyxl")
                        df.columns = [col.strip().upper() for col in df.columns]
            
                        occurrences = df[df["COMMANDE"] == commande_input]
            
                        if len(occurrences) == 0:
                            st.error("Commande introuvable.")
                        elif len(occurrences) > 1:
                            st.error(f"Il y a {len(occurrences)} lignes avec la commande '{commande_input}'. Mise à jour annulée.")
                        else:
                            # Une seule ligne → on poursuit
                            idx = occurrences.index[0]
                            df.at[idx, "DATE TRAITEMENT"] = pd.to_datetime(date_traitement_input)
            
                            dt_traitement = pd.to_datetime(df.at[idx, "DATE TRAITEMENT"], errors="coerce")
                            dt_reception = pd.to_datetime(df.at[idx, "DATE RECEPTION"], errors="coerce")
            
                            if pd.notnull(dt_traitement) and pd.notnull(dt_reception):
                                delai = (dt_traitement - dt_reception).days
                                df.at[idx, "DELAI TRAITEMENT"] = delai
                                st.success(f"Date et délai mis à jour : {delai} jours")
            
                                # Mise à jour dans le fichier Excel
                                wb = load_workbook(suivi_file_path)
                                ws = wb.active
                                headers = [cell.value.strip() if cell.value else "" for cell in ws[1]]
                                for row in ws.iter_rows(min_row=2):
                                    if row[headers.index("COMMANDE")].value == commande_input:
                                        row[headers.index("DATE TRAITEMENT")].value = date_traitement_input.strftime("%d/%m/%Y")
                                        row[headers.index("DELAI TRAITEMENT")].value = delai
                                        break
            
                                wb.save(suivi_file_path)
                                try_push_to_github()
                                st.rerun()
                            else:
                                st.warning("Dates invalides : réception ou traitement manquante.")
                    except Exception as e:
                        st.error(f"Erreur lors de la mise à jour : {e}")

    
        except Exception as e:
            st.error(f"Erreur lors de la lecture du fichier : {e}")

    
    elif section == "ajout":
        st.subheader("Ajout d'une nouvelle fiche DRI")

        if st.session_state["ligne_temporaire"] is None:
            with st.form("formulaire"):
                date_reception = st.date_input("Date de réception (jj/mm/aaaa)", format="DD/MM/YYYY")
                reseau = st.text_input("Réseau")
                type_demande_input = st.radio("Type de demande", ["1 - DEPASSEMENT DE COUT", "2 - DEMANDE DE MA"])
                demande_type = "DEPASSEMENT DE COUT" if type_demande_input.startswith("1") else "DEMANDE DE MA"
                commentaire = st.text_area("Commentaire (optionnel)")
                fiche_dri_file = st.file_uploader("Fichier Excel de la fiche DRI", type="xlsx")
                submit = st.form_submit_button("Ajouter la fiche")

            if submit and fiche_dri_file:
                try:
                    dri_wb = load_workbook(fiche_dri_file, data_only=True)
                    dri_ws = dri_wb.active

                    ligne = {
                        "DATE RECEPTION": date_reception.strftime("%d/%m/%Y"),
                        "RESEAU": reseau,
                        "RESPONSABLE PROD": dri_ws["C7"].value,
                        "COMMERCIAL": dri_ws["D16"].value,
                        "PROJET": dri_ws["D9"].value,
                        "TYPE DE DEMANDE": demande_type,
                        "COUT EXTENSION": dri_ws["D37"].value,
                        "COUT GLOBAL PROJET": dri_ws["D38"].value,
                        "OPERATEUR": dri_ws["D11"].value,
                        "GAIN DRI": dri_ws["G30"].value,
                        "ROI": round(dri_ws["G31"].value) if isinstance(dri_ws["G31"].value, (int, float)) else "ERREUR",
                        "NB CLIENTS AMORTISSEMENT": round((dri_ws["D38"].value - dri_ws["G30"].value) / 4000, 2)
                            if isinstance(dri_ws["D38"].value, (int, float)) and isinstance(dri_ws["G30"].value, (int, float)) else "ERREUR",
                        "COMMANDE": dri_ws["D10"].value,
                        "DATE TRAITEMENT": "",
                        "DELAI TRAITEMENT": "",
                        "ETAT GEOMARKETING": "",
                        "RESP GEOMARKET": "",
                        "CONCLUSION": "",
                        "COMMENTAIRE": commentaire
                    }

                    st.session_state["ligne_temporaire"] = ligne

                except Exception as e:
                    st.error(f"Erreur lors de la lecture du fichier DRI : {e}")

        else:
            st.success("Fiche prête à être enregistrée :")
            df_temp = pd.DataFrame([st.session_state["ligne_temporaire"]])
            st.dataframe(df_temp)

            col1, col2 = st.columns(2)
            with col1:
                if st.button("Enregistrer dans le fichier"):
                    try:
                        wb = load_workbook(suivi_file_path)
                        ws = wb.active
                        headers = [cell.value.strip() if cell.value else "" for cell in ws[1]]
                        new_row = [st.session_state["ligne_temporaire"].get(h, "") for h in headers]

                        last_row = max(
                            i for i, row in enumerate(ws.iter_rows(values_only=True), 1)
                            if any(cell is not None and str(cell).strip() != "" for cell in row)
                        )
                        next_row = last_row + 1

                        for col_idx, val in enumerate(new_row, start=1):
                            ws.cell(row=next_row, column=col_idx, value=val)

                        wb.save(suivi_file_path)
                        try_push_to_github()
                        st.success("Fiche enregistrée avec succès")
                        st.session_state["ligne_temporaire"] = None
                    except Exception as e:
                        st.error(f"Erreur lors de l'enregistrement : {e}")
            with col2:
                if st.button("Ajouter une autre fiche"):
                    st.session_state["ligne_temporaire"] = None
                    st.rerun()

    elif section == "tcd_ma":
        st.subheader("Analyse TCD - Marché Adressable")

        uploaded_file = st.file_uploader(
            "Déposez une archive ZIP contenant un shapefile (.shp, .shx, .dbf, etc.)",
            type=["zip"],
            key=st.session_state.get("upload_key", "upload")
        )

        if uploaded_file:
            with tempfile.TemporaryDirectory() as tmpdir:
                zip_path = os.path.join(tmpdir, "shapefile.zip")
                with open(zip_path, "wb") as f:
                    f.write(uploaded_file.read())
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(tmpdir)

                shp_files = [f for f in os.listdir(tmpdir) if f.endswith(".shp")]
                if not shp_files:
                    st.error("Aucun fichier .shp trouvé dans l'archive.")
                else:
                    shp_path = os.path.join(tmpdir, shp_files[0])
                    try:
                        gdf = gpd.read_file(shp_path)
                        required_cols = {"GAIN ENTRE", "NB_SALARIE", "SIRET"}
                        if not required_cols.issubset(gdf.columns):
                            st.error(f"Colonnes manquantes : {required_cols - set(gdf.columns)}")
                        else:
                            pivot = pd.pivot_table(
                                gdf,
                                index="NB_SALARIE",
                                columns="GAIN ENTRE",
                                values="SIRET",
                                aggfunc="count",
                                margins=True,
                                margins_name="Total général",
                                fill_value=0
                            )

                            ordered_cols = [
                                "Nouvellement forfaitaire",
                                "Changement forfaitaire",
                                "Pas de changement"
                            ]
                            existing_cols = [col for col in ordered_cols if col in pivot.columns]
                            pivot = pivot[existing_cols + ["Total général"]]
                            pivot = pivot.sort_index()

                            st.session_state["pivot"] = pivot

                            counts = gdf["GAIN ENTRE"].value_counts()
                            phrases = []
                            if "Nouvellement forfaitaire" in counts:
                                phrases.append(f"- {counts['Nouvellement forfaitaire']} entreprises de 1 salarié et plus deviennent éligibles forfaitairement")
                            if "Changement forfaitaire" in counts:
                                phrases.append(f"- {counts['Changement forfaitaire']} entreprises de 1 salarié et plus changent de zone forfaitaire positivement")

                            st.session_state["phrases"] = phrases

                            st.subheader("Tableau Croisé Dynamique")

                            def highlight_table(df):
                                styles = pd.DataFrame('', index=df.index, columns=df.columns)
                                styles.loc["Total général", :] = 'background-color: #f5e6cc'
                                styles.loc[:, "Total général"] = 'background-color: #e0f7fa'
                                styles.iloc[:, 0] = 'background-color: #f0f0f0'
                                return styles

                            styled_df = pivot.style.apply(highlight_table, axis=None)
                            st.dataframe(styled_df)

                            st.subheader("Résumé")
                            for phrase in phrases:
                                st.markdown(phrase)

                            fig, ax = plt.subplots(figsize=(11, len(pivot) * 0.6 + len(phrases) * 0.6 + 2))
                            y_offset = 0.1
                            for i, phrase in enumerate(phrases):
                                ax.text(0, 1 - y_offset * (i + 1), phrase, fontsize=10, ha='left')
                            ax.axis('off')

                            table_data = [[idx] + list(row) for idx, row in pivot.iterrows()]
                            col_labels = ["NB_SALARIE"] + pivot.columns.to_list()

                            table = ax.table(
                                cellText=table_data,
                                colLabels=col_labels,
                                loc='center',
                                cellLoc='center'
                            )
                            table.auto_set_font_size(False)
                            table.set_fontsize(10)
                            table.scale(1.2, 1.3)

                            for key, cell in table.get_celld().items():
                                row, col = key
                                cell.set_text_props(wrap=True)
                                if row == 0:
                                    cell.set_facecolor('#d3d3d3')
                                    cell.set_text_props(weight='bold')
                                elif col == 0:
                                    cell.set_facecolor('#f0f0f0')
                                if row > 0 and pivot.index[row - 1] == "Total général":
                                    cell.set_facecolor('#f5e6cc')
                                    cell.set_text_props(weight='bold')

                            buf = BytesIO()
                            plt.savefig(buf, format="png", bbox_inches='tight', dpi=300)
                            buf.seek(0)

                            st.download_button(
                                label="Télécharger l'image du TCD",
                                data=buf,
                                file_name="TCD_MA.png",
                                mime="image/png"
                            )

                            if st.button("Renouveler l'opération"):
                                st.session_state["upload_key"] = f"upload_{int(time.time())}"
                                st.session_state.pop("pivot", None)
                                st.session_state.pop("phrases", None)
                                st.rerun()
                    except Exception as e:
                        st.error(f"Erreur de lecture du shapefile : {e}")

    elif section == "analyse_donnees":
        st.subheader("Analyse des données - Synthèse Graphique et Statistique")

        try:
            import plotly.express as px
            df = pd.read_excel(suivi_file_path, engine="openpyxl")
            df.columns = [col.strip().upper() for col in df.columns]
            df["DATE RECEPTION"] = pd.to_datetime(df["DATE RECEPTION"], errors="coerce", dayfirst=True)
            df["ANNEE"] = df["DATE RECEPTION"].dt.year
            df["MOIS"] = df["DATE RECEPTION"].dt.strftime("%b")
            df["MOIS_NUM"] = df["DATE RECEPTION"].dt.month

            st.markdown("### 1. Coût global moyen & nombre de commandes par mois")
            grouped1 = df.groupby(["ANNEE", "MOIS", "MOIS_NUM"]).agg({
                "COUT GLOBAL PROJET": "mean",
                "COMMANDE": "count"
            }).reset_index().sort_values(by=["ANNEE", "MOIS_NUM"])
            grouped1 = grouped1.rename(columns={
                "COUT GLOBAL PROJET": "Coût moyen",
                "COMMANDE": "Nb commandes"
            })
            grouped1["Période"] = grouped1["ANNEE"].astype(str) + " - " + grouped1["MOIS"]
            st.dataframe(grouped1[["Période", "Coût moyen", "Nb commandes"]])


            st.markdown("### 2. Somme des coûts globaux par mois")
            grouped2 = df.groupby(["ANNEE", "MOIS", "MOIS_NUM"]).agg({
                "COUT GLOBAL PROJET": "sum",
                "COMMANDE": "count"
            }).reset_index().sort_values(by=["ANNEE", "MOIS_NUM"])
            grouped2 = grouped2.rename(columns={
                "COUT GLOBAL PROJET": "Coût total",
                "COMMANDE": "Nb commandes"
            })
            grouped2["Période"] = grouped2["ANNEE"].astype(str) + " - " + grouped2["MOIS"]
            st.dataframe(grouped2[["Période", "Coût total", "Nb commandes"]])


            st.markdown("### 3. Volume de commandes par responsable prod")
            commandes_prod = df["RESPONSABLE PROD"].value_counts().reset_index()
            commandes_prod.columns = ["Responsable", "Nb commandes"]
            st.dataframe(commandes_prod)
            st.bar_chart(commandes_prod.set_index("Responsable"))

            st.markdown("### 4. Volume de commandes par opérateur")
            commandes_operateur = df["OPERATEUR"].value_counts().reset_index()
            commandes_operateur.columns = ["Opérateur", "Nb commandes"]
            st.dataframe(commandes_operateur)
            st.bar_chart(commandes_operateur.set_index("Opérateur"))

            st.markdown("### 5. Répartition des commandes par date")
            repartition = df.groupby(["ANNEE", "MOIS", "MOIS_NUM"]).agg({"COMMANDE": "count"}).reset_index().sort_values(by=["ANNEE", "MOIS_NUM"])
            repartition.columns = ["ANNEE", "MOIS", "MOIS_NUM", "Nb commandes"]
            repartition["Période"] = repartition["ANNEE"].astype(str) + " - " + repartition["MOIS"]

            # Tableau
            st.dataframe(repartition[["Période", "Nb commandes"]])

            # Graphique
            fig = px.bar(repartition, x="Période", y="Nb commandes", title="Nb de commandes par mois")
            st.plotly_chart(fig, use_container_width=True)


            st.markdown("### 6. Délai de traitement moyen par mois")
            if "DELAI TRAITEMENT" in df.columns:
                df["DELAI TRAITEMENT"] = pd.to_numeric(df["DELAI TRAITEMENT"], errors="coerce")
                delai = df.groupby(["ANNEE", "MOIS", "MOIS_NUM"]).agg({"DELAI TRAITEMENT": "mean"}).reset_index().sort_values(by=["ANNEE", "MOIS_NUM"])
                delai["Période"] = delai["ANNEE"].astype(str) + " - " + delai["MOIS"]
                st.dataframe(delai[["Période", "DELAI TRAITEMENT"]].rename(columns={"DELAI TRAITEMENT": "Délai moyen"}))
                fig2 = px.line(delai, x="Période", y="DELAI TRAITEMENT", title="Délai de traitement moyen")
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.info("La colonne 'DELAI TRAITEMENT' est manquante dans le fichier.")

        except Exception as e:
            st.error(f"Erreur lors du chargement des analyses : {e}")

