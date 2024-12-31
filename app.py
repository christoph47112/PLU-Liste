import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

def generate_plu_list(mother_file_path, plu_week_file):
    """
    Diese Funktion erstellt eine PLU-Liste basierend auf der Mutterdatei und der hochgeladenen PLU-Woche-Datei.

    Parameter:
    - mother_file_path: Pfad zur Mutterdatei (Excel).
    - plu_week_file: Hochgeladene Excel-Datei (PLU-Woche) als BytesIO.

    Rückgabe:
    - BytesIO-Objekt mit der generierten Word-Datei.
    - DataFrame für die Excel-Ausgabe im kategorisierten Format.
    """
    # 1. Daten laden
    mother_file = pd.ExcelFile(mother_file_path)
    plu_week_df = pd.read_excel(plu_week_file, header=None)  # Keine Kopfzeile vorhanden

    # Umbenennen der Spalte in 'PLU'
    plu_week_df.columns = ['PLU']

    # Entferne ungültige Werte (NaN oder inf)
    plu_week_df = plu_week_df.dropna(subset=["PLU"])  # Entferne Zeilen mit NaN in der PLU-Spalte

    # Sicherstellen, dass die Werte numerisch sind
    plu_week_df["PLU"] = pd.to_numeric(plu_week_df["PLU"], errors='coerce')  # Konvertiere in numerische Werte
    plu_week_df = plu_week_df.dropna(subset=["PLU"])  # Entferne Zeilen, die nach der Konvertierung ungültig sind
    plu_week_df["PLU"] = plu_week_df["PLU"].astype(int)  # Konvertiere die bereinigte Spalte in Integer

    # Kategorien aus der Mutterdatei laden
    categories = mother_file.sheet_names
    filtered_data = []

    # 2. Abgleich der PLU-Nummern und Filtern der Artikel
    for category in categories:
        category_data = mother_file.parse(category)

        if "PLU" not in category_data.columns or "Artikel" not in category_data.columns:
            raise ValueError(f"Die Kategorie '{category}' in der Mutterdatei enthält nicht die benötigten Spalten 'PLU' oder 'Artikel'.")

        matched_data = pd.merge(plu_week_df, category_data, on="PLU", how="inner")
        matched_data = matched_data.sort_values(by="Artikel").reset_index(drop=True)
        matched_data["Kategorie"] = category
        filtered_data.append(matched_data)

    # Kombiniere alle Daten zu einem DataFrame
    combined_df = pd.concat(filtered_data, ignore_index=True)

    # 3. Word-Dokument erstellen
    doc = Document()

    for category, data in combined_df.groupby("Kategorie"):
        # Kategorie als Überschrift
        doc.add_heading(category, level=1)

        # PLU-Nummern und Artikel einfügen
        for _, row in data.iterrows():
            doc.add_paragraph(f"{row['PLU']}\t{row['Artikel']}")

    # 4. Dokument in BytesIO speichern
    output_word = BytesIO()
    doc.save(output_word)
    output_word.seek(0)

    # 5. Erstelle eine Pivot-ähnliche Tabelle
    pivot_table = combined_df.pivot_table(
        values="PLU",
        index=["Kategorie"],
        columns=["Artikel"],
        aggfunc=lambda x: ', '.join(map(str, x)),
        fill_value=""
    )

    # Speichere die Tabelle als Excel
    output_excel = BytesIO()
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        combined_df.to_excel(writer, index=False, sheet_name="Detailed Data")
        pivot_table.to_excel(writer, sheet_name="Pivot Table")
    output_excel.seek(0)

    return output_word, output_excel

# Streamlit App
st.title("PLU List Generator")

# Feste Mutterdatei definieren
MOTHER_FILE_PATH = "mother_file.xlsx"

# Datei-Upload für die PLU-Woche
uploaded_plu_week_file = st.file_uploader("Upload PLU Week File (Excel)", type="xlsx")

if st.button("Generate PLU List"):
    if uploaded_plu_week_file:
        try:
            with st.spinner("Processing..."):
                # PLU-Liste generieren
                output_word, output_excel = generate_plu_list(MOTHER_FILE_PATH, uploaded_plu_week_file)

            # Download-Links für die Dateien anzeigen
            st.success("PLU List successfully generated!")
            st.download_button(
                label="Download PLU List (Word)",
                data=output_word,
                file_name="plu_list.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            st.download_button(
                label="Download PLU List (Excel - Pivot Format)",
                data=output_excel,
                file_name="plu_list_pivot.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except ValueError as e:
            st.error(f"Input Error: {str(e)}")
        except Exception as e:
            st.error(f"An unexpected error occurred: {str(e)}")
    else:
        st.warning("Please upload the PLU Week File.")
