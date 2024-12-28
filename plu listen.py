import pandas as pd
from docx import Document

def generate_plu_list(mother_file_path, plu_week_file_path, output_doc_path):
    # 1. Daten laden
    mother_file = pd.ExcelFile(mother_file_path)
    plu_week_df = pd.read_excel(plu_week_file_path)

    # Sicherstellen, dass PLU-Nummern Integer sind
    plu_week_df["PLU"] = plu_week_df["PLU"].astype(int)

    # Kategorien aus der Mutterdatei laden
    categories = mother_file.sheet_names
    filtered_data = {}

    # 2. Abgleich der PLU-Nummern und Filtern der Artikel
    for category in categories:
        category_data = mother_file.parse(category)
        matched_data = pd.merge(plu_week_df, category_data, on="PLU", how="inner")
        matched_data = matched_data.sort_values(by="Artikel").reset_index(drop=True)
        filtered_data[category] = matched_data

    # 3. Word-Dokument erstellen
    doc = Document()

    for category, data in filtered_data.items():
        # Kategorie als Überschrift
        doc.add_heading(category, level=1)

        # PLU-Nummern und Artikel einfügen
        for _, row in data.iterrows():
            doc.add_paragraph(f"{row['PLU']}\t{row['Artikel']}")

    # 4. Dokument speichern
    doc.save(output_doc_path)

# Beispielaufruf
generate_plu_list(
    mother_file_path="/path/to/mutterdatei.xlsx",
    plu_week_file_path="/path/to/plu_woche.xlsx",
    output_doc_path="/path/to/output_plu_list.docx"
)
