import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import unicodedata

st.set_page_config(page_title="PLU Listen Anwendung")

def normalize_text(text):
    """Entfernt fehlerhafte Zeichen und normalisiert den Text."""
    return unicodedata.normalize("NFKD", str(text)).encode("ascii", "ignore").decode("ascii")

def generate_plu_list(mother_file_path, plu_week_file):
    """
    Erstellt eine PLU-Liste mit Kategorien auf neuen Seiten.
    """
    mother_file = pd.ExcelFile(mother_file_path)
    plu_week_df = pd.read_excel(plu_week_file, header=None, names=['PLU'])
    
    # Entferne ungültige Werte und wandle in Integer um
    plu_week_df = plu_week_df.dropna(subset=["PLU"])
    plu_week_df["PLU"] = pd.to_numeric(plu_week_df["PLU"], errors='coerce').dropna().astype(int)
    
    categories = mother_file.sheet_names
    filtered_data = []
    
    for category in categories:
        category_data = mother_file.parse(category)
        if "PLU" not in category_data.columns or "Artikel" not in category_data.columns:
            st.warning(f"Kategorie '{category}' hat keine gültige PLU oder Artikel-Spalte und wird übersprungen.")
            continue
        
        category_data["Artikel"] = category_data["Artikel"].apply(normalize_text)
        matched_data = pd.merge(plu_week_df, category_data, on="PLU", how="inner")
        matched_data = matched_data.sort_values(by="Artikel").reset_index(drop=True)
        matched_data["Kategorie"] = category
        filtered_data.append(matched_data)
    
    if not filtered_data:
        raise ValueError("Keine Übereinstimmungen gefunden. Bitte überprüfen Sie die Eingabedateien.")
    
    combined_df = pd.concat(filtered_data, ignore_index=True).drop_duplicates(subset=['PLU'])
    doc = Document()
    
    first_page = True
    for category, data in combined_df.groupby("Kategorie"):
        if not first_page:
            doc.add_page_break()
        first_page = False
        doc.add_heading(category, level=1)
        
        for _, row in data.iterrows():
            doc.add_paragraph(f"{row['PLU']} {row['Artikel']}")
    
    output_word = BytesIO()
    doc.save(output_word)
    output_word.seek(0)
    
    pivot_table = combined_df.pivot_table(
        values="PLU",
        index=["Kategorie"],
        columns=["Artikel"],
        aggfunc=lambda x: ', '.join(map(str, x)),
        fill_value=""
    )
    
    output_excel = BytesIO()
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        combined_df.to_excel(writer, index=False, sheet_name="Detailed Data")
        pivot_table.to_excel(writer, sheet_name="Pivot Table")
    output_excel.seek(0)
    
    return output_word, output_excel

st.title("PLU-Listen Generator")
MOTHER_FILE_PATH = "mother_file.xlsx"
uploaded_plu_week_file = st.file_uploader("PLU-Wochen-Datei hochladen (Excel)", type=["xlsx"])

if st.button("PLU-Liste generieren"):
    if uploaded_plu_week_file:
        try:
            with st.spinner("Processing..."):
                output_word, output_excel = generate_plu_list(MOTHER_FILE_PATH, uploaded_plu_week_file)
            
            st.success("PLU-Liste erfolgreich erstellt!")
            st.download_button(
                label="PLU-Liste herunterladen (Word)",
                data=output_word,
                file_name="plu_list.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            st.download_button(
                label="PLU-Liste herunterladen (Excel - Pivot-Format)",
                data=output_excel,
                file_name="plu_list_pivot.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except ValueError as e:
            st.error(f"Eingabefehler: {str(e)}")
        except Exception as e:
            st.error(f"Ein unerwarteter Fehler ist aufgetreten: {str(e)}")
    else:
        st.warning("Bitte laden Sie die PLU-Wochen-Datei hoch.")

# Neuer Datenschutzhinweis
st.markdown("⚠️ **Hinweis:** Diese Anwendung speichert keine Daten und hat keinen Zugriff auf Ihre Dateien.")
st.markdown("🌟 **Erstellt von Christoph R. Kaiser mit Hilfe von Künstlicher Intelligenz.**")
