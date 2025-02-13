import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import pypandoc
import os
from io import BytesIO
import zipfile
import tempfile

def generate_doc(template_bytes, context):
    """
    Erzeugt aus einem Docx-Template (als Bytes) und einem Kontext (dict)
    eine ausgefüllte .docx-Datei und gibt diese als Bytes zurück.
    """
    with open("temp_template.docx", "wb") as f:
        f.write(template_bytes)
    
    doc = DocxTemplate("temp_template.docx")
    doc.render(context)

    output_stream = BytesIO()
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_output_path = os.path.join(tmpdir, "temp_output.docx")
        doc.save(docx_output_path)
        with open(docx_output_path, "rb") as f:
            output_stream.write(f.read())

    output_stream.seek(0)
    return output_stream.read()

def convert_docx_to_pdf(docx_bytes):
    """
    Konvertiert DOCX-Bytes in PDF-Bytes mittels pypandoc.
    Erfordert Pandoc + LaTeX-Installation (z.B. TeX Live / MikTeX).
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, "temp.docx")
        pdf_path = os.path.join(tmpdir, "temp.pdf")
        
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)
        
        # Optional: extra_args=["--pdf-engine=xelatex"] für Emojis/Unicode
        pypandoc.convert_file(docx_path, "pdf", outputfile=pdf_path)
        
        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()
    
    return pdf_bytes

def expand_filename_template(template_str, context):
    """
    Ersetzt in template_str alle Vorkommen von {Spaltenname} 
    durch die entsprechenden Werte aus context.
    
    Beispiel:
      template_str: "Rechnung_{Kundennummer}_{Nachname}"
      context: {"Kundennummer": "1234", "Nachname": "Meier", ...}
      return -> "Rechnung_1234_Meier"
      
    Wenn ein Platzhalter nicht im context existiert, wird er 
    (vorsorglich) durch '' ersetzt.
    """
    result = template_str
    for key, value in context.items():
        placeholder = f"{{{key}}}"  # z. B. {Kundennummer}
        if placeholder in result:
            result = result.replace(placeholder, str(value))
    return result

def main():
    st.title("Serienbrief-Dokumentengenerierung (Dateinamen-Template)")

    st.write(
        """
        **Vorgehen**:
        1. Lade dein Word-Template (*.docx*) hoch.
        2. Lade deine CSV-Datei hoch (mit beliebigen Spaltennamen).
        3. Gib ein Dateinamen-Template ein, das Platzhalter in geschweiften Klammern nutzt.
           Beispiel: `Rechnung_{Rechnungsnummer}_{Nachname}`.
        4. Wähle aus, ob du **nur Word**, **nur PDF** oder **beides** erzeugen möchtest.
        5. Klicke auf "Dokumente generieren" und lade das ZIP-Archiv herunter.
        
        ---
        **Tipp**:
        - Die Platzhalter in deinem DOCX-Template sollten den Spaltennamen der CSV entsprechen.
        - Für den Dateinamen kannst du beliebige Spalten als `{Spaltenname}` verwenden. 
        - Wenn eine Spalte nicht existiert, wird sie durch einen leeren String ersetzt.
        """
    )

    # 1) Word-Template hochladen
    st.subheader("1) Word-Template hochladen")
    uploaded_template = st.file_uploader("Bitte lade dein DOCX-Template hoch", type=["docx"])
    
    # 2) CSV-Datei hochladen
    st.subheader("2) CSV-Datei hochladen")
    csv_file = st.file_uploader("Bitte wähle deine CSV-Datei aus", type=["csv"])

    # 3) Dateinamen-Template eingeben
    st.subheader("3) Dateinamen-Template")
    filename_template = st.text_input(
        "Verwende {Spaltenname} für Platzhalter, z. B. Rechnung_{Kundennummer}_{Nachname}",
        value="Dokument_{Nachname}_{Vorname}"
    )

    # # 4) DOCX/PDF-Auswahl
    # st.subheader("4) Welche Dokumente sollen erzeugt werden?")
    # doc_type_option = st.radio(
    #     "Dokumenten-Typ wählen",
    #     ["Nur DOCX", "Nur PDF", "DOCX und PDF"]
    # )

    if st.button("Dokumente generieren"):
        # Warnen, falls nichts hochgeladen wurde
        if not uploaded_template or not csv_file:
            st.warning("Bitte lade sowohl das Word-Template als auch eine CSV-Datei hoch.")
            return
        
        # CSV einlesen
        try:
            # Bei Bedarf anpassen: z. B. sep=';' für Semikolon-getrennte Dateien
            data = pd.read_csv(csv_file, sep=';', dtype=str)
        except Exception as e:
            st.error(f"Fehler beim Einlesen der CSV-Datei: {e}")
            return

        # ZIP-Puffer erstellen
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zf:
            # Für jede Zeile in der CSV
            for index, row in data.iterrows():
                # Kontext (alle Spalten dieser Zeile)
                context = row.to_dict()

                # Dateiname anhand der eingegebenen Template-Formel
                filename_prefix = expand_filename_template(filename_template, context)
                
                # Fallback, falls nach Ersetzen nichts übrig bleibt
                if not filename_prefix.strip():
                    filename_prefix = f"Eintrag_{index}"

                # 1) DOCX generieren
                docx_bytes = generate_doc(uploaded_template.getvalue(), context)

                # "Nur DOCX" oder "DOCX und PDF" -> DOCX ins ZIP
                # if doc_type_option in ["Nur DOCX", "DOCX und PDF"]:
                docx_filename = f"{filename_prefix}.docx"
                zf.writestr(docx_filename, docx_bytes)

                # # "Nur PDF" oder "DOCX und PDF" -> DOCX -> PDF, PDF ins ZIP
                # if doc_type_option in ["Nur PDF", "DOCX und PDF"]:
                #     pdf_bytes = convert_docx_to_pdf(docx_bytes)
                #     pdf_filename = f"{filename_prefix}.pdf"
                #     zf.writestr(pdf_filename, pdf_bytes)

        st.success("Dokumente erfolgreich generiert!")
        st.download_button(
            label="ZIP-Archiv herunterladen",
            data=zip_buffer.getvalue(),
            file_name="dokumente.zip",
            mime="application/x-zip-compressed"
        )

if __name__ == "__main__":
    main()