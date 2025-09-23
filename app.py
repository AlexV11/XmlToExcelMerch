import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from io import BytesIO
import os
import re

# --- Diccionario de reemplazos ---
# --- Cargar reemplazos desde Replacements.xlsx ---
def load_replacements_from_excel(filepath="Replacements.xlsx"):
    """
    Carga pares de reemplazo desde un archivo Excel con columnas 'old_word' y 'new_word'.
    """
    if os.path.exists(filepath):
        df = pd.read_excel(filepath)
        # Espera columnas: old_word, new_word
        return dict(zip(df['Term'].astype(str), df['New Term'].astype(str)))
    else:
        st.warning(f"No se encontró {filepath}. El diccionario de reemplazos estará vacío.")
        return {}

replace_dict = load_replacements_from_excel()

def clean_text(text, replacements):
    if not text:
        return text

    sorted_terms = sorted(replacements.keys(), key=len, reverse=True)
    pattern = r'(' + '|'.join(re.escape(term) for term in sorted_terms) + r')'

    def replace_match(match):
        found = match.group(0)
        for old, new in replacements.items():
            if found.lower() == old.lower():
                return new
        return found

    cleaned = re.sub(pattern, replace_match, text, flags=re.IGNORECASE)
    # Normalizar espacios
    return re.sub(r'\s+', ' ', cleaned).strip()

def extract_data_from_xml(file, filename):
    """
    Extrae todos los datos de Note/Text y PartType id de un archivo XML (desde UploadedFile)
    """
    try:
        tree = ET.parse(file)
        root = tree.getroot()

        data = []

        for app in root.findall('.//App'):  # App en cualquier nivel
            parttype_element = app.find('PartType')
            parttype_id = parttype_element.get('id', '') if parttype_element is not None else ""

            notes = [n.text or "" for n in app.findall('.//Note')] + [n.text or "" for n in app.findall('.//note')]
            texts = [t.text or "" for t in app.findall('.//Text')] + [t.text or "" for t in app.findall('.//text')]
            labels = [l.text or "" for l in app.findall('.//MfrLabel')] + [l.text or "" for l in app.findall('.//mfrlabel')]

            for value in notes + texts + labels:
                if value or parttype_id:
                    data.append({
                        'Note/Text/MfrLabel': value,
                        'PartType_ID': parttype_id,
                        'Source_File': filename,
                        'key': parttype_id + filename + value
                    })

        return data

    except ET.ParseError as e:
        st.error(f"Error procesando archivo {filename}: {type(e).__name__} - {str(e)}")
        return []
    except Exception as e:
        st.error(f"Error procesando archivo {filename}: {type(e).__name__} - {str(e)}")
        return []


def convert_xmls_to_excel(uploaded_files):
    """
    Convierte múltiples archivos XML subidos a un único Excel (BytesIO)
    con barra de progreso y limpieza de textos
    """
    all_data = []
    total_files = len(uploaded_files)
    progress_bar = st.progress(0)
    status_text = st.empty()

    for i, uploaded_file in enumerate(uploaded_files, start=1):
        if uploaded_file.size > 0:
            status_text.text(f"Procesando {uploaded_file.name} ({i}/{total_files})...")
            data = extract_data_from_xml(uploaded_file, uploaded_file.name)
            all_data.extend(data)
        else:
            st.warning(f"El archivo {uploaded_file.name} está vacío y será omitido.")

        # actualizar progreso
        progress = int(i / total_files * 100)
        progress_bar.progress(progress)

    status_text.text("Procesamiento completado ✅")

    if all_data:
        df = pd.DataFrame(all_data)
        df.drop_duplicates(inplace=True)

        # Aplicar limpieza exacta a la columna
        df['Note/Text/MfrLabel_Clean'] = df['Note/Text/MfrLabel'].apply(lambda x: clean_text(x, replace_dict))

        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="XML_Data")
        output.seek(0)

        return output, df
    else:
        return None, None


# --- Streamlit UI ---
st.title("XML to Excel Converter (con limpieza de textos)")

uploaded_files = st.file_uploader(
    "Sube uno o varios archivos XML",
    accept_multiple_files=True,
    type="xml"
)

if uploaded_files:
    st.success(f"Se cargaron {len(uploaded_files)} archivo(s).")

    if st.button("Convertir a Excel"):
        excel_file, df_preview = convert_xmls_to_excel(uploaded_files)

        if df_preview is not None:
            st.write(f"Unique Notes/Text/MfrLabel: {df_preview['Note/Text/MfrLabel'].nunique()}")
            st.write(f"Unique PartType_IDs: {df_preview['PartType_ID'].nunique()}")
            st.dataframe(df_preview.head(20))  # vista previa

        if excel_file:
            st.download_button(
                label="📥 Descargar Excel",
                data=excel_file,
                file_name="xml_data_converted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No se encontraron datos para convertir.")