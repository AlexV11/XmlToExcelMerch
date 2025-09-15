import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from io import BytesIO
import os

def extract_data_from_xml(file, filename):
    """
    Extrae todos los datos de Note/Text y PartType id de un archivo XML (desde UploadedFile)
    """
    try:
        tree = ET.parse(file)
        root = tree.getroot()

        data = []

        for app in root.findall('App'):
            parttype_element = app.find('PartType')
            parttype_id = parttype_element.get('id', '') if parttype_element is not None else ""

            notes = [n.text or "" for n in app.findall('.//Note')]
            texts = [t.text or "" for t in app.findall('.//Text')]

            for note in notes + texts:
                if note or parttype_id:
                    data.append({
                        'Note/Text': note,
                        'PartType_ID': parttype_id,
                        'Source_File': filename,
                        'key': parttype_id + filename + note
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
    con barra de progreso
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

        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="XML_Data")
        output.seek(0)

        return output, df
    else:
        return None, None

# Streamlit UI
st.title("XML to Excel Converter")

uploaded_files = st.file_uploader(
    "Sube uno o varios archivos XML",
    accept_multiple_files=True,
    type="xml"
)

if uploaded_files:
    st.success(f"Se cargaron {len(uploaded_files)} archivo(s).")

    if st.button("Convertir a Excel"):
        excel_file, df_preview = convert_xmls_to_excel(uploaded_files)

        st.write(f"Unique Notes/Text: {df_preview['Note/Text'].nunique()}")
        st.write(f"Unique PartType_IDs: {df_preview['PartType_ID'].nunique()}")

        if excel_file:
            st.download_button(
                label="📥 Descargar Excel",
                data=excel_file,
                file_name="xml_data_converted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No se encontraron datos para convertir.")