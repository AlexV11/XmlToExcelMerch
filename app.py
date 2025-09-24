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
        # elimina los strings que digan '#N/A' de df['New Term']
        df = df[df['New Term'].notna() & (df['New Term'].astype(str).str.upper() != '#N/A')]

        return dict(zip(df['Term'].astype(str), df['New Term'].astype(str)))
    else:
        st.warning(f"No se encontró {filepath}. El diccionario de reemplazos estará vacío.")
        return {}

replace_dict = load_replacements_from_excel()

# ---------------------------
# Helpers para limpieza
# ---------------------------

def build_pattern_from_keys(keys):
    """
    Construye un patrón alternante:
     - términos alfanuméricos -> con \b...\b (evita sustituir dentro de palabras)
     - términos con símbolos   -> literal escapado
    Ordena de mayor a menor longitud para evitar solapamientos.
    """
    keys_sorted = sorted(keys, key=lambda k: len(k), reverse=True)
    parts = []
    for k in keys_sorted:
        esc = re.escape(k.strip())
        if k.strip().isalnum():
            parts.append(rf"\b{esc}\b")
        else:
            parts.append(esc)
    return re.compile("|".join(parts), flags=re.IGNORECASE)


def clean_text(text, replacements):
    """
    Reemplaza términos y ajusta formato:
    - Aplica diccionario de reemplazos con seguridad
    - Mantiene palabras en MAYÚSCULAS (ej. OBD, ABS)
    - Para lo demás: solo la primera letra de todo el string en mayúscula,
      el resto en minúsculas
    """
    if not text:
        return text

    s = text

    # --- Reemplazos ---
    sorted_terms = sorted(replacements.keys(), key=len, reverse=True)
    for term in sorted_terms:
        repl = replacements[term] if replacements[term] is not None else ""

        if re.fullmatch(r'\w+', term):
            # términos alfanuméricos → con \b
            pattern = r'\b' + re.escape(term) + r'\b'
            s = re.sub(pattern, repl, s, flags=re.IGNORECASE)
        else:
            # términos con símbolos → buscar literal y añadir espacio si corresponde
            def _repl_nonword(m):
                after_idx = m.end()
                add_space = after_idx < len(m.string) and m.string[after_idx].isalnum()
                # si el reemplazo no termina en espacio y lo siguiente es letra/dígito, agregarlo
                return repl + (' ' if add_space and not repl.endswith(' ') else '')
            s = re.sub(re.escape(term), _repl_nonword, s, flags=re.IGNORECASE)

    # Normalizar espacios
    s = re.sub(r'\s+', ' ', s).strip()

    if not s:
        return s

    # --- Manejo de mayúsculas ---
    words = s.split()
    new_words = []
    for w in words:
        if w.isupper():
            new_words.append(w)        # conservar acrónimos
        else:
            new_words.append(w.lower())  # minúsculas temporales

    # reconstruir y aplicar capitalización global
    s = " ".join(new_words)
    s = s[0].upper() + s[1:] if len(s) > 1 else s.upper()

    return s



# ---------------------------
# Extracción de XML (igual que tenías)
# ---------------------------
def extract_data_from_xml(file, filename):
    try:
        tree = ET.parse(file)
        root = tree.getroot()

        data = []

        for app in root.findall('.//App'):
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

# ---------------------------
# Conversión y UI
# ---------------------------
def convert_xmls_to_excel(uploaded_files, replacements):
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

        progress = int(i / total_files * 100)
        progress_bar.progress(progress)

    status_text.text("Procesamiento completado ✅")

    if all_data:
        df = pd.DataFrame(all_data)
        df.drop_duplicates(inplace=True)

        # Aplicar limpieza segura
        df['Note/Text/MfrLabel_Clean'] = df['Note/Text/MfrLabel'].apply(lambda x: clean_text(x, replacements))

        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="XML_Data")
        output.seek(0)

        return output, df
    else:
        return None, None

# Streamlit UI
st.title("XML to Excel Converter (limpieza segura de términos)")

uploaded_files = st.file_uploader(
    "Sube uno o varios archivos XML",
    accept_multiple_files=True,
    type="xml"
)

# Área para probar texto suelto rápido
st.markdown("**Probar limpieza rápida:** pega una línea y pulsa el botón.")
sample_input = st.text_area("Texto de prueba", "w/HD Dana 60 Fr Axle\nCamber: ± 1.75 Degrees\nCamber: plus or minus 1.75 Degrees", height=120)
if st.button("Limpiar texto de prueba"):
    out_lines = []
    for line in sample_input.splitlines():
        out_lines.append(clean_text(line, replace_dict))
    st.text("\n".join(out_lines))

if uploaded_files:
    st.success(f"Se cargaron {len(uploaded_files)} archivo(s).")

    if st.button("Convertir a Excel"):
        excel_file, df_preview = convert_xmls_to_excel(uploaded_files, replace_dict)

        if df_preview is not None:
            st.write(f"Unique Notes/Text/MfrLabel: {df_preview['Note/Text/MfrLabel'].nunique()}")
            st.write(f"Unique PartType_IDs: {df_preview['PartType_ID'].nunique()}")
            st.dataframe(df_preview.head(20))

        if excel_file:
            st.download_button(
                label="📥 Descargar Excel",
                data=excel_file,
                file_name="xml_data_converted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No se encontraron datos para convertir.")