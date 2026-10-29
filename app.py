import streamlit as st
import pdfplumber
import pandas as pd
import re

st.set_page_config(page_title="Extractor de Facturas a Excel", page_icon="📄")

st.title("📄 Extractor de Facturas/Albaranes en PDF a Excel")
st.write("Sube tus facturas en PDF y descarga una tabla con la información extraída.")

uploaded_files = st.file_uploader("Sube una o varias facturas (PDF)", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    data = []
    
    for pdf_file in uploaded_files:
        with pdfplumber.open(pdf_file) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() or ""
            
            fecha = re.search(r'\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b', text)
            cif = re.search(r'\b[ABCDEFGHJNPQRSUVW]\d{7}[0-9A-J]\b', text, re.IGNORECASE)
            total = re.search(r'(?:total|importe)\s*:?\s*([\d.,]+)\s*€?', text, re.IGNORECASE)

            data.append({
                "Archivo": pdf_file.name,
                "Fecha": fecha.group(0) if fecha else "No detectado",
                "CIF/NIF": cif.group(0) if cif else "No detectado",
                "Total (€)": total.group(1) if total else "No detectado"
            })
            
    df = pd.DataFrame(data)
    st.write("### Datos Extraídos:")
    st.dataframe(df)

    excel_file = "facturas_extraidas.xlsx"
    df.to_excel(excel_file, index=False)
    
    with open(excel_file, "rb") as f:
        st.download_button(
            label="📥 Descargar todo en Excel (.xlsx)",
            data=f,
            file_name=excel_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
