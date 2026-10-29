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
                text += page.extract_text() or "\n"
            
            # 1. Búsqueda de número de factura
            num_factura = re.search(r'N[°º]?\s*[:.]?\s*([0-9A-Za-z_-]+)', text, re.IGNORECASE)
            
            # 2. Búsqueda de fecha (si existe)
            fecha = re.search(r'\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b', text)
            
            # 3. Búsqueda robusta de Total (tolera saltos de línea y barras de tablas)
            total = re.search(r'TOTAL[\s\S]*?([\d.,]+\s*€)', text, re.IGNORECASE)

            data.append({
                "Archivo": pdf_file.name,
                "Nº Factura": num_factura.group(1) if num_factura else "No detectado",
                "Fecha": fecha.group(0) if fecha else "No especificada",
                "Total": total.group(1).strip() if total else "No detectado"
            })
            
    df = pd.DataFrame(data)
    st.write("### Datos Extraídos:")
    st.dataframe(df)

    # Convertir DataFrame a Excel en memoria
    excel_file = "facturas_extraidas.xlsx"
    df.to_excel(excel_file, index=False)
    
    with open(excel_file, "rb") as f:
        st.download_button(
            label="📥 Descargar todo en Excel (.xlsx)",
            data=f,
            file_name=excel_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
