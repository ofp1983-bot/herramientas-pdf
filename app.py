import streamlit as st
import fitz  # PyMuPDF
import pypdf # Librería para compatibilidad cruzada
from io import BytesIO
import hashlib
import os
import subprocess
import tempfile
from datetime import datetime
import re

# ==========================================
# FUNCIONES PRINCIPALES
# ==========================================

def embed_file_in_pdf(pdf_bytes, attachment_bytes, attachment_name):
    """Anexa un archivo embebido a un PDF usando PyMuPDF."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    doc.embfile_add(attachment_name, attachment_bytes, filename=attachment_name)
    return doc.write()

def convert_to_pdfa(pdf_bytes, level="2b"):
    """
    Convierte un PDF a PDF/A usando Ghostscript con depuración en vivo.
    Tiene compatibilidad cruzada para anexos creados por pypdf y PyMuPDF.
    """
    gs_cmd = "gswin64c" if os.name == "nt" else "gs"
    part = "3" if level == "3b" else "2"
    conformance = "B"
    
    # =========================================================
    # 1. BÚSQUEDA BILINGÜE DE ADJUNTOS
    # =========================================================
    st.write("🔍 **Debug Paso 1:** Analizando el documento con motor dual (pypdf + PyMuPDF)...")
    attachments = []
    
    # Motor 1: Intentar leer con pypdf
    try:
        reader = pypdf.PdfReader(BytesIO(pdf_bytes))
        if reader.attachments:
            for name, byte_list in reader.attachments.items():
                for file_data in byte_list:
                    attachments.append((name, file_data))
    except Exception:
        pass
        
    # Motor 2: Si pypdf no encontró nada, intentar con PyMuPDF
    if not attachments:
        try:
            doc_original = fitz.open(stream=pdf_bytes, filetype="pdf")
            for name in doc_original.embfile_names():
                attachments.append((name, doc_original.embfile_get(name)))
            doc_original.close()
        except Exception as e:
            st.error(f"❌ Error al leer adjuntos: {e}")
            
    st.write(f"✅ **Debug Paso 2:** Se encontraron **{len(attachments)}** archivos adjuntos en la memoria.")
    
    # =========================================================
    # 2. CONVERSIÓN CON GHOSTSCRIPT
    # =========================================================
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_in:
        temp_in.write(pdf_bytes)
        temp_in_path = temp_in.name
        
    temp_out_path = temp_in_path.replace(".pdf", "_out.pdf")
    
    try:
        if attachments:
            st.info(f"🛠️ Procediendo a recodificar el documento y rescatar {len(attachments)} anexo(s)...")
            
        st.write("⚙️ **Debug Paso 3:** Enviando a Ghostscript para recodificación PDF/A...")
        cmd = [
            gs_cmd,
            "-dPDFA=" + part,
            "-dBATCH",
            "-dNOPAUSE",
            "-dColorConversionStrategy=/UseDeviceIndependentColor",
            "-sDEVICE=pdfwrite",
            "-dPDFACompatibilityPolicy=1",
            f"-sOutputFile={temp_out_path}",
            temp_in_path
        ]
        
        process = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        
        if process.returncode != 0:
            st.error(f"❌ Fallo en Ghostscript. Detalles: {process.stderr if process.stderr.strip() else process.stdout}")
            return None
            
        st.write("✅ **Debug Paso 4:** Ghostscript terminó. Reconstruyendo documento...")
        
        # =========================================================
        # 3. RE-INYECCIÓN Y METADATOS
        # =========================================================
        doc = fitz.open(temp_out_path)
        
        # Re-inyectar adjuntos
        if level == "3b" and attachments:
            for name, file_data in attachments:
                doc.embfile_add(name, file_data, filename=name)
            st.write("✅ **Debug Paso 5:** Los anexos fueron re-inyectados al PDF final usando PyMuPDF.")
            st.success("¡Documento ensamblado con éxito!")
        elif level == "2b" and attachments:
            st.warning("⚠️ Nota: Elegiste PDF/A-2b. Esta norma NO permite anexos. Los archivos embebidos han sido descartados definitivamente.")
        
        # Inyectar metadatos XML
        xml_metadata = f"""<?xpacket begin="﻿" id="W5M0MpCehiHzreSzNTczkc9d"?>
<x:xmpmeta xmlns:x="adobe:ns:meta/">
  <rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
    <rdf:Description rdf:about="" xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/">
      <pdfaid:part>{part}</pdfaid:part>
      <pdfaid:conformance>{conformance}</pdfaid:conformance>
    </rdf:Description>
  </rdf:RDF>
</x:xmpmeta>
<?xpacket end="w"?>"""
        
        doc.set_xml_metadata(xml_metadata)
        pdfa_bytes = doc.write()
        doc.close()
            
        return pdfa_bytes
        
    except Exception as e:
        st.error(f"❌ Error inesperado: {str(e)}")
        return None
    finally:
        if os.path.exists(temp_in_path): os.remove(temp_in_path)
        if os.path.exists(temp_out_path): os.remove(temp_out_path)

def get_file_hash(file_bytes):
    return hashlib.sha256(file_bytes).hexdigest()

def validate_pdfa(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    xml = doc.get_xml_metadata()
    if not xml:
        return False, "No es PDF/A (Sin metadatos XML)"
    
    part_match = re.search(r'<pdfaid:part>(\d)</pdfaid:part>', xml)
    conf_match = re.search(r'<pdfaid:conformance>([A-Z]+)</pdfaid:conformance>', xml)
    
    if part_match and conf_match:
        return True, f"PDF/A-{part_match.group(1)}{conf_match.group(1)}"
    
    return False, "No detectado"

def get_attachments_names(pdf_bytes):
    """Extrae los nombres de los archivos adjuntos usando un motor dual."""
    names = []
    # Motor 1: pypdf
    try:
        reader = pypdf.PdfReader(BytesIO(pdf_bytes))
        if reader.attachments:
            for name in reader.attachments.keys():
                names.append(name)
    except Exception:
        pass
        
    # Motor 2: PyMuPDF (si pypdf no encontró nada)
    if not names:
        try:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            names = list(doc.embfile_names())
            doc.close()
        except Exception:
            pass
    
    # Retorna la lista de nombres sin duplicados
    return list(set(names))

def generate_report(file_name, file_bytes):
    sha256_hash = get_file_hash(file_bytes)
    size_kb = len(file_bytes) / 1024
    is_valid, pdfa_level = validate_pdfa(file_bytes)
    
    # Búsqueda de anexos para el informe
    nombres_anexos = get_attachments_names(file_bytes)
    estado_anexos = "Sí" if nombres_anexos else "No"
    texto_nombres_anexos = ", ".join(nombres_anexos) if nombres_anexos else "N/A"
    
    report = {
        "Nombre del Archivo": file_name,
        "Hash SHA-256": sha256_hash,
        "Tamaño": f"{size_kb:.2f} KB",
        "Fecha de Análisis": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Cumplimiento PDF/A": "Válido" if is_valid else "No Cumple",
        "Nivel PDF/A Detectado": pdfa_level,
        "Contiene Anexos": estado_anexos,
        "Nombre de los Anexos": texto_nombres_anexos
    }
    return report

# ==========================================
# INTERFAZ WEB CON STREAMLIT
# ==========================================

st.set_page_config(page_title="Gestor PDF v3.1", layout="wide")

col_menu, col_main = st.columns([1, 3])

with col_menu:
    st.markdown("### ⚙️ Menú de Tareas")
    menu_option = st.radio(
        "Selecciona una acción:",
        (
            "1. Anexar Excel a PDF",
            "2. Adjuntar Cualquier Archivo",
            "3. Convertir a PDF/A-2b",
            "4. Convertir a PDF/A-3b",
            "5. Generar Informe de Preservación"
        )
    )

with col_main:
    st.title("📄 Herramienta de Preservación PDF (v3.1)")
    st.write("Sube tu documento principal y selecciona una acción.")
    
    main_pdf = st.file_uploader("Sube el archivo PDF principal", type=["pdf"])

    if main_pdf is not None:
        pdf_bytes = main_pdf.read()

        if "1. Anexar" in menu_option or "2. Adjuntar" in menu_option:
            st.subheader(menu_option)
            allowed = ["xlsx", "xls"] if "Excel" in menu_option else None
            attachment = st.file_uploader("Sube el archivo a adjuntar", type=allowed)
            
            if attachment and st.button("Embeber Archivo"):
                with st.spinner("Embebiendo archivo..."):
                    result_pdf = embed_file_in_pdf(pdf_bytes, attachment.read(), attachment.name)
                    st.success("¡Archivo adjuntado con éxito!")
                    st.download_button(label="Descargar PDF con Anexo", data=result_pdf, file_name=f"con_anexo_{main_pdf.name}", mime="application/pdf")

        elif "Convertir a PDF/A" in menu_option:
            st.subheader(menu_option)
            level = "3b" if "3b" in menu_option else "2b"
            
            if st.button(f"Ejecutar Conversión a {level}"):
                with st.spinner("Procesando conversión..."):
                    result_pdfa = convert_to_pdfa(pdf_bytes, level=level)
                    if result_pdfa:
                        st.download_button(label=f"Descargar PDF/A-{level}", data=result_pdfa, file_name=f"pdfa_{level}_{main_pdf.name}", mime="application/pdf")

        elif "5. Generar Informe" in menu_option:
            st.subheader("📊 Informe de Preservación")
            if st.button("Ejecutar Análisis"):
                with st.spinner("Analizando..."):
                    report = generate_report(main_pdf.name, pdf_bytes)
                    for key, value in report.items():
                        st.markdown(f"**{key}:** {value}")
                        
                    report_text = "\n".join([f"{k}: {v}" for k, v in report.items()])
                    st.download_button(
                        label="Descargar Informe (.txt)",
                        data=report_text,
                        file_name=f"informe_{main_pdf.name}.txt",
                        mime="text/plain"
                    )