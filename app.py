import streamlit as st
import fitz  # PyMuPDF
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
    Convierte un PDF a PDF/A usando Ghostscript. 
    Rescata los archivos adjuntos previos y muestra alertas visuales del proceso.
    """
    gs_cmd = "gswin64c" if os.name == "nt" else "gs"
    part = "3" if level == "3b" else "2"
    conformance = "B"
    
    # =========================================================
    # 1. RESCATAR ADJUNTOS ORIGINALES
    # =========================================================
    attachments = []
    try:
        doc_original = fitz.open(stream=pdf_bytes, filetype="pdf")
        for name in doc_original.embfile_names():
            attachments.append((name, doc_original.embfile_get(name)))
        doc_original.close()
    except Exception as e:
        st.warning(f"Advertencia al leer adjuntos: {e}")
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_in:
        temp_in.write(pdf_bytes)
        temp_in_path = temp_in.name
        
    temp_out_path = temp_in_path.replace(".pdf", "_out.pdf")
    
    try:
        # INDICADOR VISUAL PARA EL USUARIO
        if attachments:
            st.info(f"🛠️ Se detectaron {len(attachments)} archivo(s) anexo(s). Procediendo a rescatarlos tras la recodificación...")
            
        # =========================================================
        # 2. EJECUTAR GHOSTSCRIPT
        # =========================================================
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
            error_details = process.stderr if process.stderr.strip() else process.stdout
            st.error(f"Fallo en Ghostscript. Detalles: {error_details}")
            return None
            
        # =========================================================
        # 3. RECONSTRUIR EL DOCUMENTO
        # =========================================================
        doc = fitz.open(temp_out_path)
        
        # Re-inyectar adjuntos
        if level == "3b" and attachments:
            for name, file_data in attachments:
                doc.embfile_add(name, file_data, filename=name)
            st.success("✅ Archivos anexos re-inyectados exitosamente en el documento final.")
        elif level == "2b" and attachments:
            st.warning("⚠️ Nota: Elegiste PDF/A-2b. Esta norma NO permite anexos. Los archivos embebidos han sido descartados definitivamente. Usa PDF/A-3b si deseas conservarlos.")
        
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
        
    except FileNotFoundError:
        st.error("⚠️ No se encontró Ghostscript.")
        return None
    except Exception as e:
        st.error(f"Error inesperado: {str(e)}")
        return None
    finally:
        if os.path.exists(temp_in_path): os.remove(temp_in_path)
        if os.path.exists(temp_out_path): os.remove(temp_out_path)

def get_file_hash(file_bytes):
    """Calcula el hash SHA-256."""
    return hashlib.sha256(file_bytes).hexdigest()

def validate_pdfa(pdf_bytes):
    """
    Valida heurísticamente el cumplimiento leyendo los metadatos XML del PDF.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    xml = doc.get_xml_metadata()
    if not xml:
        return False, "No es PDF/A (Sin metadatos XML)"
    
    part_match = re.search(r'<pdfaid:part>(\d)</pdfaid:part>', xml)
    conf_match = re.search(r'<pdfaid:conformance>([A-Z]+)</pdfaid:conformance>', xml)
    
    if part_match and conf_match:
        level = f"PDF/A-{part_match.group(1)}{conf_match.group(1)}"
        return True, level
    
    return False, "No detectado"

def generate_report(file_name, file_bytes):
    """Genera un diccionario con el informe de preservación."""
    sha256_hash = get_file_hash(file_bytes)
    size_kb = len(file_bytes) / 1024
    date_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    is_valid, pdfa_level = validate_pdfa(file_bytes)
    
    report = {
        "Nombre del Archivo": file_name,
        "Hash SHA-256": sha256_hash,
        "Tamaño": f"{size_kb:.2f} KB",
        "Fecha de Análisis": date_now,
        "Cumplimiento PDF/A": "Válido" if is_valid else "No Cumple",
        "Nivel PDF/A Detectado": pdfa_level
    }
    return report

# ==========================================
# INTERFAZ WEB CON STREAMLIT
# ==========================================

st.set_page_config(page_title="Gestor de Preservación PDF", layout="wide")

# CAMBIO AQUÍ: Diseño de columnas invertido. 
# 1/4 (25%) para el menú a la izquierda, 3/4 (75%) para contenido principal a la derecha.
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
    st.title("📄 Herramienta de Preservación Documental PDF")
    st.write("Sube tu documento principal y selecciona una acción en el menú de la izquierda.")
    
    main_pdf = st.file_uploader("Sube el archivo PDF principal", type=["pdf"])

    if main_pdf is not None:
        pdf_bytes = main_pdf.read()

        # Opciones 1 y 2: Anexar archivos
        if "1. Anexar Excel" in menu_option or "2. Adjuntar Cualquier" in menu_option:
            st.subheader(menu_option)
            allowed_types = ["xlsx", "xls"] if "Excel" in menu_option else None
            attachment = st.file_uploader("Sube el archivo a adjuntar", type=allowed_types)
            
            if attachment and st.button("Embeber Archivo"):
                with st.spinner("Embebiendo archivo..."):
                    result_pdf = embed_file_in_pdf(pdf_bytes, attachment.read(), attachment.name)
                    st.success("¡Archivo adjuntado con éxito!")
                    st.download_button(
                        label="Descargar PDF con Anexo",
                        data=result_pdf,
                        file_name=f"con_anexo_{main_pdf.name}",
                        mime="application/pdf"
                    )

        # Opciones 3 y 4: Convertir a PDF/A
        elif "Convertir a PDF/A" in menu_option:
            st.subheader(menu_option)
            level = "3b" if "3b" in menu_option else "2b"
            
            st.info(f"💡 Nota: Se recomienda PDF/A-3b si tu documento lleva anexos embebidos.")
            if st.button(f"Ejecutar Conversión a {level}"):
                with st.spinner("Procesando conversión (requiere Ghostscript)..."):
                    result_pdfa = convert_to_pdfa(pdf_bytes, level=level)
                    if result_pdfa:
                        st.success(f"Conversión a PDF/A-{level} exitosa.")
                        st.download_button(
                            label=f"Descargar PDF/A-{level}",
                            data=result_pdfa,
                            file_name=f"pdfa_{level}_{main_pdf.name}",
                            mime="application/pdf"
                        )

        # Opción 5: Informe y Validación
        elif "5. Generar Informe" in menu_option:
            st.subheader("📊 Informe de Preservación y Verificación")
            if st.button("Ejecutar Análisis"):
                with st.spinner("Analizando documento..."):
                    report = generate_report(main_pdf.name, pdf_bytes)
                    
                    # Mostrar resultados visualmente
                    st.markdown("### Resultados del Análisis")
                    for key, value in report.items():
                        st.markdown(f"**{key}:** {value}")
                    
                    # Formatear informe para descarga
                    report_text = "\n".join([f"{k}: {v}" for k, v in report.items()])
                    st.download_button(
                        label="Descargar Informe (.txt)",
                        data=report_text,
                        file_name=f"informe_{main_pdf.name}.txt",
                        mime="text/plain"
                    )
