import streamlit as st
import pymupdf  # Reemplazo oficial de fitz
import pypdf
import pandas as pd
from nc_py_api import Nextcloud
from io import BytesIO
import hashlib
import os
import subprocess
import tempfile
from datetime import datetime
import re

# ==========================================
# FUNCIONES AUXILIARES
# ==========================================

def parse_pdf_date(date_str):
    """Convierte la fecha interna del PDF a YYYY-MM-DD HH:MM:SS para lectura y ordenamiento"""
    if not date_str:
        return "Desconocida"
    match = re.search(r"D:(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})", date_str)
    if match:
        y, m, d, h, mn, s = match.groups()
        return f"{y}-{m}-{d} {h}:{mn}:{s}"
    return str(date_str) 

# ==========================================
# FUNCIONES PRINCIPALES
# ==========================================

def embed_file_in_pdf(pdf_bytes, attachment_bytes, attachment_name):
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    doc.embfile_add(attachment_name, attachment_bytes, filename=attachment_name)
    return doc.write()

def convert_to_pdfa(pdf_bytes, level="2b"):
    gs_cmd = "gswin64c" if os.name == "nt" else "gs"
    part = "3" if level == "3b" else "2"
    conformance = "B"
    
    attachments = []
    try:
        reader = pypdf.PdfReader(BytesIO(pdf_bytes))
        if reader.attachments:
            for name, byte_list in reader.attachments.items():
                for file_data in byte_list:
                    attachments.append((name, file_data))
    except Exception:
        pass
        
    if not attachments:
        try:
            doc_original = pymupdf.open(stream=pdf_bytes, filetype="pdf")
            for name in doc_original.embfile_names():
                attachments.append((name, doc_original.embfile_get(name)))
            doc_original.close()
        except Exception:
            pass
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_in:
        temp_in.write(pdf_bytes)
        temp_in_path = temp_in.name
        
    temp_out_path = temp_in_path.replace(".pdf", "_out.pdf")
    
    try:
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
            st.error(f"Fallo en Ghostscript: {process.stderr}")
            return None
            
        doc = pymupdf.open(temp_out_path)
        if level == "3b" and attachments:
            for name, file_data in attachments:
                doc.embfile_add(name, file_data, filename=name)
        
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
        st.error(f"Error: {str(e)}")
        return None
    finally:
        if os.path.exists(temp_in_path): os.remove(temp_in_path)
        if os.path.exists(temp_out_path): os.remove(temp_out_path)

def get_file_hash(file_bytes):
    return hashlib.sha256(file_bytes).hexdigest()

def validate_pdfa(pdf_bytes):
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    xml = doc.get_xml_metadata()
    if not xml:
        return False, "No es PDF/A (Sin metadatos XML)"
    part_match = re.search(r'<pdfaid:part>(\d)</pdfaid:part>', xml)
    conf_match = re.search(r'<pdfaid:conformance>([A-Z]+)</pdfaid:conformance>', xml)
    if part_match and conf_match:
        return True, f"PDF/A-{part_match.group(1)}{conf_match.group(1)}"
    return False, "No detectado"

def get_attachments_info(pdf_bytes):
    names = []
    try:
        reader = pypdf.PdfReader(BytesIO(pdf_bytes))
        if reader.attachments:
            for name in reader.attachments.keys():
                names.append(name)
    except Exception:
        pass
    if not names:
        try:
            doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
            names = list(doc.embfile_names())
            doc.close()
        except Exception:
            pass
    return list(set(names))

def generate_report(file_name, file_bytes):
    sha256_hash = get_file_hash(file_bytes)
    size_kb = len(file_bytes) / 1024
    is_valid, pdfa_level = validate_pdfa(file_bytes)
    nombres_anexos = get_attachments_info(file_bytes)
    
    return {
        "Nombre del Archivo": file_name,
        "Hash SHA-256": sha256_hash,
        "Tamaño": f"{size_kb:.2f} KB",
        "Fecha de Análisis": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Cumplimiento PDF/A": "Válido" if is_valid else "No Cumple",
        "Nivel PDF/A Detectado": pdfa_level,
        "Contiene Anexos": "Sí" if nombres_anexos else "No",
        "Nombre de los Anexos": ", ".join(nombres_anexos) if nombres_anexos else "N/A"
    }

def generate_electronic_index(archivos, origen_default="Digitalizado"):
    """Procesa un lote de PDFs, los ordena cronológicamente y genera un DataFrame con el índice electrónico."""
    temp_docs = []
    
    # FASE 1: Extracción de metadatos de todos los archivos
    for file_obj in archivos:
        file_bytes = file_obj.read()
        
        # Extracción de PDF/A y Anexos
        is_valid, pdfa_level = validate_pdfa(file_bytes)
        nombres_anexos = get_attachments_info(file_bytes)
        
        tiene_anexos = "Sí" if nombres_anexos else "No"
        nombres_anexos_str = ", ".join(nombres_anexos) if nombres_anexos else "N/A"
        
        # Extraer extensiones para el "Tipo de anexo" (ej. .XLSX, .XML)
        tipos_anexos = list(set([os.path.splitext(name)[1].upper() for name in nombres_anexos if os.path.splitext(name)[1]]))
        tipos_anexos_str = ", ".join(tipos_anexos) if tipos_anexos else "N/A"
        
        try:
            doc = pymupdf.open(stream=file_bytes, filetype="pdf")
            total_pages = len(doc)
            metadata = doc.metadata
            raw_creation_date = metadata.get('creationDate', '')
            creation_date = parse_pdf_date(raw_creation_date)
            doc.close()
        except Exception:
            total_pages = 0
            creation_date = "Desconocida"
            
        # Objeto fecha temporal para usar en el ordenamiento interno
        try:
            date_obj = datetime.strptime(creation_date, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            date_obj = datetime.max # Si no hay fecha, se envían al final del índice
            
        sha256_hash = get_file_hash(file_bytes)
        size_kb = len(file_bytes) / 1024
        
        temp_docs.append({
            "Nombre_Documento": file_obj.name,
            "creation_date_str": creation_date,
            "date_obj": date_obj, # Campo oculto solo para ordenar
            "Valor_Huella": sha256_hash,
            "Pagina_Inicio": 1 if total_pages > 0 else 0,
            "Pagina_Fin": total_pages,
            "Tamano": f"{size_kb:.2f} KB",
            "Tipo_PDFA": pdfa_level if is_valid else "No detectado",
            "Tiene_Anexos": tiene_anexos,
            "Tipo_Anexo": tipos_anexos_str,
            "Nombre_Anexo": nombres_anexos_str
        })
        
    # FASE 2: Ordenamiento Cronológico (Ascendente)
    temp_docs.sort(key=lambda x: x["date_obj"])
    
    # FASE 3: Asignación de ID y estructuración del DataFrame
    index_data = []
    for idx, item in enumerate(temp_docs, start=1):
        row = {
            "Id": f"DOC-{idx:03d}",
            "Nombre_Documento": item["Nombre_Documento"],
            "Tipologia_Documental": "Documento General",
            "Fecha_Creacion_Documento": item["creation_date_str"],
            "Fecha_Incorporacion_Expediente": item["creation_date_str"],
            "Valor_Huella": item["Valor_Huella"],
            "Funcion_Resumen": "SHA-256",
            "Orden_Documento_Expediente": idx,
            "Pagina_Inicio": item["Pagina_Inicio"],
            "Pagina_Fin": item["Pagina_Fin"],
            "Formato": "PDF",
            "Tamano": item["Tamano"],
            "Origen": origen_default,
            "Tipo_PDFA": item["Tipo_PDFA"],
            "Tiene_Anexos": item["Tiene_Anexos"],
            "Tipo_Anexo": item["Tipo_Anexo"],
            "Nombre_Anexo": item["Nombre_Anexo"]
        }
        index_data.append(row)
        
    return pd.DataFrame(index_data)

# ==========================================
# INTERFAZ WEB CON STREAMLIT
# ==========================================

st.set_page_config(page_title="Gestor de Preservación PDF v5.4", layout="wide")

col_menu, col_main = st.columns([1, 3])

with col_menu:
    st.markdown("### ⚙️ Menú Principal")
    
    modulo = st.selectbox(
        "Selecciona el módulo de trabajo:",
        ["📄 Documentos Individuales", "📁 Procesamiento por Lotes"]
    )
    
    st.markdown("---")
    
    if modulo == "📄 Documentos Individuales":
        st.markdown("#### Tareas Individuales")
        menu_option = st.radio(
            "Acción a realizar:",
            (
                "1. Anexar Excel a PDF",
                "2. Adjuntar Cualquier Archivo",
                "3. Convertir a PDF/A-2b",
                "4. Convertir a PDF/A-3b",
                "5. Generar Informe de Preservación"
            )
        )
    else:
        st.markdown("#### Creación de Índices")
        menu_option = st.radio(
            "Origen de los archivos:",
            (
                "1. Subir archivos manualmente",
                "2. Conectar a Aurora Nextcloud"
            )
        )

with col_main:
    st.title("📄 Herramienta de Preservación Documental (v5.4)")
    
    if modulo == "📄 Documentos Individuales":
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
                        st.download_button("Descargar PDF con Anexo", result_pdf, file_name=f"con_anexo_{main_pdf.name}", mime="application/pdf")

            elif "Convertir a PDF/A" in menu_option:
                st.subheader(menu_option)
                level = "3b" if "3b" in menu_option else "2b"
                if st.button(f"Ejecutar Conversión a {level}"):
                    with st.spinner("Procesando conversión..."):
                        result_pdfa = convert_to_pdfa(pdf_bytes, level=level)
                        if result_pdfa:
                            st.download_button(f"Descargar PDF/A-{level}", result_pdfa, file_name=f"pdfa_{level}_{main_pdf.name}", mime="application/pdf")

            elif "5. Generar Informe" in menu_option:
                st.subheader("📊 Informe de Preservación")
                if st.button("Ejecutar Análisis"):
                    with st.spinner("Analizando..."):
                        report = generate_report(main_pdf.name, pdf_bytes)
                        for k, v in report.items():
                            st.markdown(f"**{k}:** {v}")
                        report_text = "\n".join([f"{k}: {v}" for k, v in report.items()])
                        st.download_button("Descargar Informe (.txt)", report_text, file_name=f"informe_{main_pdf.name}.txt", mime="text/plain")

    else:
        st.subheader("📁 Índice Electrónico de Expedientes")
        
        if "Subir archivos" in menu_option:
            st.write("Selecciona o arrastra múltiples archivos PDF desde tu equipo.")
            origen_opcion = st.selectbox("Origen predeterminado:", ["Digitalizado", "Electrónico", "Físico"], key="orig_manual")
            batch_files = st.file_uploader("Sube los archivos PDF del expediente", type=["pdf"], accept_multiple_files=True)
            
            if batch_files:
                st.info(f"Se han cargado {len(batch_files)} archivos para procesar.")
                if st.button("Generar Índice en Excel"):
                    with st.spinner("Procesando lote, ordenando por fecha y extrayendo metadatos..."):
                        df_index = generate_electronic_index(batch_files, origen_default=origen_opcion)
                        st.dataframe(df_index)
                        
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_index.to_excel(writer, index=False, sheet_name='Indice_Electronico')
                        
                        st.success("¡Índice electrónico generado con éxito!")
                        st.download_button(
                            label="📥 Descargar Índice Electrónico (.xlsx)",
                            data=output.getvalue(),
                            file_name="indice_electronico_ordenado.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
        elif "Aurora Nextcloud" in menu_option:
            st.write("Conéctate directamente al repositorio en la nube para indexar el expediente sin descargar los archivos localmente.")
            
            col1, col2 = st.columns(2)
            with col1:
                nc_url = st.text_input("URL del Servidor", value="https://cloud.insdeportescajica.gov.co")
                nc_user = st.text_input("Usuario Nextcloud")
            with col2:
                nc_pass = st.text_input("Contraseña de Aplicación", type="password", help="Genera esta contraseña en la configuración de seguridad de Nextcloud.")
                nc_folder = st.text_input("Ruta de la carpeta", placeholder="/Expedientes/2026/Contrato_01")
                
            origen_opcion = st.selectbox("Origen predeterminado:", ["Electrónico", "Digitalizado", "Físico"], key="orig_cloud")
            
            if st.button("Conectar y Generar Índice"):
                if not all([nc_url, nc_user, nc_pass, nc_folder]):
                    st.warning("⚠️ Por favor, completa todos los campos de conexión.")
                else:
                    with st.spinner("Conectando con Aurora Nextcloud..."):
                        try:
                            nc = Nextcloud(nextcloud_url=nc_url, nc_auth_user=nc_user, nc_auth_pass=nc_pass)
                            nodos = nc.files.listdir(nc_folder)
                            archivos_pdf = [nodo for nodo in nodos if nodo.name.lower().endswith('.pdf')]
                            
                            if not archivos_pdf:
                                st.error("No se encontraron archivos PDF en la ruta especificada.")
                            else:
                                st.info(f"✅ Conexión exitosa. Extrayendo metadatos de {len(archivos_pdf)} documentos (esto puede tardar unos momentos)...")
                                
                                batch_files = []
                                for pdf_node in archivos_pdf:
                                    contenido_bytes = nc.files.download(pdf_node)
                                    archivo_en_memoria = BytesIO(contenido_bytes)
                                    archivo_en_memoria.name = pdf_node.name 
                                    batch_files.append(archivo_en_memoria)
                                
                                df_index = generate_electronic_index(batch_files, origen_default=origen_opcion)
                                st.dataframe(df_index)
                                
                                output = BytesIO()
                                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                    df_index.to_excel(writer, index=False, sheet_name='Indice_Electronico')
                                
                                st.success("¡Índice electrónico ordenado y generado con éxito desde Aurora Nextcloud!")
                                st.download_button(
                                    label="📥 Descargar Índice Electrónico (.xlsx)",
                                    data=output.getvalue(),
                                    file_name="indice_aurora_nextcloud_ordenado.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                        except Exception as e:
                            st.error(f"❌ Error de conexión o lectura: {str(e)}")
                            st.info("Verifica que la URL sea correcta, las credenciales sean válidas y que la ruta de la carpeta empiece con '/'")
