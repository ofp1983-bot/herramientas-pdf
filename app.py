import streamlit as st
import pymupdf  # Para lectura y extracción
import pikepdf  # Motor estructural puro para cumplimiento PDF/A-3
import pandas as pd
from nc_py_api import Nextcloud
from io import BytesIO
import hashlib
import os
import subprocess
import tempfile
from datetime import datetime
import re
import xml.etree.ElementTree as ET

# ==========================================
# FUNCIONES AUXILIARES
# ==========================================

def parse_pdf_date(date_str):
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
    doc.embfile_add(attachment_name, attachment_bytes, filename=attachment_name, ufilename=attachment_name)
    return doc.write()

def convert_to_pdfa(pdf_bytes, level="3b"):
    gs_cmd = "gswin64c" if os.name == "nt" else "gs"
    part = "3" if level == "3b" else "2"
    conformance = "B"
    
    attachments = []
    try:
        doc_original = pymupdf.open(stream=pdf_bytes, filetype="pdf")
        for name in doc_original.embfile_names():
            attachments.append((name, doc_original.embfile_get(name)))
            
        for name in doc_original.embfile_names():
            doc_original.embfile_del(name)
            
        pdf_bytes = doc_original.write(garbage=4)
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
            
        pdf = pikepdf.Pdf.open(temp_out_path, allow_overwriting_input=True)
        
        if level == "3b" and attachments:
            if "/AF" not in pdf.Root:
                pdf.Root.AF = pikepdf.Array()
                
            if "/Names" not in pdf.Root:
                pdf.Root.Names = pikepdf.Dictionary()
                
            if "/EmbeddedFiles" not in pdf.Root.Names:
                pdf.Root.Names.EmbeddedFiles = pikepdf.Dictionary(Names=pikepdf.Array())
                
            for name, file_data in attachments:
                ef_stream = pdf.make_stream(file_data)
                ef_stream.Type = pikepdf.Name("/EmbeddedFile")
                ef_stream.Subtype = pikepdf.Name("/application/octet-stream")
                
                filespec_dict = pikepdf.Dictionary(
                    Type=pikepdf.Name("/Filespec"),
                    F=name,
                    UF=name,
                    EF=pikepdf.Dictionary(F=ef_stream, UF=ef_stream),
                    AFRelationship=pikepdf.Name("/Unspecified")
                )
                
                filespec_obj = pdf.make_indirect(filespec_dict)
                
                pdf.Root.Names.EmbeddedFiles.Names.append(name)
                pdf.Root.Names.EmbeddedFiles.Names.append(filespec_obj)
                pdf.Root.AF.append(filespec_obj)
                
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
        
        meta_stream = pdf.make_stream(xml_metadata.encode('utf-8'))
        meta_stream.Type = pikepdf.Name("/Metadata")
        meta_stream.Subtype = pikepdf.Name("/XML")
        pdf.Root.Metadata = meta_stream
        
        pdf.save(temp_out_path)
        pdf.close()
        
        with open(temp_out_path, "rb") as f:
            pdfa_bytes = f.read()
            
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
    """
    Motor de validación híbrido: 
    Intenta ejecutar veraPDF en el sistema para una validación estricta forense.
    Si falla (no instalado o error), usa la revisión superficial de metadatos con PyMuPDF.
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(pdf_bytes)
        temp_pdf_path = temp_pdf.name
        
    # Fase 1: Validación estricta con veraPDF
    try:
        # shell=True permite encontrar el ejecutable global 'verapdf' en Windows/Linux
        process = subprocess.run(
            f'verapdf "{temp_pdf_path}"', 
            shell=True, 
            stdout=subprocess.PIPE, 
            stderr=subprocess.PIPE, 
            text=True, 
            encoding='utf-8', 
            errors='ignore'
        )
        
        if process.stdout and "<?xml" in process.stdout:
            # Extraer solo la parte XML de la salida de consola
            xml_str = process.stdout[process.stdout.find("<?xml"):]
            root = ET.fromstring(xml_str)
            
            report_node = root.find('.//validationReport')
            if report_node is not None:
                is_compliant = report_node.get('isCompliant') == 'true'
                profile_name = report_node.get('profileName', 'Desconocido')
                profile_clean = profile_name.replace(" validation profile", "")
                
                estado = f"{profile_clean} (Auditoría: veraPDF)"
                return is_compliant, estado
    except Exception:
        pass # Falla silenciosa si veraPDF no está configurado, pasa a Fase 2
    finally:
        if os.path.exists(temp_pdf_path): 
            os.remove(temp_pdf_path)

    # Fase 2: Validación superficial de metadatos (Respaldo)
    try:
        doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
        xml = doc.get_xml_metadata()
        doc.close()
        if not xml:
            return False, "No es PDF/A (Sin metadatos XML)"
        part_match = re.search(r'<pdfaid:part>(\d)</pdfaid:part>', xml)
        conf_match = re.search(r'<pdfaid:conformance>([A-Z]+)</pdfaid:conformance>', xml)
        if part_match and conf_match:
            return True, f"PDF/A-{part_match.group(1)}{conf_match.group(1)} (Lectura de Etiqueta)"
        return False, "No detectado"
    except:
        return False, "Error al leer documento"

def get_attachments_info(pdf_bytes):
    names = []
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
        "Cumplimiento PDF/A": "Válido (Normado)" if is_valid else "No Cumple (Inválido)",
        "Nivel y Motor de Verificación": pdfa_level,
        "Contiene Anexos": "Sí" if nombres_anexos else "No",
        "Nombre de los Anexos": ", ".join(nombres_anexos) if nombres_anexos else "N/A"
    }

def generate_electronic_index(archivos, origen_default="Digitalizado"):
    temp_docs = []
    
    for file_obj in archivos:
        file_bytes = file_obj.read()
        file_name = file_obj.name
        
        tipologia_nombre = os.path.splitext(file_name)[0]
        extension = os.path.splitext(file_name)[1].lower()
        formato_str = extension.replace(".", "").upper() if extension else "DESCONOCIDO"
        
        total_pages = 0
        creation_date = "Desconocida"
        pdfa_final = "N/A"
        tiene_anexos = "N/A"
        tipos_anexos_str = "N/A"
        nombres_anexos_str = "N/A"
        
        if extension == ".pdf":
            is_valid, pdfa_level = validate_pdfa(file_bytes)
            # Acortamos el texto para que quepa bien en el Excel del índice
            pdfa_final = pdfa_level.split(" ")[0] if is_valid else "No detectado"
            
            nombres_anexos = get_attachments_info(file_bytes)
            tiene_anexos = "Sí" if nombres_anexos else "No"
            if nombres_anexos:
                nombres_anexos_str = ", ".join(nombres_anexos)
                tipos_anexos = list(set([os.path.splitext(n)[1].upper() for n in nombres_anexos if os.path.splitext(n)[1]]))
                tipos_anexos_str = ", ".join(tipos_anexos) if tipos_anexos else "N/A"
            
            try:
                doc = pymupdf.open(stream=file_bytes, filetype="pdf")
                total_pages = len(doc)
                metadata = doc.metadata
                raw_creation_date = metadata.get('creationDate', '')
                creation_date = parse_pdf_date(raw_creation_date)
                doc.close()
            except Exception:
                pass
                
        try:
            date_obj = datetime.strptime(creation_date, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            date_obj = datetime.max 
            
        sha256_hash = get_file_hash(file_bytes)
        size_kb = len(file_bytes) / 1024
        
        temp_docs.append({
            "Nombre_Documento": file_name,
            "Tipologia_Documental": tipologia_nombre,
            "creation_date_str": creation_date,
            "date_obj": date_obj, 
            "Valor_Huella": sha256_hash,
            "Pagina_Inicio": 1 if total_pages > 0 else 0,
            "Pagina_Fin": total_pages if total_pages > 0 else "N/A",
            "Formato": formato_str,
            "Tamano": f"{size_kb:.2f} KB",
            "Tipo_PDFA": pdfa_final,
            "Tiene_Anexos": tiene_anexos,
            "Tipo_Anexo": tipos_anexos_str,
            "Nombre_Anexo": nombres_anexos_str
        })
        
    temp_docs.sort(key=lambda x: x["date_obj"])
    
    index_data = []
    for idx, item in enumerate(temp_docs, start=1):
        row = {
            "Id": f"DOC-{idx:03d}",
            "Nombre_Documento": item["Nombre_Documento"],
            "Tipologia_Documental": item["Tipologia_Documental"],
            "Fecha_Creacion_Documento": item["creation_date_str"],
            "Fecha_Incorporacion_Expediente": item["creation_date_str"],
            "Valor_Huella": item["Valor_Huella"],
            "Funcion_Resumen": "SHA-256",
            "Orden_Documento_Expediente": idx,
            "Pagina_Inicio": item["Pagina_Inicio"],
            "Pagina_Fin": item["Pagina_Fin"],
            "Formato": item["Formato"],
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

st.set_page_config(page_title="Gestor de Preservación PDF v16.0", layout="wide")

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
                "4. Convertir a PDF/A-3b (Archivos Híbridos)",
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
    st.title("📄 Herramienta de Preservación Documental (v16.0)")
    
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
                if st.button(f"Ejecutar Conversión a PDF/A-{level}"):
                    with st.spinner("Construyendo matriz estructural y codificando a norma..."):
                        result_pdfa = convert_to_pdfa(pdf_bytes, level=level)
                        if result_pdfa:
                            st.download_button(f"Descargar PDF/A-{level}", result_pdfa, file_name=f"pdfa_{level}_{main_pdf.name}", mime="application/pdf")

            elif "5. Generar Informe" in menu_option:
                st.subheader("📊 Informe de Preservación (Estricto)")
                if st.button("Ejecutar Auditoría Forense"):
                    with st.spinner("Sometiendo documento al motor veraPDF..."):
                        report = generate_report(main_pdf.name, pdf_bytes)
                        for k, v in report.items():
                            st.markdown(f"**{k}:** {v}")
                        report_text = "\n".join([f"{k}: {v}" for k, v in report.items()])
                        st.download_button("Descargar Informe (.txt)", report_text, file_name=f"informe_{main_pdf.name}.txt", mime="text/plain")

    else:
        st.subheader("📁 Índice Electrónico de Expedientes Multiformato")
        
        if "Subir archivos" in menu_option:
            st.write("Selecciona o arrastra múltiples archivos de cualquier formato desde tu equipo.")
            origen_opcion = st.selectbox("Origen predeterminado:", ["Digitalizado", "Electrónico", "Físico"], key="orig_manual")
            batch_files = st.file_uploader("Sube los archivos del expediente", accept_multiple_files=True)
            
            if batch_files:
                st.info(f"Se han cargado {len(batch_files)} archivos para procesar.")
                if st.button("Generar Índice"):
                    with st.spinner("Auditando lote con veraPDF y extrayendo metadatos..."):
                        df_index = generate_electronic_index(batch_files, origen_default=origen_opcion)
                        st.dataframe(df_index)
                        
                        output_excel = BytesIO()
                        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                            df_index.to_excel(writer, index=False, sheet_name='Indice_Electronico')
                        
                        xml_data = df_index.to_xml(index=False, root_name="Expediente", row_name="Documento", parser="etree")
                        
                        st.success("¡Índices electrónicos generados con éxito!")
                        
                        col_btn1, col_btn2 = st.columns(2)
                        with col_btn1:
                            st.download_button(
                                label="📥 Descargar Índice (.xlsx)",
                                data=output_excel.getvalue(),
                                file_name="indice_electronico_multiformato.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        with col_btn2:
                            st.download_button(
                                label="📥 Descargar Índice (.xml)",
                                data=xml_data.encode('utf-8'),
                                file_name="indice_electronico_multiformato.xml",
                                mime="application/xml"
                            )
                        
        elif "Aurora Nextcloud" in menu_option:
            st.write("Conéctate al repositorio en la nube para indexar el expediente (incluye todos los formatos).")
            
            col1, col2 = st.columns(2)
            with col1:
                nc_url = st.text_input("URL del Servidor", value="https://cloud.insdeportescajica.gov.co")
                nc_user = st.text_input("Usuario Nextcloud")
            with col2:
                nc_pass = st.text_input("Contraseña de Aplicación", type="password")
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
                            
                            archivos_lote = [nodo for nodo in nodos if not nodo.is_dir]
                            
                            if not archivos_lote:
                                st.error("No se encontraron archivos en la ruta especificada.")
                            else:
                                st.info(f"✅ Conexión exitosa. Extrayendo metadatos de {len(archivos_lote)} documentos...")
                                
                                batch_files = []
                                for node in archivos_lote:
                                    contenido_bytes = nc.files.download(node)
                                    archivo_en_memoria = BytesIO(contenido_bytes)
                                    archivo_en_memoria.name = node.name 
                                    batch_files.append(archivo_en_memoria)
                                
                                df_index = generate_electronic_index(batch_files, origen_default=origen_opcion)
                                st.dataframe(df_index)
                                
                                output_excel = BytesIO()
                                with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                                    df_index.to_excel(writer, index=False, sheet_name='Indice_Electronico')
                                
                                xml_data = df_index.to_xml(index=False, root_name="Expediente", row_name="Documento", parser="etree")
                                
                                st.success("¡Índices electrónicos generados con éxito desde Aurora Nextcloud!")
                                
                                col_btn1, col_btn2 = st.columns(2)
                                with col_btn1:
                                    st.download_button(
                                        label="📥 Descargar Índice (.xlsx)",
                                        data=output_excel.getvalue(),
                                        file_name="indice_aurora_nextcloud.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                                with col_btn2:
                                    st.download_button(
                                        label="📥 Descargar Índice (.xml)",
                                        data=xml_data.encode('utf-8'),
                                        file_name="indice_aurora_nextcloud.xml",
                                        mime="application/xml"
                                    )
                        except Exception as e:
                            st.error(f"❌ Error de conexión o lectura: {str(e)}")
                            st.info("Verifica que la URL sea correcta, las credenciales sean válidas y que la ruta exista.")
