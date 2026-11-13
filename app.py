import streamlit as st
import pymupdf  
import pikepdf  
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
from pathlib import Path

# Librerías para comunicación nativa con Microsoft Word en Windows
try:
    import win32com.client
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

# ==========================================
# FUNCIONES AUXILIARES
# ==========================================

def parse_pdf_date(date_str):
    if not date_str:
        return None
    match = re.search(r"D:(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})", date_str)
    if match:
        y, m, d, h, mn, s = match.groups()
        try:
            return datetime(int(y), int(m), int(d), int(h), int(mn), int(s))
        except ValueError:
            return None
    return None

def get_file_hash(file_bytes):
    return hashlib.sha256(file_bytes).hexdigest()

def validate_pdfa(pdf_bytes):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(pdf_bytes)
        temp_pdf_path = temp_pdf.name
        
    try:
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
        pass 
    finally:
        if os.path.exists(temp_pdf_path): 
            os.remove(temp_pdf_path)

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

# ==========================================
# FUNCIONES PRINCIPALES
# ==========================================

def convert_word_to_pdfa_native(docx_bytes, original_filename):
    if not WIN32_AVAILABLE:
        st.error("La librería pywin32 no está instalada. Ejecuta: pip install pywin32")
        return None

    pythoncom.CoInitialize()
    word = None
    doc = None
    docx_path = ""
    pdf_path = ""
    
    try:
        fd_docx, docx_path = tempfile.mkstemp(suffix=".docx")
        os.close(fd_docx) 
        
        fd_pdf, pdf_path = tempfile.mkstemp(suffix=".pdf")
        os.close(fd_pdf) 
        
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)
            
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  
        word.AutomationSecurity = 3  
        
        doc = word.Documents.Open(
            FileName=docx_path, 
            ConfirmConversions=False, 
            ReadOnly=False, 
            AddToRecentFiles=False, 
            Visible=False
        )
        
        doc.ExportAsFixedFormat(
            OutputFileName=pdf_path,
            ExportFormat=17, 
            OpenAfterExport=False,
            OptimizeFor=0, 
            Range=0, 
            From=1,
            To=1,
            Item=0, 
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=1, 
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=True 
        )
        
        doc.Close(False) 
        word.Quit(False)
        
        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()
            
        return pdf_bytes
        
    except Exception as e:
        st.error(f"Error en el puente con Microsoft Word: {str(e)}")
        return None
    finally:
        try:
            if doc: doc.Close(False)
            if word: word.Quit(False)
        except:
            pass
        pythoncom.CoUninitialize()
        
        if os.path.exists(docx_path): 
            try: os.remove(docx_path) 
            except: pass
        if os.path.exists(pdf_path): 
            try: os.remove(pdf_path) 
            except: pass

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
            "-sColorConversionStrategy=RGB",
            "-sDEVICE=pdfwrite",
            f"-sOutputFile={temp_out_path}",
            temp_in_path
        ]
        
        process = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if process.returncode != 0:
            st.error(f"Fallo en Ghostscript: {process.stderr}")
            return None
            
        pdf = pikepdf.Pdf.open(temp_out_path, allow_overwriting_input=True)
        
        cal_rgb = pikepdf.Array([
            pikepdf.Name("/CalRGB"),
            pikepdf.Dictionary(
                WhitePoint=pikepdf.Array([0.9505, 1.0000, 1.0890]),
                Gamma=pikepdf.Array([2.2, 2.2, 2.2]),
                Matrix=pikepdf.Array([0.4124, 0.2126, 0.0193, 0.3576, 0.7152, 0.1192, 0.1805, 0.0722, 0.9505])
            )
        ])
        
        cal_gray = pikepdf.Array([
            pikepdf.Name("/CalGray"),
            pikepdf.Dictionary(
                WhitePoint=pikepdf.Array([0.9505, 1.0000, 1.0890]),
                Gamma=2.2
            )
        ])
        
        for page in pdf.pages:
            if "/Group" not in page:
                page.Group = pikepdf.Dictionary(
                    S=pikepdf.Name("/Transparency"),
                    CS=pikepdf.Name("/DeviceRGB")
                )
            else:
                if "/CS" not in page.Group:
                    page.Group.CS = pikepdf.Name("/DeviceRGB")
                    
            if "/Resources" not in page:
                page.Resources = pikepdf.Dictionary()
            if "/ColorSpace" not in page.Resources:
                page.Resources.ColorSpace = pikepdf.Dictionary()
                
            if "/DefaultRGB" not in page.Resources.ColorSpace:
                page.Resources.ColorSpace.DefaultRGB = cal_rgb
            if "/DefaultGray" not in page.Resources.ColorSpace:
                page.Resources.ColorSpace.DefaultGray = cal_gray
        
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
                creation_date = parse_pdf_date(raw_creation_date).strftime("%Y-%m-%d %H:%M:%S") if parse_pdf_date(raw_creation_date) else "Desconocida"
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

def generate_fuid_from_nextcloud(nc_url, nc_user, nc_pass, nc_folder):
    """Genera el Formato Único de Inventario Documental (FUID) conectándose a Nextcloud"""
    try:
        nc = Nextcloud(nextcloud_url=nc_url, nc_auth_user=nc_user, nc_auth_pass=nc_pass)
        
        # 1. Obtener nodos de la carpeta principal
        nodos_principales = nc.files.listdir(nc_folder)
        folder_name = nc_folder.strip("/").split("/")[-1]
        
        match = re.match(r"^([\w\-\.]+)\s+(.*)$", folder_name)
        if match:
            codigo_serie = match.group(1)
            nombre_serie = match.group(2)
        else:
            codigo_serie = "S.C."
            nombre_serie = folder_name

        inventory_rows = []
        
        # Filtrar subcarpetas (Unidades Documentales)
        subcarpetas = [n for n in nodos_principales if n.is_dir]
        if not subcarpetas:
            return None, "No se encontraron subcarpetas (unidades documentales) en la ruta de Nextcloud especificada."

        for idx, sub_node in enumerate(subcarpetas, start=1):
            unidad_documental = sub_node.name
            
            # Construir la ruta de la subcarpeta de manera segura utilizando su nombre
            sub_folder_path = f"{nc_folder.rstrip('/')}/{unidad_documental}"
            ubicacion_ruta = f"{nc_url}/index.php/apps/files/?dir={sub_folder_path}"
            
            # Listar contenido de la subcarpeta usando la ruta construida
            try:
                nodos_internos = nc.files.listdir(sub_folder_path)
            except:
                nodos_internos = []
                
            archivos_internos = [n for n in nodos_internos if not n.is_dir]
            cantidad_docs = len(archivos_internos)
            
            total_size_bytes = 0
            fechas = []
            
            for file_node in archivos_internos:
                try:
                    total_size_bytes += file_node.size
                except:
                    pass
                
                # Descargar en memoria para leer metadatos de fechas si es PDF
                if file_node.name.lower().endswith(".pdf"):
                    try:
                        file_bytes = nc.files.download(file_node)
                        doc = pymupdf.open(stream=file_bytes, filetype="pdf")
                        raw_date = doc.metadata.get('creationDate', '')
                        dt = parse_pdf_date(raw_date)
                        if dt:
                            fechas.append(dt)
                        doc.close()
                    except:
                        pass
                
                # Respaldar con fecha de modificación de Nextcloud si existe
                try:
                    if hasattr(file_node, 'mtime') and file_node.mtime:
                        fechas.append(datetime.fromtimestamp(file_node.mtime))
                except:
                    pass

            tamano_mb = f"{total_size_bytes / (1024 * 1024):.2f} MB"
            
            if fechas:
                fecha_inicial = min(fechas).strftime("%Y-%m-%d")
                fecha_final = max(fechas).strftime("%Y-%m-%d")
            else:
                fecha_inicial = "S.F."
                fecha_final = "S.F."

            row = {
                "NÚMERO DE ORDEN": idx,
                "CÓDIGO": codigo_serie,
                "NOMBRE DE LA SERIE, SUBSERIE O ASUNTO": nombre_serie,
                "NOMBRE DE LA UNIDAD DOCUMENTAL": unidad_documental,
                "FECHAS EXTREMAS - Inicial": fecha_inicial,
                "FECHAS EXTREMAS - Final": fecha_final,
                "SOPORTE - Electrónico": "X",
                "SOPORTE - Físico": "",
                "FÍSICO - Caja": "N/A",
                "FÍSICO - Carpeta": "N/A",
                "FÍSICO - Tomo/legajo/libro": "N/A",
                "FÍSICO - Número de folios": "N/A",
                "FÍSICO - Otro Tipo": "N/A",
                "FÍSICO - Otro Cantidad": "N/A",
                "ELECTRÓNICO - Ubicación": ubicacion_ruta,
                "ELECTRÓNICO - Cantidad de documentos": cantidad_docs,
                "ELECTRÓNICO - Tamaño": tamano_mb,
                "NOTAS": "Generado automáticamente desde Aurora Nextcloud"
            }
            inventory_rows.append(row)
            
        return pd.DataFrame(inventory_rows), None
    except Exception as e:
        return None, f"Error conectando a Nextcloud: {str(e)}"

# ==========================================
# INTERFAZ WEB CON STREAMLIT
# ==========================================

st.set_page_config(page_title="Gestor de Preservación PDF v21.0", layout="wide")

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
                "5. Generar Informe de Preservación",
                "6. Extraer Metadatos del PDF",
                "7. Transformar Word a PDF/A nativo"
            )
        )
    else:
        st.markdown("#### Procesamiento por Lotes")
        menu_option = st.radio(
            "Acción de lote:",
            (
                "1. Generar Índice Electrónico (Manual)",
                "2. Conectar a Aurora Nextcloud",
                "3. Inventario Documental (FUID - AGN)"
            )
        )

with col_main:
    st.title("📄 Herramienta de Preservación Documental (v21.0)")
    
    if modulo == "📄 Documentos Individuales":
        
        if "7. Transformar Word" in menu_option:
            st.subheader("📝 Exportación Nativa de Word a PDF/A")
            st.write("Genera un documento PDF preservado utilizando el motor original de Microsoft Word.")
            
            word_file = st.file_uploader("Sube el archivo Word (.docx, .doc)", type=["docx", "doc"])
            
            if word_file and st.button("Ejecutar Transformación Nativa"):
                with st.spinner("Conectando con Microsoft Word..."):
                    pdfa_result = convert_word_to_pdfa_native(word_file.read(), word_file.name)
                    
                    if pdfa_result:
                        st.success("¡Documento transformado exitosamente bajo el estándar de preservación!")
                        new_name = os.path.splitext(word_file.name)[0] + "_PDFA.pdf"
                        st.download_button(
                            label="Descargar PDF/A Nativo",
                            data=pdfa_result,
                            file_name=new_name,
                            mime="application/pdf"
                        )
        
        else:
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
                    st.subheader("📊 Informe de Preservación")
                    if st.button("Ejecutar Análisis"):
                        with st.spinner("Analizando..."):
                            report = generate_report(main_pdf.name, pdf_bytes)
                            for k, v in report.items():
                                st.markdown(f"**{k}:** {v}")
                            report_text = "\n".join([f"{k}: {v}" for k, v in report.items()])
                            st.download_button("Descargar Informe (.txt)", report_text, file_name=f"informe_{main_pdf.name}.txt", mime="text/plain")

                elif "6. Extraer Metadatos" in menu_option:
                    st.subheader("🗂️ Extracción Forense Absoluta de Metadatos")
                    if st.button("Ejecutar Escáner Profundo"):
                        with st.spinner("Analizando diccionarios y estructura interna..."):
                            doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
                            pdf_pike = pikepdf.Pdf.open(BytesIO(pdf_bytes))
                            
                            st.markdown("### 📄 1. Propiedades Físicas y de Seguridad")
                            col1, col2 = st.columns(2)
                            with col1:
                                st.write(f"**Total de Páginas:** {doc.page_count}")
                                st.write(f"**Versión PDF:** {doc.metadata.get('format', 'Desconocido')}")
                            with col2:
                                st.write(f"**Cifrado:** {'Sí (Protegido)' if doc.is_encrypted else 'No'}")
                                st.write(f"**Anexos detectados:** {doc.embfile_count()}")

                            st.markdown("---")
                            st.markdown("### 🏷️ 2. Diccionario de Información (`/Info`)")
                            info_dict = {}
                            if hasattr(pdf_pike, 'docinfo'):
                                for key, value in pdf_pike.docinfo.items():
                                    info_dict[str(key)] = str(value)
                            
                            if info_dict:
                                st.json(info_dict)
                            else:
                                st.info("El diccionario /Info está vacío o no existe en este documento.")

                            st.markdown("---")
                            st.markdown("### 🏗️ 3. Estructura del Catálogo Maestro (`/Root`)")
                            root_dict = {}
                            for key, value in pdf_pike.Root.items():
                                if str(key) == "/Pages":
                                    root_dict[str(key)] = f"[Árbol de Páginas Oculto: {len(pdf_pike.pages)} páginas]"
                                else:
                                    val_str = str(value)
                                    if len(val_str) > 200:
                                        val_str = val_str[:200] + "... [Truncado]"
                                    root_dict[str(key)] = val_str
                            st.json(root_dict)

                            st.markdown("---")
                            st.markdown("### 📎 4. Metadatos de Archivos Asociados (Anexos)")
                            if doc.embfile_count() > 0:
                                for anexo in doc.embfile_names():
                                    info = doc.embfile_info(anexo)
                                    size_kb = info.get('size', 0) / 1024
                                    st.write(f"**Archivo:** `{anexo}`")
                                    st.write(f"- **Tamaño:** {size_kb:.2f} KB")
                                    st.write(f"- **Descripción:** {info.get('desc', 'N/A')}")
                                    st.write(f"- **Fecha Modificación:** {info.get('modDate', 'N/A')}")
                            else:
                                st.info("No hay archivos incrustados en este documento.")

                            st.markdown("---")
                            st.markdown("### 🧩 5. Capa XMP Completa (XML Crudo)")
                            xmp_meta = doc.get_xml_metadata()
                            if xmp_meta:
                                with st.expander("Ver estructura XML completa", expanded=True):
                                    st.code(xmp_meta, language="xml")
                            else:
                                st.warning("Este documento no contiene capa XMP incrustada.")
                                
                            doc.close()
                            pdf_pike.close()

    else:
        st.subheader("📁 Procesamiento por Lotes y Gestión Archivística")
        
        if "1. Generar Índice" in menu_option:
            st.write("Selecciona o arrastra múltiples archivos de cualquier formato desde tu equipo.")
            origen_opcion = st.selectbox("Origen predeterminado:", ["Digitalizado", "Electrónico", "Físico"], key="orig_manual")
            batch_files = st.file_uploader("Sube los archivos del expediente", accept_multiple_files=True)
            
            if batch_files:
                st.info(f"Se han cargado {len(batch_files)} archivos para procesar.")
                if st.button("Generar Índice"):
                    with st.spinner("Procesando lote, ordenando por fecha y extrayendo metadatos..."):
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
                        
        elif "2. Conectar a Aurora Nextcloud" in menu_option:
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

        elif "3. Inventario Documental (FUID - AGN)" in menu_option:
            st.subheader("📋 Inventario Documental (FUID) vía Aurora Nextcloud")
            st.write("Conéctate a la nube institucional para generar el Formato Único de Inventario Documental (Acuerdo 001 de 2024 - Anexo 3) recorriendo automáticamente las subcarpetas[cite: 13, 14, 18].")
            
            col1, col2 = st.columns(2)
            with col1:
                nc_url_fuid = st.text_input("URL del Servidor", value="https://cloud.insdeportescajica.gov.co", key="fuid_url")
                nc_user_fuid = st.text_input("Usuario Nextcloud", key="fuid_user")
            with col2:
                nc_pass_fuid = st.text_input("Contraseña de Aplicación", type="password", key="fuid_pass")
                nc_folder_fuid = st.text_input("Ruta de la carpeta serie/subserie", placeholder="/Archivo Central/2021/100 DIRECCION/100.2 ACTAS", key="fuid_folder")
                
            if st.button("Conectar y Generar FUID Institucional"):
                if not all([nc_url_fuid, nc_user_fuid, nc_pass_fuid, nc_folder_fuid]):
                    st.warning("⚠️ Por favor, completa todos los campos de conexión.")
                else:
                    with st.spinner("Conectando con Aurora Nextcloud, analizando subcarpetas y leyendo metadatos de PDFs..."):
                        df_fuid, err = generate_fuid_from_nextcloud(nc_url_fuid, nc_user_fuid, nc_pass_fuid, nc_folder_fuid)
                        
                        if err:
                            st.error(f"❌ {err}")
                        else:
                            st.success("¡Inventario FUID generado exitosamente desde Aurora Nextcloud!")
                            st.dataframe(df_fuid)
                            
                            output_fuid_excel = BytesIO()
                            with pd.ExcelWriter(output_fuid_excel, engine='openpyxl') as writer:
                                df_fuid.to_excel(writer, index=False, sheet_name='FUID_Inventario')
                                
                            st.download_button(
                                label="📥 Descargar FUID Oficial (.xlsx)",
                                data=output_fuid_excel.getvalue(),
                                file_name="Inventario_FUID_Nextcloud.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
