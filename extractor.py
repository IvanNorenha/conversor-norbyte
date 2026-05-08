import streamlit as st
import pdfplumber
import pandas as pd
import io
import tempfile
import os
import zipfile
from pypdf import PdfReader, PdfWriter

# 1. CONFIGURACIÓN CORPORATIVA
st.set_page_config(page_title="Conversor Bancario | Norbyte", page_icon="📊", layout="centered")

# 2. CARGAR ESTILOS EXTERNOS (Conexión con estilos.css)
def cargar_css(archivo_css):
    try:
        with open(archivo_css) as f:
            st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)
    except FileNotFoundError:
        st.warning(f"⚠️ No se encontró el archivo {archivo_css}. El sistema funcionará, pero sin diseño corporativo.")

cargar_css("estilos.css")

# 3. ENCABEZADO Y LOGO
try:
    st.image("logo_norbyte.png", width=220)
except FileNotFoundError:
    pass

st.markdown("<h3 style='text-align: center; color: #333; margin-bottom: 20px;'>Conversor de Estados de Cuenta a Excel</h3>", unsafe_allow_html=True)

# 4. INTERFAZ
banco_seleccionado = st.radio(
    "🏦 Selecciona el Banco del Estado de Cuenta",
    ("BCP", "BBVA", "Interbank", "Scotiabank"),
    horizontal=True
)

archivos_subidos = st.file_uploader(f"📂 Sube tu(s) PDF(s) del {banco_seleccionado}", type="pdf", accept_multiple_files=True)

clave_pdf = st.text_input(
    "🔒 Contraseña del PDF (Tu DNI/RUC). Déjalo vacío si no tiene:", 
    type="password", 
    autocomplete="one-time-code"
)

titulos_estandar = ["FECHA PROC.", "FECHA VALOR", "DESCRIPCION", "CARGOS / DEBE", "ABONOS / HABER"]

def a_numero(texto):
    texto_limpio = str(texto).strip()
    if not texto_limpio: return None
    try:
        texto_limpio = texto_limpio.replace(",", "").split()[0]
        return float(texto_limpio)
    except ValueError:
        return texto_limpio

# ==========================================
# EL PURIFICADOR DE BASURA DIGITAL
# ==========================================
def limpiar_basura_bancaria(archivo_bytes):
    inicio_real = archivo_bytes.find(b"%PDF-")
    if inicio_real != -1:
        return archivo_bytes[inicio_real:]
    return archivo_bytes

# ==========================================
# EL CERRAJERO (Desencriptador Maestro)
# ==========================================
def quitar_candado(archivo_bytes, clave):
    lector = PdfReader(io.BytesIO(archivo_bytes), strict=False)
    if lector.is_encrypted:
        if not clave:
            raise ValueError("Este archivo está protegido. Necesitas ingresar la contraseña (DNI/RUC).")
        exito = lector.decrypt(clave)
        if exito == 0: 
            raise ValueError("Contraseña incorrecta. Por favor ingresa el DNI/RUC válido.")
            
    escritor = PdfWriter()
    for pagina in lector.pages:
        escritor.add_page(pagina)
        
    salida = io.BytesIO()
    escritor.write(salida)
    salida.seek(0)
    return salida.getvalue()

# ==========================================
# MOTOR 1: BCP
# ==========================================
def procesar_bcp(archivo):
    hojas_datos = {}
    with pdfplumber.open(archivo) as pdf:
        for i, pagina in enumerate(pdf.pages):
            tabla = pagina.extract_table({"vertical_strategy": "text", "horizontal_strategy": "text"})
            if not tabla: continue
            filas_limpias = []
            guardar = False
            
            for fila in tabla:
                fila_str = [str(c).strip() if c is not None else "" for c in fila]
                if not fila_str: continue
                texto_fila = " ".join(fila_str).upper()
                
                if "FECHA" in texto_fila and "PROC" in texto_fila:
                    guardar = True; continue
                    
                if not guardar:
                    if "SALDO ANTERIOR" in texto_fila:
                        guardar = True
                    elif len(fila_str) > 0:
                        inicio = fila_str[0].replace(" ", "")
                        if len(inicio) >= 5 and inicio[:2].isdigit() and inicio[2:5].isalpha():
                            guardar = True

                if "TOTAL MOVIMIENTO" in texto_fila or ("SALDO" in texto_fila and "ANTERIOR" not in texto_fila):
                    guardar = False; continue 
                    
                if guardar:
                    if not "".join(fila_str).strip(): continue
                        
                    if "SALDO ANTERIOR" in texto_fila:
                        monto = ""
                        for celda in reversed(fila_str):
                            if "." in celda and any(d.isdigit() for d in celda):
                                monto = celda; break
                        filas_limpias.append(["", "", "SALDO ANTERIOR", None, a_numero(monto)])
                        continue

                    if len(fila_str) < 2 or (not fila_str[0].strip() and not fila_str[1].strip()): continue

                    if " " in fila_str[0]:
                        partes = fila_str[0].split(" ", 1)
                        if len(partes[0]) >= 5 and partes[0][:2].isdigit():
                            fila_str[0] = partes[0]
                            fila_str.insert(1, partes[1])

                    if len(fila_str) >= 3 and fila_str[1].strip().isdigit() and len(fila_str[1].strip()) == 2:
                        texto_desc = fila_str[2].strip()
                        meses = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "SET", "OCT", "NOV", "DIC"]
                        if len(texto_desc) >= 3 and texto_desc[:3].upper() in meses:
                            fila_str[1] = fila_str[1].strip() + texto_desc[:3].upper()
                            fila_str[2] = texto_desc[3:].strip()
                            
                    while len(fila_str) < 5: fila_str.append("")
                        
                    abono = a_numero(fila_str[-1])
                    cargo = a_numero(fila_str[-2])
                    desc = " ".join([c for c in fila_str[2:-2] if c])
                    
                    filas_limpias.append([fila_str[0], fila_str[1], desc, cargo, abono])
                    
            if filas_limpias:
                hojas_datos[f"Hoja_{i+1}"] = filas_limpias

    if not hojas_datos: return None
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for hoja, filas in hojas_datos.items():
            pd.DataFrame(filas, columns=titulos_estandar).to_excel(writer, sheet_name=hoja, index=False)
    return buffer

# ==========================================
# MOTOR 2: BBVA
# ==========================================
def procesar_bbva(archivo):
    hojas_datos = {}
    titulos_bbva = ["FECHA OPER.", "FECHA VALOR", "DESCRIPCION", "CARGO / ABONO"]
    
    with pdfplumber.open(archivo) as pdf:
        for i, pagina in enumerate(pdf.pages):
            texto = pagina.extract_text()
            if not texto: continue
            
            filas_limpias = []
            lineas = texto.split('\n')
            
            current_oper = ""
            current_val = ""
            
            for linea in lineas:
                linea = linea.strip()
                if not linea: continue
                
                linea_upper = linea.upper()
                if any(x in linea_upper for x in ["BANCA POR", "SALDO A NUESTRO", "SALDO A SU", "WWW.BBVA", "EN CASO DE RECLAMOS", "ROGAMOS VERIFIQUE", "OF. JOCKEY"]):
                    break 
                if linea_upper.startswith("DNI"):
                    break 

                partes = linea.split()
                if not partes: continue
                
                es_fecha = len(partes[0]) == 5 and partes[0][:2].isdigit() and partes[0][2] == '-' and partes[0][3:].isdigit()
                
                if es_fecha:
                    current_oper = partes[0]
                    if len(partes) > 1 and len(partes[1]) == 5 and partes[1][:2].isdigit() and partes[1][2] == '-':
                        current_val = partes[1]
                        inicio_desc = 2
                    else:
                        current_val = current_oper
                        inicio_desc = 1
                else:
                    if not current_oper: continue
                    inicio_desc = 0
                    
                if "SALDO ANTERIOR" in linea.upper():
                    desc = "SALDO ANTERIOR"
                    cargo_abono = ""
                else:
                    if len(partes) >= inicio_desc + 2:
                        monto = partes[-2]
                        if any(c.isdigit() for c in monto) and ("." in monto or "," in monto):
                            cargo_abono = monto 
                            desc = " ".join(partes[inicio_desc:-2])
                        else:
                            cargo_abono = ""
                            desc = " ".join(partes[inicio_desc:])
                    else:
                        cargo_abono = ""
                        desc = " ".join(partes[inicio_desc:])
                        
                if "FECHA" in desc or "DESCRIPCION" in desc or "SALDO CONTABLE" in desc:
                    continue
                    
                filas_limpias.append([current_oper, current_val, desc, cargo_abono])
                
            if filas_limpias:
                hojas_datos[f"Hoja_{i+1}"] = filas_limpias

    if not hojas_datos: return None
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for hoja, filas in hojas_datos.items():
            pd.DataFrame(filas, columns=titulos_bbva).to_excel(writer, sheet_name=hoja, index=False)
    return buffer

# ==========================================
# MOTOR 3: INTERBANK
# ==========================================
def procesar_interbank(archivo):
    hojas_datos = {}
    titulos_ibk = ["Fecha", "Concepto y/o descripción", "Ingresos", "Gastos"]
    
    with pdfplumber.open(archivo) as pdf:
        for i, pagina in enumerate(pdf.pages):
            texto_crudo = pagina.extract_text()
            if texto_crudo:
                texto_upper = texto_crudo.upper()
                palabras_clave_ejemplo = ["TE AYUDAMOS A CONOCER", "MARÍA VARA DE GAMARRA", "MARIA VARA DE GAMARRA", "CICLO DE CONSUMO"]
                if any(key in texto_upper for key in palabras_clave_ejemplo):
                    continue 
            
            palabras = pagina.extract_words()
            if not palabras: continue

            lineas_y = {}
            for p in palabras:
                y = round(p['top'] / 4) * 4 
                if y not in lineas_y: lineas_y[y] = []
                lineas_y[y].append(p)
                
            filas_limpias = []
            
            for y in sorted(lineas_y.keys()):
                words = sorted(lineas_y[y], key=lambda w: w['x0'])
                if not words: continue
                
                w0 = words[0]['text']
                if len(w0) >= 10 and w0[2] == '/' and w0[5] == '/' and w0[:2].isdigit():
                    fecha = w0[:10]
                    words.pop(0)
                    
                    if not words: continue
                    
                    ingreso = None
                    gasto = None
                    
                    txt_ult = words[-1]['text'].replace(",", "").replace("+", "").replace("-", "")
                    es_saldo = False
                    try:
                        if "." in txt_ult and len(txt_ult.split(".")[1]) >= 2:
                            val_ult = float(txt_ult)
                            if words[-1]['x0'] > (pagina.width * 0.76): 
                                es_saldo = True
                    except ValueError:
                        pass
                        
                    if es_saldo:
                        words.pop() 
                        
                    if not words: continue
                    
                    txt_monto_raw = words[-1]['text']
                    txt_monto = txt_monto_raw.replace(",", "")
                    
                    is_neg = txt_monto.startswith("-")
                    is_pos = txt_monto.startswith("+")
                    
                    if is_neg or is_pos:
                        txt_monto = txt_monto[1:]
                        
                    try:
                        if "." in txt_monto and len(txt_monto.split(".")[1]) >= 2:
                            val = float(txt_monto)
                            x0_monto = words[-1]['x0']
                            
                            if is_neg:
                                gasto = val
                            elif is_pos:
                                ingreso = val
                            else:
                                if x0_monto < (pagina.width * 0.65):
                                    ingreso = val
                                else:
                                    gasto = val
                                    
                            words.pop() 
                    except ValueError:
                        pass 
                        
                    concepto = " ".join([w['text'] for w in words])
                    filas_limpias.append([fecha, concepto, ingreso, gasto])
                    
            if filas_limpias:
                hojas_datos[f"Hoja_{i+1}"] = filas_limpias

    if not hojas_datos: return None
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for hoja, filas in hojas_datos.items():
            pd.DataFrame(filas, columns=titulos_ibk).to_excel(writer, sheet_name=hoja, index=False)
    return buffer

# ==========================================
# MOTOR 4: SCOTIABANK
# ==========================================
def procesar_scotiabank(archivo):
    hojas_datos = {}
    titulos_scotia = ["FECHA OPER.", "FECHA VALOR", "DESCRIPCION", "CARGOS", "ABONOS"]
    
    with pdfplumber.open(archivo) as pdf:
        for i, pagina in enumerate(pdf.pages):
            palabras = pagina.extract_words()
            if not palabras: continue

            lineas_y = {}
            for p in palabras:
                y = round(p['top'] / 4) * 4  
                if y not in lineas_y: lineas_y[y] = []
                lineas_y[y].append(p)
                
            filas_limpias = []
            
            for y in sorted(lineas_y.keys()):
                words = sorted(lineas_y[y], key=lambda w: w['x0'])
                if not words: continue
                
                w0 = words[0]['text']
                if len(w0) >= 5 and w0[:2].isdigit() and w0[2] == '/' and w0[3:5].isdigit():
                    fecha_oper = w0[:5]
                    words.pop(0) 
                    
                    if not words: continue
                    
                    w1 = words[0]['text']
                    if len(w1) >= 5 and w1[:2].isdigit() and w1[2] == '/' and w1[3:5].isdigit():
                        fecha_val = w1[:5]
                        words.pop(0)
                    else:
                        fecha_val = fecha_oper
                        
                    if not words: continue
                    
                    if len(words[0]['text']) <= 4 and words[0]['text'].isdigit():
                        words.pop(0)
                        
                    if not words: continue
                    
                    words.pop() 
                    if not words: continue
                    
                    cargo = None
                    abono = None
                    monto_idx = -1
                    monto_val = None
                    
                    for idx in range(len(words)-1, -1, -1):
                        txt = words[idx]['text'].replace(",", "")
                        if txt.endswith("-"): txt = txt[:-1] 
                        
                        try:
                            val = float(txt)
                            if "." in txt and len(txt.split(".")[1]) >= 2:
                                monto_idx = idx
                                monto_val = val
                                break
                        except ValueError:
                            continue
                            
                    if monto_idx != -1:
                        if words[monto_idx]['x0'] > (pagina.width * 0.74):
                            abono = monto_val
                        else:
                            cargo = monto_val
                            
                        words.pop(monto_idx)
                        
                    desc = " ".join([w['text'] for w in words])
                    filas_limpias.append([fecha_oper, fecha_val, desc, cargo, abono])
                    
            if filas_limpias:
                hojas_datos[f"Hoja_{i+1}"] = filas_limpias

    if not hojas_datos: return None
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for hoja, filas in hojas_datos.items():
            pd.DataFrame(filas, columns=titulos_scotia).to_excel(writer, sheet_name=hoja, index=False)
    return buffer

# ==========================================
# EJECUCIÓN PRINCIPAL CON PROCESAMIENTO MASIVO
# ==========================================
if archivos_subidos:
    if st.button(f"🚀 Convertir PDF(s) de {banco_seleccionado}"):
        with st.spinner(f"Procesando {len(archivos_subidos)} archivo(s) de {banco_seleccionado}..."):
            archivos_exitosos = []
            errores = []

            for archivo in archivos_subidos:
                try:
                    bytes_puros = limpiar_basura_bancaria(archivo.getvalue())
                    archivo_limpio_bytes = quitar_candado(bytes_puros, clave_pdf)
                    
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                        tmp.write(archivo_limpio_bytes)
                        ruta_temporal = tmp.name
                    
                    buffer_excel = None
                    if banco_seleccionado == "BCP":
                        buffer_excel = procesar_bcp(ruta_temporal)
                    elif banco_seleccionado == "BBVA":
                        buffer_excel = procesar_bbva(ruta_temporal)
                    elif banco_seleccionado == "Interbank":
                        buffer_excel = procesar_interbank(ruta_temporal)
                    elif banco_seleccionado == "Scotiabank":
                        buffer_excel = procesar_scotiabank(ruta_temporal)
                    
                    try: os.remove(ruta_temporal)
                    except: pass
                    
                    if buffer_excel is not None:
                        nombre_original_sin_ext = os.path.splitext(archivo.name)[0]
                        nombre_excel = f"{nombre_original_sin_ext}.xlsx"
                        archivos_exitosos.append((nombre_excel, buffer_excel))
                    else:
                        errores.append(f"No se encontraron transacciones en: {archivo.name}")
                        
                except ValueError as ve:
                    errores.append(f"🔒 {archivo.name}: {str(ve)}")
                except Exception as e:
                    errores.append(f"❌ {archivo.name}: Error interno al procesar.")

            if errores:
                for err in errores:
                    st.warning(err)
                    
            if archivos_exitosos:
                st.success(f"¡Se convirtieron exitosamente {len(archivos_exitosos)} documento(s)!")
                
                if len(archivos_exitosos) == 1:
                    st.download_button(
                        label=f"📥 Descargar {archivos_exitosos[0][0]}",
                        data=archivos_exitosos[0][1].getvalue(),
                        file_name=archivos_exitosos[0][0],
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                        for nombre, buffer in archivos_exitosos:
                            zf.writestr(nombre, buffer.getvalue())
                    
                    st.download_button(
                        label=f"📦 Descargar {len(archivos_exitosos)} archivos en ZIP",
                        data=zip_buffer.getvalue(),
                        file_name=f"Estados_Cuenta_{banco_seleccionado}.zip",
                        mime="application/zip"
                    )

st.markdown("""
    <div class="footer-norbyte">
        <p>Potenciado por <b>Norbyte</b></p>
        <p><i>Desarrollamos soluciones de software para el usuario final, diseñadas específicamente para optimizar tu día a día.</i></p>
    </div>
""", unsafe_allow_html=True)
