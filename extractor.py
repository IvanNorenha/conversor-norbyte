# ==========================================
# IMPORTACIÓN DE HERRAMIENTAS (LIBRERÍAS)
# ==========================================
import streamlit as st
import pdfplumber
import pandas as pd
import io
import tempfile
import os
import zipfile
from pypdf import PdfReader, PdfWriter

# ==========================================
# 1. CONFIGURACIÓN DE LA PÁGINA
# ==========================================
st.set_page_config(page_title="Conversor Bancario | Norbyte", page_icon="📊", layout="centered")

# ==========================================
# 2. ESTILOS VISUALES INCRUSTADOS (CSS)
# ==========================================
estilos_nativos = """
<style>
    /* --- OCULTAR ELEMENTOS POR DEFECTO DE STREAMLIT --- */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    [data-testid="stHeader"] { display: none !important; }

    /* ==========================================
       1. AJUSTE DE ESPACIOS (SUPERIOR Y TÍTULO)
       ========================================== */
    .block-container {
        padding-top: 0rem !important;
        margin-top: -30px !important;
    }

    [data-testid="stImage"] {
        margin-bottom: -40px !important;
    }

    /* ==========================================
       2. TÍTULOS Y TARJETAS DE BANCOS
       ========================================== */
    .titulo-centrado {
        text-align: center; font-weight: 600; color: #2C3E50;
        margin-bottom: 5px; font-size: 1.1rem; margin-top: 15px;
    }
    .subtitulo-centrado {
        text-align: center; font-size: 0.85em; color: #7f8c8d; margin-bottom: 15px;
    }

    div[role="radiogroup"] {
        display: flex !important; flex-wrap: wrap !important;
        justify-content: center !important; gap: 15px !important;
    }
    div[role="radiogroup"] label div[data-baseweb="radio"] { display: none !important; }
    div[role="radiogroup"] label {
        border: 2px solid #E0E6ED !important; background-color: #FFFFFF !important;
        padding: 12px 24px !important; border-radius: 10px !important;
        cursor: pointer !important; transition: all 0.3s ease !important;
        min-width: 130px; display: flex; justify-content: center;
    }
    div[role="radiogroup"] label:hover { border-color: #F05A28 !important; background-color: #FFF2ED !important; }
    div[role="radiogroup"] label:has(input:checked) {
        border-color: #F05A28 !important; background-color: #FFF2ED !important;
        box-shadow: 0px 4px 10px rgba(240, 90, 40, 0.2) !important;
    }
    div[role="radiogroup"] label:has(input:checked) p { color: #F05A28 !important; font-weight: bold !important; }

    /* ==========================================
       3. CONTRASEÑA Y UPLOADER
       ========================================== */
    [data-testid="stTextInput"] {
        max-width: 450px !important; margin: 0 auto !important;
    }
    [data-testid="stWidgetLabel"] { display: none !important; }
    
    /* FULMINAR MENSAJE "Press Enter to apply" DEL OJO */
    [data-testid="InputInstructions"] { display: none !important; }
    
    /* Diseño del uploader (Borde naranja) */
    .stFileUploader section {
        border: 2px dashed #F05A28 !important;
        background-color: #fcfcfc !important;
        border-radius: 10px !important;
        padding: 20px !important;
        max-width: 450px !important;
        margin: 0 auto !important;
    }
    .stFileUploader section:hover {
        background-color: #FFF2ED !important;
    }

    /* ==========================================
       4. BOTONES Y ALERTAS
       ========================================== */
    [data-testid="stAlert"] > div {
        display: flex !important; justify-content: center !important; 
        align-items: center !important; text-align: center !important;
    }

    div.stButton > button { 
        background-color: #F05A28 !important; color: white !important; 
        border: none !important; border-radius: 8px !important; 
        padding: 12px 30px !important; font-weight: bold !important;
    }
    div.stButton > button:hover { background-color: #D94A1D !important; }

    div.stDownloadButton > button { 
        background-color: #28a745 !important; color: white !important; 
        border: none !important; border-radius: 8px !important; 
        padding: 12px 30px !important; font-weight: bold !important;
    }
    div.stDownloadButton > button:hover { background-color: #218838 !important; }

    /* Footer */
    .footer-norbyte {
        text-align: center; margin-top: 40px; padding-top: 20px;
        border-top: 2px solid #f0f2f6; color: #555; font-size: 0.85em;
    }
</style>
"""
st.markdown(estilos_nativos, unsafe_allow_html=True)

# ==========================================
# 3. ENCABEZADO Y LOGO
# ==========================================
col1, col2, col3 = st.columns([0.2, 2.5, 0.2])
with col2:
    try:
        st.image("logo_norbyte.png", use_container_width=True)
    except Exception:
        pass

st.markdown("<h3 style='text-align: center; color: #333; margin-top: -20px;'>Conversor de Estados de Cuenta a Excel</h3>", unsafe_allow_html=True)

# ==========================================
# 4. INTERFAZ VISUAL DEL USUARIO
# ==========================================
st.markdown("<p class='titulo-centrado'>🏦 Selecciona el Banco del Estado de Cuenta</p>", unsafe_allow_html=True)
banco_seleccionado = st.radio("Banco", ("BCP", "BBVA", "Interbank", "Scotiabank"), horizontal=True)

st.markdown(f"<p class='titulo-centrado'>📂 Sube tu(s) PDF(s) del {banco_seleccionado}</p>", unsafe_allow_html=True)
archivos_subidos = st.file_uploader("PDFs", type="pdf", accept_multiple_files=True)

st.markdown("<p class='titulo-centrado'>🔒 Contraseña del PDF (DNI/RUC)</p>", unsafe_allow_html=True)
st.markdown("<p class='subtitulo-centrado'>Déjalo vacío si no tiene contraseña</p>", unsafe_allow_html=True)
clave_pdf = st.text_input("Contraseña", type="password", autocomplete="one-time-code")


# ==========================================
# 5. FUNCIONES DE AYUDA
# ==========================================
titulos_estandar = ["FECHA PROC.", "FECHA VALOR", "DESCRIPCION", "CARGOS / DEBE", "ABONOS / HABER"]

def a_numero(texto):
    texto_limpio = str(texto).strip()
    if not texto_limpio: return None
    try:
        texto_limpio = texto_limpio.replace(",", "").split()[0]
        return float(texto_limpio)
    except ValueError:
        return texto_limpio

def limpiar_basura_bancaria(archivo_bytes):
    inicio_real = archivo_bytes.find(b"%PDF-")
    if inicio_real != -1:
        return archivo_bytes[inicio_real:]
    return archivo_bytes

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
# 6. MOTORES DE EXTRACCIÓN
# ==========================================

def procesar_bcp(archivo):
    hojas_datos = {}
    with pdfplumber.open(archivo) as pdf:
        texto_p1 = pdf.pages[0].extract_text()
        es_empresa = texto_p1 and ("CUENTA CORRIENTE" in texto_p1.upper() or "CUENTA EMPRESA" in texto_p1.upper())
        
        for i, pagina in enumerate(pdf.pages):
            if es_empresa:
                texto = pagina.extract_text()
                if not texto: continue
                filas_limpias = []
                lineas = texto.split('\n')
                for linea in lineas:
                    linea = linea.strip()
                    if not linea: continue
                    partes = linea.split()
                    if len(partes) >= 3 and len(partes[0]) == 5 and partes[0][:2].isdigit() and partes[0][2] == '-':
                        fecha = partes[0]
                        monto_raw = partes[-2].replace(',', '') 
                        is_neg = monto_raw.endswith('-')
                        monto_str = monto_raw[:-1] if is_neg else monto_raw
                        try:
                            monto_val = float(monto_str)
                        except ValueError:
                            continue 
                        cargo = monto_val if is_neg else None
                        abono = monto_val if not is_neg else None
                        desc = " ".join(partes[1:-2])
                        filas_limpias.append([fecha, fecha, desc, cargo, abono])
                if filas_limpias:
                    hojas_datos[f"Hoja_{i+1}"] = filas_limpias
            else:
                # FIX HOJA 1: Recorte inteligente
                if i == 0:
                    crop_y = 0
                    words = pagina.extract_words()
                    for idx, w in enumerate(words):
                        if w['text'] == 'FECHA' and idx + 1 < len(words) and 'PROC' in words[idx+1]['text']:
                            crop_y = max(0, w['top'] - 10)
                            break
                    if crop_y > 0:
                        pagina = pagina.crop((0, crop_y, pagina.width, pagina.height))

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
                        if "SALDO ANTERIOR" in texto_fila: guardar = True
                        elif len(fila_str) > 0:
                            inicio = fila_str[0].replace(" ", "")
                            if len(inicio) >= 5 and inicio[:2].isdigit() and inicio[2:5].isalpha(): guardar = True
                            
                    if "TOTAL MOVIMIENTO" in texto_fila or ("SALDO" in texto_fila and "ANTERIOR" not in texto_fila):
                        guardar = False; continue
                        
                    if guardar:
                        if not "".join(fila_str).strip(): continue
                        if "SALDO ANTERIOR" in texto_fila:
                            monto = ""
                            for celda in reversed(fila_str):
                                if "." in celda and any(d.isdigit() for d in celda): monto = celda; break
                            filas_limpias.append(["", "", "SALDO ANTERIOR", None, a_numero(monto)])
                            continue
                            
                        if len(fila_str) < 2 or (not fila_str[0].strip() and not fila_str[1].strip()): continue
                        
                        if " " in fila_str[0]:
                            partes = fila_str[0].split(" ", 1)
                            if len(partes[0]) >= 5 and partes[0][:2].isdigit():
                                fila_str[0] = partes[0]; fila_str.insert(1, partes[1])
                                
                        if len(fila_str) >= 3 and fila_str[1].strip().isdigit() and len(fila_str[1].strip()) == 2:
                            texto_desc = fila_str[2].strip()
                            meses = ["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","SET","OCT","NOV","DIC"]
                            if len(texto_desc) >= 3 and texto_desc[:3].upper() in meses:
                                fila_str[1] = fila_str[1].strip() + texto_desc[:3].upper()
                                fila_str[2] = texto_desc[3:].strip()
                                
                        while len(fila_str) < 5: fila_str.append("")
                        
                        # ==================================================
                        # NUEVO FILTRO ANTI-INVASIÓN DE TEXTO
                        # ==================================================
                        abono_raw = str(fila_str[-1]).strip()
                        cargo_raw = str(fila_str[-2]).strip()
                        desc_parts = [str(c).strip() for c in fila_str[2:-2] if str(c).strip()]
                        
                        cargo_val = None
                        abono_val = None
                        
                        # Evaluar Columna Cargo
                        if cargo_raw:
                            c_clean = cargo_raw.replace(',', '')
                            # Si tiene punto decimal, es dinero. Si no, es texto invasor.
                            if '.' in c_clean:
                                try: cargo_val = float(c_clean)
                                except ValueError: desc_parts.append(cargo_raw)
                            else:
                                desc_parts.append(cargo_raw)
                                
                        # Evaluar Columna Abono
                        if abono_raw:
                            a_clean = abono_raw.replace(',', '')
                            if '.' in a_clean:
                                try: abono_val = float(a_clean)
                                except ValueError: desc_parts.append(abono_raw)
                            else:
                                desc_parts.append(abono_raw)
                                
                        # Volvemos a pegar la descripción sanada
                        desc = " ".join(desc_parts)
                        
                        filas_limpias.append([fila_str[0], fila_str[1], desc, cargo_val, abono_val])
                        
                if filas_limpias: 
                    hojas_datos[f"Hoja_{i+1}"] = filas_limpias
                    
    if not hojas_datos: return None
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for hoja, filas in hojas_datos.items():
            pd.DataFrame(filas, columns=titulos_estandar).to_excel(writer, sheet_name=hoja, index=False)
    return buffer

def procesar_bbva(archivo):
    hojas_datos = {}
    titulos_bbva = ["FECHA OPER.", "FECHA VALOR", "DESCRIPCION", "CARGOS", "ABONOS"]
    
    with pdfplumber.open(archivo) as pdf:
        for i, pagina in enumerate(pdf.pages):
            texto = pagina.extract_text()
            if not texto: continue
            filas_limpias = []
            lineas = texto.split('\n')
            current_oper = ""; current_val = ""
            
            for linea in lineas:
                linea = linea.strip()
                if not linea: continue
                linea_upper = linea.upper()
                if any(x in linea_upper for x in ["BANCA POR","SALDO A NUESTRO","SALDO A SU","WWW.BBVA","EN CASO DE RECLAMOS","ROGAMOS VERIFIQUE","OF. JOCKEY"]): break
                if linea_upper.startswith("DNI"): break
                
                partes = linea.split()
                if not partes: continue
                
                es_fecha = len(partes[0]) == 5 and partes[0][:2].isdigit() and partes[0][2] == '-' and partes[0][3:].isdigit()
                if es_fecha:
                    current_oper = partes[0]
                    if len(partes) > 1 and len(partes[1]) == 5 and partes[1][:2].isdigit() and partes[1][2] == '-':
                        current_val = partes[1]; inicio_desc = 2
                    else:
                        current_val = current_oper; inicio_desc = 1
                else:
                    if not current_oper: continue
                    inicio_desc = 0
                
                cargo = None
                abono = None
                
                if "SALDO ANTERIOR" in linea.upper():
                    desc = "SALDO ANTERIOR"
                else:
                    if len(partes) >= inicio_desc + 2:
                        monto_raw = partes[-2]
                        if any(c.isdigit() for c in monto_raw) and ("." in monto_raw or "," in monto_raw):
                            desc = " ".join(partes[inicio_desc:-2])
                            monto_clean = monto_raw.replace(',', '')
                            is_neg = monto_clean.endswith('-')
                            monto_str = monto_clean[:-1] if is_neg else monto_clean
                            try:
                                monto_val = float(monto_str)
                                if is_neg: cargo = monto_val
                                else: abono = monto_val
                            except ValueError: pass
                        else: desc = " ".join(partes[inicio_desc:])
                    else: desc = " ".join(partes[inicio_desc:])
                        
                if "FECHA" in desc or "DESCRIPCION" in desc or "SALDO CONTABLE" in desc: continue
                filas_limpias.append([current_oper, current_val, desc, cargo, abono])
                
            if filas_limpias: hojas_datos[f"Hoja_{i+1}"] = filas_limpias
            
    if not hojas_datos: return None
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for hoja, filas in hojas_datos.items(): 
            pd.DataFrame(filas, columns=titulos_bbva).to_excel(writer, sheet_name=hoja, index=False)
    return buffer

def procesar_interbank(archivo):
    hojas_datos = {}
    titulos_ibk = ["Fecha", "Concepto y/o descripción", "Gastos", "Ingresos"]
    
    with pdfplumber.open(archivo) as pdf:
        texto_p1 = pdf.pages[0].extract_text()
        es_empresa = texto_p1 and ("ESTADO DE CUENTA NEGOCIOS" in texto_p1.upper() or "PARA EMPRESAS" in texto_p1.upper())

        for i, pagina in enumerate(pdf.pages):
            if es_empresa:
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
                    
                    if len(w0) == 5 and w0[:2].isdigit() and w0[2] == '/' and w0[3:].isdigit():
                        fecha = w0; words.pop(0)
                        if not words: continue
                        w1 = words[0]['text']
                        if len(w1) == 5 and w1[:2].isdigit() and w1[2] == '/' and w1[3:].isdigit():
                            words.pop(0)
                        if not words: continue
                        
                        ingreso = None; gasto = None
                        txt_ult = words[-1]['text'].replace(",","").replace("-","")
                        if "." in txt_ult and len(txt_ult.split(".")[1]) >= 2:
                            if words[-1]['x0'] > (pagina.width * 0.75): words.pop()
                        if not words: continue
                        
                        txt_monto_raw = words[-1]['text']
                        txt_monto_clean = txt_monto_raw.replace(',', '')
                        is_neg = txt_monto_clean.startswith('-') or txt_monto_clean.endswith('-')
                        monto_str = txt_monto_clean.replace('-', '') 
                        try:
                            if "." in monto_str and len(monto_str.split(".")[1]) >= 2:
                                val = float(monto_str)
                                if is_neg: gasto = val 
                                else: ingreso = val 
                                words.pop() 
                        except ValueError: pass
                            
                        concepto = " ".join([w['text'] for w in words])
                        filas_limpias.append([fecha, concepto, gasto, ingreso])
                if filas_limpias: hojas_datos[f"Hoja_{i+1}"] = filas_limpias

            else:
                texto_crudo = pagina.extract_text()
                if texto_crudo:
                    texto_upper = texto_crudo.upper()
                    palabras_clave_ejemplo = ["TE AYUDAMOS A CONOCER","MARÍA VARA DE GAMARRA","MARIA VARA DE GAMARRA","CICLO DE CONSUMO"]
                    if any(key in texto_upper for key in palabras_clave_ejemplo): continue
                
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
                        fecha = w0[:10]; words.pop(0)
                        if not words: continue
                        ingreso = None; gasto = None
                        txt_ult = words[-1]['text'].replace(",","").replace("+","").replace("-","")
                        es_saldo = False
                        try:
                            if "." in txt_ult and len(txt_ult.split(".")[1]) >= 2:
                                if words[-1]['x0'] > (pagina.width * 0.76): es_saldo = True
                        except ValueError: pass
                        if es_saldo: words.pop()
                        if not words: continue
                        txt_monto_raw = words[-1]['text']
                        txt_monto = txt_monto_raw.replace(",","")
                        is_neg = txt_monto.startswith("-"); is_pos = txt_monto.startswith("+")
                        if is_neg or is_pos: txt_monto = txt_monto[1:]
                        try:
                            if "." in txt_monto and len(txt_monto.split(".")[1]) >= 2:
                                val = float(txt_monto); x0_monto = words[-1]['x0']
                                if is_neg: gasto = val
                                elif is_pos: ingreso = val
                                else:
                                    if x0_monto < (pagina.width * 0.65): ingreso = val
                                    else: gasto = val
                                words.pop()
                        except ValueError: pass
                        concepto = " ".join([w['text'] for w in words])
                        filas_limpias.append([fecha, concepto, gasto, ingreso])
                if filas_limpias: hojas_datos[f"Hoja_{i+1}"] = filas_limpias
                
    if not hojas_datos: return None
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for hoja, filas in hojas_datos.items(): pd.DataFrame(filas, columns=titulos_ibk).to_excel(writer, sheet_name=hoja, index=False)
    return buffer

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
                    fecha_oper = w0[:5]; words.pop(0)
                    if not words: continue
                    w1 = words[0]['text']
                    if len(w1) >= 5 and w1[:2].isdigit() and w1[2] == '/' and w1[3:5].isdigit():
                        fecha_val = w1[:5]; words.pop(0)
                    else: fecha_val = fecha_oper
                    if not words: continue
                    if len(words[0]['text']) <= 4 and words[0]['text'].isdigit(): words.pop(0)
                    if not words: continue
                    words.pop()
                    if not words: continue
                    cargo = None; abono = None; monto_idx = -1; monto_val = None
                    for idx in range(len(words)-1, -1, -1):
                        txt = words[idx]['text'].replace(",","")
                        if txt.endswith("-"): txt = txt[:-1]
                        try:
                            val = float(txt)
                            if "." in txt and len(txt.split(".")[1]) >= 2:
                                monto_idx = idx; monto_val = val; break
                        except ValueError: continue
                    if monto_idx != -1:
                        if words[monto_idx]['x0'] > (pagina.width * 0.74): abono = monto_val
                        else: cargo = monto_val
                        words.pop(monto_idx)
                    desc = " ".join([w['text'] for w in words])
                    filas_limpias.append([fecha_oper, fecha_val, desc, cargo, abono])
            if filas_limpias: hojas_datos[f"Hoja_{i+1}"] = filas_limpias
    if not hojas_datos: return None
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for hoja, filas in hojas_datos.items(): pd.DataFrame(filas, columns=titulos_scotia).to_excel(writer, sheet_name=hoja, index=False)
    return buffer


# ==========================================
# 7. EJECUCIÓN PRINCIPAL Y CENTRADO DE BOTONES
# ==========================================
if archivos_subidos:
    
    col_btn_izq, col_btn_cen, col_btn_der = st.columns([1, 2, 1])
    with col_btn_cen:
        ejecutar = st.button(f"🚀 Convertir PDF(s) de {banco_seleccionado}", use_container_width=True)

    if ejecutar:
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
                    if banco_seleccionado == "BCP":        buffer_excel = procesar_bcp(ruta_temporal)
                    elif banco_seleccionado == "BBVA":     buffer_excel = procesar_bbva(ruta_temporal)
                    elif banco_seleccionado == "Interbank":buffer_excel = procesar_interbank(ruta_temporal)
                    elif banco_seleccionado == "Scotiabank":buffer_excel = procesar_scotiabank(ruta_temporal)
                    try: os.remove(ruta_temporal)
                    except: pass
                    if buffer_excel is not None:
                        nombre_excel = f"{os.path.splitext(archivo.name)[0]}.xlsx"
                        archivos_exitosos.append((nombre_excel, buffer_excel))
                    else:
                        errores.append(f"No se encontraron transacciones legibles en: {archivo.name}")
                except ValueError as ve:
                    errores.append(f"🔒 {archivo.name}: {str(ve)}")
                except Exception:
                    errores.append(f"❌ {archivo.name}: Error interno al procesar. Verifica el archivo.")

            if errores:
                for err in errores: st.warning(err)

            if archivos_exitosos:
                col_alerta_izq, col_alerta_cen, col_alerta_der = st.columns([1, 2, 1])
                with col_alerta_cen:
                    st.success(f"✅ ¡Se convirtieron {len(archivos_exitosos)} documento(s)!")
                    if len(archivos_exitosos) == 1:
                        st.download_button(
                            label=f"📥 Descargar {archivos_exitosos[0][0]}",
                            data=archivos_exitosos[0][1].getvalue(),
                            file_name=archivos_exitosos[0][0],
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
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
                            mime="application/zip",
                            use_container_width=True
                        )

# ==========================================
# 8. FOOTER
# ==========================================
st.markdown("""
    <div style="text-align: center; margin-top: 60px; padding-top: 30px; border-top: 2px solid #f0f2f6;" class="footer-norbyte">
        <p style="font-size: 1.15em; color: #333; margin-bottom: 8px;">🚀 Potenciado por <b>Norbyte</b></p>
        <p style="font-size: 1em; color: #444; line-height: 1.6; max-width: 600px; margin: 0 auto;">
            <i>En <b>Norbyte</b> transformamos la complejidad en agilidad. Desarrollamos soluciones de software a medida para clientes finales que buscan optimizar cada minuto de su gestión. <br>
            <b>¿Tienes un proceso manual que te quita tiempo? Nosotros lo automatizamos.</b> ¡Lleva tu empresa al siguiente nivel!</i>
        </p>
    </div>
""", unsafe_allow_html=True)
