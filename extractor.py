import streamlit as st
import pdfplumber
import pandas as pd
import io

# 1. CONFIGURACIÓN CORPORATIVA
st.set_page_config(page_title="Conversor Bancario | Norbyte", page_icon="📊", layout="centered")

# 2. ESTILOS VISUALES (CSS)
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .block-container { padding-bottom: 100px; padding-top: 2rem; }
    
    /* 1. Ocultar el molesto texto "Press Enter to apply" que tapa el ojo de la contraseña */
    [data-testid="InputInstructions"] {
        display: none !important;
    }
    
    div.stButton > button:first-child {
        background-color: #F05A28; color: white; border: none; border-radius: 6px;
        padding: 10px 24px; font-weight: bold; width: 100%;
    }
    div.stButton > button:first-child:hover { background-color: #D94A1D; color: white; }
    div.stDownloadButton > button:first-child {
        background-color: #28a745; color: white; border: none; font-weight: bold; width: 100%;
    }
    div.stDownloadButton > button:first-child:hover { background-color: #218838; color: white; }
    .footer-norbyte {
        text-align: center; margin-top: 40px; padding-top: 20px;
        border-top: 2px solid #f0f2f6; color: #555; font-size: 0.85em;
    }
    </style>
""", unsafe_allow_html=True)

# 3. ENCABEZADO Y LOGO
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    try:
        st.image("logo_norbyte.png", use_container_width=True)
    except FileNotFoundError:
        pass

st.markdown("<h3 style='text-align: center; color: #333; margin-bottom: 30px;'>Conversor de Estados de Cuenta a Excel</h3>", unsafe_allow_html=True)

# 4. INTERFAZ MEJORADA (Botones Horizontales y Anti-Gestor de Claves)

# Ahora mostramos los bancos como botones visibles directamente
banco_seleccionado = st.radio(
    "🏦 Selecciona el Banco del Estado de Cuenta",
    ("BCP", "BBVA", "Interbank", "Scotiabank"),
    horizontal=True
)

archivo_subido = st.file_uploader(f"📂 Sube tu PDF del {banco_seleccionado}", type="pdf")

# Se agregó autocomplete="new-password" para bloquear el popup de Google/Chrome
clave_pdf = st.text_input(
    "🔒 Contraseña del PDF (Tu DNI/RUC). Déjalo vacío si no tiene:", 
    type="password", 
    autocomplete="one-time-code"  # <--- Este es el truco anti-Chrome
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
# MOTOR 1: BCP
# ==========================================
def procesar_bcp(archivo, clave):
    hojas_datos = {}
    clave_args = {"password": clave} if clave else {}
    
    with pdfplumber.open(archivo, **clave_args) as pdf:
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
def procesar_bbva(archivo, clave):
    hojas_datos = {}
    clave_args = {"password": clave} if clave else {}
    titulos_bbva = ["FECHA OPER.", "FECHA VALOR", "DESCRIPCION", "CARGO / ABONO"]
    
    with pdfplumber.open(archivo, **clave_args) as pdf:
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
def procesar_interbank(archivo, clave):
    hojas_datos = {}
    clave_args = {"password": clave} if clave else {}
    titulos_ibk = ["Fecha", "Concepto y/o descripción", "Ingresos", "Gastos"]
    
    with pdfplumber.open(archivo, **clave_args) as pdf:
        for i, pagina in enumerate(pdf.pages):
            # --- DETECTOR Y BLOQUEADOR DE EJEMPLOS/PUBLICIDAD ---
            texto_crudo = pagina.extract_text()
            if texto_crudo:
                texto_upper = texto_crudo.upper()
                # AHORA SÍ: Usamos las palabras exactas que aparecen en la hoja 5
                palabras_clave_ejemplo = ["TE AYUDAMOS A CONOCER", "MARÍA VARA DE GAMARRA", "MARIA VARA DE GAMARRA", "CICLO DE CONSUMO"]
                if any(key in texto_upper for key in palabras_clave_ejemplo):
                    continue # Destruye la hoja de ejemplo y pasa a la siguiente
            # -----------------------------------------------------------
            
            # Intento 1: Tabla clásica
            tabla = pagina.extract_table({"vertical_strategy": "text", "horizontal_strategy": "text"})
            filas_limpias = []
            
            if tabla:
                for fila in tabla:
                    row = [str(c).strip().replace("\n", " ") if c is not None else "" for c in fila]
                    if not row or not row[0]: continue
                    
                    fecha_str = row[0][:10] 
                    if len(fecha_str) == 10 and fecha_str[2] == '/' and fecha_str[5] == '/':
                        fecha = fecha_str
                        concepto = row[1] if len(row) > 1 else ""
                        ingreso = None; gasto = None
                        
                        if len(row) >= 5:
                            if row[2]: ingreso = a_numero(row[2])
                            if row[3]: gasto = a_numero(row[3])
                        elif len(row) == 4:
                            monto_str = row[2].replace(",", "").strip()
                            if monto_str.startswith("-"): gasto = a_numero(monto_str.replace("-", ""))
                            elif monto_str.startswith("+"): ingreso = a_numero(monto_str.replace("+", ""))
                            else: 
                                num = a_numero(monto_str)
                                if num is not None: ingreso = num
                        
                        filas_limpias.append([fecha, concepto, ingreso, gasto])
            
            # Intento 2: Radar de Respaldo Forzado (Si es una hoja rara como la 4)
            if not filas_limpias:
                palabras = pagina.extract_words()
                lineas_y = {}
                for p in palabras:
                    y = round(p['top'] / 4) * 4 
                    if y not in lineas_y: lineas_y[y] = []
                    lineas_y[y].append(p)
                    
                for y in sorted(lineas_y.keys()):
                    words = sorted(lineas_y[y], key=lambda w: w['x0'])
                    if not words: continue
                    texto_linea = " ".join([w['text'] for w in words])
                    
                    if len(texto_linea) >= 10 and texto_linea[2] == '/' and texto_linea[5] == '/':
                        fecha = words[0]['text'][:10]
                        if not (fecha[:2].isdigit() and fecha[3:5].isdigit()): continue
                        
                        if len(words) >= 3:
                            ult_text = words[-1]['text'].replace(",", "")
                            pen_text = words[-2]['text'].replace(",", "")
                            
                            ult = a_numero(ult_text)
                            pen = a_numero(pen_text)
                            
                            if ult is not None and pen is not None:
                                x0_monto = words[-2]['x0']
                                monto = pen
                                concepto = " ".join([w['text'] for w in words[1:-2]])
                            elif ult is not None:
                                x0_monto = words[-1]['x0']
                                monto = ult
                                concepto = " ".join([w['text'] for w in words[1:-1]])
                            else:
                                continue
                                
                            ingreso = None; gasto = None
                            
                            if x0_monto < (pagina.width * 0.68):
                                ingreso = monto
                            else:
                                gasto = monto
                                
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
def procesar_scotiabank(archivo, clave):
    hojas_datos = {}
    clave_args = {"password": clave} if clave else {}
    # Las 5 columnas exactas solicitadas
    titulos_scotia = ["FECHA OPER.", "FECHA VALOR", "DESCRIPCION", "CARGOS", "ABONOS"]
    
    with pdfplumber.open(archivo, **clave_args) as pdf:
        for i, pagina in enumerate(pdf.pages):
            palabras = pagina.extract_words()
            if not palabras: continue

            # Agrupamos las palabras en filas exactas usando su coordenada vertical (Y)
            lineas_y = {}
            for p in palabras:
                y = round(p['top'] / 4) * 4  # Tolerancia milimétrica para alinear el texto
                if y not in lineas_y: lineas_y[y] = []
                lineas_y[y].append(p)
                
            filas_limpias = []
            
            for y in sorted(lineas_y.keys()):
                words = sorted(lineas_y[y], key=lambda w: w['x0'])
                if not words: continue
                
                # 1. Validar que la fila empiece con una FECHA (Ej: "06/03")
                w0 = words[0]['text']
                if len(w0) >= 5 and w0[:2].isdigit() and w0[2] == '/' and w0[3:5].isdigit():
                    fecha_oper = w0[:5]
                    words.pop(0) # Retiramos la Fecha de Operación de la lista
                    
                    if not words: continue
                    
                    # 2. Atrapar FECHA VALOR (si existe)
                    w1 = words[0]['text']
                    if len(w1) >= 5 and w1[:2].isdigit() and w1[2] == '/' and w1[3:5].isdigit():
                        fecha_val = w1[:5]
                        words.pop(0)
                    else:
                        fecha_val = fecha_oper
                        
                    if not words: continue
                    
                    # 3. Eliminar la columna ORIG (usualmente 3 dígitos, ej: "784" o "001")
                    if len(words[0]['text']) <= 4 and words[0]['text'].isdigit():
                        words.pop(0)
                        
                    if not words: continue
                    
                    # 4. El último dato de la derecha en Scotiabank siempre es el SALDO. Lo eliminamos.
                    words.pop() 
                    if not words: continue
                    
                    cargo = None
                    abono = None
                    monto_idx = -1
                    monto_val = None
                    
                    # 5. Escáner de Derecha a Izquierda para cazar el Dinero escondido
                    for idx in range(len(words)-1, -1, -1):
                        txt = words[idx]['text'].replace(",", "")
                        if txt.endswith("-"): txt = txt[:-1] # Limpiar signos negativos pegados
                        
                        try:
                            val = float(txt)
                            # Asegurar que es dinero real (tiene punto y 2 decimales exactos)
                            if "." in txt and len(txt.split(".")[1]) >= 2:
                                monto_idx = idx
                                monto_val = val
                                break
                        except ValueError:
                            continue
                            
                    if monto_idx != -1:
                        # Geometría pura: Si el X0 del número está más allá del 74% de la hoja, es ABONO.
                        if words[monto_idx]['x0'] > (pagina.width * 0.74):
                            abono = monto_val
                        else:
                            cargo = monto_val
                            
                        # Retiramos el dinero de la lista para que no ensucie el texto
                        words.pop(monto_idx)
                        
                    # 6. Lo que sobra es PURA DESCRIPCIÓN (Concepto + Referencia unidos limpiamente)
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
# EJECUCIÓN PRINCIPAL
# ==========================================
if archivo_subido is not None:
    if st.button(f"🚀 Convertir PDF de {banco_seleccionado}"):
        with st.spinner(f"Procesando estructura de {banco_seleccionado}..."):
            try:
                buffer_excel = None
                
                if banco_seleccionado == "BCP":
                    buffer_excel = procesar_bcp(archivo_subido, clave_pdf)
                elif banco_seleccionado == "BBVA":
                    buffer_excel = procesar_bbva(archivo_subido, clave_pdf)
                elif banco_seleccionado == "Interbank":
                    buffer_excel = procesar_interbank(archivo_subido, clave_pdf)
                elif banco_seleccionado == "Scotiabank":
                    buffer_excel = procesar_scotiabank(archivo_subido, clave_pdf)
                
                if buffer_excel is not None:
                    st.success("¡Conversión exitosa!")
                    st.download_button(
                        label="📥 Descargar archivo Excel",
                        data=buffer_excel.getvalue(),
                        file_name=f"Estado_Cuenta_{banco_seleccionado}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning(f"⚠️ No se encontraron transacciones legibles. Verifica que el archivo sea un Estado de Cuenta de {banco_seleccionado}.")
                    
            except pdfplumber.pdfminer.pdfdocument.PDFPasswordIncorrect:
                st.error("🔒 Contraseña incorrecta. Por favor ingresa el DNI/RUC válido para abrir este PDF.")
            except Exception as e:
                st.error(f"Error interno al procesar el archivo: {str(e)}")
                st.info("💡 Consejo: Verifica que hayas seleccionado el banco correcto en el menú de arriba.")

st.markdown("""
    <div class="footer-norbyte">
        <p>Potenciado por <b>Norbyte</b></p>
        <p><i>Desarrollamos soluciones de software para el usuario final, diseñadas específicamente para optimizar tu día a día.</i></p>
    </div>
""", unsafe_allow_html=True)