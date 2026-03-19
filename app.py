import streamlit as st
import pandas as pd
import io
from datetime import datetime

def reset_analisis():
    st.session_state.analisis_completado = False
    st.session_state.excel_buffer = None
    st.session_state.num_asignados = 0
    st.session_state.total_titulares = 0
    st.session_state.total_suplentes = 0
    st.session_state.df_titulares = None
    st.session_state.last_max = 0

# 1. Detección del dispositivo
ua = st.context.headers.get("User-Agent", "").lower()
es_movil = any(x in ua for x in ["iphone", "android", "mobile", "ipad"])

# Configuración de página
st.set_page_config(page_title="Tallaje de Costaleros", page_icon="logo.png", layout="wide", initial_sidebar_state="collapsed")

año_actual = datetime.now().year

st.title("Tallaje de Costaleros")
st.divider()

# --- 1. ESTADO DE LA SESIÓN ---
if "analisis_completado" not in st.session_state: st.session_state.analisis_completado = False
if "excel_buffer" not in st.session_state: st.session_state.excel_buffer = None
if "num_asignados" not in st.session_state: st.session_state.num_asignados = 0
if "total_titulares" not in st.session_state: st.session_state.total_titulares = 0
if "total_suplentes" not in st.session_state: st.session_state.total_suplentes = 0
if "df_titulares" not in st.session_state: st.session_state.df_titulares = None

if es_movil:
    col1 = st.container()
    st.divider()
    col2 = st.container()
    st.divider()
    col3 = st.container()
else:
    col1, col2, col3 = st.columns(3, gap="large")

# ==========================================
# COLUMNA 1: MAX COSTALEROS Y PLANTILLA
# ==========================================
with col1:

    st.markdown("#### Plantilla")
    st.write("")

    st.markdown("###### Número máximo de costaleros")
    val_act_max = st.session_state.get("max_port", 60)
    paso_max = 1 if val_act_max % 2 != 0 else 2

    max_portadores = st.number_input(
        "Número máximo de costaleros", 
        min_value=4, 
        value=60, 
        step=paso_max, 
        key="max_port", 
        label_visibility="collapsed"
    )
    
    st.write("")    
    
    # Condición para cambiar el mensaje
    if max_portadores % 2 != 0:
        st.error("✖ &nbsp;&nbsp;&nbsp;El número máximo de costaleros debe ser par.")
    else:
        st.info("ⓘ &nbsp;&nbsp;&nbsp;Descarga esta plantilla y rellénala con los datos de cada costalero.")

    df_plantilla = pd.DataFrame(columns=[
        "Nombre", "Preferencia de Hombro", "Altura Hombro Izquierdo (cm)", "Altura Hombro Derecho (cm)"
    ])
    
    buffer_plantilla = io.BytesIO()
    with pd.ExcelWriter(buffer_plantilla, engine='xlsxwriter') as writer:
        pestana_1 = 'Titulares' 
        pestana_2 = 'Suplentes'
        
        df_plantilla.to_excel(writer, index=False, sheet_name=pestana_1, startrow=1)
        df_plantilla.to_excel(writer, index=False, sheet_name=pestana_2, startrow=1)
        
        titulo_fmt = writer.book.add_format({'bold': True, 'bg_color': '#FF2B2B', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 14})
        titulo_supl_fmt = writer.book.add_format({'bold': True, 'bg_color': "#FF2B2B", 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 14})
        header_fmt = writer.book.add_format({'bold': True, 'bg_color': "#8583FF", 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 13})
        cell_fmt = writer.book.add_format({'border': 1, 'valign': 'vcenter', 'font_size': 12}) 
        
        num_filas = max_portadores + 2 
        opciones_hombro = ['Indiferente', 'Izquierdo', 'Derecho', 'Solo Izquierdo', 'Solo Derecho']

        configuracion = [
            (pestana_1, 'Costaleros Titulares', titulo_fmt), 
            (pestana_2, 'Costaleros Suplentes', titulo_supl_fmt)
        ]

        for nombre_pestana, titulo_fila_1, formato_titulo in configuracion:
            ws = writer.sheets[nombre_pestana]
            
            ws.set_row(0, 35)
            ws.merge_range(0, 0, 0, 3, titulo_fila_1, formato_titulo)

            ws.set_row(1, 30)
            for i, col in enumerate(df_plantilla.columns):
                ws.set_column(i, i, 40)
                ws.write(1, i, col, header_fmt)

            for fila in range(2, num_filas):
                ws.set_row(fila, 25) 
                for col in range(4): 
                    ws.write_blank(fila, col, "", cell_fmt)
            
            ws.data_validation(2, 1, num_filas - 1, 1, {
                'validate': 'list', 'source': opciones_hombro,
                'error_title': 'Opción no válida', 'error_message': 'Por favor, elige una opción de la lista.'
            })
            ws.data_validation(2, 2, num_filas, 3, {
                'validate': 'decimal', 'criteria': '>', 'value': 0,
                'error_title': 'Dato no válido', 'error_message': 'Introduce un valor numérico.'
            })

    st.download_button(
        label="Descargar", data=buffer_plantilla.getvalue(), 
        file_name="plantilla_costaleros.xlsx", use_container_width=True,
        disabled=(max_portadores % 2 != 0) # Aquí se bloquea si es impar
    )

    st.divider()

    st.markdown("#### Carga de datos")
    st.write("")

    st.info("ⓘ &nbsp;&nbsp;&nbsp;Sube aquí la plantilla rellenada.")
    
    archivo = st.file_uploader("Sube el Excel", type=["xlsx"], label_visibility="collapsed", on_change=reset_analisis)

    # Leemos el Excel automáticamente al subirlo para detectar cantidades
    # dropna(subset=['Nombre']) asegura que solo se cuenten las filas donde se ha escrito un nombre
    if archivo is not None and st.session_state.df_titulares is None:
        try:
            df_t = pd.read_excel(archivo, sheet_name='Titulares', skiprows=1).dropna(subset=['Nombre'])
            df_s = pd.read_excel(archivo, sheet_name='Suplentes', skiprows=1).dropna(subset=['Nombre'])
            st.session_state.total_titulares = len(df_t)
            st.session_state.total_suplentes = len(df_s)
            st.session_state.df_titulares = df_t
        except:
            st.session_state.total_titulares = 0

    

# ==========================================
# COLUMNA 2: ARCHIVO -> PARÁMETROS -> ANÁLISIS
# ==========================================
with col2:
    st.markdown("#### Parámetros generales")
    st.write("")

    if archivo is not None:
        st.write(f"Costaleros detectados &nbsp; --- &nbsp; Titulares: **{st.session_state.total_titulares}** &nbsp;|&nbsp; Suplentes: **{st.session_state.total_suplentes}**")
        st.write("")

    total_t = st.session_state.total_titulares
    parametros_validos = False
    varales_config = []

    # BARRERA: Solo mostramos la configuración si hay archivo subido y detecta costaleros
    if st.session_state.df_titulares is None:
        st.warning("⚠ &nbsp;&nbsp;&nbsp;Carga los datos para poder configurar el paso.")
    elif total_t == 0:
        st.error("⚠ &nbsp;&nbsp;&nbsp;No se han detectado costaleros titulares en el archivo.")
    elif total_t % 2 != 0:
        st.error(f"⚠ &nbsp;&nbsp;&nbsp;El número de titulares ({total_t}) debe ser par.")
    else:
        def actualizar_desde_interior():
            st.session_state.err_int = False
            st.session_state.err_ext = False
            c_int = st.session_state.cap_int_edit
            n_var = st.session_state.num_var
            max_p = st.session_state.total_titulares
            c_ext = (max_p - (n_var - 2) * c_int) // 2
            if c_ext >= 2: st.session_state.cap_ext_edit = c_ext
            else: st.session_state.err_int = True

        def actualizar_desde_exterior():
            st.session_state.err_int = False
            st.session_state.err_ext = False
            c_ext = st.session_state.cap_ext_edit
            n_var = st.session_state.num_var
            max_p = st.session_state.total_titulares
            if n_var > 2:
                resto = max_p - 2 * c_ext
                c_int = resto / (n_var - 2)
                if c_int >= 2 and c_int % 2 == 0 and c_int.is_integer(): st.session_state.cap_int_edit = int(c_int)
                else: st.session_state.err_ext = True

        if "err_int" not in st.session_state: st.session_state.err_int = False
        if "err_ext" not in st.session_state: st.session_state.err_ext = False

        st.markdown("###### Número de varales")
        val_act_var = st.session_state.get("num_var", 4)
        paso_var = 1 if val_act_var % 2 != 0 else 2
        num_varales = st.number_input("Número de varales", min_value=2, max_value=16, value=4, step=paso_var, key="num_var", label_visibility="collapsed")

        if num_varales == 2: sug_i, sug_e = 0, total_t // 2
        else:
            ideal_i = total_t / (num_varales + 2)
            sug_i = max(2, int(round(ideal_i / 2) * 2))
            sug_e = int((total_t - (num_varales - 2) * sug_i) / 2)

        if "last_max" not in st.session_state or st.session_state.last_max != total_t or \
           "last_var" not in st.session_state or st.session_state.last_var != num_varales:
            st.session_state.cap_ext_edit = max(2, sug_e) # Evitamos negativos por seguridad
            st.session_state.cap_int_edit = max(2, sug_i)
            st.session_state.err_int = False
            st.session_state.err_ext = False
            st.session_state.last_max = total_t
            st.session_state.last_var = num_varales

        if num_varales >= 4:
            st.write("")
            st.markdown("###### Capacidad de varales exteriores")
            capacidad_ext = st.number_input("Capacidad de varales exteriores", min_value=2, key="cap_ext_edit", step=1, on_change=actualizar_desde_exterior, label_visibility="collapsed")
            if st.session_state.err_ext: st.error("✖ &nbsp;&nbsp;&nbsp;Combinación no válida.")
            
            st.write("")
            st.markdown("###### Capacidad de varales interiores")
            val_act_int = st.session_state.get("cap_int_edit", 2)
            paso_int = 1 if val_act_int % 2 != 0 else 2
            capacidad_int = st.number_input("Capacidad de varales interiores", min_value=2, key="cap_int_edit", step=paso_int, on_change=actualizar_desde_interior, label_visibility="collapsed")
            if st.session_state.err_int: st.error("✖ &nbsp;&nbsp;&nbsp;Combinación no válida.")
            elif capacidad_int % 2 != 0: st.error("✖ &nbsp;&nbsp;&nbsp;La capacidad interior debe ser par.")
        else:
            capacidad_ext = total_t // 2
            capacidad_int = 0

        # Validación final cruzada
        suma_calculada = (num_varales - 2) * capacidad_int + 2 * capacidad_ext if num_varales > 2 else 2 * capacidad_ext
        cuadra_total = (suma_calculada == total_t)
        
        if not cuadra_total and not st.session_state.err_int and not st.session_state.err_ext:
            st.error(f"✖ &nbsp;&nbsp;&nbsp;La capacidad total configurada ({suma_calculada}) no coincide con los Titulares del Excel ({total_t}).")

        parametros_validos = (total_t % 2 == 0) and (num_varales % 2 == 0) and \
                             (capacidad_int % 2 == 0 or num_varales == 2) and \
                             not st.session_state.err_int and not st.session_state.err_ext and \
                             cuadra_total

        if parametros_validos:
            varales_por_lado = num_varales // 2
            for lado in ["Izquierdo", "Derecho"]:
                for i in range(varales_por_lado):
                    if i == 0: varales_config.append({"Nombre": f"Varal {lado} Exterior", "Lado": lado, "Capacidad": capacidad_ext, "Tipo": "Exterior"})
                    else:
                        sufijo = "" if varales_por_lado == 2 else f" {i}"
                        varales_config.append({"Nombre": f"Varal {lado} Interior{sufijo}", "Lado": lado, "Capacidad": capacidad_int, "Tipo": "Interior"})

        st.write("")
        st.write("")
    if st.session_state.df_titulares is None:
        pass # Mensaje manejado arriba
    elif not parametros_validos: 
        st.warning("⚠ &nbsp;&nbsp;&nbsp;Corrige los parámetros para poder analizar el archivo.")
    else: 
        st.info("ⓘ &nbsp;&nbsp;&nbsp;Parámetros correctos. Pulsa 'Analizar'.")

    btn_analizar = st.button("Analizar", use_container_width=True, disabled=(not parametros_validos or archivo is None))

    # --- LÓGICA DE ANÁLISIS ---
    if btn_analizar:
        import itertools
        df = st.session_state.df_titulares.copy()
        
        # Parche de seguridad: si un hombro está en blanco, copia el otro
        df['Altura Hombro Izquierdo (cm)'] = pd.to_numeric(df['Altura Hombro Izquierdo (cm)'].astype(str).str.replace(',', '.'), errors='coerce')
        df['Altura Hombro Derecho (cm)'] = pd.to_numeric(df['Altura Hombro Derecho (cm)'].astype(str).str.replace(',', '.'), errors='coerce')
        df['Altura Hombro Izquierdo (cm)'] = df['Altura Hombro Izquierdo (cm)'].fillna(df['Altura Hombro Derecho (cm)']).fillna(0)
        df['Altura Hombro Derecho (cm)'] = df['Altura Hombro Derecho (cm)'].fillna(df['Altura Hombro Izquierdo (cm)']).fillna(0)
        
        # 1. Agrupar en filas por Altura Media
        df['Altura_Media'] = (df['Altura Hombro Izquierdo (cm)'] + df['Altura Hombro Derecho (cm)']) / 2.0
        pool = df.sort_values(by='Altura_Media', ascending=False).to_dict('records')
        
        max_filas = max(capacidad_ext, capacidad_int)
        asignaciones_por_varal = {v["Nombre"]: [] for v in varales_config}
        
        for fila in range(1, max_filas + 1):
            huecos_fila = []
            for varal in varales_config:
                if varal["Tipo"] == "Exterior" and fila <= capacidad_ext:
                    huecos_fila.append(varal)
                elif varal["Tipo"] == "Interior" and (fila <= capacidad_int // 2 or fila > max_filas - capacidad_int // 2):
                    huecos_fila.append(varal)
            
            req_total = len(huecos_fila)
            if req_total == 0: continue
            
            req_vizq = sum(1 for h in huecos_fila if h["Lado"] == "Izquierdo")
            req_vder = sum(1 for h in huecos_fila if h["Lado"] == "Derecho")
            
            # 2. Extraer los X portadores más altos
            candidatos = []
            pool_restante = []
            temp_pool = pool.copy()
            
            while len(candidatos) < req_total and temp_pool:
                necesarios = req_total - len(candidatos)
                candidatos.extend(temp_pool[:necesarios])
                temp_pool = temp_pool[necesarios:]
                
                solo_der = [c for c in candidatos if c['Preferencia de Hombro'] == 'Solo Derecho'] 
                solo_izq = [c for c in candidatos if c['Preferencia de Hombro'] == 'Solo Izquierdo'] 
                
                while len(solo_der) > req_vizq:
                    peor = solo_der[-1]
                    candidatos.remove(peor)
                    pool_restante.append(peor)
                    solo_der.remove(peor)
                
                while len(solo_izq) > req_vder:
                    peor = solo_izq[-1]
                    candidatos.remove(peor)
                    pool_restante.append(peor)
                    solo_izq.remove(peor)
                    
            pool = pool_restante + temp_pool
            pool.sort(key=lambda x: x['Altura_Media'], reverse=True) 
            
            # 3. COMPENSACIÓN MILIMÉTRICA Y FORMA DEL PASO (USANDO ALTURAS REALES)
            best_diff = float('inf')
            best_shape_score = -1
            best_pref_score = -1
            best_perm = None
            
            for perm in itertools.permutations(candidatos):
                valid = True
                sum_izq = 0
                sum_der = 0
                pref_score = 0
                shape_score = 0
                alturas_carga = {}
                
                for i, c in enumerate(perm):
                    hueco = huecos_fila[i]
                    lado = hueco['Lado']
                    pref = c['Preferencia de Hombro']
                    
                    if pref == 'Solo Izquierdo' and lado == 'Izquierdo': valid = False; break
                    if pref == 'Solo Derecho' and lado == 'Derecho': valid = False; break
                    
                    # Altura real en el hombro exacto de carga
                    h_carga = c['Altura Hombro Derecho (cm)'] if lado == 'Izquierdo' else c['Altura Hombro Izquierdo (cm)']
                    alturas_carga[hueco['Nombre']] = h_carga
                    
                    if lado == 'Izquierdo': sum_izq += h_carga
                    else: sum_der += h_carga
                        
                    if pref == 'Derecho' and lado == 'Izquierdo': pref_score += 1
                    if pref == 'Izquierdo' and lado == 'Derecho': pref_score += 1

                if not valid: continue
                
                # REGLA DE ORO FÍSICA: El exterior debe ser más alto (o igual) que el interior en el mismo lado
                if 'Varal Izquierdo Exterior' in alturas_carga and 'Varal Izquierdo Interior' in alturas_carga:
                    if alturas_carga['Varal Izquierdo Exterior'] >= alturas_carga['Varal Izquierdo Interior']:
                        shape_score += 1
                        
                if 'Varal Derecho Exterior' in alturas_carga and 'Varal Derecho Interior' in alturas_carga:
                    if alturas_carga['Varal Derecho Exterior'] >= alturas_carga['Varal Derecho Interior']:
                        shape_score += 1
                
                diff = abs(sum_izq - sum_der)
                
                # 1º Mantener la forma (Exterior > Interior) | 2º Nivelar pesos | 3º Gustos
                if shape_score > best_shape_score:
                    best_shape_score = shape_score
                    best_diff = diff
                    best_pref_score = pref_score
                    best_perm = perm
                elif shape_score == best_shape_score:
                    if diff < best_diff - 0.1: 
                        best_diff = diff
                        best_pref_score = pref_score
                        best_perm = perm
                    elif abs(diff - best_diff) <= 0.1:
                        if pref_score > best_pref_score:
                            best_diff = diff
                            best_pref_score = pref_score
                            best_perm = perm

            if best_perm is None: best_perm = candidatos
            
            for i, c in enumerate(best_perm):
                asignaciones_por_varal[huecos_fila[i]["Nombre"]].append(c)

        # 4. GUARDADO DIRECTO
        resultado = []
        for varal in varales_config:
            nombre = varal["Nombre"]
            lado = varal["Lado"]
            tipo = varal["Tipo"]
            costaleros_varal = asignaciones_por_varal[nombre]
                
            filas_validas = [f for f in range(1, max_filas + 1) if (tipo == "Exterior" and f <= capacidad_ext) or (tipo == "Interior" and (f <= capacidad_int // 2 or f > max_filas - capacidad_int // 2))]
                    
            for i, c in enumerate(costaleros_varal):
                fila_real = filas_validas[i]
                altura_real = c['Altura Hombro Derecho (cm)'] if lado == "Izquierdo" else c['Altura Hombro Izquierdo (cm)']
                resultado.append({'Varal': nombre, 'Fila': fila_real, 'Nombre': c['Nombre'], 'Altura': str(altura_real).replace('.', ',')})

        # --- GENERACIÓN DEL EXCEL ---
        columnas_nombres = [v["Nombre"] for v in varales_config]
        izq_e = [v["Nombre"] for v in varales_config if v["Lado"] == "Izquierdo" and v["Tipo"] == "Exterior"]
        izq_i = [v["Nombre"] for v in varales_config if v["Lado"] == "Izquierdo" and v["Tipo"] == "Interior"][::-1]
        der_i = [v["Nombre"] for v in varales_config if v["Lado"] == "Derecho" and v["Tipo"] == "Interior"]
        der_e = [v["Nombre"] for v in varales_config if v["Lado"] == "Derecho" and v["Tipo"] == "Exterior"]
        orden_final = izq_e + izq_i + der_i + der_e
        
        df_res = pd.DataFrame("", index=range(max_filas), columns=orden_final)
        for r in resultado:
            df_res.at[r['Fila']-1, r['Varal']] = f"{r['Nombre']} ({r['Altura']})"
            
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False, sheet_name='Cuadrante')
            workbook, worksheet = writer.book, writer.sheets['Cuadrante']
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#8583FF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 14})
            cell_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 12})
            
            worksheet.set_row(0, 30)
            for i, col in enumerate(df_res.columns):
                worksheet.set_column(i, i, 40)
                worksheet.write(0, i, col, header_fmt)
                for f in range(len(df_res)):
                    worksheet.set_row(f+1, 25)
                    val = df_res.iloc[f, i]
                    if val != "": worksheet.write(f+1, i, val, cell_fmt)
            
            if num_varales >= 4:
                centro_fmt = workbook.add_format({'bg_color': "#BEBEBE", 'border': 1})
                m_ini, m_fin = (capacidad_int // 2) + 1, max_filas - (capacidad_int // 2)
                if m_ini <= m_fin: worksheet.merge_range(m_ini, 1, m_fin, len(orden_final)-2, "", centro_fmt)

        st.session_state.excel_buffer = output.getvalue()
        st.session_state.num_asignados = len(resultado)
        st.session_state.analisis_completado = True

# ==========================================
# COLUMNA 3: RESULTADOS Y DESCARGA
# ==========================================
with col3:
    st.markdown("#### Resultados")
    st.write("")

    boton_descarga_bloqueado = not st.session_state.analisis_completado or archivo is None

    if st.session_state.analisis_completado and archivo is not None:
        st.success("✔ &nbsp;&nbsp;&nbsp;¡Tallaje completado! Ya puedes descargar el cuadrante.")
    else:
        st.info("ⓘ &nbsp;&nbsp;&nbsp;Aún no se ha generado el cuadrante.")

    st.download_button(
        label="Descargar", 
        data=st.session_state.excel_buffer if st.session_state.analisis_completado else b"", 
        file_name=f"cuadrante_costaleros_{año_actual}.xlsx", 
        use_container_width=True, 
        type="primary",
        disabled=boton_descarga_bloqueado
    )