import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import io
from datetime import datetime

def reset_analisis():
    st.session_state.analisis_completado = False
    st.session_state.excel_buffer = None
    st.session_state.num_asignados = 0

# 1. Detección del dispositivo
ua = st.context.headers.get("User-Agent", "").lower()
es_movil = any(x in ua for x in ["iphone", "android", "mobile", "ipad"])

# Configuración de página
st.set_page_config(page_title="Tallaje de Costaleros", page_icon="logo.png", layout="wide", initial_sidebar_state="collapsed")

components.html(
    """
    <script>
        window.parent.document.title = "Tallaje de Costaleros";
    </script>
    """,
    height=0,
)

año_actual = datetime.now().year

st.title("Tallaje de Costaleros")
st.divider()

# --- 1. ESTADO DE LA SESIÓN ---
if "analisis_completado" not in st.session_state:
    st.session_state.analisis_completado = False
if "excel_buffer" not in st.session_state:
    st.session_state.excel_buffer = None
if "num_asignados" not in st.session_state:
    st.session_state.num_asignados = 0

if es_movil:
    # En móvil, 'engañamos' al código creando contenedores simples en lugar de columnas.
    # Así el código de abajo (col1, col2, col3) sigue funcionando sin dar error.
    col1 = st.container()
    st.divider()
    col2 = st.container()
    st.divider()
    col3 = st.container()
else:
    # En ordenador, usamos tus 3 columnas con espacio grande
    col1, col2, col3 = st.columns(3, gap="large")

with col1:
    # --- FUNCIONES DE CÁLCULO BIDIRECCIONAL ---
    def actualizar_desde_interior():
        st.session_state.err_int = False
        st.session_state.err_ext = False
        c_int = st.session_state.cap_int_edit
        n_var = st.session_state.num_var
        max_p = st.session_state.max_port
        
        # E = (Total - I * (N-2)) / 2
        c_ext = (max_p - (n_var - 2) * c_int) // 2
        
        if c_ext >= 2:
            st.session_state.cap_ext_edit = c_ext
        else:
            st.session_state.err_int = True

    def actualizar_desde_exterior():
        st.session_state.err_int = False
        st.session_state.err_ext = False
        c_ext = st.session_state.cap_ext_edit
        n_var = st.session_state.num_var
        max_p = st.session_state.max_port
        
        # I = (Total - 2 * E) / (N-2)
        if n_var > 2:
            resto = max_p - 2 * c_ext
            c_int = resto / (n_var - 2)
            
            if c_int >= 2 and c_int % 2 == 0 and c_int.is_integer():
                st.session_state.cap_int_edit = int(c_int)
            else:
                st.session_state.err_ext = True

    # Inicialización de errores
    if "err_int" not in st.session_state: st.session_state.err_int = False
    if "err_ext" not in st.session_state: st.session_state.err_ext = False

    # --- 2. PARÁMETROS GENERALES ---
    st.markdown("#### Parámetros generales")
    st.write("")

    # A. Número de Varales
    st.markdown("###### Número de varales")
    val_act_var = st.session_state.get("num_var", 4)
    paso_var = 1 if val_act_var % 2 != 0 else 2

    num_varales = st.number_input(
        "Número de varales", 
        min_value=2, max_value=16, value=4, step=paso_var, 
        key="num_var", label_visibility="collapsed"
    )
    if num_varales % 2 != 0:
        st.error("✖ &nbsp;&nbsp;&nbsp;El número de varales debe ser par.")

    st.write("")

    # B. Número máximo de costaleros
    st.markdown("###### Número máximo de costaleros")
    min_c_absoluto = num_varales * 2
    val_act_max = st.session_state.get("max_port", max(60, min_c_absoluto))
    paso_max = 1 if val_act_max % 2 != 0 else 2

    max_portadores = st.number_input(
        "Número máximo de costaleros", 
        min_value=min_c_absoluto, 
        value=max(60, min_c_absoluto), 
        step=paso_max, 
        key="max_port", 
        label_visibility="collapsed"
    )
    if max_portadores % 2 != 0:
        st.error("✖ &nbsp;&nbsp;&nbsp;El número máximo de costaleros debe ser par.")

    # --- C. CÁLCULO DE SUGERENCIAS Y REINICIO ---
    if num_varales == 2:
        sug_i, sug_e = 0, max_portadores // 2
    else:
        ideal_i = max_portadores / (num_varales + 2)
        sug_i = max(2, int(round(ideal_i / 2) * 2))
        sug_e = int((max_portadores - (num_varales - 2) * sug_i) / 2)

    # Si cambia el total o los varales, forzamos las sugerencias ideales y limpiamos errores
    if "last_max" not in st.session_state or st.session_state.last_max != max_portadores or \
       "last_var" not in st.session_state or st.session_state.last_var != num_varales:
        st.session_state.cap_ext_edit = sug_e
        st.session_state.cap_int_edit = sug_i
        st.session_state.err_int = False
        st.session_state.err_ext = False
        st.session_state.last_max = max_portadores
        st.session_state.last_var = num_varales

    # --- D. INPUTS PARA CAPACIDADES ---
    if num_varales >= 4:
        st.write("")
        st.markdown("###### Capacidad de varales exteriores")
        capacidad_ext = st.number_input(
            "Capacidad de varales exteriores", 
            min_value=2, 
            key="cap_ext_edit", 
            step=1, 
            on_change=actualizar_desde_exterior,
            label_visibility="collapsed"
        )
        if st.session_state.err_ext:
            st.error("✖ &nbsp;&nbsp;&nbsp;Combinación no válida.")
        
        st.write("")
        st.markdown("###### Capacidad de varales interiores")
        val_act_int = st.session_state.get("cap_int_edit", 2)
        paso_int = 1 if val_act_int % 2 != 0 else 2
        
        capacidad_int = st.number_input(
            "Capacidad de varales interiores", 
            min_value=2, 
            key="cap_int_edit", 
            step=paso_int, 
            on_change=actualizar_desde_interior,
            label_visibility="collapsed"
        )
        if st.session_state.err_int:
            st.error("✖ &nbsp;&nbsp;&nbsp;Combinación no válida.")
        elif capacidad_int % 2 != 0:
            st.error("✖ &nbsp;&nbsp;&nbsp;La capacidad interior debe ser par.")
    else:
        capacidad_ext = max_portadores // 2
        capacidad_int = 0

    # Validación final (bloquea el análisis si hay algún error de cálculo)
    parametros_validos = (max_portadores % 2 == 0) and (num_varales % 2 == 0) and \
                         (capacidad_int % 2 == 0 or num_varales == 2) and \
                         not st.session_state.err_int and not st.session_state.err_ext

    # Estructura del paso
    varales_config = []
    varales_por_lado = num_varales // 2
    for lado in ["Izquierdo", "Derecho"]:
        for i in range(varales_por_lado):
            if i == 0:
                varales_config.append({"Nombre": f"Varal {lado} Exterior", "Lado": lado, "Capacidad": capacidad_ext, "Tipo": "Exterior"})
            else:
                sufijo = "" if varales_por_lado == 2 else f" {i}"
                varales_config.append({"Nombre": f"Varal {lado} Interior{sufijo}", "Lado": lado, "Capacidad": capacidad_int, "Tipo": "Interior"})

with col2:

    # --- 3. DESCARGA DE PLANTILLA ---
    st.markdown("#### Plantilla")
    st.write("")
    st.info("ⓘ &nbsp;&nbsp;&nbsp;Descarga esta plantilla y rellénala con los datos de cada costalero.")
    st.write("")
    df_plantilla = pd.DataFrame(columns=[
        "Nombre", "Preferencia de Hombro", "Altura Hombro Izquierdo (cm)", "Altura Hombro Derecho (cm)"
    ])
    buffer_plantilla = io.BytesIO()
    with pd.ExcelWriter(buffer_plantilla, engine='xlsxwriter') as writer:
        df_plantilla.to_excel(writer, index=False, sheet_name='Costaleros')
        workbook, worksheet = writer.book, writer.sheets['Costaleros']
        
        # 1. Tamaños de fuente añadidos
        header_fmt = workbook.add_format({'bold': True, 'bg_color': "#8583FF", 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 14})
        cell_fmt = workbook.add_format({'border': 1, 'valign': 'vcenter', 'font_size': 12}) 
        
        worksheet.set_row(0, 30)
        for i, col in enumerate(df_plantilla.columns):
            worksheet.set_column(i, i, 40)
            worksheet.write(0, i, col, header_fmt)
        
        num_filas = int(max_portadores) + 1
        for fila in range(1, num_filas):
            worksheet.set_row(fila, 25) 
            for col in range(len(df_plantilla.columns)):
                worksheet.write_blank(fila, col, "", cell_fmt)
        
        # 2. Título y mensaje de error en la validación
        worksheet.data_validation(1, 1, num_filas, 1, {
            'validate': 'list', 
            'source': ['Izquierdo', 'Derecho', 'Ambos'],
            'error_title': 'Opción no válida',
            'error_message': 'Por favor, elige una opción de la lista.'
        })

        # 3. Validación numérica para las alturas (Columnas C y D)
        worksheet.data_validation(1, 2, num_filas, 3, {
            'validate': 'decimal',
            'criteria': '>',
            'value': 0,
            'error_title': 'Dato no válido',
            'error_message': 'Por favor, introduce un valor numérico.'
        })

    st.download_button(
        label="Descargar", data=buffer_plantilla.getvalue(), 
        file_name="plantilla_costaleros.xlsx", use_container_width=True, disabled=not parametros_validos
    )

    st.divider()

    # --- 4. CARGA Y ANÁLISIS ---
    st.markdown("#### Carga de datos")
    st.write("")
    if not parametros_validos:
        st.warning("⚠ &nbsp;&nbsp;&nbsp;Corrige los parámetros generales para poder subir y analizar el archivo.")
    else:
        st.info("ⓘ &nbsp;&nbsp;&nbsp;Sube aquí la plantilla completada y pulsa 'Analizar'.")

    archivo = st.file_uploader("Sube el Excel", type=["xlsx"], label_visibility="collapsed", disabled=not parametros_validos, on_change=reset_analisis)

    st.write("")

    if st.button("Analizar", use_container_width=True, disabled=(not parametros_validos or archivo is None)):
        df = pd.read_excel(archivo)
        
        # Lógica de reparto por hombro y altura
        izq = df[df['Preferencia de Hombro'] == 'Izquierdo'][['Nombre', 'Altura Hombro Izquierdo (cm)']].rename(columns={'Altura Hombro Izquierdo (cm)': 'Altura'})
        der = df[df['Preferencia de Hombro'] == 'Derecho'][['Nombre', 'Altura Hombro Derecho (cm)']].rename(columns={'Altura Hombro Derecho (cm)': 'Altura'})
        ambos = df[df['Preferencia de Hombro'] == 'Ambos']
        
        for _, row in ambos.iterrows():
            if len(izq) <= len(der):
                izq = pd.concat([izq, pd.DataFrame([{'Nombre': row['Nombre'], 'Altura': row['Altura Hombro Izquierdo (cm)']}])])
            else:
                der = pd.concat([der, pd.DataFrame([{'Nombre': row['Nombre'], 'Altura': row['Altura Hombro Derecho (cm)']}])])
        
        izq = izq.sort_values(by='Altura', ascending=False).to_dict('records')
        der = der.sort_values(by='Altura', ascending=False).to_dict('records')
        
        # Asignación a varales
        resultado = []
        ptr_izq, ptr_der = 0, 0
        max_filas = max(capacidad_ext, capacidad_int)
        
        for fila in range(1, max_filas + 1):
            for varal in varales_config:
                asignar = False
                if varal["Tipo"] == "Exterior":
                    if fila <= capacidad_ext: asignar = True
                else:
                    mitad = capacidad_int // 2
                    if fila <= mitad or fila > (max_filas - mitad): asignar = True
                
                if asignar:
                    if varal["Lado"] == "Izquierdo" and ptr_der < len(der):
                        resultado.append({'Varal': varal["Nombre"], 'Fila': fila, **der[ptr_der]})
                        ptr_der += 1
                    elif varal["Lado"] == "Derecho" and ptr_izq < len(izq):
                        resultado.append({'Varal': varal["Nombre"], 'Fila': fila, **izq[ptr_izq]})
                        ptr_izq += 1

        # Crear Excel de salida
        columnas_nombres = [v["Nombre"] for v in varales_config]
        # Orden visual: Ext Izq -> Int Izq -> Int Der -> Ext Der
        izq_e = [v["Nombre"] for v in varales_config if v["Lado"] == "Izquierdo" and v["Tipo"] == "Exterior"]
        izq_i = [v["Nombre"] for v in varales_config if v["Lado"] == "Izquierdo" and v["Tipo"] == "Interior"][::-1]
        der_i = [v["Nombre"] for v in varales_config if v["Lado"] == "Derecho" and v["Tipo"] == "Interior"]
        der_e = [v["Nombre"] for v in varales_config if v["Lado"] == "Derecho" and v["Tipo"] == "Exterior"]
        orden_final = izq_e + izq_i + der_i + der_e
        
        df_res = pd.DataFrame("", index=range(max_filas), columns=orden_final)
        for r in resultado:
            df_res.at[r['Fila']-1, r['Varal']] = f"{r['Nombre']} ({str(r['Altura']).replace('.', ',')})"
            
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False, sheet_name='Cuadrante')
            workbook, worksheet = writer.book, writer.sheets['Cuadrante']
            
            # Formatos idénticos a la plantilla
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#8583FF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 14})
            cell_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 12})
            
            # Altura de la cabecera
            worksheet.set_row(0, 30)
            
            for i, col in enumerate(df_res.columns):
                worksheet.set_column(i, i, 40)
                worksheet.write(0, i, col, header_fmt)
                for f in range(len(df_res)):
                    # Altura de las filas
                    worksheet.set_row(f+1, 25)
                    val = df_res.iloc[f, i]
                    if val != "": worksheet.write(f+1, i, val, cell_fmt)
            
            # Dibujar el centro vacío si hay interiores
            if num_varales >= 4:
                centro_fmt = workbook.add_format({'bg_color': "#BEBEBE", 'border': 1})
                m_ini, m_fin = (capacidad_int // 2) + 1, max_filas - (capacidad_int // 2)
                if m_ini <= m_fin:
                    worksheet.merge_range(m_ini, 1, m_fin, len(orden_final)-2, "", centro_fmt)

        st.session_state.excel_buffer = output.getvalue()
        st.session_state.num_asignados = len(resultado)
        st.session_state.analisis_completado = True

with col3:
    # --- 5. RESULTADOS ---
    st.markdown("#### Resultados")
    st.write("")

    # Lógica de bloqueo: Se deshabilita si NO se ha analizado O si el archivo se ha eliminado
    boton_descarga_bloqueado = not st.session_state.analisis_completado or archivo is None

    # Mostramos el mensaje de éxito solo si hay un análisis válido Y el archivo sigue presente
    if st.session_state.analisis_completado and archivo is not None:
        st.success("✔ &nbsp;&nbsp;&nbsp;¡Tallaje completado! Ya puedes descargar el cuadrante.")
    else:
        st.info("ⓘ &nbsp;&nbsp;&nbsp;Aún no se ha generado el cuadrante.")

    st.write("")

    # El botón siempre es visible, pero su estado depende de 'boton_descarga_bloqueado'
    st.download_button(
        label="Descargar", 
        data=st.session_state.excel_buffer if st.session_state.analisis_completado else b"", 
        file_name=f"cuadrante_costaleros_{año_actual}.xlsx", 
        use_container_width=True, 
        type="primary",
        disabled=boton_descarga_bloqueado
    )