import streamlit as st
import pandas as pd
import os
import io
from datetime import datetime
import time # Importar para simular carga

# Importar las funciones y clases del script modificado
from Evaluacion_Comercial import (
    Configuracion,
    validar_archivos,
    cargar_archivos,
    preparar_datos,
    convertir_ciudades,
    procesar_cotizaciones,
    calcular_costo_handling_final,
    calcular_costo_ultimamilla_final,
    preparar_dataframe_para_exportar,
    generar_nombre_archivo,
    COSTO_INHOUSE_FIJO,
    COSTO_PRIMERA_MILLA_FIJO
)

# --- CONFIGURACIÓN DE PÁGINA Y ESTILO STREAMLIT ---
st.set_page_config(
    page_title="Cotizador Comercial Starken",
    page_icon="📦",
    layout="centered", # Centered layout para un aspecto más compacto y enfocado
    initial_sidebar_state="auto"
)

# Rutas
DATA_FOLDER = "data/"
TEMPLATE_FILE = os.path.join(DATA_FOLDER, "Cotizar.xlsx") # Aseguramos el nombre correcto

# Instancia de configuración con la ruta de datos
config = Configuracion(base_path=DATA_FOLDER)

# --- ESTILO CSS PERSONALIZADO (MÁS PROFUNDO) ---
st.markdown(
    """
    <style>
    /* Asegurarse que el fondo principal sea blanco */
    .stApp {
        background-color: #1A3B15; /* Un gris muy claro para un contraste suave */
    }
    .main {
        background-color: #FFFFFF; /* Contenido principal blanco */
        border-radius: 15px; /* Bordes más redondeados para el área principal */
        padding: 30px; /* Más padding para el contenido */
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08); /* Sombra más pronunciada */
    }

    /* Estilo para los botones */
    .stButton>button {
        background-color: #FF6600; /* Naranja Starken */
        color: white;
        border-radius: 10px; /* Bordes muy redondeados */
        border: none;
        padding: 15px 30px; /* Más padding para botones grandes */
        font-size: 1.1em; /* Texto un poco más grande */
        font-weight: bold;
        cursor: pointer;
        transition: all 0.3s ease;
        box-shadow: 3px 3px 10px rgba(0,0,0,0.2); /* Sombra más definida */
        margin-top: 15px; /* Espacio superior */
    }
    .stButton>button:hover {
        background-color: #E65C00; /* Naranja más oscuro al pasar el ratón */
        transform: translateY(-4px); /* Efecto 3D al pasar el ratón */
        box-shadow: 5px 5px 15px rgba(0,0,0,0.3);
    }

    /* Estilo para el cargador de archivos */
    .stFileUploader>div>div>button {
        background-color: #4CAF50; /* Verde para el cargador de archivos */
        color: white; /* Texto blanco en el botón de subir */
        border-radius: 10px;
        border: none;
        padding: 12px 25px;
        font-size: 1em;
        cursor: pointer;
        transition: all 0.3s ease;
        box-shadow: 2px 2px 8px rgba(0,0,0,0.15);
    }
    .stFileUploader>div>div>button:hover {
        background-color: #45a049;
        transform: translateY(-2px);
    }

    /* Título principal de la aplicación */
    h1 {
        color: #000000 !important; /* NEGRO para el título principal, con !important */
        text-align: center;
        font-size: 3.8em; /* Título aún más grande */
        margin-bottom: 0.3em;
        text-shadow: 2px 2px 5px rgba(0,0,0,0.1);
        font-family: 'Arial Black', Gadget, sans-serif; /* Fuente más impactante */
    }
    /* Asegurarse que el texto genérico de párrafos sea oscuro */
    .stMarkdown p { 
        text-align: center;
        font-size: 1.3em;
        color: #000000; /* Gris oscuro para el texto de descripción */
        margin-bottom: 2em;
    }

    /* Subtítulos de sección (como los de "Subir archivo", "Procesar", "Descargar") */
    h3 {
        color: #000000 !important; /* NEGRO para los subtítulos de sección, con !important */
        font-size: 2.2em; /* Título de sección más grande */
        margin-bottom: 1em;
        font-weight: bold;
        text-align: center; /* Centrar títulos de sección */
        padding-bottom: 10px;
        border-bottom: 2px solid #EEEEEE;
    }

    /* Eliminar el logo */
    .logo-top-right {
        display: none !important;
    }

    /* Contenedores de sección con diseño de tarjeta */
    .st-emotion-cache-nahz7x { /* Esta clase es el contenedor principal de Streamlit para el contenido */
        border-radius: 15px !important;
        box-shadow: 0 4px 20px rgba(0,0,0,0.1); /* Sombra más fuerte para los "cards" */
        padding: 30px !important;
        margin-bottom: 30px !important; /* Espacio entre secciones */
        background-color: #FFFFFF; /* Fondo blanco para las tarjetas */
    }
    /* Estilo para los mensajes de alerta, éxito, error (asegurando contraste) */
    .stAlert { 
        border-radius: 10px;
        font-size: 1.1em;
        padding: 1rem 1.2rem;
        margin-bottom: 1rem;
    }
    .stSuccess {
        color: #006400 !important; /* Verde oscuro para texto de éxito */
        background-color: #D4EDDA !important; /* Fondo verde claro para éxito */
        border: 1px solid #C3E6CB !important;
    }
    .stInfo {
        color: #004085 !important; /* Azul oscuro para texto de información */
        background-color: #CCE5FF !important; /* Fondo azul claro para info */
        border: 1px solid #B8DAFF !important;
    }
    .stError {
        color: #721C24 !important; /* Rojo oscuro para texto de error */
        background-color: #F8D7DA !important; /* Fondo rojo claro para error */
        border: 1px solid #F5C6CB !important;
    }
    .stWarning {
        color: #856404 !important; /* Amarillo oscuro para texto de advertencia */
        background-color: #FFF3CD !important; /* Fondo amarillo claro para advertencia */
        border: 1px solid #FFEEDB !important;
    }

    .stTextInput>div>div>input { /* Estilo para el campo de texto de empresa */
        border-radius: 10px;
        border: 1px solid #CCCCCC;
        padding: 12px;
        font-size: 1.1em;
    }
    /* Asegurar que el texto dentro de los status messages también sea legible */
    .st-emotion-cache-1f819w0 { /* Clases genéricas de los mensajes de estado */
        padding: 1rem 1.2rem;
        border-radius: 0.5rem;
        margin-bottom: 1rem;
    }
    /* Asegurar color de texto para st.write normal */
    div.stMarkdown {
        color: #000000; /* Color oscuro para el texto normal */
    }
    /* También apuntar a p directamente dentro del st.markdown para textos genéricos */
    .stMarkdown p {
        color: #000000;
    }


    </style>
    """,
    unsafe_allow_html=True
)

# --- CABECERA SIN LOGO Y TÍTULO ---
st.title("Cotizador Comercial Starken 📦")
st.markdown("<p>Optimiza tus evaluaciones de rentabilidad de servicios de transporte de forma rápida y sencilla.</p>", unsafe_allow_html=True)
st.markdown("---") # Separador visual

# --- SECCIÓN: SUBIR ARCHIVO DE COTIZACIÓN ---
with st.container(border=True): # El borde de Streamlit combinado con nuestro CSS
    st.markdown("<h3>1. Sube tu Archivo de Cotización ⬆️</h3>", unsafe_allow_html=True)
    st.write("Sube aquí el archivo Excel (`Cotizar.xlsx`) que contiene los datos de las cotizaciones que deseas procesar.")
    st.info("💡 **Importante:** Tu archivo debe seguir el formato de la plantilla oficial. Si no la tienes, descárgala en el Paso 3.")
    
    uploaded_file = st.file_uploader(
        "Haz clic para seleccionar tu archivo de cotizaciones:",
        type=["xlsx"],
        accept_multiple_files=False,
        help="Solo se acepta un archivo Excel (.xlsx)."
    )
    if uploaded_file:
        st.success("🎉 ¡Archivo cargado exitosamente! Ahora puedes ir al Paso 2 para procesarlo.")
    st.markdown("---") # Separador interno

# --- SECCIÓN: PROCESAR COTIZACIÓN ---
with st.container(border=True):
    st.markdown("<h3>2. Procesa y Genera el Informe 🚀</h3>", unsafe_allow_html=True)
    st.write("Ingresa el nombre de la empresa para personalizar el informe de salida y luego haz clic en 'Procesar'.")

    nombre_empresa_input = st.text_input(
        "📝 **Nombre de la empresa:**",
        placeholder="Ej: Logística Rápida S.A.",
        help="Este nombre se incluirá en el informe y en el nombre del archivo de salida."
    )

    process_button = st.button("🚀 Procesar Cotización y Generar Informe")

    if process_button:
        if uploaded_file is None:
            st.error("🚨 **Error:** Por favor, sube un archivo de cotización en el 'Paso 1' antes de procesar.")
        elif not nombre_empresa_input:
            st.error("🚨 **Error:** Por favor, ingresa el nombre de la empresa para el informe de salida.")
        else:
            st.info("Iniciando el proceso de cálculo... ¡Esto puede tomar un momento! Por favor, espera. ⏳")
            
            # Placeholder para los mensajes de progreso
            progress_container = st.empty() 
            
            try:
                with progress_container.status("🔍 Validando archivos auxiliares del sistema...", expanded=True) as status_validar:
                    time.sleep(0.5)
                    archivos_ok, archivos_faltantes = validar_archivos(config)
                    if not archivos_ok:
                        status_validar.update(label="❌ Validación fallida.", state="error", expanded=True)
                        st.error(f"🚨 **Error crítico:** Faltan archivos maestros en la carpeta `{DATA_FOLDER}`. Asegúrate de tener todos:")
                        for f in archivos_faltantes:
                            st.write(f"- `{f}`")
                        st.stop()
                    status_validar.update(label="✅ Archivos auxiliares validados.", state="complete", expanded=False)
                
                with progress_container.status("📖 Leyendo archivo de cotización subido...", expanded=True) as status_lectura:
                    time.sleep(0.5)
                    cotizar_df_input = pd.read_excel(uploaded_file)
                    if cotizar_df_input.empty:
                        status_lectura.update(label="❌ Archivo vacío.", state="error", expanded=True)
                        st.error("🚨 **Error:** El archivo Excel subido está vacío o no contiene datos válidos.")
                        st.stop()
                    status_lectura.update(label="✅ Archivo de cotización leído.", state="complete", expanded=False)

                with progress_container.status("⚙️ Preparando y unificando datos para el cálculo...", expanded=True) as status_preparacion:
                    time.sleep(0.5)
                    archivos = cargar_archivos(config, cotizar_df_input)
                    archivos = preparar_datos(archivos, config)
                    archivos, origen_problemas, destino_problemas = convertir_ciudades(archivos, config)

                    if origen_problemas:
                        st.warning(f"⚠️ **Alerta:** Algunas ciudades de ORIGEN no fueron mapeadas correctamente (mostrando las primeras 10): {', '.join(origen_problemas)}")
                    if destino_problemas:
                        st.warning(f"⚠️ **Alerta:** Algunas ciudades de DESTINO no fueron mapeadas correctamente (mostrando las primeras 10): {', '.join(destino_problemas)}")
                    status_preparacion.update(label="✅ Datos preparados y ubicaciones mapeadas.", state="complete", expanded=False)

                with progress_container.status("🔄 Calculando cotizaciones y analizando rentabilidad... (esto puede tardar unos segundos)", expanded=True) as status_calculo:
                    time.sleep(2) # Simula procesamiento pesado
                    resultados_df = procesar_cotizaciones(archivos)
                    resultados_df = calcular_costo_handling_final(resultados_df, archivos['ma_costo_handling'])
                    resultados_df = calcular_costo_ultimamilla_final(resultados_df, archivos['ma_costo_ultimamilla'])
                    status_calculo.update(label="✅ Cotizaciones calculadas y costos finales aplicados.", state="complete", expanded=False)
                
                with progress_container.status("📊 Organizando resultados para el informe final...", expanded=True) as status_exportacion:
                    time.sleep(0.5)
                    final_df_to_export, resumen_valores = preparar_dataframe_para_exportar(resultados_df.copy(), nombre_empresa_input)
                    status_exportacion.update(label="✅ Informe listo para descarga.", state="complete", expanded=False)
                
                # Ocultar el último mensaje de progreso antes de mostrar el botón de descarga
                progress_container.empty()
                st.success("🎉 ¡Proceso completado exitosamente! Tu informe está listo para descargar.")

                # Generar el archivo Excel en memoria
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_df_to_export.to_excel(writer, sheet_name='Evaluacion Comercial', index=False)

                    # Hoja de resumen (el código de formatos y llenado de resumen es el mismo que antes)
                    worksheet_resumen = writer.book.add_worksheet('Resumen Cotizacion')
                    workbook = writer.book

                    # === DEFINICIÓN DE FORMATOS ===
                    header_merge_format = workbook.add_format({
                        'bold': True, 'align': 'center', 'valign': 'vcenter',
                        'bg_color': '#D9D9D9', 'border': 1
                    })
                    label_format = workbook.add_format({'align': 'left', 'valign': 'vcenter'})
                    value_format = workbook.add_format({'align': 'right', 'valign': 'vcenter'})
                    currency_value_format = workbook.add_format({
                        'align': 'right', 'valign': 'vcenter', 'num_format': '$#,##0'
                    })
                    percent_value_format = workbook.add_format({
                        'align': 'right', 'valign': 'vcenter', 'num_format': '0%'
                    })
                    total_label_format = workbook.add_format({
                        'bold': True, 'align': 'left', 'valign': 'vcenter', 'top': 1, 'bottom': 1
                    })
                    total_currency_format = workbook.add_format({
                        'bold': True, 'align': 'right', 'valign': 'vcenter',
                        'top': 1, 'bottom': 1, 'num_format': '$#,##0'
                    })
                    margin_format = workbook.add_format({
                        'bold': True, 'align': 'right', 'valign': 'vcenter',
                        'top': 1, 'bottom': 1, 'num_format': '0.0%'
                    })
                    sub_header_format = workbook.add_format({
                        'bold': True, 'align': 'right', 'valign': 'vcenter',
                        'bg_color': '#D9D9D9', 'top': 1, 'bottom': 1, 'left': 1, 'right': 1
                    })
                    ingreso_label_format = workbook.add_format({
                        'bold': True, 'align': 'left', 'valign': 'vcenter', 'top': 1
                    })
                    ingreso_value_format = workbook.add_format({
                        'bold': True, 'align': 'right', 'valign': 'vcenter', 'top': 1, 'num_format': '$#,##0'
                    })

                    # Ancho de columnas para la hoja de resumen
                    worksheet_resumen.set_column('A:A', 25)
                    worksheet_resumen.set_column('B:B', 15)
                    row_offset = 0

                    # === SECCIÓN COTIZACIÓN ===
                    worksheet_resumen.merge_range(row_offset, 0, row_offset, 1, 'Cotización', header_merge_format)
                    row_offset += 1
                    worksheet_resumen.write(row_offset, 0, 'Envios Mensuales', label_format)
                    worksheet_resumen.write(row_offset, 1, resumen_valores['total_envios'], value_format)
                    row_offset += 1
                    worksheet_resumen.write(row_offset, 0, 'Peso Promedio', label_format)
                    worksheet_resumen.write(row_offset, 1, resumen_valores['peso_promedio'], value_format)
                    row_offset += 1
                    worksheet_resumen.write(row_offset, 0, 'Recorrido Promedio (km)', label_format)
                    worksheet_resumen.write(row_offset, 1, resumen_valores['recorrido_promedio'], value_format)
                    row_offset += 2 # Espacio

                    # === SECCIÓN INGRESOS ===
                    worksheet_resumen.merge_range(row_offset, 0, row_offset, 1, 'Ingresos', header_merge_format)
                    row_offset += 1
                    worksheet_resumen.write(row_offset, 0, 'Valor Base (Tarifa Cliente)', label_format)
                    worksheet_resumen.write(row_offset, 1, resumen_valores['total_valor_tarifa_cliente'], currency_value_format)
                    row_offset += 1
                    worksheet_resumen.write(row_offset, 0, 'Cargo Adicional', label_format)
                    worksheet_resumen.write(row_offset, 1, resumen_valores['total_cargo_adicional'], currency_value_format)
                    row_offset += 1
                    worksheet_resumen.write(row_offset, 0, 'Valor Handling', label_format)
                    worksheet_resumen.write(row_offset, 1, resumen_valores['total_costo_handling'], currency_value_format)
                    row_offset += 1
                    worksheet_resumen.write(row_offset, 0, 'Valor Última Milla', label_format)
                    worksheet_resumen.write(row_offset, 1, resumen_valores['total_costo_ultimamilla'], currency_value_format)
                    row_offset += 1
                    worksheet_resumen.write(row_offset, 0, 'Ingreso Bruto Mensual', ingreso_label_format)
                    worksheet_resumen.write(row_offset, 1, resumen_valores['ingreso_bruto_mensual'], ingreso_value_format)
                    row_offset += 2 # Espacio

                    # === SECCIÓN COSTOS VARIABLES ===
                    worksheet_resumen.merge_range(row_offset, 0, row_offset, 1, 'Costos Variables (Mensual)', header_merge_format)
                    row_offset += 1
                    worksheet_resumen.write(row_offset, 0, 'Costo Troncal', label_format)
                    worksheet_resumen.write(row_offset, 1, resumen_valores['total_costo_troncal'], currency_value_format)
                    row_offset += 1
                    worksheet_resumen.write(row_offset, 0, 'Costo Primera Milla', label_format)
                    worksheet_resumen.write(row_offset, 1, resumen_valores['total_costo_primera_milla'], currency_value_format)
                    row_offset += 1
                    worksheet_resumen.write(row_offset, 0, 'Costo Última Milla', label_format)
                    worksheet_resumen.write(row_offset, 1, resumen_valores['total_costo_ultimamilla_costo'], currency_value_format)
                    row_offset += 1
                    worksheet_resumen.write(row_offset, 0, 'Costo Handling', label_format)
                    worksheet_resumen.write(row_offset, 1, resumen_valores['total_costo_handling_costo'], currency_value_format)
                    row_offset += 1
                    worksheet_resumen.write(row_offset, 0, 'Costo Total Variable', total_label_format)
                    worksheet_resumen.write(row_offset, 1, resumen_valores['costo_total_variable'], total_currency_format)
                    row_offset += 2 # Espacio

                    # === SECCIÓN COSTOS FIJOS ===
                    worksheet_resumen.merge_range(row_offset, 0, row_offset, 1, 'Costos Fijos (Mensual)', header_merge_format)
                    row_offset += 1
                    worksheet_resumen.write(row_offset, 0, 'InHouse', label_format)
                    worksheet_resumen.write(row_offset, 1, COSTO_INHOUSE_FIJO, currency_value_format)
                    row_offset += 1
                    worksheet_resumen.write(row_offset, 0, 'Costo Total Fijo', total_label_format)
                    worksheet_resumen.write(row_offset, 1, COSTO_INHOUSE_FIJO, total_currency_format)
                    row_offset += 2 # Espacio

                    # === SECCIÓN RESUMEN FINAL ===
                    worksheet_resumen.write(row_offset, 0, 'UTILIDAD', total_label_format)
                    worksheet_resumen.write(row_offset, 1, resumen_valores['utilidad_mensual'], total_currency_format)
                    row_offset += 1
                    worksheet_resumen.write(row_offset, 0, 'MARGEN (%)', total_label_format)
                    worksheet_resumen.write(row_offset, 1, resumen_valores['margen_porcentaje'], margin_format)

                processed_data = output.getvalue()
                
                st.download_button(
                    label="⬇️ Descargar Informe de Evaluación Comercial",
                    data=processed_data,
                    file_name=generar_nombre_archivo(nombre_empresa_input),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Haz clic para descargar el informe de evaluación comercial procesado en formato Excel."
                )
                
            except ValueError as ve:
                st.error(f"🚨 **Error de configuración o datos:** {ve}")
                st.info("Por favor, revisa la estructura de tus archivos auxiliares o el archivo de cotización subido.")
            except Exception as e:
                st.error(f"🚨 **Ocurrió un error inesperado durante el procesamiento:** {e}")
                st.info("Si el problema persiste, contacta al soporte técnico o verifica el formato del archivo de entrada.")
    st.markdown("---") # Separador interno

# --- SECCIÓN: DESCARGAR PLANTILLA TIPO ---
with st.container(border=True):
    st.markdown("<h3>3. Descarga la Plantilla Tipo 📁</h3>", unsafe_allow_html=True)
    st.write("Si necesitas la plantilla vacía con los encabezados correctos para tu archivo de cotización, descárgala aquí. Es esencial para el correcto funcionamiento del sistema.")
    
    try:
        if not os.path.exists(TEMPLATE_FILE):
            st.warning(f"⚠️ **Alerta:** La plantilla tipo no se encontró en `{TEMPLATE_FILE}`. Para que el botón funcione, por favor, ejecuta el script `generar_plantilla_vacia.py` primero.")
        else:
            with open(TEMPLATE_FILE, "rb") as file:
                st.download_button(
                    label="📥 Descargar Plantilla 'Cotizar.xlsx' Vacía",
                    data=file,
                    file_name="Cotizar.xlsx", # Nombre de descarga
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Haz clic para descargar la plantilla de entrada de cotizaciones, solo con encabezados."
                )
    except Exception as e:
        st.error(f"❌ **Error al preparar la descarga de la plantilla:** {e}")
    st.markdown("---") # Separador interno

st.markdown("<br><br><p style='text-align: center; color: #AAAAAA; font-size: 0.9em;'>© 2025 Cotizador Comercial. Todos los derechos reservados.</p>", unsafe_allow_html=True)