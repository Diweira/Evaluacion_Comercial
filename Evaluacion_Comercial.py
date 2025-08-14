import pandas as pd
import numpy as np
import os
from datetime import datetime

# --- CONSTANTES GLOBALES (pueden ser cargadas desde un archivo de configuración si es necesario) ---
# Costos fijos (ejemplo, ajustar según realidad)
COSTO_INHOUSE_FIJO = 2000000  # Costo fijo mensual de InHouse
COSTO_PRIMERA_MILLA_FIJO = 2000000 # Costo fijo mensual de Primera Milla

# Nombres de archivos esperados en la carpeta de datos
MAESTROS_ESPERADOS = [
    "MA_REGION.xlsx",
    "MA_CIUDAD.xlsx",
    "MA_TRONCAL.xlsx",
    "MA_SERVICIO.xlsx",
    "MA_CARGO_ADICIONAL.xlsx",
    "MA_TARIFA_PESO.xlsx",
    "MA_COSTO_HANDLING.xlsx",
    "MA_COSTO_ULTIMAMILLA.xlsx",
    "MA_TIPO_ENTREGA.xlsx" # Agregamos este porque se usa
]

# Columnas esperadas en el archivo de cotización de entrada
COLUMNAS_COTIZACION_ENTRADA = [
    "ORIGEN",
    "DESTINO",
    "TARIFARIO",
    "PESO",
    "TIPO ENTREGA",
    "TIPO SERVICIO"
]

# Columnas finales esperadas en el DataFrame de salida
COLUMNAS_RESULTADO_FINAL = [
    "REGION ORIGEN", "REGION DESTINO", "COMUNA ORIGEN", "COMUNA DESTINO",
    "CODIGO POSTAL ORIGEN", "CODIGO POSTAL DESTINO", "TARIFARIO", "PESO",
    "TIPO ENTREGA", "TIPO SERVICIO", "ID_CIUDAD_ORIGEN", "ID_CIUDAD_DESTINO",
    "ID_REGION_ORIGEN", "ID_REGION_DESTINO", "ID_SERVICIO", "ID_TIPO_ENTREGA",
    "VALOR TARIFA CLIENTE", "CARGO ADICIONAL", "VALOR HANDLING", "VALOR ULTIMA MILLA",
    "VALOR NETO", "COSTO TRONCAL", "COSTO PRIMERA MILLA", "COSTO ULTIMA MILLA",
    "COSTO HANDLING", "COSTO TOTAL", "UTILIDAD NETA", "MARGEN %", "KM_RECORRIDO"
]


class Configuracion:
    """Clase para manejar la configuración de rutas y archivos."""
    def __init__(self, base_path="data/"):
        self.base_path = base_path
        self.rutas = {
            "ma_region": os.path.join(base_path, "MA_REGION.xlsx"),
            "ma_ciudad": os.path.join(base_path, "MA_CIUDAD.xlsx"),
            "ma_troncal": os.path.join(base_path, "MA_TRONCAL.xlsx"),
            "ma_servicio": os.path.join(base_path, "MA_SERVICIO.xlsx"),
            "ma_cargo_adicional": os.path.join(base_path, "MA_CARGO_ADICIONAL.xlsx"),
            "ma_tarifa_peso": os.path.join(base_path, "MA_TARIFA_PESO.xlsx"),
            "ma_costo_handling": os.path.join(base_path, "MA_COSTO_HANDLING.xlsx"),
            "ma_costo_ultimamilla": os.path.join(base_path, "MA_COSTO_ULTIMAMILLA.xlsx"),
            "ma_tipo_entrega": os.path.join(base_path, "MA_TIPO_ENTREGA.xlsx")
        }

def validar_archivos(config: Configuracion) -> tuple[bool, list[str]]:
    """
    Valida que todos los archivos maestros esperados existan en la ruta configurada.

    Args:
        config (Configuracion): Instancia de configuración con las rutas de los archivos.

    Returns:
        tuple[bool, list[str]]: True si todos los archivos existen, False en caso contrario,
                                junto con una lista de los archivos faltantes.
    """
    archivos_faltantes = []
    todos_ok = True
    for key, path in config.rutas.items():
        if not os.path.exists(path):
            archivos_faltantes.append(os.path.basename(path))
            todos_ok = False
    return todos_ok, archivos_faltantes

def cargar_archivos(config: Configuracion, cotizar_df_input: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """
    Carga todos los DataFrames maestros y el DataFrame de cotización en un diccionario.

    Args:
        config (Configuracion): Instancia de configuración con las rutas de los archivos.
        cotizar_df_input (pd.DataFrame): DataFrame de cotización cargado desde la UI.

    Returns:
        dict[str, pd.DataFrame]: Diccionario con todos los DataFrames cargados.
    """
    archivos = {
        "cotizar": cotizar_df_input.copy(), # Usamos la copia del DF de entrada
        "ma_region": pd.read_excel(config.rutas["ma_region"]),
        "ma_ciudad": pd.read_excel(config.rutas["ma_ciudad"]),
        "ma_troncal": pd.read_excel(config.rutas["ma_troncal"]),
        "ma_servicio": pd.read_excel(config.rutas["ma_servicio"]),
        "ma_cargo_adicional": pd.read_excel(config.rutas["ma_cargo_adicional"]),
        "ma_tarifa_peso": pd.read_excel(config.rutas["ma_tarifa_peso"]),
        "ma_costo_handling": pd.read_excel(config.rutas["ma_costo_handling"]),
        "ma_costo_ultimamilla": pd.read_excel(config.rutas["ma_costo_ultimamilla"]),
        "ma_tipo_entrega": pd.read_excel(config.rutas["ma_tipo_entrega"])
    }
    return archivos

def preparar_datos(archivos: dict[str, pd.DataFrame], config: Configuracion) -> dict[str, pd.DataFrame]:
    """
    Prepara y estandariza los DataFrames cargados para el procesamiento.

    Args:
        archivos (dict): Diccionario de DataFrames cargados.
        config (Configuracion): Instancia de configuración.

    Returns:
        dict: Diccionario de DataFrames preparados.
    """
    # Estandarizar nombres de columnas y convertir a mayúsculas para uniones
    for key in archivos:
        if key != "cotizar": # No modificar cotizar_df_input en este paso
            archivos[key].columns = archivos[key].columns.str.upper().str.strip()

    # Pre-procesamiento del DataFrame de cotización
    archivos['cotizar'].columns = archivos['cotizar'].columns.str.upper().str.strip()
    archivos['cotizar']['PESO'] = pd.to_numeric(archivos['cotizar']['PESO'], errors='coerce').fillna(0) # Asegurar tipo numérico

    # Validar que las columnas esperadas estén en el DataFrame de cotización
    for col in COLUMNAS_COTIZACION_ENTRADA:
        if col not in archivos['cotizar'].columns:
            raise ValueError(f"La columna '{col}' no se encontró en el archivo de cotización. "
                             f"Asegúrate de que el archivo 'Cotizar.xlsx' tenga las columnas correctas.")

    # Asegurar que 'MA_TARIFA_PESO' tenga las columnas necesarias y el tipo de dato correcto para 'PESO_KG'
    if 'PESO_KG' in archivos['ma_tarifa_peso'].columns:
        archivos['ma_tarifa_peso']['PESO_KG'] = pd.to_numeric(archivos['ma_tarifa_peso']['PESO_KG'], errors='coerce')
        archivos['ma_tarifa_peso'] = archivos['ma_tarifa_peso'].dropna(subset=['PESO_KG'])
    else:
        raise ValueError("La columna 'PESO_KG' no se encontró en 'MA_TARIFA_PESO.xlsx'.")

    return archivos

def convertir_ciudades(archivos: dict[str, pd.DataFrame], config: Configuracion) -> tuple[dict[str, pd.DataFrame], list[str], list[str]]:
    """
    Mapea nombres de ciudades a sus IDs correspondientes en el DataFrame de cotización.

    Args:
        archivos (dict): Diccionario de DataFrames.
        config (Configuracion): Instancia de configuración.

    Returns:
        tuple[dict, list, list]: Diccionario de DataFrames actualizado, y listas de problemas en origen y destino.
    """
    cotizar_df = archivos['cotizar']
    ma_ciudad_df = archivos['ma_ciudad']
    ma_region_df = archivos['ma_region']

    # Unir MA_CIUDAD con MA_REGION para obtener el nombre de la región
    ma_ciudad_completa = pd.merge(
        ma_ciudad_df,
        ma_region_df[['ID_REGION', 'REGION']],
        left_on='ID_REGION',
        right_on='ID_REGION',
        how='left'
    )

    # Convertir nombres de ciudades a mayúsculas y limpiar espacios para uniones
    ma_ciudad_completa['COMUNA_UPPER'] = ma_ciudad_completa['COMUNA'].str.upper().str.strip()
    cotizar_df['ORIGEN_UPPER'] = cotizar_df['ORIGEN'].str.upper().str.strip()
    cotizar_df['DESTINO_UPPER'] = cotizar_df['DESTINO'].str.upper().str.strip()

    # Mapear ORIGEN
    cotizar_df = pd.merge(
        cotizar_df,
        ma_ciudad_completa[['COMUNA_UPPER', 'ID_CIUDAD', 'ID_REGION', 'REGION', 'CODIGO_POSTAL']].rename(
            columns={
                'ID_CIUDAD': 'ID_CIUDAD_ORIGEN',
                'ID_REGION': 'ID_REGION_ORIGEN',
                'REGION': 'REGION ORIGEN',
                'CODIGO_POSTAL': 'CODIGO POSTAL ORIGEN',
                'COMUNA_UPPER': 'COMUNA_ORIGEN_MAP' # Para seguimiento
            }
        ),
        left_on='ORIGEN_UPPER',
        right_on='COMUNA_ORIGEN_MAP',
        how='left'
    )

    # Mapear DESTINO
    cotizar_df = pd.merge(
        cotizar_df,
        ma_ciudad_completa[['COMUNA_UPPER', 'ID_CIUDAD', 'ID_REGION', 'REGION', 'CODIGO_POSTAL']].rename(
            columns={
                'ID_CIUDAD': 'ID_CIUDAD_DESTINO',
                'ID_REGION': 'ID_REGION_DESTINO',
                'REGION': 'REGION DESTINO',
                'CODIGO_POSTAL': 'CODIGO POSTAL DESTINO',
                'COMUNA_UPPER': 'COMUNA_DESTINO_MAP' # Para seguimiento
            }
        ),
        left_on='DESTINO_UPPER',
        right_on='COMUNA_DESTINO_MAP',
        how='left',
        suffixes=('_origen', '_destino') # Resolver posibles conflictos de nombres
    )

    # Identificar problemas
    origen_problemas = cotizar_df[cotizar_df['ID_CIUDAD_ORIGEN'].isnull()]['ORIGEN'].unique().tolist()
    destino_problemas = cotizar_df[cotizar_df['ID_CIUDAD_DESTINO'].isnull()]['DESTINO'].unique().tolist()

    # Eliminar columnas temporales de mapeo
    cotizar_df = cotizar_df.drop(columns=['ORIGEN_UPPER', 'DESTINO_UPPER', 'COMUNA_ORIGEN_MAP', 'COMUNA_DESTINO_MAP'])

    # Añadir columna 'COMUNA ORIGEN' y 'COMUNA DESTINO' basándose en 'ORIGEN' y 'DESTINO'
    # Esto es útil si los nombres originales son preferidos para visualización
    cotizar_df['COMUNA ORIGEN'] = cotizar_df['ORIGEN']
    cotizar_df['COMUNA DESTINO'] = cotizar_df['DESTINO']

    # Eliminar columnas ORIGINALES de "ORIGEN" y "DESTINO" para evitar duplicidad o confusión
    # después de mapearlas a 'COMUNA ORIGEN' y 'COMUNA DESTINO'
    cotizar_df = cotizar_df.drop(columns=['ORIGEN', 'DESTINO'])

    archivos['cotizar'] = cotizar_df
    return archivos, origen_problemas[:10], destino_problemas[:10] # Limitar a 10 para no sobrecargar el mensaje

def procesar_cotizaciones(archivos: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """
    Procesa las cotizaciones, uniendo con maestros y calculando valores.

    Args:
        archivos (dict): Diccionario de DataFrames cargados y preparados.

    Returns:
        pd.DataFrame: DataFrame con las cotizaciones procesadas y valores calculados.
    """
    cotizar_df = archivos['cotizar'].copy()
    ma_servicio_df = archivos['ma_servicio']
    ma_tarifa_peso_df = archivos['ma_tarifa_peso']
    ma_cargo_adicional_df = archivos['ma_cargo_adicional']
    ma_tipo_entrega_df = archivos['ma_tipo_entrega']
    ma_troncal_df = archivos['ma_troncal']

    # --- Unir con MA_SERVICIO ---
    cotizar_df = pd.merge(
        cotizar_df,
        ma_servicio_df[['ID_SERVICIO', 'TIPO SERVICIO']].rename(columns={'ID_SERVICIO': 'ID_SERVICIO_lookup'}),
        left_on='TIPO SERVICIO',
        right_on='TIPO SERVICIO',
        how='left'
    ).rename(columns={'ID_SERVICIO_lookup': 'ID_SERVICIO'})

    # --- Unir con MA_TIPO_ENTREGA ---
    cotizar_df = pd.merge(
        cotizar_df,
        ma_tipo_entrega_df[['ID_TIPO_ENTREGA', 'TIPO ENTREGA']].rename(columns={'ID_TIPO_ENTREGA': 'ID_TIPO_ENTREGA_lookup'}),
        left_on='TIPO ENTREGA',
        right_on='TIPO ENTREGA',
        how='left'
    ).rename(columns={'ID_TIPO_ENTREGA_lookup': 'ID_TIPO_ENTREGA'})

    # --- Calcular VALOR TARIFA CLIENTE ---
    # Unir para obtener la tarifa base por tarifario
    cotizar_df = pd.merge(
        cotizar_df,
        ma_tarifa_peso_df[['TARIFARIO', 'PESO_KG', 'VALOR_KG']].rename(columns={'PESO_KG': 'PESO_KG_lookup'}),
        on='TARIFARIO',
        how='left'
    )

    # Filtrar para encontrar el VALOR_KG correcto según el peso del envío
    def get_valor_tarifa(row):
        tarifario = row['TARIFARIO']
        peso = row['PESO']
        # Obtener tarifas para el tarifario específico
        tarifas_disponibles = ma_tarifa_peso_df[ma_tarifa_peso_df['TARIFARIO'] == tarifario]

        if tarifas_disponibles.empty:
            return np.nan # No se encontró tarifario

        # Encontrar el Peso_KG más cercano y menor o igual al peso del envío
        tarifas_filtradas = tarifas_disponibles[tarifas_disponibles['PESO_KG'] >= peso]

        if not tarifas_filtradas.empty:
            # Si hay pesos mayores o iguales, toma el menor de ellos
            return tarifas_filtradas['VALOR_KG'].min()
        else:
            # Si no hay pesos mayores o iguales, toma el VALOR_KG del mayor peso_kg disponible para ese tarifario
            if not tarifas_disponibles.empty:
                return tarifas_disponibles['VALOR_KG'].max()
            else:
                return np.nan # En caso de que no haya tarifas para ese tarifario

    cotizar_df['VALOR_KG_APLICADO'] = cotizar_df.apply(get_valor_tarifa, axis=1)
    cotizar_df['VALOR TARIFA CLIENTE'] = cotizar_df['VALOR_KG_APLICADO'] * cotizar_df['PESO']
    cotizar_df.drop(columns=['VALOR_KG_APLICADO', 'PESO_KG_lookup'], inplace=True) # Limpiar columnas auxiliares

    # --- Calcular CARGO ADICIONAL ---
    # Unir con MA_CARGO_ADICIONAL usando ID_SERVICIO y ID_TIPO_ENTREGA
    cotizar_df = pd.merge(
        cotizar_df,
        ma_cargo_adicional_df[['ID_SERVICIO', 'ID_TIPO_ENTREGA', 'CARGO_ADICIONAL']],
        on=['ID_SERVICIO', 'ID_TIPO_ENTREGA'],
        how='left'
    )
    cotizar_df['CARGO ADICIONAL'] = cotizar_df['CARGO_ADICIONAL'].fillna(0)
    cotizar_df.drop(columns=['CARGO_ADICIONAL'], inplace=True) # Limpiar columna auxiliar

    # --- Calcular COSTO TRONCAL y KM_RECORRIDO ---
    # Unir con MA_TRONCAL usando ID_REGION_ORIGEN y ID_REGION_DESTINO
    cotizar_df = pd.merge(
        cotizar_df,
        ma_troncal_df[['ID_REGION_ORIGEN', 'ID_REGION_DESTINO', 'COSTO_TRONCAL', 'KM_RECORRIDO']],
        on=['ID_REGION_ORIGEN', 'ID_REGION_DESTINO'],
        how='left'
    )
    cotizar_df['COSTO TRONCAL'] = cotizar_df['COSTO_TRONCAL'].fillna(0) * cotizar_df['PESO'] # Costo por KG * Peso
    cotizar_df['KM_RECORRIDO'] = cotizar_df['KM_RECORRIDO'].fillna(0)
    cotizar_df.drop(columns=['COSTO_TRONCAL'], inplace=True)

    # --- Calcular COSTO PRIMERA MILLA ---
    # Este es un costo fijo por envío que se sumará al costo variable total
    cotizar_df['COSTO PRIMERA MILLA'] = COSTO_PRIMERA_MILLA_FIJO / len(cotizar_df) if len(cotizar_df) > 0 else 0

    # --- Calcular VALOR NETO (Ingreso Bruto) ---
    cotizar_df['VALOR NETO'] = cotizar_df['VALOR TARIFA CLIENTE'] + cotizar_df['CARGO ADICIONAL']

    return cotizar_df

def calcular_costo_handling_final(df: pd.DataFrame, ma_costo_handling: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula el costo de handling para cada registro.

    Args:
        df (pd.DataFrame): DataFrame con los datos de cotización procesados.
        ma_costo_handling (pd.DataFrame): Maestro de costos de handling.

    Returns:
        pd.DataFrame: DataFrame con el costo de handling calculado.
    """
    # Unir con MA_COSTO_HANDLING
    df = pd.merge(
        df,
        ma_costo_handling[['ID_SERVICIO', 'ID_TIPO_ENTREGA', 'COSTO_HANDLING']].rename(columns={'COSTO_HANDLING': 'COSTO_HANDLING_LOOKUP'}),
        on=['ID_SERVICIO', 'ID_TIPO_ENTREGA'],
        how='left'
    )
    df['VALOR HANDLING'] = df['COSTO_HANDLING_LOOKUP'].fillna(0)
    df['COSTO HANDLING'] = df['COSTO_HANDLING_LOOKUP'].fillna(0)
    df.drop(columns=['COSTO_HANDLING_LOOKUP'], inplace=True)
    return df

def calcular_costo_ultimamilla_final(df: pd.DataFrame, ma_costo_ultimamilla: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula el costo de última milla para cada registro.

    Args:
        df (pd.DataFrame): DataFrame con los datos de cotización procesados.
        ma_costo_ultimamilla (pd.DataFrame): Maestro de costos de última milla.

    Returns:
        pd.DataFrame: DataFrame con el costo de última milla calculado.
    """
    # Unir con MA_COSTO_ULTIMAMILLA
    df = pd.merge(
        df,
        ma_costo_ultimamilla[['ID_REGION', 'ID_CIUDAD', 'COSTO_ULTIMAMILLA']].rename(columns={
            'ID_REGION': 'ID_REGION_LOOKUP',
            'ID_CIUDAD': 'ID_CIUDAD_LOOKUP',
            'COSTO_ULTIMAMILLA': 'COSTO_ULTIMAMILLA_LOOKUP'
        }),
        left_on=['ID_REGION_DESTINO', 'ID_CIUDAD_DESTINO'], # Se usa el destino para última milla
        right_on=['ID_REGION_LOOKUP', 'ID_CIUDAD_LOOKUP'],
        how='left'
    )
    df['VALOR ULTIMA MILLA'] = df['COSTO_ULTIMAMILLA_LOOKUP'].fillna(0)
    df['COSTO ULTIMA MILLA'] = df['COSTO_ULTIMAMILLA_LOOKUP'].fillna(0)
    df.drop(columns=['ID_REGION_LOOKUP', 'ID_CIUDAD_LOOKUP', 'COSTO_ULTIMAMILLA_LOOKUP'], inplace=True)
    return df

def preparar_dataframe_para_exportar(df: pd.DataFrame, nombre_empresa: str) -> tuple[pd.DataFrame, dict]:
    """
    Calcula los totales y prepara el DataFrame final para exportación, incluyendo un resumen.

    Args:
        df (pd.DataFrame): DataFrame procesado.
        nombre_empresa (str): Nombre de la empresa para el resumen.

    Returns:
        tuple[pd.DataFrame, dict]: DataFrame final para exportar y un diccionario de valores de resumen.
    """
    # Asegurar que todas las columnas esperadas estén presentes
    for col in COLUMNAS_RESULTADO_FINAL:
        if col not in df.columns:
            df[col] = np.nan # Añadir columnas faltantes con NaN

    # Ordenar y seleccionar solo las columnas finales
    df_final = df[COLUMNAS_RESULTADO_FINAL]

    # Calcular Costo Total por envío
    df_final['COSTO TOTAL'] = df_final['COSTO TRONCAL'] + df_final['COSTO PRIMERA MILLA'] + \
                              df_final['COSTO ULTIMA MILLA'] + df_final['COSTO HANDLING']

    # Calcular Utilidad Neta
    df_final['UTILIDAD NETA'] = (
        df_final['VALOR TARIFA CLIENTE'] +
        df_final['CARGO ADICIONAL'] +
        df_final['VALOR HANDLING'] +
        df_final['VALOR ULTIMA MILLA']
    ) - df_final['COSTO TOTAL']

    # Calcular Margen
    df_final['MARGEN %'] = df_final['UTILIDAD NETA'] / (
        df_final['VALOR TARIFA CLIENTE'] +
        df_final['CARGO ADICIONAL'] +
        df_final['VALOR HANDLING'] +
        df_final['VALOR ULTIMA MILLA']
    )
    df_final['MARGEN %'] = df_final['MARGEN %'].fillna(0) # Manejar división por cero

    # --- Calcular valores de resumen para la segunda hoja del Excel ---
    total_envios = len(df_final)
    total_valor_tarifa_cliente = df_final['VALOR TARIFA CLIENTE'].sum()
    total_cargo_adicional = df_final['CARGO ADICIONAL'].sum()
    total_costo_handling = df_final['VALOR HANDLING'].sum() # Ingreso por handling
    total_costo_ultimamilla = df_final['VALOR ULTIMA MILLA'].sum() # Ingreso por última milla

    ingreso_bruto_mensual = (total_valor_tarifa_cliente + total_cargo_adicional +
                            total_costo_handling + total_costo_ultimamilla)

    total_costo_troncal = df_final['COSTO TRONCAL'].sum()
    total_costo_primera_milla = df_final['COSTO PRIMERA MILLA'].sum()
    total_costo_ultimamilla_costo = df_final['COSTO ULTIMA MILLA'].sum() # Costo por última milla
    total_costo_handling_costo = df_final['COSTO HANDLING'].sum() # Costo por handling

    costo_total_variable = (total_costo_troncal + total_costo_primera_milla +
                            total_costo_ultimamilla_costo + total_costo_handling_costo)

    utilidad_mensual = ingreso_bruto_mensual - costo_total_variable - COSTO_INHOUSE_FIJO
    margen_porcentaje = utilidad_mensual / ingreso_bruto_mensual if ingreso_bruto_mensual != 0 else 0

    peso_promedio = df_final['PESO'].mean() if total_envios > 0 else 0
    recorrido_promedio = df_final['KM_RECORRIDO'].mean() if total_envios > 0 else 0

    resumen_valores = {
        'nombre_empresa': nombre_empresa,
        'fecha_generacion': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        'total_envios': total_envios,
        'peso_promedio': round(peso_promedio, 2),
        'recorrido_promedio': round(recorrido_promedio, 2),
        'total_valor_tarifa_cliente': total_valor_tarifa_cliente,
        'total_cargo_adicional': total_cargo_adicional,
        'total_costo_handling': total_costo_handling,
        'total_costo_ultimamilla': total_costo_ultimamilla,
        'ingreso_bruto_mensual': ingreso_bruto_mensual,
        'total_costo_troncal': total_costo_troncal,
        'total_costo_primera_milla': total_costo_primera_milla,
        'total_costo_ultimamilla_costo': total_costo_ultimamilla_costo,
        'total_costo_handling_costo': total_costo_handling_costo,
        'costo_total_variable': costo_total_variable,
        'costo_inhouse_fijo': COSTO_INHOUSE_FIJO,
        'utilidad_mensual': utilidad_mensual,
        'margen_porcentaje': margen_porcentaje
    }

    return df_final, resumen_valores

def generar_nombre_archivo(nombre_empresa: str) -> str:
    """
    Genera un nombre de archivo para el informe de salida.

    Args:
        nombre_empresa (str): Nombre de la empresa ingresado por el usuario.

    Returns:
        str: Nombre del archivo de salida.
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_limpio = "".join(c for c in nombre_empresa if c.isalnum() or c.isspace()).strip().replace(" ", "_")
    return f"Informe_Evaluacion_Comercial_{nombre_limpio}_{timestamp}.xlsx"