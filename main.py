import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd

# 1. Configurar la ventana oculta de Tkinter
root = tk.Tk()
root.withdraw()

# 2. Abrir el explorador de archivos
print("Por favor, selecciona el archivo Excel en la ventana emergente...")
ruta_archivo = filedialog.askopenfilename(
    title="Selecciona el archivo Excel",
    filetypes=[("Archivos de Excel", "*.xlsx *.xls")],
)

# 3. Procesar el archivo si fue seleccionado
if ruta_archivo:
    print(f"Archivo seleccionado: {os.path.basename(ruta_archivo)}\n")

    # Leer el archivo Excel
    df = pd.read_excel(ruta_archivo)

    # Contar total de filas generales
    total_filas = len(df)
    # print(f"El archivo tiene un total de {total_filas} filas generales.\n")

    # 4. Buscar la columna 'tipificacion' y contar 'colaboracion'
    columna_tipificacion = "TIPIFICACION"
    valor_a_buscar = "COLABORACION"
    columna_contraventores = "CONTRAVENTORES"
    columna_detenidos = "DETENIDOS"
    columna_cierre = "TIPO DE CIERRE"
    columna_pc = "PC"
    suma = 0

    # Verificar si la columna existe en el archivo
    if columna_tipificacion in df.columns:
        # Contar cuántas celdas contienen exactamente el valor (ignorando mayúsculas/minúsculas)
        total_colaboracion = df[
            df[columna_tipificacion].astype(str).str.lower() == valor_a_buscar.lower()
        ].shape[0]

        suma = df[columna_contraventores].sum() + df[columna_detenidos].sum()

        # print(f"-> Total de celdas que contienen '{valor_a_buscar}': {total_colaboracion}")
        # print(f"-> Total de filas generales: {total_filas}")
        print(f"-> DETENIDOS TOTALES PROPIOS Y EN COLABORACION: {suma}\n")

    else:
        print(
            f"Error: No se encontró la columna '{columna_tipificacion}' en este archivo."
        )
        print(f"Las columnas disponibles son: {list(df.columns)}")

    columnas_requeridas = [
        columna_tipificacion,
        columna_detenidos,
        columna_contraventores,
        columna_pc
    ]
    columnas_faltantes = [c for c in columnas_requeridas if c not in df.columns]

    if not columnas_faltantes:
        # Asegurar que las columnas numéricas no tengan errores (reemplaza vacíos o texto por 0)
        df[columna_detenidos] = pd.to_numeric(
            df[columna_detenidos], errors="coerce"
        ).fillna(0)
        df[columna_contraventores] = pd.to_numeric(
            df[columna_contraventores], errors="coerce"
        ).fillna(0)

        # 4. Filtrar filas que contengan 'colaboración' o 'colaboracion' (con o sin tilde, ignorando mayúsculas)
        # El patrón permite buscar ambas variantes de forma segura
        patron_busqueda = r"^colaboraci[oó]n$"
        filtro_colaboracion = (
            df[columna_tipificacion]
            .astype(str)
            .str.strip()
            .str.lower()
            .str.match(patron_busqueda)
        )

        df_filtrado = df[filtro_colaboracion]
        total_coincidencias = len(df_filtrado)

        # --- INTERVENCIONES ---
        # PC con contenido: no es nula y tampoco es texto vacío o solo espacios
        tiene_pc = df[columna_pc].notna() & (
            df[columna_pc].astype(str).str.strip() != ""
        )

        # Intervención = NO es colaboración Y tiene contenido en PC
        filtro_intervenciones = (~filtro_colaboracion) & tiene_pc
        total_intervenciones = int(filtro_intervenciones.sum())

        print(f"==> INTERVENCIONES: {total_intervenciones}")

        # print(f"Resultados para '{valor_a_buscar}':")
        print(f"==> COLABORACION: {total_coincidencias}")

        if total_coincidencias > 0:
            # 5. Sumar los valores de las columnas en las filas filtradas
            suma_detenidos = int(df_filtrado[columna_detenidos].sum())
            suma_contraventores = int(df_filtrado[columna_contraventores].sum())
            suma_total_combinada = suma_detenidos + suma_contraventores

            # print(f"-> Total de '{columna_detenidos}': {suma_detenidos}")
            # print(f"-> Total de '{columna_contraventores}': {suma_contraventores}")
            print(f"==> DETENIDOS EN COLABORACION: {suma_total_combinada}")
            print(f"==> DETENIDOS TOTALES PROPIOS: {suma - suma_total_combinada}")
            print(f"==> TOTAL DE EVENTOS: {total_filas}")
        else:
            print(
                "No se calcularon sumas porque no se encontraron filas con ese criterio."
            )
    else:
        print(f"Error: Faltan columnas en el archivo Excel: {columnas_faltantes}")
        print(f"Columnas detectadas en tu archivo: {list(df.columns)}")


else:
    print("No se seleccionó ningún archivo.")
