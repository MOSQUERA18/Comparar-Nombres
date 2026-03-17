# file: CompararNombresMasivo.py

from pathlib import Path
import logging
import pandas as pd
from tkinter import filedialog, Tk
from fuzzywuzzy import fuzz

# IMPORTAMOS TUS FUNCIONES
from unidos2 import (
    leer_pdf_completo,
    extraer_nombre_desde_texto,
    leer_columna_excel,
    comparar_nombres_fuzzy,
    exportar_dataframe_con_formato
)

logging.basicConfig(level=logging.INFO)


def procesar_carpeta(carpeta: Path, excel_path: Path):

    nombres_pdf = []

    for item in carpeta.iterdir():

        if item.suffix.lower() == ".pdf":

            texto = leer_pdf_completo(item)

            nombre = extraer_nombre_desde_texto(texto)

            if nombre:

                nombres_pdf.append(nombre)

    nombres_excel = leer_columna_excel(excel_path, index_col=1)

    resultados = comparar_nombres_fuzzy(
        nombres_pdf,
        nombres_excel
    )

    df = pd.DataFrame(
        resultados,
        columns=[
            "Nombre PDF",
            "Mejor Coincidencia Excel",
            "Porcentaje Similitud",
            "Estado"
        ]
    )

    ruta_salida = carpeta / "comparacion_nombres.xlsx"

    exportar_dataframe_con_formato(
        df,
        ruta_salida,
        {4: (1, 4)}
    )

    logging.info("✔ Resultado guardado en %s", ruta_salida)


def procesar_masivo(carpeta_raiz: Path):

    total = 0

    for subcarpeta in carpeta_raiz.iterdir():

        if not subcarpeta.is_dir():

            continue

        logging.info("Procesando carpeta %s", subcarpeta.name)

        excels = list(subcarpeta.glob("*.xlsx"))

        if not excels:

            logging.warning(
                "⚠ No hay Excel en %s",
                subcarpeta
            )

            continue

        excel_path = excels[0]

        try:

            procesar_carpeta(
                subcarpeta,
                excel_path
            )

            total += 1

        except Exception as e:

            logging.error(
                "Error en %s: %s",
                subcarpeta,
                e
            )

    print(f"\n✅ Carpetas procesadas: {total}")


def seleccionar_carpeta():

    root = Tk()

    root.withdraw()

    carpeta = filedialog.askdirectory(
        title="Seleccionar carpeta MATRÍCULAS"
    )

    return Path(carpeta)


if __name__ == "__main__":

    carpeta = seleccionar_carpeta()

    if carpeta:

        procesar_masivo(carpeta)

    else:

        print("No se seleccionó carpeta")