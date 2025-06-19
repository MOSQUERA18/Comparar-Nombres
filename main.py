import os
import re
import pandas as pd
import pdfplumber
from fuzzywuzzy import fuzz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import unicodedata
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

CARPETA_DESCARGAS = str(Path.home() / "Downloads")


CARPETA_PDFS = ""
EXCEL_VALIDACION = ""
SALIDA_EXCEL = "comparacion_nombres.xlsx"
PATRON_GENERAL = re.compile(r"A nombre de:\s*(.+?)\s*Estado:", re.IGNORECASE)
PATRON_REGISTRO_CIVIL = re.compile(r"Registro Civil,\s*(.*?)\s*tiene inscrito", re.IGNORECASE | re.DOTALL)
PATRON_MIGRACION = re.compile(
    r"el migrante venezolano\s+([A-ZÁÉÍÓÚÑ ]{5,})\s+surtió",
    re.IGNORECASE | re.DOTALL
)


VERDE = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
ROJO = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

def normalizar(texto):
    texto = texto.strip().upper()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    texto = re.sub(r"\s+", " ", texto)
    return texto

def invertir_nombre(nombre):
    partes = nombre.strip().split()
    if len(partes) >= 2:
        mitad = len(partes) // 2
        return " ".join(partes[mitad:] + partes[:mitad])
    return nombre

def extraer_nombres_desde_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        texto_completo = ""
        for pagina in pdf.pages:
            texto_completo += pagina.extract_text() + "\n"

        match1 = PATRON_GENERAL.search(texto_completo)
        if match1:
            return match1.group(1).strip()

        match2 = PATRON_REGISTRO_CIVIL.search(texto_completo)
        if match2:
            return invertir_nombre(match2.group(1).strip())

        match3 = PATRON_MIGRACION.search(texto_completo)
        if match3:
            return match3.group(1).strip()

    return None


def leer_nombres_desde_excel(ruta_excel):
    df = pd.read_excel(ruta_excel, usecols=[1], skiprows=6, engine='xlrd')
    return df.iloc[:, 0].dropna().astype(str).tolist()

def comparar_nombres(nombres_pdf, nombres_excel):
    resultados = []
    nombres_excel_norm = [normalizar(n) for n in nombres_excel]

    for nombre_pdf in nombres_pdf:
        nombre_pdf_norm = normalizar(nombre_pdf)
        mejor_match = ""
        mejor_score = 0
        for i, nombre_excel_norm in enumerate(nombres_excel_norm):
            score = fuzz.ratio(nombre_pdf_norm, nombre_excel_norm)
            if score > mejor_score:
                mejor_score = score
                mejor_match = nombres_excel[i]
        estado = "✔" if mejor_score == 100 else "❌"
        resultados.append([nombre_pdf, mejor_match, mejor_score, estado])
    return resultados

def exportar_con_colores(resultados):
    df_resultados = pd.DataFrame(resultados, columns=["Nombre PDF", "Nombre Excel", "Similitud", "Estado"])
    salida = os.path.join(CARPETA_DESCARGAS, SALIDA_EXCEL)
    df_resultados.to_excel(salida, index=False)
    wb = load_workbook(salida)

    ws = wb.active
    for row in range(2, ws.max_row + 1):
        estado = ws[f"D{row}"].value
        fill = VERDE if estado == "✔" else ROJO
        for col in ["A", "B", "C", "D"]:
            ws[f"{col}{row}"].fill = fill
        
    wb.save(salida)
    messagebox.showinfo("Listo", f"Comparación exportada en:\n{salida}")


def extraer_nombres():
    carpeta = filedialog.askdirectory(title="Seleccionar carpeta de PDFs")
    if not carpeta:
        return
    nombres = []
    no_leidos = []

    for archivo in os.listdir(carpeta):
        if archivo.lower().endswith(".pdf"):
            ruta_pdf = os.path.join(carpeta, archivo)
            nombre = extraer_nombres_desde_pdf(ruta_pdf)
            if nombre:
                nombres.append(nombre)
            else:
                no_leidos.append(archivo)

    nombres_ordenados = sorted(nombres, key=lambda n: normalizar(n))
    df = pd.DataFrame(nombres_ordenados, columns=["Nombres Extraídos"])

    salida = os.path.join(CARPETA_DESCARGAS, "nombres_extraidos.xlsx")
    df.to_excel(salida, index=False)

    msg = f"Nombres extraídos en: {salida}"
    if no_leidos:
        msg += f"\n⚠ No se pudo leer nombre de los siguientes archivos:\n- " + "\n- ".join(no_leidos)
    messagebox.showinfo("Resultado de extracción", msg)
def comparar_nombres_por_documento(pdf_datos, excel_datos):
    resultados = []
    excel_dict = {doc.strip(): nombre for nombre, doc in excel_datos}

    for nombre_pdf, doc_pdf in pdf_datos:
        doc_pdf = doc_pdf.strip() if doc_pdf else ""
        nombre_excel = excel_dict.get(doc_pdf)

        if nombre_excel:
            score = fuzz.ratio(normalizar(nombre_pdf), normalizar(nombre_excel))
            estado = "✔" if score == 100 else "❌"
        else:
            nombre_excel = "NO ENCONTRADO"
            score = 0
            estado = "❌"

        resultados.append([
            nombre_pdf, doc_pdf,
            nombre_excel, doc_pdf,
            score, estado
        ])
    return resultados



def comparar():
    carpeta = filedialog.askdirectory(title="Seleccionar carpeta de PDFs")
    if not carpeta:
        return
    archivo_excel = filedialog.askopenfilename(title="Seleccionar archivo Excel", filetypes=[("Excel files", "*.xls")])
    if not archivo_excel:
        return

    nombres_pdf = []
    no_leidos = []

    for archivo in os.listdir(carpeta):
        if archivo.lower().endswith(".pdf"):
            ruta_pdf = os.path.join(carpeta, archivo)
            nombre = extraer_nombres_desde_pdf(ruta_pdf)
            if nombre:
                nombres_pdf.append(nombre)
            else:
                no_leidos.append(archivo)

    nombres_excel = leer_nombres_desde_excel(archivo_excel)
    resultados = comparar_nombres(nombres_pdf, nombres_excel)
    exportar_con_colores(resultados)

    if no_leidos:
        messagebox.showwarning("Atención", "No se extrajo nombre de:\n\n" + "\n".join(no_leidos))

def lanzar_gui():
    ventana = tk.Tk()
    ventana.title("Validación de Nombres desde PDFs")
    ventana.geometry("400x200")

    boton_extraer = tk.Button(ventana, text="Extraer nombres de PDFs", command=extraer_nombres, height=2, width=30)
    boton_extraer.pack(pady=20)

    boton_comparar = tk.Button(ventana, text="Comparar con Excel", command=comparar, height=2, width=30)
    boton_comparar.pack(pady=10)

    ventana.mainloop()

if __name__ == "__main__":
    lanzar_gui()