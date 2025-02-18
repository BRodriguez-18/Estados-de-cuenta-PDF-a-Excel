
import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import Alignment

def es_fecha_valida(texto):
    """
    Verifica si el texto coincide con algo como 02/ENE o 15/FEB.
    Ajusta el patrón si tu PDF maneja otros formatos.
    """
    patron = r'^\d{1,2}/[A-Z]{3}$'  # dd/XXX (ENE, FEB, MAR, etc.)
    return bool(re.match(patron, texto.strip()))

def es_linea_movimiento(linea):
    """
    Determina si la línea es un 'movimiento' nuevo.
    Regresará True si los primeros 2 'tokens' son fechas tipo dd/ENE.
    """
    # Separamos la línea por espacios
    tokens = linea.split()
    if len(tokens) < 2:
        return False
    return es_fecha_valida(tokens[0]) and es_fecha_valida(tokens[1])

def parse_linea_movimiento(linea):
    """
    Extrae:
      - Fecha operación (tokens[0])
      - Fecha liquidación (tokens[1])
      - Un cargo al final (ej: 100,923.30) si existe
      - El resto a 'Con. Descripción'
    """
    tokens = linea.split()
    fecha_op = tokens[0]
    fecha_liq = tokens[1]

    # Buscar un valor numérico (con decimales) al final
    cargo = None
    indice_cargo = None
    for i in reversed(range(len(tokens))):
        # Busca algo como 100,923.30
        if re.search(r'[\d,]+\.\d{2}$', tokens[i]):
            cargo = tokens[i]
            indice_cargo = i
            break

    if cargo and indice_cargo > 2:
        con_desc_list = tokens[2:indice_cargo]
    else:
        con_desc_list = tokens[2:]

    con_desc = " ".join(con_desc_list)

    return {
        "Fecha operación": fecha_op,
        "Fecha liquidación": fecha_liq,
        "Con. Descripción": con_desc,
        "Referencia": None,
        "Cargos": cargo,
        "Abonos": None,
        "Saldo operación": None,
        "Saldo liquidación": None
    }

def cargar_archivo():
    global pdf_path
    archivo = filedialog.askopenfilename(
        title="Selecciona un archivo PDF",
        filetypes=(("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*"))
    )

    if archivo:
        entry_archivo.config(state=tk.NORMAL)
        entry_archivo.delete(0, tk.END)
        entry_archivo.insert(0, archivo)
        entry_archivo.config(state=tk.DISABLED)
        pdf_path = archivo

def procesar_pdf():
    global pdf_path
    if not pdf_path:
        messagebox.showwarning("Advertencia", "No se ha seleccionado un archivo PDF.")
        return

    try:
        with pdfplumber.open(pdf_path) as pdf:
            todos_los_movimientos = []
            movimiento_actual = None
            stop_phrases = ["Total de Movimientos"]
            stop_reading = False
            start_reading = False

            # Frases o palabras para ignorar la línea completa
            # (por ejemplo "Ciudad de México", "La GAT Real", etc.)
            skip_phrases = [
                "Ciudad de México",
                "Av. Paseo de la Reforma",
                "R.F.C.",
                "La GAT Real",
                "Información Financiera",
                "SUCURSAL",
                "DIRECCION",
                "PLAZA",
                "TELEFONO",
                "N/A",
                "Saldo Promedio",
                "Rendimiento",
                "Comisiones de la cuenta",
                "Cargos Objetados",
                "Saldo de Liquidación Inicial",
                "Saldo Final",
                "Estimado Cliente",
                "Estado de Cuenta",
                "MAESTRA",
                "PAGINA",
                "No. Cuenta",
                "No. Cliente",
                "FECHA",
                "SALDO",
                "OPER",
                "LIQ",
                "COD.",
                "DESCRIPCION",
                "REFERENCIA",
                "CARGOS",
                "ABONOS",
                "OPERACION",
                "LIQUIDACION",
                "También le informamos",
                "el cual puede",
                "Con BBVA",
                "BBVA MEXICO, S.A.",

            ]

            for page in pdf.pages:
                # Extraemos TODO el texto de la página como un bloque
                if stop_reading:
                    break
                
                texto_pagina = page.extract_text()
                if not texto_pagina:
                    continue

                # Separamos en líneas
                lineas = texto_pagina.split("\n")

                # Vamos a construir los movimientos de ESTA página
                # movimientos_pagina = []
                # movimiento_actual = None

                for linea in lineas:

                    # Si encontramos alguna de las frases de paro, detenemos la lectura
                    if stop_reading:
                        break

                    if any(sp in linea for sp in stop_phrases):
                        stop_reading = True
                        break  

                    # Si encontramos alguna de las frases de inicio, comenzamos a leer
                    if not start_reading:
                        tokens = linea.split()
                        found_date = any(es_fecha_valida(t) for t in tokens)
                        if found_date:
                            start_reading = True
                        else:
                            continue

                    # 1) Omitimos si contiene alguna de las skip_phrases
                    if any(sp in linea for sp in skip_phrases):
                        continue

                    # 2) Detectamos si es línea de movimiento
                    if es_linea_movimiento(linea):
                        # Guardamos el movimiento anterior
                        if movimiento_actual:
                            todos_los_movimientos.append(movimiento_actual)
                            # movimientos_pagina.append(movimiento_actual)
                        # Creamos uno nuevo
                        movimiento_actual = parse_linea_movimiento(linea)
                    else:
                        # Si no es movimiento, podría ser "referencia" o continuación
                        if not movimiento_actual:
                            # No hay movimiento anterior, creamos uno "vacío"
                            movimiento_actual = {
                                "Fecha operación": None,
                                "Fecha liquidación": None,
                                "Con. Descripción": "",
                                "Referencia": None,
                                "Cargos": None,
                                "Abonos": None,
                                "Saldo operación": None,
                                "Saldo liquidación": None
                            }
                        # Unimos esta línea al final de "Con. Descripción"
                        movimiento_actual["Con. Descripción"] += " " + linea.strip()

                # Al terminar la página, si quedó un movimiento abierto, lo guardamos
                if movimiento_actual:
                    # movimientos_pagina.append(movimiento_actual)
                    todos_los_movimientos.append(movimiento_actual)

                # Unimos con la lista global
                # todos_los_movimientos.extend(movimientos_pagina)

        # Convertimos a DataFrame
        df = pd.DataFrame(todos_los_movimientos, columns=[
            "Fecha operación",
            "Fecha liquidación",
            "Con. Descripción",
            "Cargos",
            "Abonos",
            "Saldo operación",
            "Saldo liquidación"
        ])

        # Guardamos en Excel
        ruta_salida = "movimientos_text_based.xlsx"
        df.to_excel(ruta_salida, index=False)

        # Ajustar ancho de columnas con openpyxl
        wb = load_workbook(ruta_salida)
        ws = wb.active

        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value is not None:
                    length = len(str(cell.value))
                    if length > max_length:
                        max_length = length
            ws.column_dimensions[col_letter].width = max_length + 2

        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(wrap_text=True)

        wb.save(ruta_salida)

        messagebox.showinfo("Éxito", f"Archivo Excel generado: {ruta_salida}")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al procesar el PDF:\n{e}")


# Interfaz gráfica con tkinter
root = tk.Tk()
root.title("Extracción Movimientos - Texto y Patrones (Comenzar desde primera fecha)")
root.geometry("600x250")

pdf_path = ""

btn_cargar = tk.Button(root, text="Cargar PDF", command=cargar_archivo, width=30)
btn_cargar.pack(pady=10)

entry_archivo = tk.Entry(root, width=80, state=tk.DISABLED)
entry_archivo.pack(padx=10, pady=10)

btn_procesar = tk.Button(root, text="Extraer Movimientos (Texto Completo)", command=procesar_pdf, width=30)
btn_procesar.pack(pady=10)

root.mainloop()
