import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

def es_fecha_valida(texto):
    """
    Verifica si algo como '2/ENE' o '15/FEB' coincide con el patrón de fecha.
    """
    patron = r'^\d{1,2}/[A-Z]{3}$'
    return bool(re.match(patron, texto.strip()))

def es_linea_movimiento(linea):
    """
    Determina si la línea inicia un 'movimiento' nuevo.
    Regresará True si los primeros 2 'tokens' son fechas tipo dd/ENE.
    """
    tokens = linea.split()
    if len(tokens) < 2:
        return False
    return es_fecha_valida(tokens[0]) and es_fecha_valida(tokens[1])

def es_numero_monetario(texto):
    """
    Determina si un texto es un número tipo '100,923.30'.
    Ajusta si tu PDF usa otro formato (p.ej. 100.923,30).
    """
    return bool(re.match(r'^[\d,]+\.\d{2}$', texto.strip()))

def dist(a, b):
    """Distancia absoluta entre dos valores."""
    return abs(a - b)

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
            if len(pdf.pages) == 0:
                messagebox.showinfo("Info", "El PDF está vacío.")
                return

            # 1) Detectar las posiciones X de los encabezados en la primera página
            page0_words = pdf.pages[0].extract_words()
            col_positions = {}  # dict { "CARGOS": x_center, "ABONOS": x_center, ... }

            # Ajusta estos nombres según tu PDF:
            encabezados_buscar = ["CARGOS", "ABONOS", "OPERACIÓN", "LIQUIDACIÓN"]

            for w in page0_words:
                txt_upper = w['text'].strip().upper()
                center_x = (w['x0'] + w['x1']) / 2
                if txt_upper in encabezados_buscar:
                    col_positions[txt_upper] = center_x

            # Si no detectaste todos, podrías asignar manualmente:
            # col_positions["CARGOS"] = 350
            # col_positions["ABONOS"] = 420
            # col_positions["SALDO OPERACIÓN"] = 490
            # col_positions["SALDO LIQUIDACIÓN"] = 560

            # Ordenamos las columnas por su x_center
            columnas_ordenadas = sorted(col_positions.items(), key=lambda x: x[1])
            # columnas_ordenadas = [("CARGOS", 350), ("ABONOS", 420), ...]

            # 2) Definir skip_phrases, stop_phrases, etc.
            skip_phrases = [
                "Ciudad de México", "Av. Paseo de la Reforma", "R.F.C.",
                "La GAT Real", "Información Financiera", "SUCURSAL", "DIRECCION",
                "PLAZA", "TELEFONO", "N/A", "Saldo Promedio", "Rendimiento",
                "Comisiones de la cuenta", "Cargos Objetados", "Saldo de Liquidación Inicial",
                "Saldo Final", "Estimado Cliente", "Estado de Cuenta", "MAESTRA",
                "PAGINA", "No. Cuenta", "No. Cliente", "FECHA", "SALDO", "OPER",
                "LIQ", "COD.", "DESCRIPCION", "REFERENCIA", "CARGOS", "ABONOS",
                "OPERACION", "LIQUIDACION", "También le informamos", "el cual puede",
                "Con BBVA", "BBVA MEXICO, S.A."
            ]
            stop_phrases = ["Total de Movimientos"]

            start_reading = False
            stop_reading = False

            todos_los_movimientos = []
            movimiento_actual = None

            # 3) Recorremos todas las páginas
            for page_index, page in enumerate(pdf.pages):
                if stop_reading:
                    break

                # Extraemos las words con sus coordenadas
                words = page.extract_words()

                # Agrupamos por 'top' aproximado para formar líneas
                lineas_dict = {}
                for w in words:
                    top_approx = int(w['top'])  # redondeamos
                    if top_approx not in lineas_dict:
                        lineas_dict[top_approx] = []
                    lineas_dict[top_approx].append(w)

                # Ordenamos las líneas de arriba hacia abajo
                lineas_ordenadas = sorted(lineas_dict.items(), key=lambda x: x[0])

                for top_val, words_in_line in lineas_ordenadas:
                    if stop_reading:
                        break

                    # Convertimos la línea a string (para skip_phrases, etc.)
                    line_text = " ".join(w['text'] for w in words_in_line)

                    # Checar stop_phrases
                    if any(sp in line_text for sp in stop_phrases):
                        stop_reading = True
                        break

                    # Hasta que no encontremos la primera fecha, ignoramos
                    if not start_reading:
                        tokens = line_text.split()
                        found_date = any(es_fecha_valida(t) for t in tokens)
                        if found_date:
                            start_reading = True
                        else:
                            continue

                    # skip_phrases
                    if any(sp in line_text for sp in skip_phrases):
                        continue

                    # 4) Ver si la línea inicia un nuevo movimiento
                    if es_linea_movimiento(line_text):
                        # Cerramos el movimiento anterior si existe
                        if movimiento_actual:
                            todos_los_movimientos.append(movimiento_actual)

                        # Creamos un nuevo dict para el movimiento
                        tokens_line = line_text.split()
                        movimiento_actual = {
                            "Fecha operación": tokens_line[0],
                            "Fecha liquidación": tokens_line[1],
                            "Con. Descripción": "",
                            "Referencia": None,
                            "Cargos": None,
                            "Abonos": None,
                            "Saldo operación": None,
                            "Saldo liquidación": None
                        }

                        # Quitamos esos 2 tokens de la línea para que no estorben
                        # y luego asignaremos montos a columnas
                        # tokens_line = tokens_line[2:]
                        # Pero usaremos words_in_line para la asignación con x0

                    else:
                        # Es continuación
                        if not movimiento_actual:
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

                    # 5) Asignar montos por coordenadas
                    for w in words_in_line:
                        txt = w['text'].strip()
                        center_w = (w['x0'] + w['x1']) / 2

                        # Si es un número monetario, lo ubicamos en la columna más cercana
                        if es_numero_monetario(txt):
                            if columnas_ordenadas:
                                col_name, col_center = min(
                                    columnas_ordenadas,
                                    key=lambda x: dist(x[1], center_w)
                                )
                                # col_name es algo como "CARGOS", "ABONOS", etc.
                                if col_name == "CARGOS":
                                    movimiento_actual["Cargos"] = txt
                                elif col_name == "ABONOS":
                                    movimiento_actual["Abonos"] = txt
                                elif col_name == "OPERACIÓN":
                                    movimiento_actual["Saldo operación"] = txt
                                elif col_name == "LIQUIDACIÓN":
                                    movimiento_actual["Saldo liquidación"] = txt
                            else:
                                # Si no detectamos encabezados, por defecto a "Cargos"
                                movimiento_actual["Cargos"] = txt
                        else:
                            # Si no es monetario y no es la fecha de la línea,
                            # considerarlo parte de la descripción
                            # Evitamos duplicar las 2 fechas iniciales
                            if es_fecha_valida(txt):
                                continue
                            movimiento_actual["Con. Descripción"] += " " + txt

            # Al terminar TODAS las páginas
            if movimiento_actual:
                todos_los_movimientos.append(movimiento_actual)

        # Convertir a DataFrame
        df = pd.DataFrame(todos_los_movimientos, columns=[
            "Fecha operación",
            "Fecha liquidación",
            "Con. Descripción",
            "Referencia",
            "Cargos",
            "Abonos",
            "Saldo operación",
            "Saldo liquidación"
        ])

        ruta_salida = "movimientos_asignados_por_columnas.xlsx"
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
root.title("Extracción Movimientos - Con es_linea_movimiento + Columnas por Coordenadas")
root.geometry("600x250")

pdf_path = ""

btn_cargar = tk.Button(root, text="Cargar PDF", command=cargar_archivo, width=30)
btn_cargar.pack(pady=10)

entry_archivo = tk.Entry(root, width=80, state=tk.DISABLED)
entry_archivo.pack(padx=10, pady=10)

btn_procesar = tk.Button(root, text="Procesar PDF", command=procesar_pdf, width=30)
btn_procesar.pack(pady=10)

root.mainloop()
