import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

def es_linea_movimiento(linea):
    """
    Determina si la línea inicia un 'movimiento' nuevo
    en formato 'dd mmm' en dos tokens separados:
    - tokens[0] = día (1-2 dígitos)
    - tokens[1] = mes (3 letras mayúsculas, p. ej. DIC, ENE)
    """
    tokens = linea.split()
    if len(tokens) < 2:
        return False

    # Verificar si tokens[0] es dd y tokens[1] es mmm
    if not re.match(r'^\d{1,2}$', tokens[0]):
        return False
    if not re.match(r'^[A-Z]{3}$', tokens[1]):
        return False

    return True

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

            # Variables de encabezado
            periodo_str = ""
            no_cuenta_str = ""
            empresa_str = ""
            no_cliente_str = ""
            rfc_str = ""

            # 1) Detectar las posiciones X de los encabezados en la primera página
            page0_words = pdf.pages[0].extract_words()
            col_positions = {}  # dict { "RETIROS": x_center, "DEPOSITOS": x_center, "SALDO": x_center }

            # Ajusta estos nombres según tu PDF
            encabezados_buscar = ["RETIROS", "DEPOSITOS", "SALDO"]

            for w in page0_words:
                txt_upper = w['text'].strip().upper()
                center_x = (w['x0'] + w['x1']) / 2
                if txt_upper in encabezados_buscar:
                    col_positions[txt_upper] = center_x

            # Ordenamos las columnas por su x_center
            columnas_ordenadas = sorted(col_positions.items(), key=lambda x: x[1])

            # 2) Definir skip_phrases y stop_phrases
            skip_phrases = [
                "ESTADO DE CUENTA AL",
                "CLIENTE:",
                "REGISTRO FEDERAL DE CONTRIBUYENTES:",
                "PÁGINA:",
                "SUC.",
                "CUENTA DE CHEQUES MONEDA NACIONAL",
                "GAT NOMINAL",
                "GAT REAL",
                "COMISIONES EFECTIVAMENTE COBRADAS",
                "LA GAT REAL ES EL RENDIMIENTO",
                "RESUMEN GENERAL",
                "PRODUCTO/SERVICIO",
                "CONTRATO",
                "CLABE INTERBANCARIA",
                "INVERSION EMPRESARIAL",
                "DOMICILIACIÓN BANAMEX",
                "RESUMEN DEL:",
                "SALDO ANTERIOR",
                "SALDO AL",
                "SALDO PROMEDIO",
                "DÍAS TRANSCURRIDOS",
                "CHEQUES GIRADOS",
                "CHEQUES EXENTOS",
                "RESUMEN POR MEDIOS DE ACCESO",
                "DETALLE DE OPERACIONES",
                "FECHA CONCEPTO RETIROS DEPOSITOS SALDO",
                "SUCURSAL",
                "REFORMA",
                "CENTRO",
                "LA FECHA DE CORTE ES LA INDICADA",
                # ... Agrega más frases si se requieren
            ]
            skip_phrases = [s.upper() for s in skip_phrases]

            stop_phrases = [
                "SALDO MINIMO REQUERIDO",  # Si lo pones aquí, cortas la lectura
                # "COMISIONES EFECTIVAMENTE COBRADAS"  # Si lo pones aquí, cortas la lectura
            ]

            start_reading = False
            stop_reading = False

            todos_los_movimientos = []
            movimiento_actual = None

            # 3) Recorrer todas las páginas
            for page_index, page in enumerate(pdf.pages):
                if stop_reading:
                    break

                words = page.extract_words()
                # Agrupamos por 'top' aproximado
                lineas_dict = {}
                for w in words:
                    top_approx = int(w['top'])
                    if top_approx not in lineas_dict:
                        lineas_dict[top_approx] = []
                    lineas_dict[top_approx].append(w)

                # Ordenar las líneas
                lineas_ordenadas = sorted(lineas_dict.items(), key=lambda x: x[0])

                for top_val, words_in_line in lineas_ordenadas:
                    if stop_reading:
                        break

                    line_text = " ".join(w['text'] for w in words_in_line)
                    line_text_upper = line_text.upper()

                    # -- DEBUG: imprimir la línea para ver cómo se extrae
                    print(f"[DEBUG] LINEA: '{line_text}'")

                    # Ejemplo de detección de período: "RESUMEN DEL: 01/DIC/2023 AL 31/DIC/2023"
                    if "RESUMEN" in line_text_upper and "DEL:" in line_text_upper:
                        tokens_line = line_text.split()
                        # Busca tokens dd/mm/yyyy
                        fechas = [t for t in tokens_line if re.match(r'^\d{1,2}/\d{1,2}/\d{4}$', t)]
                        if len(fechas) == 2:
                            periodo_str = f"{fechas[0]} al {fechas[1]}"
                        else:
                            periodo_str = line_text
                        continue

                    # Extraer número de cuenta => "CONTRATO 12300415060"
                    if "CONTRATO" in line_text_upper and not no_cuenta_str:
                        tokens_line = line_text.split()
                        no_cuenta_str = tokens_line[-1]
                        print(f"[DEBUG] Se detectó No. de Cuenta: {no_cuenta_str}")
                        continue

                    # Extraer cliente => "CLIENTE: 2855558"
                    if "CLIENTE:" in line_text_upper and not no_cliente_str:
                        tokens_line = line_text.split()
                        no_cliente_str = tokens_line[-1]
                        print(f"[DEBUG] Se detectó No. de Cliente: {no_cliente_str}")
                        continue

                    # Extraer RFC => "Registro Federal de Contribuyentes: HST781101TJ8"
                    if "REGISTRO FEDERAL DE CONTRIBUYENTES:" in line_text_upper and not rfc_str:
                        tokens_line = line_text.split()
                        rfc_str = tokens_line[-1]
                        print(f"[DEBUG] Se detectó RFC: {rfc_str}")
                        continue

                    # Extraer empresa => p. ej. "HERRAMIENTAS STANLEY SA DE CV"
                    # Evita "ESTADO DE CUENTA AL" en skip
                    # if not empresa_str and top_val < 150 not in line_text_upper:
                    #     empresa_str = line_text.strip()
                    #     print(f"[DEBUG] Se detectó Empresa: {empresa_str}")

                    # Revisar stop_phrases
                    if any(sp in line_text_upper for sp in stop_phrases):
                        print(f"[DEBUG] Se encontró stop_phrase en: {line_text}")
                        stop_reading = True
                        break

                    # Iniciar lectura cuando aparezca la primera fecha (p. ej. "07 DIC")
                    if not start_reading:
                        tokens_line = line_text.split()
                        # Buscamos si la línea tiene "dd" y "mmm"
                        # (No es 100% confiable, pero ejemplo)
                        found_day = any(re.match(r'^\d{1,2}$', t) for t in tokens_line)
                        found_month = any(re.match(r'^[A-Z]{3}$', t) for t in tokens_line)
                        if found_day and found_month:
                            print(f"[DEBUG] Se inicia lectura a partir de: '{line_text}'")
                            start_reading = True
                        else:
                            print("[DEBUG] -> No se detecta fecha aún, se omite esta línea")
                            continue

                    # Omitir líneas con skip_phrases
                    if any(sp in line_text_upper for sp in skip_phrases):
                        print(f"[DEBUG] -> Omitido por skip_phrases: {line_text}")
                        continue

                    # ¿Es un nuevo movimiento? (tokens[0] = dd, tokens[1] = mmm)
                    if es_linea_movimiento(line_text_upper):
                        print("[DEBUG] -> Se detectó nuevo movimiento!")
                        if movimiento_actual:
                            todos_los_movimientos.append(movimiento_actual)

                        tokens_line = line_text_upper.split()
                        # Unimos día y mes en "Fecha"
                        movimiento_actual = {
                            "Fecha": f"{tokens_line[0]} {tokens_line[1]}",  # p.ej. "07 DIC"
                            "Concepto": "",
                            "Retiros": None,
                            "Depositos": None,
                            "Saldo": None
                        }
                    else:
                        print("[DEBUG] -> Continuación (o no es movimiento).")
                        if not movimiento_actual:
                            movimiento_actual = {
                                "Fecha": None,
                                "Concepto": "",
                                "Retiros": None,
                                "Depositos": None,
                                "Saldo": None
                            }

                    # 5) Asignar montos por coordenadas
                    for w in words_in_line:
                        txt = w['text'].strip()
                        center_w = (w['x0'] + w['x1']) / 2

                        # ¿Es número monetario?
                        if es_numero_monetario(txt):
                            if columnas_ordenadas:
                                col_name, col_center = min(
                                    columnas_ordenadas,
                                    key=lambda x: dist(x[1], center_w)
                                )
                                if col_name == "RETIROS":
                                    movimiento_actual["Retiros"] = txt
                                elif col_name == "DEPOSITOS":
                                    movimiento_actual["Depositos"] = txt
                                elif col_name == "SALDO":
                                    movimiento_actual["Saldo"] = txt
                                print(f"[DEBUG] Asignado '{txt}' a columna '{col_name}'")
                            else:
                                movimiento_actual["Retiros"] = txt
                                print(f"[DEBUG] Asignado '{txt}' a Retiros por defecto")
                        else:
                            # Agregar texto al concepto (omitir el día y mes repetidos)
                            if re.match(r'^\d{1,2}$', txt) or re.match(r'^[A-Z]{3}$', txt):
                                continue
                            movimiento_actual["Concepto"] += " " + txt

            # Agregar el último movimiento
            if movimiento_actual:
                todos_los_movimientos.append(movimiento_actual)

        # Convertir a DataFrame
        df = pd.DataFrame(todos_los_movimientos, columns=[
            "Fecha",
            "Concepto",
            "Retiros",
            "Depositos",
            "Saldo",
        ])

        ruta_salida = "movimientos_citibanamex_dd_mmm.xlsx"
        df.to_excel(ruta_salida, index=False)

        # Ajustes en Excel
        wb = load_workbook(ruta_salida)
        ws = wb.active

        # Insertar 6 filas para encabezado
        ws.insert_rows(1, 6)

        # Encabezado
        ws["A1"] = f"Banco: Citibanamex"
        ws["A2"] = f"Empresa: {empresa_str}"
        ws["A3"] = f"No. Cuenta: {no_cuenta_str}"
        ws["A4"] = f"No. Cliente: {no_cliente_str}"
        ws["A5"] = f"Periodo: {periodo_str}"
        ws["A6"] = f"RFC: {rfc_str}"

        thin_side = Side(border_style="thin")
        thin_border = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side)

        header_fill = PatternFill(start_color="000080", end_color="000080", fill_type="solid")
        white_font = Font(color="FFFFFF", bold=True)

        max_row = ws.max_row
        max_col = ws.max_column

        # Estilo para la fila de encabezados de la tabla (fila 7)
        for col in range(1, max_col + 1):
            cell = ws.cell(row=7, column=col)
            cell.fill = header_fill
            cell.font = white_font
            cell.alignment = Alignment(horizontal="center")
            cell.border = thin_border

        # Estilizar filas de datos
        for row in range(8, max_row + 1):
            for col in range(1, max_col + 1):
                cell = ws.cell(row=row, column=col)
                cell.border = thin_border
                # if col in [3, 4, 5]:  # Retiros, Depósitos, Saldo
                #     cell.alignment = Alignment(horizontal="right")

        # Ajustar ancho de columnas
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value is not None:
                    length = len(str(cell.value))
                    if length > max_length:
                        max_length = length
            ws.column_dimensions[col_letter].width = max_length + 2

        # Ajustar alineación con wrap_text
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(wrap_text=True)

        wb.save(ruta_salida)
        messagebox.showinfo("Éxito", f"Archivo Excel generado: {ruta_salida}")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al procesar el PDF:\n{e}")

# Interfaz gráfica con tkinter
root = tk.Tk()
root.title("Extracción Movimientos - Formato dd mmm (Debug)")
root.geometry("600x250")

pdf_path = ""

btn_cargar = tk.Button(root, text="Cargar PDF", command=cargar_archivo, width=30)
btn_cargar.pack(pady=10)

entry_archivo = tk.Entry(root, width=80, state=tk.DISABLED)
entry_archivo.pack(padx=10, pady=10)

btn_procesar = tk.Button(root, text="Procesar PDF", command=procesar_pdf, width=30)
btn_procesar.pack(pady=10)

root.mainloop()
