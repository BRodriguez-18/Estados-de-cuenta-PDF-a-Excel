import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

# Palabras clave en el PDF
SALDO_INICIAL_TXT = "SALDO INICIAL:"
SALDO_FINAL_TXT   = "SALDO FINAL:"

def agrupar_por_top_con_tolerancia(words, tolerancia=2): 
    lineas_dict = {}
    for w in words:
        top_val = w['top']
        top_encontrado = None
        for tv in lineas_dict.keys():
            if abs(tv - top_val) <= tolerancia:
                top_encontrado = tv
                break
        if top_encontrado is not None:
            lineas_dict[top_encontrado].append(w)
        else:
            lineas_dict[top_val] = [w]
    lineas_ordenadas = sorted(lineas_dict.items(), key=lambda x: x[0])
    return lineas_ordenadas

def es_numero_monetario(txt):
    """
    Revisa si 'txt' es un número estilo 100,234.50
    (sin contemplar negativos aquí; ajústalo si manejas negativos).
    """
    return bool(re.match(r'^\d{1,3}(,\d{3})*\.\d{2}$', txt.strip()))

def limpiar_descripcion(texto):
    """
    Elimina de 'texto' los tokens que sean montos (ej. 100,234.50)
    o fechas (formato dd/mmm), dejando solo el texto descriptivo.
    """
    tokens = texto.split()
    tokens_filtrados = [
        tok for tok in tokens
        if not es_numero_monetario(tok) and not re.match(r'^\d{1,2}/[A-Za-z]{3}$', tok)
    ]
    return " ".join(tokens_filtrados)

def extraer_movimientos_por_linea(pdf_path):
    """
    - Empezamos en PESOS (isPesosActive = True).
    - Al encontrar la secuencia de DÓLAR => cerramos PESOS y abrimos DÓLAR.
    - Al encontrar la secuencia de EURO => cerramos DÓLAR y abrimos EURO.

    Retorna movements_by_currency = { "PESOS": [...], "DOLAR": [...], "EURO": [] }
    """

    # Secuencias para detectar:
    secuenciaDolares = ["CUENTA VISTA", "RESUMEN CUENTA", "DÓLAR AMERICANO", "MOVIMIENTOS"]
    secuenciaEuros   = ["CUENTA VISTA", "RESUMEN CUENTA", "EURO", "MOVIMIENTOS"]
    idx_dolar = 0
    idx_euro  = 0

    # Estados para saber quién está activo
    isPesosActive = True   # <--- Arrancamos en PESOS
    isDolarActive = False
    isEuroActive  = False

    leyendo_tabla = True   # Queremos leer movimientos de PESOS desde el inicio
    movimiento_actual = None

    # Columnas para asignar montos
    columnas_ordenadas = [
        ("Abonos", 320),
        ("Cargos", 400),
        ("Movimiento garantia", 480),
        ("Saldo en garantia", 550),
        ("Saldo disponible", 620),
        ("Saldo total", 680)
    ]
    ultimas_3_columnas = [
        ("Saldo en garantia", 550),
        ("Saldo disponible", 620),
        ("Saldo total", 680)
    ]

    # Diccionario final
    movements_by_currency = {
        "PESOS": [],
        "DOLAR": [],
        "EURO": []
    }

    # Moneda actual (iniciamos en PESOS)
    current_currency = "PESOS"

    with pdfplumber.open(pdf_path) as pdf:
        for page_index, page in enumerate(pdf.pages):
            words = page.extract_words()
            lineas_ordenadas = agrupar_por_top_con_tolerancia(words, tolerancia=1)

            for top_val, words_in_line in lineas_ordenadas:
                line_text = " ".join(w['text'] for w in words_in_line)
                line_text_upper = line_text.upper()

                # --------------------------
                # 1) Si DÓLAR no está activo, intentamos completar su secuencia
                #    (pero solo si ya no estamos en EURO, claro).
                # --------------------------
                if not isDolarActive and not isEuroActive:
                    palabra_esperada = secuenciaDolares[idx_dolar]
                    if palabra_esperada in line_text_upper:
                        idx_dolar += 1
                        if idx_dolar == len(secuenciaDolares):
                            # ¡Secuencia de Dólar completa!
                            print(">>> Activando DÓLAR en la página", page_index+1)
                            # Cerrar PESOS (si estaba leyendo)
                            if movimiento_actual:
                                movements_by_currency[current_currency].append(movimiento_actual)
                                movimiento_actual = None
                            # Cambiamos la moneda
                            isPesosActive = False
                            isDolarActive = True
                            current_currency = "DOLAR"
                            leyendo_tabla = True
                    # Mientras no se active, seguimos en PESOS
                else:
                    # --------------------------
                    # 2) Si DÓLAR está activo y EURO no, buscamos la secuencia de Euro
                    # --------------------------
                    if isDolarActive and not isEuroActive:
                        palabra_euro_esperada = secuenciaEuros[idx_euro]
                        if palabra_euro_esperada in line_text_upper:
                            idx_euro += 1
                            if idx_euro == len(secuenciaEuros):
                                # ¡Secuencia de Euro completa!
                                print(">>> Activando EURO en la página", page_index+1)
                                if movimiento_actual:
                                    movements_by_currency[current_currency].append(movimiento_actual)
                                    movimiento_actual = None
                                isDolarActive = False
                                isEuroActive  = True
                                current_currency = "EURO"
                                leyendo_tabla = True

                # --------------------------
                # 3) Parseamos movimientos si leyendo_tabla está True
                # --------------------------
                if not leyendo_tabla:
                    continue

                # -- SALDO INICIAL
                if SALDO_INICIAL_TXT in line_text_upper:
                    saldo_inicial_mov = {
                        "Fecha": None,
                        "Descripción": SALDO_INICIAL_TXT,
                        "Referencia": "",
                        "Abonos": None,
                        "Cargos": None,
                        "Movimiento garantia": None,
                        "Saldo en garantia": None,
                        "Saldo disponible": None,
                        "Saldo total": None
                    }
                    numeric_count = 0
                    for w in words_in_line:
                        txt = w['text'].strip()
                        if es_numero_monetario(txt):
                            center_x = (w['x0'] + w['x1']) / 2
                            col_name, _ = min(ultimas_3_columnas, key=lambda x: abs(x[1] - center_x))
                            saldo_inicial_mov[col_name] = txt
                            numeric_count += 1
                            if numeric_count == 3:
                                break
                    movements_by_currency[current_currency].append(saldo_inicial_mov)
                    continue

                # -- SALDO FINAL
                if SALDO_FINAL_TXT in line_text_upper:
                    saldo_final_mov = {
                        "Fecha": None,
                        "Descripción": SALDO_FINAL_TXT,
                        "Referencia": "",
                        "Abonos": None,
                        "Cargos": None,
                        "Movimiento garantia": None,
                        "Saldo en garantia": None,
                        "Saldo disponible": None,
                        "Saldo total": None
                    }
                    numeric_count = 0
                    for w in words_in_line:
                        txt = w['text'].strip()
                        if es_numero_monetario(txt):
                            center_x = (w['x0'] + w['x1']) / 2
                            col_name, _ = min(ultimas_3_columnas, key=lambda x: abs(x[1] - center_x))
                            saldo_final_mov[col_name] = txt
                            numeric_count += 1
                            if numeric_count == 3:
                                break
                    movements_by_currency[current_currency].append(saldo_final_mov)
                    # Cerramos la tabla actual
                    if movimiento_actual:
                        movements_by_currency[current_currency].append(movimiento_actual)
                        movimiento_actual = None
                    leyendo_tabla = False
                    print(f">>> Se cerró la tabla de {current_currency} en la página {page_index+1}")
                    continue

                # -- Buscar fecha (dd/mmm)
                tokens_line = line_text.split()
                fecha_token = None
                for tok in tokens_line:
                    if re.match(r'^\d{1,2}/[A-Za-z]{3}$', tok):
                        fecha_token = tok
                        break

                if fecha_token:
                    # Guardar el anterior
                    if movimiento_actual:
                        movements_by_currency[current_currency].append(movimiento_actual)
                    movimiento_actual = {
                        "Fecha": fecha_token,
                        "Descripción": limpiar_descripcion(line_text),
                        "Referencia": "",
                        "Abonos": None,
                        "Cargos": None,
                        "Movimiento garantia": None,
                        "Saldo en garantia": None,
                        "Saldo disponible": None,
                        "Saldo total": None
                    }
                else:
                    # Línea de continuación
                    if not movimiento_actual:
                        movimiento_actual = {
                            "Fecha": None,
                            "Descripción": "",
                            "Referencia": "",
                            "Abonos": None,
                            "Cargos": None,
                            "Movimiento garantia": None,
                            "Saldo en garantia": None,
                            "Saldo disponible": None,
                            "Saldo total": None
                        }
                    movimiento_actual["Descripción"] += " " + limpiar_descripcion(line_text)

                # -- Asignar montos
                if movimiento_actual:
                    for w in words_in_line:
                        txt = w['text'].strip()
                        center_x = (w['x0'] + w['x1']) / 2
                        if es_numero_monetario(txt):
                            col_name, _ = min(columnas_ordenadas, key=lambda x: abs(x[1] - center_x))
                            movimiento_actual[col_name] = txt

        # Al terminar, si quedó un movimiento pendiente
        if leyendo_tabla and movimiento_actual:
            movements_by_currency[current_currency].append(movimiento_actual)

    return movements_by_currency

def cargar_archivo():
    global pdf_paths
    archivos = filedialog.askopenfilenames(
        title="Selecciona uno o más archivos PDF",
        filetypes=[("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*")]
    )
    if archivos:
        entry_archivo.config(state=tk.NORMAL)
        entry_archivo.delete(0, tk.END)
        files_text = " ; ".join(archivos)
        entry_archivo.insert(0, files_text)
        entry_archivo.config(state=tk.DISABLED)
        pdf_paths = archivos

def procesar_pdf():
    global pdf_paths, output_folder
    if not pdf_paths:
        messagebox.showwarning("Advertencia", "No se ha seleccionado un archivo PDF.")
        return

    for pdf_path in pdf_paths:
        try:
            movements_by_currency = extraer_movimientos_por_linea(pdf_path)

            pdf_name = os.path.basename(pdf_path)
            pdf_stem, _ = os.path.splitext(pdf_name)
            excel_name = pdf_stem + ".xlsx"
            ruta_salida = os.path.join(output_folder, excel_name)

            with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
                for currency in ["PESOS", "DOLAR", "EURO"]:
                    df = pd.DataFrame(
                        movements_by_currency[currency],
                        columns=[
                            "Fecha","Descripción","Referencia","Abonos","Cargos","Movimiento garantia",
                            "Saldo en garantia","Saldo disponible","Saldo total"
                        ]
                    )
                    df.to_excel(writer, sheet_name=currency, index=False)

            # Estilizar
            wb = load_workbook(ruta_salida)
            for sheet_name in ["PESOS","DOLAR","EURO"]:
                if sheet_name not in wb.sheetnames:
                    continue
                ws = wb[sheet_name]

                ws.insert_rows(1, 6)
                ws["A1"] = "Banco: Monex"
                ws["A2"] = "Empresa: (por definir)"
                ws["A3"] = "No. Cuenta: (por definir)"
                ws["A4"] = "No. Cliente: (por definir)"
                ws["A5"] = "Periodo: (por definir)"
                ws["A6"] = "RFC: (por definir)"

                thin_side = Side(border_style="thin")
                thin_border = Border(top=thin_side, left=thin_side,
                                     right=thin_side, bottom=thin_side)

                header_fill = PatternFill(start_color="000080",
                                          end_color="000080",
                                          fill_type="solid")
                white_font = Font(color="FFFFFF", bold=True)

                max_row = ws.max_row
                max_col = ws.max_column

                for col in range(1, max_col + 1):
                    cell = ws.cell(row=7, column=col)
                    cell.fill = header_fill
                    cell.font = white_font
                    cell.alignment = Alignment(horizontal="center")
                    cell.border = thin_border

                for row in range(8, max_row + 1):
                    for col in range(1, max_col + 1):
                        cell = ws.cell(row=row, column=col)
                        cell.border = thin_border

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
            messagebox.showerror("Error", f"Ocurrió un error:\n{e}")

def main():
    global pdf_paths, output_folder
    if len(sys.argv) < 2:
        print("Uso: python procesar_pdf_monex.py <carpeta_salida>")
        return
    output_folder = sys.argv[1]
    pdf_paths = ""

    root = tk.Tk()
    root.title("Extracción Movimientos - Monex")
    root.geometry("600x250")
    root.update()
    root.lift()
    root.focus_force()
    root.attributes("-topmost", True)
    root.after(10, lambda: root.attributes("-topmost", False))

    btn_cargar = tk.Button(root, text="Cargar PDF", command=cargar_archivo, width=30)
    btn_cargar.pack(pady=10)

    global entry_archivo
    entry_archivo = tk.Entry(root, width=80, state=tk.DISABLED)
    entry_archivo.pack(padx=10, pady=10)

    btn_procesar = tk.Button(root, text="Procesar PDF", command=procesar_pdf, width=30)
    btn_procesar.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
