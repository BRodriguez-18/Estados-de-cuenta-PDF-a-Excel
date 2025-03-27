import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

##########################################################
# 1) FUNCIONES DE UTILIDAD
##########################################################

REGEX_FECHA = re.compile(r'^\d{1,2}/\d{1,2}/\d{4}$')  # Ej. "17/07/2024"

def es_linea_movimiento(token):
    """Verifica si un token corresponde a fecha dd/mm/aaaa => marca inicio de movimiento."""
    return bool(REGEX_FECHA.match(token.strip()))

def es_numero_monetario(texto):
    """
    Verifica si 'texto' es un número monetario tipo 100,923.30.
    Ajusta si requieres negativos, paréntesis, etc.
    """
    return bool(re.match(r'^[\d,]+\.\d{2}$', texto.strip()))

def parse_monetario(txt):
    """
    Convierte '100,923.30' en float => 100923.30.
    Quita comas y ajusta si usas otro formato.
    """
    txt = txt.strip().replace(",", "")
    return float(txt)

def dist(a, b):
    """Distancia absoluta, para asignar columnas por cercanía."""
    return abs(a - b)

##########################################################
# 2) INTERFAZ: SELECCIONAR PDF
##########################################################

def cargar_archivo():
    global pdf_paths
    archivos = filedialog.askopenfilenames(
        title="Selecciona uno o más archivos PDF",
        filetypes=(("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*"))
    )
    if archivos:
        entry_archivo.config(state=tk.NORMAL)
        entry_archivo.delete(0, tk.END)
        files_text = " ; ".join(archivos)
        entry_archivo.insert(0, files_text)
        entry_archivo.config(state=tk.DISABLED)
        pdf_paths = archivos

##########################################################
# 3) PROCESO PRINCIPAL
##########################################################

def procesar_pdf():
    global pdf_paths, output_folder
    if not pdf_paths:
        messagebox.showwarning("Advertencia", "No se ha seleccionado un archivo PDF.")
        return

    for pdf_path in pdf_paths:
        try:
            with pdfplumber.open(pdf_path) as pdf:
                pdf_name = os.path.basename(pdf_path)
                pdf_stem, pdf_ext = os.path.splitext(pdf_name)
                excel_name = pdf_stem + "_MIFEL.xlsx"

                if len(pdf.pages) == 0:
                    messagebox.showinfo("Info", "El PDF está vacío.")
                    return

                ##############################################
                # 3.1 LEER ENCABEZADOS EN LA SEGUNDA PÁGINA
                ##############################################
                if len(pdf.pages) < 2:
                    messagebox.showinfo("Info", "El PDF no tiene segunda página.")
                    return

                page1 = pdf.pages[1]  # la segunda página
                words_page1 = page1.extract_words()

                col_positions = {}
                headers_needed = ["RETIROS", "DEPÓSITOS", "SALDO"]

                lineas_dict_page1 = {}
                for w in words_page1:
                    top_approx = int(w['top'])
                    if top_approx not in lineas_dict_page1:
                        lineas_dict_page1[top_approx] = []
                    lineas_dict_page1[top_approx].append(w)

                lineas_ordenadas_page1 = sorted(lineas_dict_page1.items(), key=lambda x: x[0])

                for top_val, words_in_line in lineas_ordenadas_page1:
                    line_text_upper = " ".join(w['text'].strip().upper() for w in words_in_line)
                    if all(h in line_text_upper for h in headers_needed):
                        for w in words_in_line:
                            w_text_up = w['text'].strip().upper()
                            if w_text_up in headers_needed:
                                center_x = (w['x0'] + w['x1']) / 2
                                col_positions[w_text_up] = center_x
                        break

                columnas_ordenadas = sorted(col_positions.items(), key=lambda x: x[1])
                if not columnas_ordenadas:
                    messagebox.showwarning(
                        "Advertencia", 
                        "No se detectaron encabezados en la segunda página. Revisa nombres y acentos."
                    )
                    return

                ##############################################
                # 3.2 DATOS DEL ENCABEZADO (OPCIONAL)
                ##############################################
                page0 = pdf.pages[0]
                text_page0_str = page0.extract_text() or ""

                numero_cliente = ""
                match_nc = re.search(r'(?i)Número de cliente\s+(\S+)', text_page0_str)
                if match_nc:
                    numero_cliente = match_nc.group(1)

                rfc_cliente = ""
                match_rfc = re.search(r'(?i)RFC\s+([A-Z0-9]{12,13})', text_page0_str)
                if match_rfc:
                    rfc_cliente = match_rfc.group(1)

                periodo_str = ""
                match_periodo = re.search(
                    r'(?i)DEL\s+(\d{1,2}/\d{1,2}/\d{4})\s+AL\s+(\d{1,2}/\d{1,2}/\d{4})',
                    text_page0_str
                )
                if match_periodo:
                    periodo_str = match_periodo.group(0)

                ##############################################
                # 3.3 FRASES A OMITIR / DETENER LECTURA
                ##############################################
                skip_phrases = [
                    "ESTADO DE CUENTA",
                    "PÁGINA",
                    "CARGOS OBJETADOS POR EL CLIENTE",
                    "FONDOS DE INVERSIÓN",
                    "INFORMACIÓN IMPORTANTE",
                    "COMPROBANTE FISCAL",
                    "CUENTA A LA VISTA"
                ]
                skip_phrases = [s.upper() for s in skip_phrases]

                stop_phrases = [
                    "SUMA DE RETIROS Y DEPÓSITOS",
                    "CARGOS OBJETADOS POR EL CLIENTE"
                ]
                stop_phrases = [s.upper() for s in stop_phrases]

                header_phrases = [
                    "FECHA REFERENCIA DESCRIPCIÓN IMPORTE"
                ]
                header_phrases = [s.upper() for s in header_phrases]

                footer_phrases = [
                    "BANCA MIFEL, S.A., INSTITUCIÓN DE BANCA MÚLTIPLE"
                ]
                footer_phrases = [s.upper() for s in footer_phrases]

                start_reading = False
                stop_reading = False

                todos_los_movimientos = []
                movimiento_actual = None

                ##############################################
                # 3.4 RECORRER TODAS LAS PÁGINAS
                ##############################################
                for page_index, page in enumerate(pdf.pages):
                    if stop_reading:
                        break

                    words = page.extract_words()
                    lineas_dict = {}
                    for w in words:
                        top_approx = int(w['top'])
                        if top_approx not in lineas_dict:
                            lineas_dict[top_approx] = []
                        lineas_dict[top_approx].append(w)

                    lineas_ordenadas = sorted(lineas_dict.items(), key=lambda x: x[0])

                    for top_val, words_in_line in lineas_ordenadas:
                        if stop_reading:
                            break

                        line_text = " ".join(w['text'] for w in words_in_line)
                        line_text_upper = line_text.upper().strip()
                        tokens_line = line_text.split()

                        # 1) FOOTER => saltamos resto
                        if any(fp in line_text_upper for fp in footer_phrases):
                            break

                        # 2) STOP => paramos
                        if any(sp in line_text_upper for sp in stop_phrases):
                            stop_reading = True
                            break

                        # 3) Buscar la primera fecha para iniciar la lectura
                        if not start_reading:
                            if len(tokens_line) > 0 and es_linea_movimiento(tokens_line[0]):
                                has_monetary = any(es_numero_monetario(tk) for tk in tokens_line[1:])
                                if has_monetary and not any(sp in line_text_upper for sp in skip_phrases):
                                    start_reading = True
                                else:
                                    continue
                            else:
                                continue

                        # 4) Omitir skip y encabezados repetidos
                        if any(sp in line_text_upper for sp in skip_phrases):
                            continue
                        if any(hp in line_text_upper for hp in header_phrases):
                            continue

                        # 5) Omitir totales
                        if "SUMA DE RETIROS" in line_text_upper or "SALDO A FECHA DE CORTE" in line_text_upper:
                            continue

                        # 6) Detectar nueva fecha => nuevo movimiento
                        if len(tokens_line) > 0 and es_linea_movimiento(tokens_line[0]):
                            has_monetary = any(es_numero_monetario(tk) for tk in tokens_line[1:])
                            if not has_monetary:
                                continue
                            if movimiento_actual:
                                todos_los_movimientos.append(movimiento_actual)
                            movimiento_actual = {
                                "Fecha": tokens_line[0],
                                "Referencia": None,
                                "Descripción": "",
                                "Retiros": None,
                                "Depositos": None,
                                "Saldo": None,
                            }
                            tokens_line = tokens_line[1:]
                        else:
                            if not movimiento_actual:
                                movimiento_actual = {
                                    "Fecha": "",
                                    "Referencia": None,
                                    "Descripción": "",
                                    "Retiros": None,
                                    "Depositos": None,
                                    "Saldo": None,
                                }

                        # 7) Asignar importes por cercanía (bounding box)
                        for w in words_in_line:
                            raw_txt = w['text'].strip()
                            if es_numero_monetario(raw_txt):
                                val = parse_monetario(raw_txt)
                                center_w = (w['x0'] + w['x1']) / 2
                                if columnas_ordenadas:
                                    col_name, col_xc = min(
                                        columnas_ordenadas,
                                        key=lambda c: dist(c[1], center_w)
                                    )
                                    if col_name == "RETIROS":
                                        movimiento_actual["Retiros"] = val
                                    elif col_name == "DEPÓSITOS":
                                        movimiento_actual["Depositos"] = val
                                    elif col_name == "SALDO":
                                        movimiento_actual["Saldo"] = val
                                else:
                                    # Si no detectó encabezados => default deposit
                                    movimiento_actual["Depositos"] = val

                        # (Cambio para limpiar montos de la desc)
                        # 8) Quitamos montos de tokens_line antes de formar Descripción
                        clean_tokens = [t for t in tokens_line if not es_numero_monetario(t)]

                        # 9) Primer token => Referencia (si no está definida)
                        if movimiento_actual["Referencia"] is None and len(clean_tokens) > 0:
                            movimiento_actual["Referencia"] = clean_tokens[0]
                            desc_tokens = clean_tokens[1:]
                        else:
                            desc_tokens = clean_tokens

                        # 10) Resto => Descripción
                        if desc_tokens:
                            movimiento_actual["Descripción"] += " " + " ".join(desc_tokens)

                    # fin for lineas_ordenadas

                # Agregamos el último
                if movimiento_actual:
                    todos_los_movimientos.append(movimiento_actual)

            ##############################################
            # 3.5 GUARDAR EN EXCEL
            ##############################################
            df = pd.DataFrame(todos_los_movimientos, columns=[
                "Fecha", "Referencia", "Descripción", "Retiros", "Depositos", "Saldo"
            ])

            ruta_salida = os.path.join(output_folder, excel_name)
            df.to_excel(ruta_salida, index=False)

            # Formato con openpyxl
            wb = load_workbook(ruta_salida)
            ws = wb.active

            ws.insert_rows(1, 6)
            ws["A1"] = f"Banco: MIFEL"
            ws["A2"] = f"Núm. Cliente: {numero_cliente}"
            ws["A3"] = f"RFC: {rfc_cliente}"
            ws["A4"] = f"Periodo: {periodo_str}"

            thin_side = Side(border_style="thin")
            thin_border = Border(
                top=thin_side, left=thin_side, right=thin_side, bottom=thin_side
            )
            header_fill = PatternFill(start_color="000080", end_color="000080", fill_type="solid")
            white_font = Font(color="FFFFFF", bold=True)

            max_row = ws.max_row
            max_col = ws.max_column

            # Encabezados en fila 7
            for col in range(1, max_col + 1):
                cell = ws.cell(row=7, column=col)
                cell.fill = header_fill
                cell.font = white_font
                cell.alignment = Alignment(horizontal="center")
                cell.border = thin_border

            # Bordes en datos
            for row in range(8, max_row + 1):
                for col in range(1, max_col + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.border = thin_border

            # Ajustar ancho
            for col in ws.columns:
                max_length = 0
                col_letter = col[0].column_letter
                for cell in col:
                    if cell.value is not None:
                        length = len(str(cell.value))
                        if length > max_length:
                            max_length = length
                ws.column_dimensions[col_letter].width = max_length + 2

            # Wrap text
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True)

            wb.save(ruta_salida)
            messagebox.showinfo("Éxito", f"Archivo Excel generado: {ruta_salida}")

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al procesar el PDF:\n{e}")

def main():
    global pdf_paths, output_folder
    if len(sys.argv) < 2:
        print("Uso: python procesar_pdf_mifel.py <carpeta_salida>")
        return
    output_folder = sys.argv[1]
    pdf_paths = ""

    root = tk.Tk()
    root.title("Extracción Movimientos - MIFEL")
    win_width = 600
    win_height = 250

    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - win_width) // 2
    y = (screen_height - win_height) // 2
    root.geometry(f"{win_width}x{win_height}+{x}+{y}")
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
