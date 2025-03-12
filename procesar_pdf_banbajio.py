import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

# Conjunto de meses en forma abreviada
MESES_CORTOS = {"ENE", "FEB", "MAR", "ABR", "MAY", "JUN",
                "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"}

def es_linea_movimiento(linea):  # ok
    """
    Determina si la línea inicia un movimiento nuevo en formato 'dd mmm'.
    """
    tokens = linea.split()
    if len(tokens) < 2:
        # DEBUG: No hay suficientes tokens para detectar un movimiento
        return False
    if not re.match(r'^\d{1,2}$', tokens[0]):
        # DEBUG: El primer token no es un número válido (día)
        return False
    if tokens[1] not in MESES_CORTOS:
        # DEBUG: El segundo token no coincide con ningún mes esperado
        return False
    return True

def es_numero_monetario(texto): # ok
    """
    Valida montos con/sin '$', con/sin 'USD' al final.
    Ej: 100,923.30 / $ 100,923.30 / $100,923.30 USD
    """
    patron = r'^\$?\s?[\d,]+\.\d{2}(?:\s?USD)?$'
    resultado = bool(re.match(patron, texto.strip()))
    # DEBUG: Descomenta la siguiente línea para ver si se detecta correctamente un número monetario
    # print(f"DEBUG: es_numero_monetario('{texto}') => {resultado}")
    return resultado

def dist(a, b): # ok
    """
    Calcula la distancia absoluta entre dos valores.
    """
    return abs(a - b)

def merge_tokens(token1, token2): # ok $ y 100.00  //falta 
    """
    Une dos tokens (por ejemplo, '$' y '100.00') en uno solo, combinando sus coordenadas.
    """
    new_text = token1['text'].strip() + token2['text'].strip()
    new_x0 = min(token1['x0'], token2['x0'])
    new_x1 = max(token1['x1'], token2['x1'])
    new_top = token1['top']
    merged = {
        'text': new_text,
        'x0': new_x0,
        'x1': new_x1,
        'top': new_top
    }
    # DEBUG: Ver el resultado de combinar tokens
    # print(f"DEBUG: merge_tokens: {token1['text']} + {token2['text']} => {merged}")
    return merged

def detect_headers(page0):
    """
    Recorre todos los tokens de la primera página para detectar la posición (x central)
    de cada encabezado de interés.
    Se buscan:
      - "NO. REF." (o parte de él, por ejemplo "NO. REF")
      - "DESCRIPCION DE LA OPERACION"
      - "DEPOSITOS"
      - "RETIROS"
      - "SALDO"
    """
    headers_to_find = {
        "NO. REF.": None,
        "DESCRIPCION DE LA OPERACION": None,
        "DEPOSITOS": None,
        "RETIROS": None,
        "SALDO": None
    }
    words_page0 = page0.extract_words()
    # DEBUG: Descomenta para ver todos los tokens de la primera página
    # print("DEBUG: Tokens de la primera página:", words_page0)

    saldo_count = 0
    depositos_count = 0

    for w in words_page0:
        txt = w['text'].upper().strip()
        center_x = (w['x0'] + w['x1']) / 2
        # DEBUG: Mostrar token y posición
        # print(f"DEBUG: Token='{txt}', center_x={center_x}")
        if txt == "DOCTO": # ok
            if headers_to_find["NO. REF."] is None:
                headers_to_find["NO. REF."] = center_x
                # print(f"DEBUG: Se detectó 'NO. REF.' en x={center_x}")
        elif "DESCRIPCION" in txt: 
            if headers_to_find["DESCRIPCION DE LA OPERACION"] is None:
                headers_to_find["DESCRIPCION DE LA OPERACION"] = center_x
                print(f"DEBUG: Se detectó 'DESCRIPCION DE LA OPERACION' en x={center_x}")
        elif "DEPOSITOS" in txt:
            depositos_count += 1
            if depositos_count == 2:
                headers_to_find["DEPOSITOS"] = center_x
                print(f"DEBUG: Se detectó 'DEPOSITOS' en x={center_x}")
        elif "RETIROS" in txt:
            if headers_to_find["RETIROS"] is None:
                headers_to_find["RETIROS"] = center_x
                print(f"DEBUG: Se detectó 'RETIROS' en x={center_x}")
        elif "SALDO" in txt:
            saldo_count += 1
            if saldo_count == 5:
                headers_to_find["SALDO"] = center_x
                print(f"DEBUG: Se detectó 'SALDO' en x={center_x}")
    return headers_to_find

def cargar_archivo():
    """
    Abre un cuadro de diálogo para seleccionar archivos PDF y guarda la lista.
    """
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
        # DEBUG: Mostrar archivos seleccionados
        # print("DEBUG: Archivos seleccionados:", pdf_paths)

def procesar_pdf():
    """
    Procesa cada PDF seleccionado: extrae movimientos, asigna campos basándose en la
    posición de los encabezados y genera un archivo Excel con dos hojas (MXN y USD).
    Se agregan prints de depuración en puntos clave para detectar errores.
    """
    global pdf_paths, output_folder
    if not pdf_paths:
        messagebox.showwarning("Advertencia", "No se ha seleccionado un archivo PDF.")
        return

    for pdf_path in pdf_paths:
        try:
            # print(f"DEBUG: Procesando PDF: {pdf_path}")
            pdf_name = os.path.basename(pdf_path)
            pdf_stem, pdf_ext = os.path.splitext(pdf_name)
            excel_name = pdf_stem + ".xlsx"

            with pdfplumber.open(pdf_path) as pdf:
                if len(pdf.pages) == 0:
                    messagebox.showinfo("Info", "El PDF está vacío.")
                    return

                # --- DETECTAR POSICIÓN DE LOS ENCABEZADOS (PRIMERA PÁGINA) ---
                page0 = pdf.pages[0]
                header_positions = detect_headers(page0)
                # print("DEBUG: Encabezados detectados:", header_positions)

                # --- PREPARAR ESTRUCTURAS PARA MOVIMIENTOS ---
                mxn_movimientos = []
                usd_movimientos = []
                current_moneda = "MXN"

                # Frases para omitir o detener lectura
                skip_phrases = ["ESTADO DE CUENTA AL", "PÁGINA"]
                skip_phrases = [s.upper() for s in skip_phrases]
                stop_phrases = ["SALDO MINIMO REQUERIDO"]
                stop_phrases = [s.upper() for s in stop_phrases]

                start_reading = False
                stop_reading = False
                movimiento_actual = None

                # Datos de encabezado extra (se llenan según se detecten)
                periodo_str = ""
                no_cuenta_str = ""
                empresa_str = ""
                no_cliente_str = ""
                rfc_str = ""

                # --- RECORRER TODAS LAS PÁGINAS ---
                for page_index, page in enumerate(pdf.pages):
                    # print(f"DEBUG: Procesando página {page_index + 1}")
                    if stop_reading:
                        # print("DEBUG: stop_reading activado, saliendo del bucle de páginas")
                        break

                    words = page.extract_words()
                    if not words:
                        # print("DEBUG: No se encontraron palabras en esta página")
                        continue

                    # Agrupar tokens por su coordenada 'top'
                    lineas_dict = {}
                    for w in words:
                        top_approx = int(w['top'])
                        if top_approx not in lineas_dict:
                            lineas_dict[top_approx] = []
                        lineas_dict[top_approx].append(w)
                    lineas_ordenadas = sorted(lineas_dict.items(), key=lambda x: x[0])

                    for top_val, words_in_line in lineas_ordenadas:
                        # DEBUG: Mostrar la línea (valor top y contenido)
                        line_debug = " ".join(w['text'] for w in words_in_line)
                        # print(f"DEBUG: Línea en top={top_val}: {line_debug}")

                        if stop_reading:
                            break

                        # --- Unir tokens, por ejemplo, "$" y el número siguiente ---
                        joined_tokens = []
                        i = 0
                        while i < len(words_in_line):
                            token = words_in_line[i]
                            txt = token['text'].strip()
                            if txt == "$" and (i + 1 < len(words_in_line)):
                                next_token = words_in_line[i+1]
                                combined_text = txt + next_token['text'].strip()
                                if es_numero_monetario(combined_text):
                                    merged = merge_tokens(token, next_token)
                                    joined_tokens.append(merged)
                                    # print(f"DEBUG: Tokens combinados: {token['text']} + {next_token['text']} -> {merged['text']}")
                                    i += 2
                                    continue
                                else:
                                    joined_tokens.append(token)
                                    i += 1
                            else:
                                joined_tokens.append(token)
                                i += 1

                        # Construir la línea de texto a partir de los tokens unidos
                        line_text = " ".join(t['text'].strip() for t in joined_tokens)
                        line_text_upper = line_text.upper()
                        # print(f"DEBUG: Línea procesada: {line_text}")

                        # --- Omitir líneas con frases que se deben saltar o detener ---
                        if any(sp in line_text_upper for sp in skip_phrases):
                            # print("DEBUG: Línea omitida por skip_phrases:", line_text)
                            continue
                        if any(sp in line_text_upper for sp in stop_phrases):
                            # print("DEBUG: Línea detectada para detener lectura:", line_text)
                            stop_reading = True
                            break

                        # Detectar cambio de moneda a USD
                        if "CUENTA DE CHEQUES EN DOLARES" in line_text_upper:
                            current_moneda = "USD"
                            # print("DEBUG: Cambio de moneda a USD detectado.")

                        # --- Capturar datos de encabezado extra ---
                        if "RESUMEN" in line_text_upper and "DEL:" in line_text_upper:
                            tokens_line = line_text.split()
                            fechas = [t for t in tokens_line if re.match(r'^\d{1,2}/\d{1,2}/\d{4}$', t)]
                            if len(fechas) == 2:
                                periodo_str = f"{fechas[0]} al {fechas[1]}"
                            else:
                                periodo_str = line_text
                            # print("DEBUG: Período detectado:", periodo_str)
                            continue
                        if "CONTRATO" in line_text_upper and not no_cuenta_str:
                            tokens_line = line_text.split()
                            no_cuenta_str = tokens_line[-1]
                            # print("DEBUG: No. Cuenta detectado:", no_cuenta_str)
                            continue
                        if "NUMERO DE CLIENTE:" in line_text_upper and not no_cliente_str:
                            tokens_line = line_text.split()
                            no_cliente_str = tokens_line[-1]
                            # print("DEBUG: No. Cliente detectado:", no_cliente_str)
                            continue
                        if "R.F.C." in line_text_upper and not rfc_str:
                            tokens_line = line_text.split()
                            rfc_str = tokens_line[-1]
                            # print("DEBUG: RFC detectado:", rfc_str)
                            continue

                        # --- Iniciar lectura de movimientos ---
                        if not start_reading:
                            tokens_line = line_text.split()
                            found_day = any(re.match(r'^\d{1,2}$', t) for t in tokens_line)
                            found_month = any(t in MESES_CORTOS for t in tokens_line)
                            if found_day and found_month:
                                start_reading = True
                                # print("DEBUG: Inicio de lectura de movimientos detectado.")
                            else:
                                continue

                        # --- Detectar nueva línea de movimiento (por fecha) ---
                        if es_linea_movimiento(line_text_upper):
                            if movimiento_actual:
                                if current_moneda == "USD":
                                    usd_movimientos.append(movimiento_actual)
                                else:
                                    mxn_movimientos.append(movimiento_actual)
                                # print("DEBUG: Movimiento anterior guardado:", movimiento_actual)
                            tokens_line = line_text_upper.split()
                            movimiento_actual = {
                                "FECHA": f"{tokens_line[0]} {tokens_line[1]}",
                                "NO. REF.": "",
                                "DESCRIPCION DE LA OPERACION": "",
                                "DEPOSITOS": None,
                                "RETIROS": None,
                                "SALDO": None
                            }
                            # print("DEBUG: Nuevo movimiento iniciado con fecha:", movimiento_actual["FECHA"])
                        else:
                            if not movimiento_actual:
                                movimiento_actual = {
                                    "FECHA": None,
                                    "NO. REF.": "",
                                    "DESCRIPCION DE LA OPERACION": "",
                                    "DEPOSITOS": None,
                                    "RETIROS": None,
                                    "SALDO": None
                                }

                        # --- Procesar cada token de la línea para asignar campos ---
                        for tk in joined_tokens:
                            txt_joined = tk['text'].strip()
                            center_w = (tk['x0'] + tk['x1']) / 2

                            if es_numero_monetario(txt_joined):
                                # Asignar montos según la cercanía a los encabezados de DEPOSITOS, RETIROS o SALDO
                                dist_depositos = dist(center_w, header_positions["DEPOSITOS"]) if header_positions["DEPOSITOS"] is not None else float('inf')
                                dist_retiros   = dist(center_w, header_positions["RETIROS"])   if header_positions["RETIROS"] is not None else float('inf')
                                dist_saldo     = dist(center_w, header_positions["SALDO"])     if header_positions["SALDO"] is not None else float('inf')
                                min_distance = min(dist_depositos, dist_retiros, dist_saldo)
                                if min_distance == dist_depositos:
                                    movimiento_actual["DEPOSITOS"] = txt_joined
                                    # print(f"DEBUG: Asignado {txt_joined} a DEPOSITOS")
                                elif min_distance == dist_retiros:
                                    movimiento_actual["RETIROS"] = txt_joined
                                    # print(f"DEBUG: Asignado {txt_joined} a RETIROS")
                                elif min_distance == dist_saldo:
                                    movimiento_actual["SALDO"] = txt_joined
                                    # print(f"DEBUG: Asignado {txt_joined} a SALDO")
                            elif re.match(r'^\d+$', txt_joined) and not movimiento_actual["NO. REF."]:
                                movimiento_actual["NO. REF."] = txt_joined
                                # print(f"DEBUG: Asignado {txt_joined} a NO. REF.")

                                # # Evitar tokens que formen parte de la fecha
                                # if re.match(r'^\d{1,2}$', txt_joined) or txt_joined.upper() in MESES_CORTOS:
                                #     continue
                                # # Asignar texto según la cercanía a "NO. REF." o "DESCRIPCION DE LA OPERACION"
                                # dist_no_ref = dist(center_w, header_positions["NO. REF."]) if header_positions["NO. REF."] is not None else float('inf')
                                # dist_desc   = dist(center_w, header_positions["DESCRIPCION DE LA OPERACION"]) if header_positions["DESCRIPCION DE LA OPERACION"] is not None else float('inf')
                                # if dist_no_ref < dist_desc:
                                #     movimiento_actual["NO. REF."] = (movimiento_actual["NO. REF."].strip() + " " + txt_joined).strip()
                                #     # print(f"DEBUG: Asignado {txt_joined} a NO. REF.")
    


                            else:
                                movimiento_actual["DESCRIPCION DE LA OPERACION"] = (movimiento_actual["DESCRIPCION DE LA OPERACION"].strip() + " " + txt_joined).strip()
                                # print(f"DEBUG: Asignado {txt_joined} a DESCRIPCION DE LA OPERACION")

                if movimiento_actual:
                    if current_moneda == "USD":
                        usd_movimientos.append(movimiento_actual)
                    else:
                        mxn_movimientos.append(movimiento_actual)
                    # print("DEBUG: Último movimiento agregado:", movimiento_actual)

            # --- CREAR EXCEL CON DOS HOJAS (MXN y USD) ---
            df_mxn = pd.DataFrame(mxn_movimientos, columns=[
                "FECHA", "NO. REF.", "DESCRIPCION DE LA OPERACION", "DEPOSITOS", "RETIROS", "SALDO"
            ])
            df_usd = pd.DataFrame(usd_movimientos, columns=[
                "FECHA", "NO. REF.", "DESCRIPCION DE LA OPERACION", "DEPOSITOS", "RETIROS", "SALDO"
            ])

            ruta_salida = os.path.join(output_folder, excel_name)
            # print("DEBUG: Guardando Excel en:", ruta_salida)

            with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
                df_mxn.to_excel(writer, sheet_name="Movs_MXN", index=False)
                df_usd.to_excel(writer, sheet_name="Movs_USD", index=False)

                wb = writer.book
                ws_mxn = wb["Movs_MXN"]
                ws_usd = wb["Movs_USD"]

                # Encabezado extra en la hoja MXN
                ws_mxn.insert_rows(1, 6)
                ws_mxn["A1"] = f"Banco: BanBajío"
                ws_mxn["A2"] = f"Empresa: {empresa_str}"
                ws_mxn["A3"] = f"No. Cuenta: {no_cuenta_str}"
                ws_mxn["A4"] = f"No. Cliente: {no_cliente_str}"
                ws_mxn["A5"] = f"Periodo: {periodo_str}"
                ws_mxn["A6"] = f"RFC: {rfc_str}"

                # Encabezado extra en la hoja USD
                ws_usd.insert_rows(1, 6)
                ws_usd["A1"] = f"Banco: BanBajío"
                ws_usd["A2"] = f"Empresa: {empresa_str}"
                ws_usd["A3"] = f"No. Cuenta: {no_cuenta_str}"
                ws_usd["A4"] = f"No. Cliente: {no_cliente_str}"
                ws_usd["A5"] = f"Periodo: {periodo_str}"
                ws_usd["A6"] = f"RFC: {rfc_str}"

                # Definir formato de celdas para encabezados y datos
                thin_side = Side(border_style="thin")
                thin_border = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side)
                header_fill = PatternFill(start_color="000080", end_color="000080", fill_type="solid")
                white_font = Font(color="FFFFFF", bold=True)

                # --- Formatear hoja Movs_MXN ---
                max_row_mxn = ws_mxn.max_row
                max_col_mxn = ws_mxn.max_column
                for col in range(1, max_col_mxn + 1):
                    cell = ws_mxn.cell(row=7, column=col)
                    cell.fill = header_fill
                    cell.font = white_font
                    cell.alignment = Alignment(horizontal="center")
                    cell.border = thin_border
                for row in range(8, max_row_mxn + 1):
                    for col in range(1, max_col_mxn + 1):
                        cell = ws_mxn.cell(row=row, column=col)
                        cell.border = thin_border
                for col in ws_mxn.columns:
                    max_length = 0
                    col_letter = col[0].column_letter
                    for cell in col:
                        if cell.value is not None:
                            length = len(str(cell.value))
                            if length > max_length:
                                max_length = length
                    ws_mxn.column_dimensions[col_letter].width = max_length + 2
                for row in ws_mxn.iter_rows():
                    for cell in row:
                        cell.alignment = Alignment(wrap_text=True)

                # --- Formatear hoja Movs_USD ---
                max_row_usd = ws_usd.max_row
                max_col_usd = ws_usd.max_column
                for col in range(1, max_col_usd + 1):
                    cell = ws_usd.cell(row=7, column=col)
                    cell.fill = header_fill
                    cell.font = white_font
                    cell.alignment = Alignment(horizontal="center")
                    cell.border = thin_border
                for row in range(8, max_row_usd + 1):
                    for col in range(1, max_col_usd + 1):
                        cell = ws_usd.cell(row=row, column=col)
                        cell.border = thin_border
                for col in ws_usd.columns:
                    max_length = 0
                    col_letter = col[0].column_letter
                    for cell in col:
                        if cell.value is not None:
                            length = len(str(cell.value))
                            if length > max_length:
                                max_length = length
                    ws_usd.column_dimensions[col_letter].width = max_length + 2
                for row in ws_usd.iter_rows():
                    for cell in row:
                        cell.alignment = Alignment(wrap_text=True)

                wb.save(ruta_salida)
                # print("DEBUG: Excel guardado correctamente.")

            messagebox.showinfo("Éxito", f"Archivo Excel generado con 2 hojas: {ruta_salida}")

        except Exception as e:
            # print("DEBUG: Error en procesar_pdf:", e)
            messagebox.showerror("Error", f"Ocurrió un error al procesar el PDF:\n{e}")

def main():
    global pdf_paths, output_folder
    if len(sys.argv) < 2:
        # print("Uso: python procesar_pdf_banamex.py <carpeta_salida>")
        return
    output_folder = sys.argv[1]
    pdf_paths = ""

    # Configuración de la ventana principal de Tkinter
    root = tk.Tk()
    root.title("Extracción Movimientos - banamex")
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
