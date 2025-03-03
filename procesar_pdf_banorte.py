import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import re
import string
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

# Conjunto de meses cortos
MESES_CORTOS = {"ENE", "FEB", "MAR", "ABR", "MAY", "JUN",
                "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"}

def ajusta_fechas_en_linea(line_text):
    """
    Inserta un espacio si se detecta una fecha en formato dd-MES-yy
    inmediatamente seguida de letras o dígitos, p.ej. '06-ENE-25ABC' => '06-ENE-25 ABC'.

    Solo hace una pasada, evitando bucles infinitos.
    """
    pattern = r'(\d{1,2}-[A-Z]{3}-\d{2})([A-Za-z0-9])'
    # Usamos re.subn para ver cuántas sustituciones se hacen (útil para debug)
    new_text, num_subs = re.subn(pattern, r'\1 \2', line_text)
    # print(f"[DEBUG] Se hicieron {num_subs} sustituciones en esta línea.")
    return new_text

def es_linea_movimiento(linea):
    """
    Determina si la línea inicia un 'movimiento' nuevo
    con el primer token en formato dd-MES-yy (p.ej. '03-ENE-25').
    """
    tokens = linea.split()
    if not tokens:
        return False
    
    primer_token = tokens[0]
    partes = primer_token.split("-")
    if len(partes) != 3:
        return False
    
    dia, mes, anio = partes
    if not re.match(r'^\d{1,2}$', dia):
        return False
    if mes not in MESES_CORTOS:
        return False
    if not re.match(r'^\d{2}$', anio):
        return False

    return True

def es_numero_monetario(texto):
    """
    Determina si un texto es un número tipo '100,923.30'.
    Ajusta si tu PDF usa otro formato.
    """
    pattern = r'^\d{1,3}(,\d{3})*\.\d{2}$'
    return bool(re.match(pattern, texto.strip()))

def dist(a, b):
    """Distancia absoluta entre dos valores."""
    return abs(a - b)

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


def agrupar_por_top_con_tolerancia(words, tolerancia=2):
    """
    Recibe una lista de 'words' extraídas por pdfplumber y 
    las agrupa en diccionarios de la forma {top_agrupado: [words_en_esa_linea]}.
    La clave 'top_agrupado' es un float o int representativo de la línea.
    
    'tolerancia' indica cuán cerca (en unidades de 'top') deben estar las palabras
    para considerarlas parte de la misma línea.
    """
    lineas_dict = {}

    for w in words:
        actual_top = w['top']
        # Buscamos si hay algún top ya registrado que esté dentro de la tolerancia
        top_encontrado = None
        for top_existente in lineas_dict.keys():
            if abs(top_existente - actual_top) <= tolerancia:
                top_encontrado = top_existente
                break

        if top_encontrado is not None:
            # Agregamos la palabra a la línea existente
            lineas_dict[top_encontrado].append(w)
        else:
            # Creamos una nueva línea
            lineas_dict[actual_top] = [w]

    # Retornamos las líneas ordenadas por el valor de top
    lineas_ordenadas = sorted(lineas_dict.items(), key=lambda x: x[0])
    return lineas_ordenadas

def procesar_pdf():
    global pdf_paths
    if not pdf_paths:
        messagebox.showwarning("Advertencia", "No se ha seleccionado un archivo PDF.")
        return
    for pdf_path in pdf_paths:

        try:
            with pdfplumber.open(pdf_path) as pdf:

                pdf_name = os.path.basename(pdf_path)
                pdf_stem, pdf_ext = os.path.splitext(pdf_name)
                excel_name = pdf_stem + ".xlsx"

                if len(pdf.pages) == 0:
                    messagebox.showinfo("Info", "El PDF está vacío.")
                    return

                # 1) DETECTAR ENCABEZADOS EN LA 1RA PÁGINA
                page0 = pdf.pages[0]
                words_page0 = page0.extract_words()

                encabezados_buscar = ["DEPOSITO", "RETIRO", "SALDO"]
                col_positions = {}

                # Agrupamos las palabras de la primera página por 'top' para formar líneas
                lineas_ordenadas_page0 = agrupar_por_top_con_tolerancia(words_page0, tolerancia=2)

                # Buscamos la línea que contenga los 3 encabezados
                for top_val, words_in_line in lineas_ordenadas_page0:
                    line_text_upper = " ".join(w['text'].strip().upper() for w in words_in_line)
                    if all(h in line_text_upper for h in encabezados_buscar):
                        # Extraemos la coordenada de cada encabezado
                        for w in words_in_line:
                            w_text_upper = w['text'].strip().upper()
                            if w_text_upper in encabezados_buscar:
                                center_x = (w['x0'] + w['x1']) / 2
                                col_positions[w_text_upper] = center_x
                        break

                # Ordenamos por la coordenada X
                columnas_ordenadas = sorted(col_positions.items(), key=lambda x: x[1])

                # 2) VARIABLES PARA ENCABEZADOS DEL EXCEL
                periodo_str = ""
                empresa_str = ""
                no_cliente_str = ""
                rfc_str = ""

                # 3) FRASES A OMITIR (skip) Y A DETENER (stop)
                skip_phrases = [ 
                    "ESTADO DE CUENTA",
                    "FECHA DE CORTE",
                    "Línea Directa para su empresa:",
                    "CIUDAD DE MÉXICO: (55)",
                    "(81)8156 9640",
                    "CIUDAD DE MÉXICO:",
                    "Guadalajara",
                    "Monterrey",
                    "BANCO MERCANTIL DEL NORTE S.A. INSTITUCIÓN DE BANCA MÚLTIPLE GRUPO FINANCIERO BANORTE.",
                    "Nuevo Leon. RFC BMN930209927",
                    "Resto del país:",
                    "DETALLE DE MOVIMIENTOS",
                    "Enlace Negocios Basica",
                    "Visita nuestra página:",
                    "FECHA",
                    "DESCRIPCIÓN / ESTABLECIMIENTO",
                    "MONTO DEL DEPOSITO",
                    "MONTO DEL RETIRO",
                    "SALDO",
                    "Banco Mercantil del Norte",
                ]
                # stop_phrases: solo se detiene si aparece en páginas >= 2
                stop_phrases = ["OTROS"]

                start_reading = False
                stop_reading = False
                todos_los_movimientos = []
                movimiento_actual = None

                # 4) RECORRER TODAS LAS PÁGINAS
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

                        # Construir la línea
                        line_text = " ".join(w['text'] for w in words_in_line)

                        # (¡NUEVO!) Ajusta fechas pegadas en la línea, evitando bucles
                        line_text = ajusta_fechas_en_linea(line_text)

                        # Mantén tu regex que inserta espacio si ve letras pegadas a la fecha
                        # (ahora abarca parte del caso, pero la mantenemos por compatibilidad)
                        line_text = re.sub(
                            r'(\d{1,2}-[A-Z]{3}-\d{2})([A-Za-z])',
                            r'\1 \2',
                            line_text
                        )

                        # Convertir a mayúsculas para comparaciones
                        line_text_upper = line_text.upper()

                        # Detectar periodo
                        if "PERIODO" in line_text_upper and not periodo_str:
                            tokens_line = line_text.split()
                            fechas = [t for t in tokens_line if re.match(r'^\d{1,2}/\d{1,2}/\d{4}$', t)]
                            if len(fechas) == 2:
                                periodo_str = f"{fechas[0]} al {fechas[1]}"
                            else:
                                periodo_str = line_text
                            continue

                        # Detectar No. de Cliente
                        if "NO. DE CLIENTE:" in line_text_upper and not no_cliente_str:
                            tokens_line = line_text.split()
                            no_cliente_str = tokens_line[-1]
                            continue

                        # Detectar RFC
                        if "RFC:" in line_text_upper and not rfc_str:
                            tokens_line = line_text.split()
                            rfc_str = tokens_line[-1]
                            continue

                        # Stop phrases (solo en páginas >= 2)
                        if page_index >= 1 and any(sp in line_text_upper for sp in stop_phrases):
                            stop_reading = True
                            break

                        # Empezar a leer movimientos
                        if not start_reading:
                            if es_linea_movimiento(line_text_upper):
                                start_reading = True
                            else:
                                continue

                        # Omitir la línea si contiene skip_phrases
                        if any(sp in line_text for sp in skip_phrases):
                            continue    

                        # Omitir si contiene algo como 2/17
                        if re.search(r'\b\d+/\d+\b', line_text_upper):
                            continue

                        # ¿Es nuevo movimiento?
                        if es_linea_movimiento(line_text_upper):
                            # Guardar el anterior, si existe
                            if movimiento_actual:
                                todos_los_movimientos.append(movimiento_actual)

                            tokens_line = line_text_upper.split()
                            movimiento_actual = {
                                "Fecha": tokens_line[0],
                                "Descripción / Establecimiento": "",
                                "Monto del deposito": None,
                                "Monto del retiro": None,
                                "Saldo": None
                            }
                        else:
                            # Continuación de un movimiento
                            if not movimiento_actual:
                                movimiento_actual = {
                                    "Fecha": None,
                                    "Descripción / Establecimiento": "",
                                    "Monto del deposito": None,
                                    "Monto del retiro": None,
                                    "Saldo": None
                                }

                        # Procesar cada token de la línea
                        for w in words_in_line:
                            token_upper = w['text'].upper()
                            # Si el token contiene alguna de las skip phrases, detener
                            if any(sp in token_upper for sp in skip_phrases):
                                break

                            txt = w['text'].strip()
                            center_w = (w['x0'] + w['x1']) / 2

                            if es_numero_monetario(txt):
                                if columnas_ordenadas:
                                    col_name, col_center = min(
                                        columnas_ordenadas,
                                        key=lambda x: dist(x[1], center_w)
                                    )
                                    # Observa que col_name es la clave exacta, p.ej. "MONTO DEL DEPOSITO"
                                    print(f"[DEBUG] -> Token '{txt}' en columna '{col_name}'")
                                    if col_name in "MONTO DEL DEPOSITO":
                                        movimiento_actual["Monto del deposito"] = txt
                                    elif col_name in "MONTO DEL RETIRO":
                                        movimiento_actual["Monto del retiro"] = txt
                                    elif col_name in "SALDO":
                                        movimiento_actual["Saldo"] = txt
                                else:
                                    movimiento_actual["Monto del retiro"] = txt
                            else:
                                # Si el token inicia con un formato fecha, separamos la parte de fecha y el resto
                                m = re.match(r'^(\d{1,2}-[A-Z]{3}-\d{2})(.*)$', txt)
                                if m:
                                    date_part = m.group(1)
                                    rest = m.group(2)
                                    if movimiento_actual["Fecha"] and date_part.upper() == movimiento_actual["Fecha"]:
                                        txt = rest.strip()
                                clean_txt = txt.strip(string.punctuation)
                                # Omitir tokens que sean solo dígitos, meses o fecha completa
                                if re.match(r'^\d{1,2}$', clean_txt):
                                    continue
                                if clean_txt in MESES_CORTOS:
                                    continue
                                if re.match(r'^\d{1,2}-[A-Z]{3}-\d{2}$', clean_txt):
                                    continue
                                movimiento_actual["Descripción / Establecimiento"] += txt + " "

                # Al terminar, guardar el último movimiento (con limpieza de skip phrases)
                if movimiento_actual:
                    for sp in skip_phrases:
                        movimiento_actual["Descripción / Establecimiento"] = re.sub(
                            re.escape(sp), "", movimiento_actual["Descripción / Establecimiento"], flags=re.IGNORECASE
                        )
                    todos_los_movimientos.append(movimiento_actual)

            # 5) GUARDAR EN EXCEL
            df = pd.DataFrame(todos_los_movimientos, columns=[
                "Fecha",
                "Descripción / Establecimiento",
                "Monto del deposito",
                "Monto del retiro",
                "Saldo"
            ])

            ruta_salida = os.path.join(output_folder, excel_name)
            df.to_excel(ruta_salida, index=False)

            # Ajustes de estilo con openpyxl
            wb = load_workbook(ruta_salida)
            ws = wb.active

            # Insertar filas para encabezado
            ws.insert_rows(1, 6)

            # Encabezado
            ws["A1"] = f"Banco: Banorte"
            ws["A2"] = f"Empresa: {empresa_str}"
            ws["A3"] = f"No. Cliente: {no_cliente_str}"
            ws["A4"] = f"Periodo: {periodo_str}"
            ws["A5"] = f"RFC: {rfc_str}"

            thin_side = Side(border_style="thin")
            thin_border = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side)

            header_fill = PatternFill(start_color="000080", end_color="000080", fill_type="solid")
            white_font = Font(color="FFFFFF", bold=True)

            max_row = ws.max_row
            max_col = ws.max_column

            # Estilo para la fila de encabezados (fila 7)
            for col in range(1, max_col + 1):
                cell = ws.cell(row=7, column=col)
                cell.fill = header_fill
                cell.font = white_font
                cell.alignment = Alignment(horizontal="center")
                cell.border = thin_border

            # Estilo para filas de datos
            for row in range(8, max_row + 1):
                for col in range(1, max_col + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.border = thin_border

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

            # Alineación wrap_text
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True)

            wb.save(ruta_salida)
            messagebox.showinfo("Éxito", f"Archivo Excel generado: {ruta_salida}")

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al procesar el PDF:\n{e}")

def main():
    # Interfaz gráfica con tkinter
    global pdf_paths, output_folder
    if len(sys.argv) < 2:
        print("Uso: python procesar_pdf_banorte.py <carpeta_salida>")
        return
    output_folder = sys.argv[1]

    pdf_paths = ""

    root = tk.Tk()
    root.title("Extracción Movimientos - Banorte")
    # root.geometry("600x250")

    win_width = 600
    win_height = 250

    # Obtenemos dimensiones de la pantalla
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

        # Calculamos coordenadas x e y
    x = (screen_width - win_width) // 2
    y = (screen_height - win_height) // 2
    # Ajustamos la geometría: ancho x alto + x + y
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