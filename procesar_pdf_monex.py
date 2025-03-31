# import sys
# import os
# import tkinter as tk
# from tkinter import filedialog, messagebox
# import pdfplumber
# import pandas as pd
# import re
# from openpyxl import load_workbook
# from openpyxl.styles import Alignment, Font, PatternFill, Border, Side


# # Texto clave en mayúsculas (cuando extraemos la línea, solemos convertirlo a upper())
# SALDO_INICIAL_TXT = "SALDO INICIAL:"
# SALDO_FINAL_TXT   = "SALDO FINAL:"

# ##############################################################################
# # 1) Función para DETECTAR "SALDO" + "INICIAL" con x0 y x1 dentro de ciertos rangos
# ##############################################################################
# def within_range(value, low, high):
#     """Retorna True si 'value' está entre 'low' y 'high'."""
#     return low <= value <= high

# def detectar_saldo_inicial_en_rango(words_in_line,
#                                     x0_saldo_low=490,
#                                     x0_saldo_high=492,
#                                     x1_inicial_low=532,
#                                     x1_inicial_high=534):
#     """
#     Revisa si en 'words_in_line' aparece la palabra 'SALDO' con x0 entre [x0_saldo_low, x0_saldo_high]
#     seguida inmediatamente de 'INICIAL' con x1 entre [x1_inicial_low, x1_inicial_high].
#     Retorna True si lo encuentra, False en caso contrario.
#     """
#     for i, w in enumerate(words_in_line):
#         texto_actual = w['text'].strip().upper()
#         if texto_actual == "SALDO":
#             # Revisamos las coordenadas x0
#             if within_range(w['x0'], x0_saldo_low, x0_saldo_high):
#                 # Vemos si hay siguiente palabra
#                 if i + 1 < len(words_in_line):
#                     w_next = words_in_line[i + 1]
#                     texto_siguiente = w_next['text'].strip().upper()
#                     if texto_siguiente.startswith("INICIAL"):
#                         # Revisamos las coordenadas x1
#                         if within_range(w_next['x1'], x1_inicial_low, x1_inicial_high):
#                             return True
#     return False

# ##############################################################################
# # 2) Funciones auxiliares para agrupar líneas y para detectar montos, etc.
# ##############################################################################
# def agrupar_por_top_con_tolerancia(words, tolerancia=2):
#     """
#     Agrupa 'words' (palabras extraídas por pdfplumber) que tienen valores 'top' muy parecidos.
#     Esto nos ayuda a considerar un mismo 'renglón' aun cuando las coordenadas no coincidan perfectamente.
#     La 'tolerancia' define cuán cerca en Y deben estar para considerarse en la misma línea.
#     """
#     lineas_dict = {}
#     for w in words:
#         top_val = w['top']
#         top_encontrado = None
#         for tv in lineas_dict.keys():
#             if abs(tv - top_val) <= tolerancia:
#                 top_encontrado = tv
#                 break
#         if top_encontrado is not None:
#             lineas_dict[top_encontrado].append(w)
#         else:
#             lineas_dict[top_val] = [w]

#     lineas_ordenadas = sorted(lineas_dict.items(), key=lambda x: x[0])
#     return lineas_ordenadas

# def es_numero_monetario(txt):
#     patron = r'^(-?\d{1,3}(,\d{3})*\.\d{2}|\(\d{1,3}(,\d{3})*\.\d{2}\))$'
#     return bool(re.match(patron, txt.strip()))

# def parse_monetario(txt):
#     txt = txt.strip()
#     sign = 1
#     if txt.startswith("(") and txt.endswith(")"):
#         sign = -1
#         txt = txt[1:-1].strip()
#     elif txt.startswith("-"):
#         sign = -1
#         txt = txt[1:].strip()
#     txt = txt.replace(",", "")
#     return sign * float(txt)

# def limpiar_descripcion(texto):
#     """
#     Elimina de 'texto' los tokens que sean montos monetarios o fechas dd/mmm,
#     para dejar sólo la parte descriptiva.
#     """
#     tokens = texto.split()
#     tokens_filtrados = [
#         tok for tok in tokens
#         if not es_numero_monetario(tok)
#         and not re.match(r'^\d{1,2}/[A-Za-z]{3}$', tok)
#     ]
#     return " ".join(tokens_filtrados)

# def unir_linea_superior_si_cerca(lineas_ordenadas, i, umbral_distancia=11):
#     """
#     Revisa si la línea i-1 está cerca en 'top'. Si sí, concatena sus textos. Si no, regresa la línea actual.
#     """
#     if i <= 0:
#         top_val, words_in_line = lineas_ordenadas[i]
#         return " ".join(w['text'] for w in words_in_line)
#     top_val_actual, words_actual = lineas_ordenadas[i]
#     top_val_prev, words_prev = lineas_ordenadas[i - 1]
#     if abs(top_val_actual - top_val_prev) <= umbral_distancia:
#         texto_prev = " ".join(w['text'] for w in words_prev)
#         texto_actual = " ".join(w['text'] for w in words_actual)
#         return texto_prev + " " + texto_actual
#     else:
#         return " ".join(w['text'] for w in words_actual)

# ##############################################################################
# # 3) Función para MOSTRAR TODAS las apariciones de SALDO+INICIAL en esas coords
# ##############################################################################
# def mostrar_todos_saldos_iniciales_en_rango(pdf_path):
#     """
#     Abre el PDF y para cada página y línea, detecta si existe la secuencia
#     SALDO x0 ~ [490..492], INICIAL x1 ~ [532..534].
#     Imprime en consola cada vez que lo encuentre.
#     """
#     with pdfplumber.open(pdf_path) as pdf:

#         if len(pdf.pages) == 0:
#             messagebox.showinfo("Info", "El PDF está vacío.")
#             return
        
#         page0 = pdf.pages[3]
#         words_page0 = page0.extract_words()

#         for w in words_page0:
#             print(f"Texto: {w['text']}, x0: {w['x0']}, x1: {w['x1']}, top: {w['top']}, bottom: {w['bottom']}")

#         for page_index, page in enumerate(pdf.pages):
#             words = page.extract_words()
#             lineas_ordenadas = agrupar_por_top_con_tolerancia(words, tolerancia=1)



#             for line_index, (top_val, words_in_line) in enumerate(lineas_ordenadas):
#                 if detectar_saldo_inicial_en_rango(words_in_line):
#                     print(f"Encontrado 'SALDO INICIAL' en la página {page_index + 1}, línea {line_index}.")

# ##############################################################################
# # 4) Función que extrae la información (tu lógica previa de parseo de movimientos)
# ##############################################################################
# def extraer_movimientos_por_linea(pdf_path):
#     """
#     Tu lógica principal para armar los movimientos e identificarlos en PESOS, DOLAR, EURO.
#     """
#     SALDO_INICIAL_TXT = "SALDO INICIAL:"
#     SALDO_FINAL_TXT   = "SALDO FINAL:"
#     secuenciaDolares = ["CUENTA VISTA", "RESUMEN CUENTA", "DÓLAR AMERICANO", "MOVIMIENTOS"]
#     secuenciaEuros   = ["CUENTA VISTA", "RESUMEN CUENTA", "EURO", "MOVIMIENTOS"]
#     idx_dolar = 0
#     idx_euro  = 0
#     isPesosActive = True
#     isDolarActive = False
#     isEuroActive  = False
#     leyendo_tabla = True
#     movimiento_actual = None

#     columnas_ordenadas = [
#         ("Descripción", 130),
#         ("Referencia",  260),
#         ("Abonos",      320),
#         ("Cargos",      400),
#         ("Movimiento garantia", 480),
#         ("Saldo en garantia",   550),
#         ("Saldo disponible",    620),
#         ("Saldo total",         680)
#     ]
#     ultimas_3_columnas = [
#         ("Saldo en garantia", 550),
#         ("Saldo disponible", 620),
#         ("Saldo total", 680)
#     ]
#     movements_by_currency = {
#         "PESOS": [],
#         "DOLAR": [],
#         "EURO": []
#     }
#     current_currency = "PESOS"

#     def es_entero_puro(txt):
#         return bool(re.match(r'^\d+$', txt.strip()))

#     with pdfplumber.open(pdf_path) as pdf:

#         if len(pdf.pages) == 0:
#             messagebox.showinfo("Info", "El PDF está vacío.")
#             return movements_by_currency

#         # Si deseas leer la página 4 (índice 3), verifica si existe:
#         # (comentar o ajustar según tu PDF)
#         if len(pdf.pages) > 3:
#             page3 = pdf.pages[3]
#             words_page3 = page3.extract_words()
#             # (debug) para imprimir coords si gustas:
#             # for w in words_page3:
#             #     print(f"Texto: {w['text']}, x0: {w['x0']}, x1: {w['x1']}")

#         for page_index, page in enumerate(pdf.pages):
#             words = page.extract_words()
#             lineas_ordenadas = agrupar_por_top_con_tolerancia(words, tolerancia=1)

#             for i, (top_val, words_in_line) in enumerate(lineas_ordenadas):

#                 # Unimos la línea actual con la superior si es muy cercana
#                 line_text = unir_linea_superior_si_cerca(lineas_ordenadas, i, umbral_distancia=11)
#                 line_text_upper = line_text.upper()

#                 # 1) Verificar si se activa la lectura de DOLAR/EURO
#                 if not isDolarActive and not isEuroActive:
#                     palabra_esperada = secuenciaDolares[idx_dolar]
#                     if palabra_esperada in line_text_upper:
#                         idx_dolar += 1
#                         if idx_dolar == len(secuenciaDolares):
#                             if movimiento_actual:
#                                 movements_by_currency[current_currency].append(movimiento_actual)
#                                 movimiento_actual = None
#                             isPesosActive = False
#                             isDolarActive = True
#                             current_currency = "DOLAR"
#                             leyendo_tabla = True
#                 else:
#                     if isDolarActive and not isEuroActive:
#                         palabra_euro_esperada = secuenciaEuros[idx_euro]
#                         if palabra_euro_esperada in line_text_upper:
#                             idx_euro += 1
#                             if idx_euro == len(secuenciaEuros):
#                                 if movimiento_actual:
#                                     movements_by_currency[current_currency].append(movimiento_actual)
#                                     movimiento_actual = None
#                                 isDolarActive = False
#                                 isEuroActive  = True
#                                 current_currency = "EURO"
#                                 leyendo_tabla = True

#                 if not leyendo_tabla:
#                     continue

#                 # 2.1) SALDO INICIAL
#                 if SALDO_INICIAL_TXT in line_text_upper:
#                     saldo_inicial_mov = {
#                         "Fecha": None,
#                         "Descripción": SALDO_INICIAL_TXT,
#                         "Referencia": "",
#                         "Abonos": None,
#                         "Cargos": None,
#                         "Movimiento garantia": None,
#                         "Saldo en garantia": None,
#                         "Saldo disponible": None,
#                         "Saldo total": None
#                     }
#                     numeric_count = 0
#                     for w in words_in_line:
#                         txt = w['text'].strip()
#                         if es_numero_monetario(txt):
#                             center_x = (w['x0'] + w['x1']) / 2
#                             col_name, _ = min(ultimas_3_columnas, key=lambda x: abs(x[1] - center_x))
#                             val = parse_monetario(txt)
#                             saldo_inicial_mov[col_name] = val
#                             numeric_count += 1
#                             if numeric_count == 3:
#                                 break
#                     movements_by_currency[current_currency].append(saldo_inicial_mov)
#                     continue

#                 # 2.2) SALDO FINAL
#                 if SALDO_FINAL_TXT in line_text_upper:
#                     saldo_final_mov = {
#                         "Fecha": None,
#                         "Descripción": SALDO_FINAL_TXT,
#                         "Referencia": "",
#                         "Abonos": None,
#                         "Cargos": None,
#                         "Movimiento garantia": None,
#                         "Saldo en garantia": None,
#                         "Saldo disponible": None,
#                         "Saldo total": None
#                     }
#                     numeric_count = 0
#                     for w in words_in_line:
#                         txt = w['text'].strip()
#                         if es_numero_monetario(txt):
#                             center_x = (w['x0'] + w['x1']) / 2
#                             col_name, _ = min(ultimas_3_columnas, key=lambda x: abs(x[1] - center_x))
#                             val = parse_monetario(txt)
#                             saldo_final_mov[col_name] = val
#                             numeric_count += 1
#                             if numeric_count == 3:
#                                 break
#                     movements_by_currency[current_currency].append(saldo_final_mov)
#                     if movimiento_actual:
#                         movements_by_currency[current_currency].append(movimiento_actual)
#                         movimiento_actual = None

#                     leyendo_tabla = False
#                     continue

#                 # 2.3) Buscar fecha
#                 tokens_line = line_text.split()
#                 fecha_token = None
#                 for tok in tokens_line:
#                     if re.match(r'^\d{1,2}/[A-Za-z]{3}$', tok):
#                         fecha_token = tok
#                         break

#                 if fecha_token:
#                     if movimiento_actual:
#                         movements_by_currency[current_currency].append(movimiento_actual)
#                     movimiento_actual = {
#                         "Fecha": fecha_token,
#                         "Descripción": limpiar_descripcion(line_text),
#                         "Referencia": "",
#                         "Abonos": None,
#                         "Cargos": None,
#                         "Movimiento garantia": None,
#                         "Saldo en garantia": None,
#                         "Saldo disponible": None,
#                         "Saldo total": None
#                     }
#                 else:
#                     if not movimiento_actual:
#                         movimiento_actual = {
#                             "Fecha": None,
#                             "Descripción": "",
#                             "Referencia": "",
#                             "Abonos": None,
#                             "Cargos": None,
#                             "Movimiento garantia": None,
#                             "Saldo en garantia": None,
#                             "Saldo disponible": None,
#                             "Saldo total": None
#                         }
#                     movimiento_actual["Descripción"] += " " + limpiar_descripcion(line_text)

#                 # 2.4) Extraer montos en la línea
#                 if movimiento_actual:
#                     for w in words_in_line:
#                         txt = w['text'].strip()
#                         center_x = (w['x0'] + w['x1']) / 2
#                         col_name, _ = min(columnas_ordenadas, key=lambda x: abs(x[1] - center_x))

#                         if es_numero_monetario(txt):
#                             if col_name in ("Abonos", "Cargos", "Movimiento garantia",
#                                             "Saldo en garantia", "Saldo disponible", "Saldo total"):
#                                 movimiento_actual[col_name] = parse_monetario(txt)

#                         elif es_entero_puro(txt):
#                             if col_name == "Referencia":
#                                 if movimiento_actual["Referencia"]:
#                                     movimiento_actual["Referencia"] += " " + txt
#                                 else:
#                                     movimiento_actual["Referencia"] = txt

#                         elif col_name == "Descripción":
#                             if (not es_numero_monetario(txt)
#                                 and not re.match(r'^\d{1,2}/[A-Za-z]{3}$', txt)):
#                                 movimiento_actual["Descripción"] += " " + txt

#         # Al final, si quedara uno en construcción...
#         if leyendo_tabla and movimiento_actual:
#             movements_by_currency[current_currency].append(movimiento_actual)

#     return movements_by_currency

# ##############################################################################
# # 5) Lógica de la interfaz Tkinter (seleccionar archivos, procesar, generar Excel)
# ##############################################################################
# def cargar_archivo():
#     global pdf_paths
#     archivos = filedialog.askopenfilenames(
#         title="Selecciona uno o más archivos PDF",
#         filetypes=[("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*")]
#     )
#     if archivos:
#         entry_archivo.config(state=tk.NORMAL)
#         entry_archivo.delete(0, tk.END)
#         files_text = " ; ".join(archivos)
#         entry_archivo.insert(0, files_text)
#         entry_archivo.config(state=tk.DISABLED)
#         pdf_paths = archivos
#     else:
#         pdf_paths = []

# def procesar_pdf():
#     global pdf_paths, output_folder

#     if not pdf_paths:
#         messagebox.showwarning("Advertencia", "No se ha seleccionado un archivo PDF.")
#         return

#     for pdf_path in pdf_paths:
#         try:
#             # 1) Primero mostramos todas las apariciones de SALDO INICIAL en los rangos x0, x1
#             mostrar_todos_saldos_iniciales_en_rango(pdf_path)

#             # 2) Extraemos movimientos
#             movements_by_currency = extraer_movimientos_por_linea(pdf_path)

#             # 3) Construimos nombre de archivo Excel
#             pdf_name = os.path.basename(pdf_path)
#             pdf_stem, _ = os.path.splitext(pdf_name)
#             excel_name = pdf_stem + ".xlsx"
#             ruta_salida = os.path.join(output_folder, excel_name)

#             # 4) Guardamos DataFrame en Excel (3 hojas: PESOS, DOLAR, EURO)
#             with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
#                 for currency in ["PESOS", "DOLAR", "EURO"]:
#                     df = pd.DataFrame(
#                         movements_by_currency[currency],
#                         columns=[
#                             "Fecha","Descripción","Referencia","Abonos","Cargos","Movimiento garantia",
#                             "Saldo en garantia","Saldo disponible","Saldo total"
#                         ]
#                     )
#                     df.to_excel(writer, sheet_name=currency, index=False)

#             # 5) Abrimos el Excel para aplicar estilos
#             wb = load_workbook(ruta_salida)
#             for sheet_name in ["PESOS","DOLAR","EURO"]:
#                 if sheet_name not in wb.sheetnames:
#                     continue
#                 ws = wb[sheet_name]

#                 # Insertamos 6 filas al inicio
#                 ws.insert_rows(1, 6)
#                 ws["A1"] = "Banco: Monex"
#                 ws["A2"] = "Empresa: (por definir)"
#                 ws["A3"] = "No. Cuenta: (por definir)"
#                 ws["A4"] = "No. Cliente: (por definir)"
#                 ws["A5"] = "Periodo: (por definir)"
#                 ws["A6"] = "RFC: (por definir)"

#                 thin_side = Side(border_style="thin")
#                 thin_border = Border(top=thin_side, left=thin_side,
#                                      right=thin_side, bottom=thin_side)

#                 header_fill = PatternFill(start_color="000080",
#                                           end_color="000080",
#                                           fill_type="solid")
#                 white_font = Font(color="FFFFFF", bold=True)

#                 max_row = ws.max_row
#                 max_col = ws.max_column

#                 # Pintamos la fila de cabecera (fila 7) con fondo azul, letra blanca
#                 for col in range(1, max_col + 1):
#                     cell = ws.cell(row=7, column=col)
#                     cell.fill = header_fill
#                     cell.font = white_font
#                     cell.alignment = Alignment(horizontal="center")
#                     cell.border = thin_border

#                 # Bordes para todas las filas siguientes
#                 for row in range(8, max_row + 1):
#                     for col in range(1, max_col + 1):
#                         cell = ws.cell(row=row, column=col)
#                         cell.border = thin_border

#                 # Ajustar ancho de columnas según su contenido
#                 for col in ws.columns:
#                     max_length = 0
#                     col_letter = col[0].column_letter
#                     for cell in col:
#                         if cell.value is not None:
#                             length = len(str(cell.value))
#                             if length > max_length:
#                                 max_length = length
#                     ws.column_dimensions[col_letter].width = max_length + 2

#                 # Wrap text en todas las celdas
#                 for row in ws.iter_rows():
#                     for cell in row:
#                         cell.alignment = Alignment(wrap_text=True)

#             wb.save(ruta_salida)
#             messagebox.showinfo("Éxito", f"Archivo Excel generado: {ruta_salida}")

#         except Exception as e:
#             messagebox.showerror("Error", f"Ocurrió un error:\n{e}")

# def main():
#     global pdf_paths, output_folder
#     if len(sys.argv) < 2:
#         return  # Si no hay carpeta de salida, no hacemos nada

#     output_folder = sys.argv[1]
#     pdf_paths = ""

#     root = tk.Tk()
#     root.title("Extracción Movimientos - Monex")
#     root.geometry("600x250")

#     root.update()
#     root.lift()
#     root.focus_force()
#     root.attributes("-topmost", True)
#     root.after(10, lambda: root.attributes("-topmost", False))

#     btn_cargar = tk.Button(root, text="Cargar PDF", command=cargar_archivo, width=30)
#     btn_cargar.pack(pady=10)

#     global entry_archivo
#     entry_archivo = tk.Entry(root, width=80, state=tk.DISABLED)
#     entry_archivo.pack(padx=10, pady=10)

#     btn_procesar = tk.Button(root, text="Procesar PDF", command=procesar_pdf, width=30)
#     btn_procesar.pack(pady=10)

#     root.mainloop()

# if __name__ == "__main__":
#     main()


import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pdfplumber
import unicodedata
import re
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

class ProcesadorPDF:
    def __init__(self, output_folder):
        # Secuencias completas que deben aparecer EN UNA MISMA PÁGINA
        self.SECUENCIA_COMPLETA_DOLAR = [
            "cuenta vista",
            "resumen cuenta", 
            "dolar americano",
            "movimientos"
        ]
        self.SECUENCIA_COMPLETA_EURO = [
            "cuenta vista",
            "resumen cuenta",
            "euro",
            "movimientos"
        ]
        self.columnas = [
            "Fecha", "Descripción", "Referencia", "Abonos", "Cargos",
            "Movimiento garantia", "Saldo en garantia", 
            "Saldo disponible", "Saldo total"
        ]
        self.output_folder = output_folder
        # Definir regiones para recorte (ajusta estas coordenadas según tu PDF)
        self.regionEmpresa = (485.0, 288.0, 745.0, 297.0)
        self.regionPeriodo = (588, 532, 725, 540.3)
        self.regionRfc = (588, 515, 636, 523.3)
        self.regionCliente = (588, 430, 630, 438.25)
    
    def normalizar_texto(self, texto):
        """Normaliza texto para comparación insensible a mayúsculas/tildes"""
        texto = texto.lower()
        texto = unicodedata.normalize('NFD', texto).encode('ascii', 'ignore').decode('utf-8')
        return texto
    
    def buscar_secuencia_completa(self, texto_pagina, secuencia):
        """Verifica si TODOS los elementos de la secuencia están en la misma página"""
        texto_normalizado = self.normalizar_texto(texto_pagina)
        return all(
            re.search(r'\b' + re.escape(palabra) + r'\b', texto_normalizado)
            for palabra in secuencia
        )

    def leer_palabras_paginas(self, pdf_path, pagina_inicial=4):
        """
        Lee palabras desde la página 1 para debug y luego 
        desde página_inicial hasta el final, ordenadas por posición.
        """
        palabras_por_pagina = []
        with pdfplumber.open(pdf_path) as pdf:
            # Verificar que el PDF tenga páginas
            if len(pdf.pages) == 0:
                messagebox.showwarning("Advertencia", "El PDF está vacío")
                return []

            # ==============
            # LECTURA DE LA PÁGINA 1 CON DEPURACIÓN
            # ==============
            page1 = pdf.pages[3]
            
            # OJO: Ajusta la tolerancia si lo deseas
            words_page1 = page1.extract_words(
                x_tolerance=2,  # <-- CAMBIO: tolerancia horizontal
                y_tolerance=2   # <-- CAMBIO: tolerancia vertical
            )
            
            # Imprimimos cuántas palabras hay en la página 1
            print("\n=== PALABRAS PÁGINA 1 ===")
            print(f"Total de palabras encontradas: {len(words_page1)}")  # <-- CAMBIO

            # Imprimir cada palabra con sus coordenadas
            for word in words_page1:
                print(
                    f"Texto: '{word['text']}', "
                    f"x0: {word['x0']:.2f}, x1: {word['x1']:.2f}, "
                    f"top: {word['top']:.2f}, bottom: {word['bottom']:.2f}"
                )

            # Extraer información de regiones específicas
            croppedEmpresa = page1.within_bbox(self.regionEmpresa)
            croppedPeriodo = page1.within_bbox(self.regionPeriodo)
            croppedRfc = page1.within_bbox(self.regionRfc)
            croppedCliente = page1.within_bbox(self.regionCliente)
            empresa_str = croppedEmpresa.extract_text() or "(No encontrado)"
            periodo_str = croppedPeriodo.extract_text() or "(No encontrado)"
            rfc_str = croppedRfc.extract_text() or "(No encontrado)"
            noCliente_str = croppedCliente.extract_text() or "(No encontrado)"
            
            print(f"\nEmpresa (regionEmpresa): {empresa_str}")
            print(f"Período (regionPeriodo): {periodo_str}")
            print("==========================\n")
            
            # Verificar si hay suficientes páginas en el PDF
            total_paginas = len(pdf.pages)
            if total_paginas < pagina_inicial:
                messagebox.showwarning("Advertencia", f"El PDF no tiene página {pagina_inicial}")
                return []
            
            # Procesar desde la página que desees (por defecto 4) hasta el final
            for page_num in range(pagina_inicial - 1, total_paginas):
                page = pdf.pages[page_num]
                
                # De nuevo, extrae las palabras con las tolerancias
                words = page.extract_words(
                    x_tolerance=2,  # <-- CAMBIO (opcional)
                    y_tolerance=2   # <-- CAMBIO (opcional)
                )

                # Ordenar palabras por la posición vertical (top) y luego horizontal (x0)
                words_sorted = sorted(words, key=lambda w: (w['top'], w['x0']))
                palabras_por_pagina.append((page_num + 1, words_sorted))
        return palabras_por_pagina

    def es_linea_encabezado(self, words_in_line, divisa_actual):
        """Determina si la línea actual es parte de los encabezados"""
        textos = [w['text'].strip().lower() for w in words_in_line]
        
        # Patrones de encabezado para cada divisa
        patrones = {
            'PESOS': ['fechas', 'descripción', 'referencia', 'abonos', 'cargos'],
            'DOLAR': ['fechas', 'descripción', 'referencia', 'saldo', 'total'],
            'EURO': ['fechas', 'descripción', 'referencia', 'saldo', 'total']
        }
        
        # Verificar si coincide con algún patrón de encabezado
        for patron in patrones[divisa_actual]:
            if patron in textos:
                return True
        
        # Caso especial para encabezados multilínea
        if divisa_actual != 'PESOS':
            if 'saldo' in textos and any(t in textos for t in ['total', 'general']):
                return True
        
        return False

    def detectar_estructura(self, pdf_path):
        """Detecta cambios de divisa, encabezados y saldos iniciales en el PDF"""
        palabras_por_pagina = self.leer_palabras_paginas(pdf_path)
        if not palabras_por_pagina:
            return [], [], []

        estado = {
            'divisa_actual': "PESOS",
            'cambios_divisa': [],
            'primer_encabezado_encontrado': False,
            'saldo_inicial_encontrado': False
        }

        resultados = []
        saldos_iniciales = []
        
        for page_num, words_sorted in palabras_por_pagina:
            # Convertir palabras de la página a texto para detección
            texto_pagina = " ".join(w['text'] for w in words_sorted).lower()
            
            # 1. Detectar cambio de divisa
            if estado['divisa_actual'] == "PESOS":
                if self.buscar_secuencia_completa(texto_pagina, self.SECUENCIA_COMPLETA_DOLAR):
                    estado['divisa_actual'] = "DOLAR"
                    estado['primer_encabezado_encontrado'] = False
                    estado['saldo_inicial_encontrado'] = False
                    cambio = f"Página {page_num}: Cambio a DÓLAR"
                    estado['cambios_divisa'].append(cambio)
                    print(cambio)
            
            elif estado['divisa_actual'] == "DOLAR":
                if self.buscar_secuencia_completa(texto_pagina, self.SECUENCIA_COMPLETA_EURO):
                    estado['divisa_actual'] = "EURO"
                    estado['primer_encabezado_encontrado'] = False
                    estado['saldo_inicial_encontrado'] = False
                    cambio = f"Página {page_num}: Cambio a EURO"
                    estado['cambios_divisa'].append(cambio)
                    print(cambio)

            # 2. Procesar líneas (detectar encabezados y saldo inicial)
            lineas = []
            current_line = []
            current_top = None
            
            for word in words_sorted:
                if current_top is None or abs(word['top'] - current_top) <= 5:
                    current_line.append(word)
                    current_top = word['top']
                else:
                    lineas.append(current_line)
                    current_line = [word]
                    current_top = word['top']
            if current_line:
                lineas.append(current_line)

            for linea in lineas:
                texto_linea = " ".join(w['text'] for w in linea).lower()
                
                if not estado['primer_encabezado_encontrado']:
                    if self.es_linea_encabezado(linea, estado['divisa_actual']):
                        estado['primer_encabezado_encontrado'] = True
                        resultados.append({
                            'pagina': page_num,
                            'tipo': 'inicio_datos',
                            'divisa': estado['divisa_actual']
                        })
                        print(f"Encabezado encontrado en página {page_num} para {estado['divisa_actual']}")
                else:
                    # Buscar "saldo inicial" después de los encabezados
                    if "saldo inicial" in texto_linea and not estado['saldo_inicial_encontrado']:
                        estado['saldo_inicial_encontrado'] = True
                        saldos_iniciales.append(f"Se encontró 'saldo inicial' en página {page_num}")
                        print(f"Saldo inicial encontrado en página {page_num}")

        return resultados, estado['cambios_divisa'], saldos_iniciales

    def generar_excel(self, pdf_path):
        """Genera el archivo Excel con la estructura requerida"""
        try:
            # Primero analizamos la estructura
            resultados, cambios_divisa, saldos_iniciales = self.detectar_estructura(pdf_path)
            
            # Extraer información de la página 1
            with pdfplumber.open(pdf_path) as pdf:
                if len(pdf.pages) > 0:
                    page1 = pdf.pages[1]
                    croppedEmpresa = page1.within_bbox(self.regionEmpresa)
                    croppedPeriodo = page1.within_bbox(self.regionPeriodo)
                    croppedCliente = page1.within_bbox(self.regionCliente)
                    croppedRfc = page1.within_bbox(self.regionRfc)
                    empresa_str = croppedEmpresa.extract_text() or "(No definido)"
                    periodo_str = croppedPeriodo.extract_text() or "(No definido)"
                    rfc_str = croppedRfc.extract_text() or "(No definido)"
                    noCliente_str = croppedCliente.extract_text() or "(No definido)"
                
            
            # Construir nombre de archivo Excel
            pdf_name = os.path.basename(pdf_path)
            pdf_stem, _ = os.path.splitext(pdf_name)
            excel_name = pdf_stem + ".xlsx"
            ruta_salida = os.path.join(self.output_folder, excel_name)
            
            # Crear un nuevo Workbook
            wb = Workbook()
            
            # Crear las 3 hojas que necesitamos
            for divisa in ["PESOS", "DOLAR", "EURO"]:
                if divisa in wb.sheetnames:
                    ws = wb[divisa]
                else:
                    ws = wb.create_sheet(title=divisa)
                
                # Eliminar la hoja por defecto si existe
                if "Sheet" in wb.sheetnames:
                    del wb["Sheet"]
                
                # Agregar encabezado del documento con info extraída
                ws['A1'] = f"Banco: Monex"
                ws['A2'] = f"Empresa: {empresa_str}"
                ws['A3'] = "No. Cuenta: (por definir)"
                ws['A4'] = f"No. Cliente: {noCliente_str}"
                ws['A5'] = f"Periodo: {periodo_str}"
                ws['A6'] = f"RFC: {rfc_str}"
                
                # Agregar encabezados de columnas
                for col_num, col_name in enumerate(self.columnas, 1):
                    ws.cell(row=7, column=col_num, value=col_name)
                
                # Aplicar estilos
                self._aplicar_estilos(ws)
            
            # Guardar el archivo
            wb.save(ruta_salida)
            return ruta_salida
            
        except Exception as e:
            raise Exception(f"Error al crear Excel: {str(e)}")
    
    def _aplicar_estilos(self, worksheet):
        """Aplica los estilos a la hoja de trabajo"""
        # Configurar estilos
        thin_side = Side(border_style="thin")
        thin_border = Border(top=thin_side, left=thin_side,
                             right=thin_side, bottom=thin_side)
        
        header_fill = PatternFill(start_color="000080",
                                 end_color="000080",
                                 fill_type="solid")
        white_font = Font(color="FFFFFF", bold=True)
        
        # Aplicar estilo a la fila de encabezados (fila 7)
        for col in range(1, len(self.columnas) + 1):
            cell = worksheet.cell(row=7, column=col)
            cell.fill = header_fill
            cell.font = white_font
            cell.alignment = Alignment(horizontal="center")
            cell.border = thin_border
        
        # Ajustar ancho de columnas
        for col in worksheet.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if cell.value is not None:
                        length = len(str(cell.value))
                        if length > max_length:
                            max_length = length
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            worksheet.column_dimensions[col_letter].width = adjusted_width
        
        # Wrap text para todas las celdas
        for row in worksheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(wrap_text=True)

class Aplicacion:
    def __init__(self, root, output_folder):
        self.root = root
        self.procesador = ProcesadorPDF(output_folder)
        self.configurar_interfaz()
        
    def configurar_interfaz(self):
        self.root.title("Analizador de Estados de Cuenta")
        self.root.geometry("700x200")
        
        # Centrar ventana
        window_width = 700
        window_height = 200
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        center_x = int(screen_width/2 - window_width/2)
        center_y = int(screen_height/2 - window_height/2)
        self.root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

        # Widgets
        frame = tk.Frame(self.root)
        frame.pack(pady=20)

        lbl_instruccion = tk.Label(frame, text="Seleccione PDF(s) para analizar desde página 4 hasta final:")
        lbl_instruccion.pack()

        btn_cargar = tk.Button(frame, text="Seleccionar PDF(s)", command=self.cargar_archivo, width=20)
        btn_cargar.pack(pady=5)

        self.entry_archivo = tk.Entry(frame, width=80, state=tk.DISABLED)
        self.entry_archivo.pack(pady=5)

        btn_procesar = tk.Button(frame, text="Analizar estructura", command=self.procesar_pdf, width=20)
        btn_procesar.pack(pady=5)

        btn_generar_excel = tk.Button(frame, text="Generar Excel", command=self.generar_excel, width=20)
        btn_generar_excel.pack(pady=5)

    def cargar_archivo(self):
        archivos = filedialog.askopenfilenames(
            title="Selecciona uno o más archivos PDF",
            filetypes=[("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*")]
        )
        if archivos:
            self.entry_archivo.config(state=tk.NORMAL)
            self.entry_archivo.delete(0, tk.END)
            self.entry_archivo.insert(0, " ; ".join(archivos))
            self.entry_archivo.config(state=tk.DISABLED)
            self.procesador.pdf_paths = archivos

    def procesar_pdf(self):
        if not hasattr(self.procesador, 'pdf_paths') or not self.procesador.pdf_paths:
            messagebox.showwarning("Advertencia", "No se ha seleccionado un archivo PDF.")
            return

        for pdf_path in self.procesador.pdf_paths:
            try:
                nombre_pdf = os.path.basename(pdf_path)
                resultados, cambios_divisa, saldos_iniciales = self.procesador.detectar_estructura(pdf_path)
                
                # Preparar resultados para mostrar
                resultado_texto = f"Archivo: {nombre_pdf}\n\n"
                
                # Mostrar cambios de divisa
                resultado_texto += "--- CAMBIOS DE DIVISA DETECTADOS ---\n"
                if cambios_divisa:
                    resultado_texto += "\n".join(cambios_divisa) + "\n"
                else:
                    resultado_texto += "No se detectaron cambios de divisa (solo PESOS)\n"
                
                # Mostrar puntos de inicio de datos
                resultado_texto += "\n--- INICIO DE DATOS DETECTADOS ---\n"
                if resultados:
                    for res in resultados:
                        resultado_texto += (
                            f"Página {res['pagina']}: Datos comienzan en {res['divisa']}\n"
                        )
                else:
                    resultado_texto += "No se encontró el inicio de los datos\n"
                
                # Mostrar saldos iniciales encontrados
                resultado_texto += "\n--- SALDOS INICIALES ---\n"
                if saldos_iniciales:
                    resultado_texto += "\n".join(saldos_iniciales) + "\n"
                else:
                    resultado_texto += "No se encontró la frase 'saldo inicial'\n"
                
                # Mostrar en ventana de resultados
                self.mostrar_resultados(resultado_texto)
                
            except Exception as e:
                messagebox.showerror("Error", f"Error procesando {pdf_path}:\n{str(e)}")

    def generar_excel(self):
        if not hasattr(self.procesador, 'pdf_paths') or not self.procesador.pdf_paths:
            messagebox.showwarning("Advertencia", "No se ha seleccionado un archivo PDF.")
            return

        for pdf_path in self.procesador.pdf_paths:
            try:
                ruta_excel = self.procesador.generar_excel(pdf_path)
                messagebox.showinfo("Éxito", f"Archivo Excel generado:\n{ruta_excel}")
                
                # Preguntar si quieren abrir el archivo
                if messagebox.askyesno("Abrir archivo", "¿Desea abrir el archivo generado?"):
                    os.startfile(ruta_excel)
                    
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo generar el archivo Excel:\n{str(e)}")

    def mostrar_resultados(self, texto):
        """Muestra resultados en una ventana con scroll."""
        result_window = tk.Toplevel(self.root)
        result_window.title("Resultados de Análisis")
        result_window.geometry("900x600")
        
        frame = tk.Frame(result_window)
        frame.pack(fill=tk.BOTH, expand=True)
        
        text_area = scrolledtext.ScrolledText(frame, wrap=tk.WORD, width=100, height=30)
        text_area.pack(fill=tk.BOTH, expand=True)
        text_area.insert(tk.END, texto)
        text_area.configure(state=tk.DISABLED)
        
        # Botón para copiar
        btn_copiar = tk.Button(frame, text="Copiar Resultados", 
                              command=lambda: self.copiar_al_portapapeles(texto))
        btn_copiar.pack(pady=5)

    def copiar_al_portapapeles(self, texto):
        """Copia el texto al portapapeles."""
        self.root.clipboard_clear()
        self.root.clipboard_append(texto)
        messagebox.showinfo("Copiado", "Resultados copiados al portapapeles")

def main():
    if len(sys.argv) < 2:
        print("Error: Se requiere especificar la carpeta de salida")
        print("Uso: python script.py <carpeta_salida>")
        return
    
    output_folder = sys.argv[1]
    if not os.path.isdir(output_folder):
        print(f"Error: La carpeta de salida no existe: {output_folder}")
        return
    
    root = tk.Tk()
    app = Aplicacion(root, output_folder)
    root.mainloop()

if __name__ == "__main__":
    main()
