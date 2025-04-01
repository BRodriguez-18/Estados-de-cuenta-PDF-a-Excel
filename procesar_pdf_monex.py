import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pdfplumber
import unicodedata
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

class ProcesadorPDF:
    def __init__(self, output_folder):
        # Palabras clave para cambio de divisa
        self.SECUENCIA_COMPLETA_DOLAR = [
            "cuenta vista", "resumen cuenta", 
            "dolar americano", "movimientos"
        ]
        self.SECUENCIA_COMPLETA_EURO = [
            "cuenta vista", "resumen cuenta", 
            "euro", "movimientos"
        ]
        self.columnas = [
            "Fecha","Descripción","Referencia","Abonos","Cargos",
            "Movimiento garantia","Saldo en garantia",
            "Saldo disponible","Saldo total"
        ]
        self.output_folder = output_folder

        # Regiones, si las usas
        self.regionEmpresa = (485.0, 288.0, 745.0, 297.0)
        self.regionPeriodo = (588, 532, 725, 540.3)
        self.regionRfc = (588, 515, 636, 523.3)
        self.regionCliente = (588, 430, 630, 438.25)

    def normalizar_texto(self, texto):
        texto = texto.lower()
        texto = unicodedata.normalize('NFD', texto).encode('ascii','ignore').decode('utf-8')
        return texto

    def buscar_secuencia_completa(self, texto_pagina, secuencia):
        texto_normalizado = self.normalizar_texto(texto_pagina)
        return all(re.search(r'\b' + re.escape(palabra) + r'\b', texto_normalizado)
                   for palabra in secuencia)

    def leer_palabras_paginas(self, pdf_path, pagina_inicial=4):
        """
        Lee las palabras de cada página (desde pagina_inicial) y retorna 
        una lista de (num_pagina, [words_sorted]).
        """
        palabras_por_pagina = []
        with pdfplumber.open(pdf_path) as pdf:
            total_paginas = len(pdf.pages)
            if total_paginas < pagina_inicial:
                print(f"El PDF no tiene la página {pagina_inicial}")
                return []
            for page_num in range(pagina_inicial - 1, total_paginas):
                page = pdf.pages[page_num]
                words = page.extract_words(x_tolerance=2,y_tolerance=2)
                words_sorted = sorted(words, key=lambda w: (w['top'], w['x0']))
                palabras_por_pagina.append((page_num+1, words_sorted))
        return palabras_por_pagina

    def es_linea_encabezado(self, words_in_line, divisa_actual):
        """
        Verifica si la línea es un encabezado, dependiendo de la divisa actual.
        """
        textos = [w['text'].strip().lower() for w in words_in_line]
        patrones = {
            'PESOS': ['fechas', 'descripción', 'referencia', 'abonos', 'cargos'],
            'DOLAR': ['fechas', 'descripción', 'referencia', 'saldo', 'total'],
            'EURO':  ['fechas', 'descripción', 'referencia', 'saldo', 'total']
        }
        for patron in patrones[divisa_actual]:
            if patron in textos:
                return True
        # Caso especial DOLAR/EURO multilinea
        if divisa_actual != 'PESOS':
            if 'saldo' in textos and any(t in textos for t in ['total','general']):
                return True
        return False

    ##########################################################################
    def procesar_movimientos_divisas(self, pdf_path):
        """
        1) Empieza en divisa PESOS.
        2) No lee movimientos hasta que:
           a) Encuentre el encabezado de la divisa
           b) Encuentre "saldo inicial" (solo en la primera página de esa divisa)
        3) Al llegar "saldo final", conservar esa línea en el último movimiento, 
           y omitir todo lo que venga después para la misma divisa.
           Solo reanudar al cambiar de divisa.
        """
        palabras_por_pagina = self.leer_palabras_paginas(pdf_path, pagina_inicial=4)
        if not palabras_por_pagina:
            print("No hay páginas desde la 4 en adelante.")
            return

        # Estado
        divisa_actual = "PESOS"
        encabezado_encontrado = False
        saldo_inicial_encontrado = False
        es_primera_pagina_divisa = True

        # Bandera para NO leer más líneas en esa divisa tras "saldo final"
        omitir_hasta_cambio_divisa = False

        # Donde guardamos los movimientos finales
        movimientos_divisas = {"PESOS":[], "DOLAR":[], "EURO":[]}

        def agrupar_en_lineas(words):
            lineas_res = []
            current_line = []
            current_top = None
            for w in words:
                if current_top is None or abs(w['top'] - current_top)<=5:
                    current_line.append(w)
                    current_top = w['top']
                else:
                    lineas_res.append(current_line)
                    current_line = [w]
                    current_top = w['top']
            if current_line:
                lineas_res.append(current_line)
            return lineas_res

        for page_num, words_sorted in palabras_por_pagina:
            # 1) Revisar si en esta página cambia la divisa:
            texto_pagina = " ".join(w['text'].lower() for w in words_sorted)

            if divisa_actual=="PESOS" and self.buscar_secuencia_completa(texto_pagina, self.SECUENCIA_COMPLETA_DOLAR):
                # PESOS -> DOLAR
                divisa_actual = "DOLAR"
                print(f"[INFO] Cambio a DÓLAR en página {page_num}")
                # Reset:
                encabezado_encontrado = False
                saldo_inicial_encontrado = False
                es_primera_pagina_divisa = True
                omitir_hasta_cambio_divisa = False

            elif divisa_actual=="DOLAR" and self.buscar_secuencia_completa(texto_pagina, self.SECUENCIA_COMPLETA_EURO):
                # DOLAR -> EURO
                divisa_actual = "EURO"
                print(f"[INFO] Cambio a EURO en página {page_num}")
                # Reset:
                encabezado_encontrado = False
                saldo_inicial_encontrado = False
                es_primera_pagina_divisa = True
                omitir_hasta_cambio_divisa = False

            # 2) Si omitir_hasta_cambio_divisa es True, no procesamos nada
            if omitir_hasta_cambio_divisa:
                print(f"[DEBUG] Omitiendo divisa {divisa_actual} en pág {page_num} (esperando cambio de divisa).")
                continue

            # 3) Agrupar en líneas
            lineas = agrupar_en_lineas(words_sorted)

            # Filtrar "Hoja X de Y"
            omitir_linea_patron = re.compile(r'^hoja\s+\d+\s+de\s+\d+$', re.IGNORECASE)
            lineas_filtradas = []
            for linea in lineas:
                texto_lin = " ".join(w['text'].lower() for w in linea).strip()
                if omitir_linea_patron.match(texto_lin):
                    continue
                lineas_filtradas.append(linea)

            # 4) Separamos en movimientos, pero SÓLO si se cumplieron las condiciones:
            #    a) encabezado_encontrado
            #    b) saldo_inicial_encontrado (solamente en la primera página de la divisa)
            # De lo contrario, podemos seguir buscando encabezado y saldo_inicial

            movimientos_pagina = []
            movimiento_actual = []

            def get_line_bounds(lin):
                tops = [x['top'] for x in lin]
                bots = [x['bottom'] for x in lin]
                return min(tops), max(bots)

            # Recorremos línea a línea
            for lin in lineas_filtradas:
                texto_lin = " ".join(w['text'].lower() for w in lin)

                # A) Buscar encabezado si no lo tenemos:
                if not encabezado_encontrado:
                    if self.es_linea_encabezado(lin, divisa_actual):
                        encabezado_encontrado = True
                        print(f"[INFO] Encabezado de {divisa_actual} en pág {page_num}")
                    # seguimos con la siguiente línea
                    continue

                # B) Si es la primera página de esta divisa y no hemos visto saldo_inicial:
                if es_primera_pagina_divisa and not saldo_inicial_encontrado:
                    if "saldo inicial" in texto_lin:
                        saldo_inicial_encontrado = True
                        print(f"[INFO] Saldo inicial en {divisa_actual}, pág {page_num}")
                        # No agregamos la línea de "saldo inicial" a un movimiento
                    # sea como sea, si no lo encontramos, no iniciamos movimientos
                    continue

                # Una vez que ya hay encabezado y (si es la primera página) saldo inicial, 
                # procesamos para separar en movimientos.

                # C) Detectar "saldo final"
                if "saldo final" in texto_lin:
                    # Mantenemos ESTA línea en el último movimiento
                    if not movimiento_actual:
                        # si no había movimiento_actual, iniciamos uno para poner la línea
                        movimiento_actual = lin
                    else:
                        # Revisar la distancia con la línea anterior
                        _, bottom_actual = get_line_bounds(movimiento_actual)
                        top_nueva, _ = get_line_bounds(lin)
                        if (top_nueva - bottom_actual) > 11.33:
                            movimientos_pagina.append(movimiento_actual)
                            movimiento_actual = lin
                        else:
                            movimiento_actual.extend(lin)

                    # YA tenemos la línea con "saldo final". La guardamos:
                    if movimiento_actual:
                        movimientos_pagina.append(movimiento_actual)
                    movimiento_actual = []

                    # Activamos la bandera de omitir:
                    omitir_hasta_cambio_divisa = True
                    print(f"[INFO] Se encontró SALDO FINAL en {divisa_actual}, en pág {page_num}. "
                          "Ya no leeremos más movimientos de esta divisa hasta cambio.")
                    # Terminamos el loop de línea, no seguimos en esta página
                    break

                # D) Lógica normal para unir la línea al movimiento
                if not movimiento_actual:
                    # Arrancamos un movimiento
                    movimiento_actual = lin
                else:
                    _, bottom_actual = get_line_bounds(movimiento_actual)
                    top_nueva, _ = get_line_bounds(lin)
                    if (top_nueva - bottom_actual) > 11.33:
                        # Nuevo movimiento
                        movimientos_pagina.append(movimiento_actual)
                        movimiento_actual = lin
                    else:
                        # Mismo movimiento
                        movimiento_actual.extend(lin)

            # Al terminar las líneas de la página:
            if movimiento_actual and not omitir_hasta_cambio_divisa:
                # si no hemos activado la omisión por saldo final, 
                # agregamos el movimiento en curso
                movimientos_pagina.append(movimiento_actual)

            # Agregar los movimientos detectados en esta página a la divisa actual
            movimientos_divisas[divisa_actual].extend(movimientos_pagina)
            # Si es la primera página de la divisa, ya la procesamos
            es_primera_pagina_divisa = False

        # Al final, imprimimos
        for d in ["PESOS","DOLAR","EURO"]:
            movs = movimientos_divisas[d]
            if not movs:
                print(f"\n[INFO] No hubo movimientos en {d}.")
                continue
            print(f"\n====== MOVIMIENTOS EN {d} ======")
            for idx, mov in enumerate(movs, start=1):
                texto = " ".join(w['text'] for w in mov)
                print(f"\n--- {d} - Movimiento {idx} ---")
                print(texto)

    ##########################################################################

    def generar_excel(self, pdf_path):
        # Tu lógica para Excel, si la implementas
        pass

    def _aplicar_estilos(self, worksheet):
        # ...
        pass


###############################################################################
class Aplicacion:
    def __init__(self, root, output_folder):
        self.root = root
        self.procesador = ProcesadorPDF(output_folder)
        self.configurar_interfaz()

    def configurar_interfaz(self):
        self.root.title("Analizador de Estados de Cuenta")
        self.root.geometry("700x200")
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
                print(f"\n=== PROCESANDO {os.path.basename(pdf_path)} ===\n")
                self.procesador.procesar_movimientos_divisas(pdf_path)
                messagebox.showinfo("OK", f"Procesado {os.path.basename(pdf_path)}")
            except Exception as e:
                messagebox.showerror("Error", f"Error procesando {pdf_path}:\n{str(e)}")

    def generar_excel(self):
        if not hasattr(self.procesador, 'pdf_paths') or not self.procesador.pdf_paths:
            messagebox.showwarning("Advertencia", "No se ha seleccionado un archivo PDF.")
            return
        for pdf_path in self.procesador.pdf_paths:
            try:
                self.procesador.generar_excel(pdf_path)
                messagebox.showinfo("OK", "Archivo Excel generado.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo generar el archivo Excel:\n{str(e)}")

def main():
    if len(sys.argv)<2:
        print("Uso: python script.py <carpeta_salida>")
        return
    output_folder = sys.argv[1]
    if not os.path.isdir(output_folder):
        print(f"No existe la carpeta: {output_folder}")
        return
    root = tk.Tk()
    app = Aplicacion(root, output_folder)
    root.mainloop()

if __name__=="__main__":
    main()
