import tkinter as tk
from tkinter import filedialog, messagebox
import tabula
import pandas as pd

# Lista global para almacenar las rutas de los PDFs seleccionados
pdf_files = []

def cargar_archivos():
    global pdf_files
    # Abre el diálogo para seleccionar múltiples archivos PDF
    archivos = filedialog.askopenfilenames(
        title="Selecciona archivos PDF", 
        filetypes=(("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*"))
    )
    # Convierte la tupla a lista
    pdf_files = list(archivos)
    
    # Actualiza el campo de texto para mostrar los archivos seleccionados
    entry_archivos.config(state=tk.NORMAL)
    entry_archivos.delete(0, tk.END)
    if pdf_files:
        entry_archivos.insert(0, ", ".join(pdf_files))
    else:
        entry_archivos.insert(0, "No se seleccionaron archivos")
    entry_archivos.config(state=tk.DISABLED)

def procesar_pdfs():
    if not pdf_files:
        messagebox.showwarning("Advertencia", "No se han seleccionado archivos PDF.")
        return

    dataframes = []
    for pdf_file in pdf_files:
        try:
            # Extrae todas las tablas de todas las páginas del PDF
            tablas = tabula.read_pdf(pdf_file, pages='all', multiple_tables=True)
            # Añade cada tabla extraída a la lista
            for tabla in tablas:
                dataframes.append(tabla)
        except Exception as e:
            messagebox.showerror("Error", f"Error al procesar {pdf_file}:\n{e}")
            return

    if dataframes:
        # Combina todos los DataFrames en uno solo
        df_combinado = pd.concat(dataframes, ignore_index=True)
        # Exporta el DataFrame combinado a un archivo Excel
        df_combinado.to_excel('movimientos_combinados.xlsx', index=False)
        messagebox.showinfo("Éxito", "Archivo Excel generado exitosamente.")
    else:
        messagebox.showinfo("Información", "No se encontraron tablas en los archivos seleccionados.")

# Configuración de la ventana principal
root = tk.Tk()
root.title("Procesar PDFs a Excel")
root.geometry("600x200")

# Botón para cargar archivos
btn_cargar = tk.Button(root, text="Seleccionar archivos PDF", command=cargar_archivos, width=30)
btn_cargar.pack(pady=10)

# Campo de texto (Entry) para mostrar la lista de archivos seleccionados
entry_archivos = tk.Entry(root, width=80, state=tk.DISABLED)
entry_archivos.pack(padx=10, pady=10)

# Botón para procesar los PDFs
btn_procesar = tk.Button(root, text="Procesar archivos PDF", command=procesar_pdfs, width=30)
btn_procesar.pack(pady=10)

root.mainloop()
