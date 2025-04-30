import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
from santander_f1 import convertir_santander_f1
from santander_f2 import convertir_santander_f2

FRASES_F1 = ("ESTADO DE CUENTA INTEGRAL","INFORMACION A CLIENTES")

def detectar_formato(pdf_path: str) -> str:
    with pdfplumber.open(pdf_path) as pdf:
        text = (pdf.pages[0].extract_text() or "").upper()
    if all(f in text for f in FRASES_F1): 
        print("Formato F1")
        return "f1"
    if "ESTADO DE CUENTA" in text: 
        print("Formato F2")
        return "f2"
    return "desconocido"

def cargar():
    archivos = filedialog.askopenfilenames(title="Selecciona PDFs", filetypes=[("PDF","*.pdf")])
    if archivos:
        app.pdfs = archivos
        entry.config(state="normal")
        entry.delete(0, tk.END)
        entry.insert(0, " ; ".join(archivos))
        entry.config(state="disabled")

def procesar():
    if not getattr(app, "pdfs", None): return messagebox.showwarning("Aviso","No hay PDFs seleccionados")
    out_dir = sys.argv[1] if len(sys.argv)>=2 else os.getcwd()
    for p in app.pdfs:
        fmt = detectar_formato(p)
        try:
            if fmt=="f1": ruta = convertir_santander_f1(p, out_dir)
            elif fmt=="f2": ruta = convertir_santander_f2(p, out_dir)
            else: raise ValueError("Formato desconocido")
            messagebox.showinfo("Éxito", f"Generado:\n{ruta}")
        except Exception as e:
            messagebox.showerror("Error", f"{p}\n{e}")

if __name__=="__main__":
    app = tk.Tk(); app.title("Convertir Santander F1/F2")
    tk.Button(app,text="Cargar PDF",command=cargar).pack(pady=5)
    entry = tk.Entry(app,width=80,state="disabled"); entry.pack(padx=5,pady=5)
    tk.Button(app,text="Procesar",command=procesar).pack(pady=5)
    app.mainloop()