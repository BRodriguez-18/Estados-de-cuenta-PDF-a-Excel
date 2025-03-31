import tkinter as tk
from tkinter import messagebox, filedialog
import subprocess
import os
from PIL import Image, ImageTk  # Para manejar JPG/WEBP/PNG con Pillow

# Variable global para almacenar la carpeta de salida
RUTA_TXT = "ruta.txt"
output_dir = None

def load_saved_route():
    "lee la ultima carpeta de salida de ruta.txt (si existe) y la asigna a output_dir."
    global output_dir
    if os.path.exists(RUTA_TXT):
        with open(RUTA_TXT, "r", encoding="utf-8") as f:
            folder = f.read().strip()
            if folder:
                output_dir.set(folder)

def save_route_to_file(folder):
    "Guarda la nueva ruta en el archivo ruta.tx"
    with open(RUTA_TXT, "w", encoding="utf-8") as f:
        f.write(folder)

def select_output_folder():
    """Permite al usuario seleccionar la carpeta de salida."""
    global output_dir
    folder = filedialog.askdirectory(title="Selecciona carpeta de salida")
    if folder:
        output_dir.set(folder)
        save_route_to_file(folder)

def run_banamex():
    """Ejecuta procesar_pdf_banamex.py, pasando la carpeta de salida como argumento."""
    global output_dir
    folder = output_dir.get()
    if not folder:
        messagebox.showwarning("Advertencia", "No se ha seleccionado carpeta de salida.")
        return
    try:
        # Llamamos al script Banamex con la carpeta de salida
        subprocess.run(["python3", "procesar_pdf_banamex.py", folder])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ejecutar procesar_pdf_banamex.py\n{e}")

def run_algo2():
    """Ejecuta procesar_pdf_banorte.py (sin carpeta de salida)."""
    global output_dir
    folder = output_dir.get()
    if not folder:
        messagebox.showwarning("Advertencia", "No se ha seleccionado carpeta de salida.")
        return    
    try:
        subprocess.run(["python3", "procesar_pdf_banorte.py", folder])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ejecutar procesar_pdf_banorte.py\n{e}")

def run_algo3():
    """Ejecuta procesar_pdf_bbva.py."""
    global output_dir
    folder = output_dir.get()
    if not folder:
        messagebox.showwarning("Advertencia", "No se ha seleccionado carpeta de salida.")
        return    
    try:
        subprocess.run(["python3", "procesar_pdf_bbva.py", folder])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ejecutar procesar_pdf_bbva.py\n{e}")

def run_algo4():
    """Ejecuta procesar_pdf_multiva.py."""
    global output_dir
    folder = output_dir.get()
    if not folder:
        messagebox.showwarning("Advertencia", "No se ha seleccionado carpeta de salida.")
        return  
    try:
        subprocess.run(["python3", "procesar_pdf_multiva.py", folder])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ejecutar procesar_pdf_multiva.py\n{e}")

def run_banbajio():
    """Ejecuta el script de banbajio"""
    global output_dir
    folder = output_dir.get()
    if not folder:
        messagebox.showwarning("Advertencia", "No se ha seleccionado carpeta de salida.")
        return
    try:
        subprocess.run(["python3", "procesar_pdf_banbajio.py", folder])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ejecutar procesar_pdf_banbajio.py\n{e}")    

def run_santander():
    """Ejecuta el script de santander"""
    global output_dir
    folder = output_dir.get()
    if not folder:
        messagebox.showwarning("Advertencia", "No se ha seleccionado carpeta de salida.")
        return
    try:
        subprocess.run(["python3", "procesar_pdf_santander.py", folder])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ejecutar procesar_pdf_santander.py\n{e}")    

def run_monex():
    """Ejecuta el script de monex"""
    global output_dir
    folder = output_dir.get()
    if not folder:
        messagebox.showwarning("Advertencia", "No se ha seleccionado carpeta de salida.")
        return
    try:
        subprocess.run(["python3", "procesar_pdf_monex.py", folder])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ejecutar procesar_pdf_monex.py\n{e}")

def run_mifel():
    global output_dir
    folder = output_dir.get()
    if not folder:
        messagebox.showwarning("Advertencia", "No se ha seleccionado carpeta de salida.")
        return
    try:
        subprocess.run(["python3", "procesar_pdf_mifel.py", folder])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ejecutar procesar_pdf_mifel.py\n{e}")


def main():
    global output_dir
    root = tk.Tk()
    root.title("Menú de Algoritmos")

    # Dimensiones deseadas de la ventana
    win_width = 800
    win_height = 400

    # Obtenemos dimensiones de la pantalla
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # Calculamos coordenadas x e y para centrar
    x = (screen_width - win_width) // 2
    y = (screen_height - win_height) // 2

    # Ajustamos la geometría: ancho x alto + x + y
    root.geometry(f"{win_width}x{win_height}+{x}+{y}")

    # Definimos la variable global como tk.StringVar
    output_dir = tk.StringVar(value="")

    #Cargamos la ultima ruta guardada
    load_saved_route()

    # Botón para seleccionar carpeta de salida
    btn_select_folder = tk.Button(root, text="Seleccionar Carpeta de Salida", command=select_output_folder)
    btn_select_folder.pack(pady=10)

    # Label para mostrar la carpeta elegida
    lbl_folder = tk.Label(root, textvariable=output_dir, fg="blue")
    lbl_folder.pack()

    # Frame para los botones de bancos
    frame = tk.Frame(root)
    frame.pack(expand=True)

    # Ruta de la carpeta Logos
    logos_path = "Logos"
    # Tamaño deseado de cada imagen
    img_size = (100, 100)

    # 1) banamex
    img1_path = os.path.join(logos_path, "banamex.jpg")
    img1_original = Image.open(img1_path)
    img1_resized = img1_original.resize(img_size, Image.Resampling.LANCZOS)
    logo1 = ImageTk.PhotoImage(img1_resized)

    # 2) banorte
    img2_path = os.path.join(logos_path, "banorte.webp")
    img2_original = Image.open(img2_path)
    img2_resized = img2_original.resize(img_size, Image.Resampling.LANCZOS)
    logo2 = ImageTk.PhotoImage(img2_resized)

    # 3) bbva
    img3_path = os.path.join(logos_path, "bbva.png")
    img3_original = Image.open(img3_path)
    img3_resized = img3_original.resize(img_size, Image.Resampling.LANCZOS)
    logo3 = ImageTk.PhotoImage(img3_resized)

    # 4) multiva
    img4_path = os.path.join(logos_path, "multiva.png")
    img4_original = Image.open(img4_path)
    img4_resized = img4_original.resize(img_size, Image.Resampling.LANCZOS)
    logo4 = ImageTk.PhotoImage(img4_resized)

    #5) banbajio
    img5_path = os.path.join(logos_path, "banbajio.jpg")
    img5_original = Image.open(img5_path)
    img5_resized = img5_original.resize(img_size, Image.Resampling.LANCZOS)
    logo5 = ImageTk.PhotoImage(img5_resized)

    #6) santander
    img6_path = os.path.join(logos_path, "santander.png")
    img6_original = Image.open(img6_path)
    img6_resized = img6_original.resize(img_size, Image.Resampling.LANCZOS)
    logo6 = ImageTk.PhotoImage(img6_resized)

    #7) monex
    img7_path = os.path.join(logos_path, "monex.png")
    img7_original = Image.open(img7_path)
    img7_resized = img7_original.resize(img_size, Image.Resampling.LANCZOS)
    logo7 = ImageTk.PhotoImage(img7_resized)

    img8_path = os.path.join(logos_path, "mifel.png")
    img8_original = Image.open(img8_path)
    img8_resized = img8_original.resize(img_size, Image.Resampling.LANCZOS)
    logo8 = ImageTk.PhotoImage(img8_resized)

    # Creamos 4 botones con sus imágenes
    btn1 = tk.Button(frame, image=logo1, command=run_banamex)
    btn2 = tk.Button(frame, image=logo2, command=run_algo2)
    btn3 = tk.Button(frame, image=logo3, command=run_algo3)
    btn4 = tk.Button(frame, image=logo4, command=run_algo4)
    btn5 = tk.Button(frame, image=logo5, command=run_banbajio)
    btn6 = tk.Button(frame, image=logo6, command=run_santander)
    btn7 = tk.Button(frame, image=logo7, command=run_monex)
    btn8 = tk.Button(frame, image=logo8, command=run_mifel)

    # Ubicamos en cuadrícula 2x2
    btn1.grid(row=0, column=0, padx=20, pady=20)
    btn2.grid(row=0, column=1, padx=20, pady=20)
    btn3.grid(row=1, column=0, padx=20, pady=20)
    btn4.grid(row=1, column=1, padx=20, pady=20)
    btn5.grid(row=0, column=2, padx=20, pady=20)
    btn6.grid(row=1, column=2, padx=20, pady=20)
    btn7.grid(row=0, column=3, padx=20, pady=20)   
    btn8.grid(row=1, column=3, padx=20, pady=20) 

    # Evitar que Python limpie las imágenes
    root.logo1 = logo1
    root.logo2 = logo2
    root.logo3 = logo3
    root.logo4 = logo4
    root.logo5 = logo5
    root.logo6 = logo6
    root.logo7 = logo7
    root.logo8 = logo8

    root.mainloop()

if __name__ == "__main__":
    main()
