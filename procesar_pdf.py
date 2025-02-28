import tkinter as tk
from tkinter import messagebox, filedialog
import subprocess
import os
from PIL import Image, ImageTk  # Para manejar JPG/WEBP/PNG con Pillow

# Variable global para almacenar la carpeta de salida
output_dir = None

def select_output_folder():
    """Permite al usuario seleccionar la carpeta de salida."""
    global output_dir
    folder = filedialog.askdirectory(title="Selecciona carpeta de salida")
    if folder:
        output_dir.set(folder)

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
    try:
        subprocess.run(["python3", "procesar_pdf_banorte.py"])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ejecutar procesar_pdf_banorte.py\n{e}")

def run_algo3():
    """Ejecuta procesar_pdf_bbva.py."""
    try:
        subprocess.run(["python3", "procesar_pdf_bbva.py"])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ejecutar procesar_pdf_bbva.py\n{e}")

def run_algo4():
    """Ejecuta procesar_pdf_multiva.py."""
    try:
        subprocess.run(["python3", "procesar_pdf_multiva.py"])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ejecutar procesar_pdf_multiva.py\n{e}")

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

    # Creamos 4 botones con sus imágenes
    btn1 = tk.Button(frame, image=logo1, command=run_banamex)
    btn2 = tk.Button(frame, image=logo2, command=run_algo2)
    btn3 = tk.Button(frame, image=logo3, command=run_algo3)
    btn4 = tk.Button(frame, image=logo4, command=run_algo4)

    # Ubicamos en cuadrícula 2x2
    btn1.grid(row=0, column=0, padx=20, pady=20)
    btn2.grid(row=0, column=1, padx=20, pady=20)
    btn3.grid(row=1, column=0, padx=20, pady=20)
    btn4.grid(row=1, column=1, padx=20, pady=20)

    # Evitar que Python limpie las imágenes
    root.logo1 = logo1
    root.logo2 = logo2
    root.logo3 = logo3
    root.logo4 = logo4

    root.mainloop()

if __name__ == "__main__":
    main()
