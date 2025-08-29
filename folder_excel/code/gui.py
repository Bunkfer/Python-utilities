import tkinter as tk
from tkinter import filedialog, messagebox
import os
from package import FolderProcessor, Exporter


class AppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Buscar archivos")
        self.root.geometry("400x200")

        # Variable interna de Tkinter
        self.path_var = tk.StringVar()

        # Interfaz
        tk.Label(root, text="Ruta").pack()
        tk.Entry(root, textvariable=self.path_var, width=50).pack(pady=10)
        tk.Button(root, text="Seleccionar ruta", command=self.select_folder).pack(
            pady=10
        )
        tk.Button(root, text="Generar Excel", command=self.generate_excel).pack(pady=10)

    def select_folder(self):
        folder = filedialog.askdirectory(title="Selecciona la carpeta")
        if folder:
            self.path_var.set(folder)

    def generate_excel(self):
        base_path = self.path_var.get()

        if not base_path:
            messagebox.showerror("Error", "Selecciona una carpeta primero.")
            return

        processor = FolderProcessor(base_path)
        data = processor.scan()

        file_path = os.path.join(base_path, "FolderTOexcel.xlsx")
        Exporter(data).to_excel(file_path)

        messagebox.showinfo("Éxito", f"Archivo generado en:\n{file_path}")
