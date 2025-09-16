import tkinter as tk
from tkinter import filedialog, messagebox
from paquetes import ReporteConsolidado


class AppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Pago al plan PP")
        self.root.geometry("540x140")

        # Variable interna de Tkinter
        self.path_var = tk.StringVar()
        self.file_var = tk.StringVar()

        # Interfaz con grid
        tk.Label(root, text="Ruta").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        tk.Entry(root, textvariable=self.path_var, width=50).grid(
            row=0, column=1, padx=10, pady=10
        )
        tk.Button(root, text="Seleccionar ruta", command=self.select_folder).grid(
            row=0, column=2, padx=10, pady=10
        )

        tk.Label(root, text="Archivo Base").grid(
            row=1, column=0, sticky="w", padx=10, pady=10
        )
        tk.Entry(root, textvariable=self.file_var, width=50).grid(
            row=1, column=1, padx=10, pady=10
        )
        tk.Button(root, text="Seleccionar base", command=self.select_file).grid(
            row=1, column=2, padx=10, pady=10
        )

        # Botón grande debajo
        tk.Button(
            root, text="Consolidar reporte", command=self.generate_excel, width=50
        ).grid(row=2, column=0, columnspan=3, pady=10)

    def select_folder(self):
        folder = filedialog.askdirectory(title="Selecciona la carpeta")
        if folder:
            self.path_var.set(folder)

    def select_file(self):
        filetypes = [
            ("Archivos Excel", "*.xlsx"),
        ]
        filepath = filedialog.askopenfilename(
            title="Selecciona el archivo Excel", filetypes=filetypes
        )
        if filepath:
            self.file_var.set(filepath)

    def generate_excel(self):
        base_path = self.path_var.get()
        archivo_base = self.file_var.get()

        if not base_path:
            messagebox.showerror("Error", "Selecciona una carpeta primero.")
            return

        if not archivo_base:
            messagebox.showerror("Error", "Selecciona una base primero.")
            return

        reporte = ReporteConsolidado(base_path, archivo_base)
        reporte.consolidar_csvs()
        reporte.cargar_equivalencias_y_acreedores()
        reporte.generar_excel()

        messagebox.showinfo("Éxito", f"Archivo generado")
