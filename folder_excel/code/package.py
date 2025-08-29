import os
import pandas as pd


class FolderProcessor:
    def __init__(self, base_path: str):
        self.base_path = base_path

    def scan(self):
        data = []

        for base in os.listdir(self.base_path):
            full_path = os.path.join(self.base_path, base)

            if os.path.isdir(full_path):
                for item in os.listdir(full_path):
                    sub_path = os.path.join(full_path, item)
                    tipo = "Carpeta" if os.path.isdir(sub_path) else "Archivo"
                    data.append([base, item, tipo])

        return data


class Exporter:
    def __init__(self, data):
        self.data = data

    def to_excel(self, file_path: str):
        df = pd.DataFrame(self.data, columns=["Carpeta", "Nombre", "Tipo"])
        df.to_excel(file_path, index=False)

    def to_csv(self, file_path: str):
        df = pd.DataFrame(self.data, columns=["Carpeta", "Nombre", "Tipo"])
        df.to_csv(file_path, index=False)
