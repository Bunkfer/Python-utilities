# Folder to Excel

This project provides a simple Python desktop application (with Tkinter) that scans a selected folder and exports its structure (subfolders and files) into an Excel file.  

It is designed with **Object-Oriented Programming (OOP)** and follows **SOLID principles** to keep the code organized, reusable, and easy to maintain.  

## Features
- Select a folder using a graphical interface.
- List subfolders and files, classifying them as *File* or *Folder*.
- Export results to Excel (.xlsx) or CSV.

---

## Setup

### 1. Create and activate a virtual environment

```bash
python -m venv .venv
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the application
```bash
python main.py
```

### 4. Build the executable (.exe)
```bash
pyinstaller --onefile --noconsole main.py
```

---

## Example Output

When you select a folder, the program generates an Excel file named:

```bash
FolderTOexcel.xlsx
```

containing the folder structure with the following columns:
- **Carpeta** → Parent folder name
- **Nombre** → File or subfolder name
- **Tipo** → "Carpeta" (Folder) or "Archivo" (File)




