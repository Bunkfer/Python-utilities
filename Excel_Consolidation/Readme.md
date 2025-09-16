# Excel Consolidation

This project provides a simple Python desktop application (with Tkinter) that consolidates multiple Excel/CSV files into a single Consolidado.xlsx report.

The process requires:

1. Selecting a folder → where the input Excel/CSV files are located.

2. Selecting a base file → which contains the reference sheets Equivalencias and Acreedores.

3. The program automatically processes all files in the folder, merges the data, enriches it with information from the base file, and generates a final consolidated Excel file.

It is designed with Object-Oriented Programming (OOP) and structured to keep the code reusable and easy to maintain.

## Features
- Graphical interface (Tkinter) to select folder and base file.

- Consolidates multiple Excel/CSV files into one structured output.

- Includes additional sheets (Equivalencias and Acreedores) from the base file.

- Automatically generates formulas and calculated fields in the consolidated Excel file.

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
python Pago_plan.py
```

### 4. Build the executable (.exe)
```bash
pyinstaller --onefile --noconsole Pago_plan.py
```

---

## Example Output

When you select a folder, the program generates an Excel file named:

```bash
Consolidado.xlsx
```

with the following sheets:

- **Consolidado** → merged data from all Excel/CSV files.
- **Equivalencias** → copied from the base file.
- **Acreedores** → copied from the base file.




