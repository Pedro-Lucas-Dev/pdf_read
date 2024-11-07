import pdfplumber
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

def extract_table_data(pdf_path, csv_path):
    extracted_data = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table[1:]:
                    product_info = row[0]
                    unit = row[1]
                    etiquetagem = row[2]
                    preparation_instructions = row[3]

                    pattern = r'Código ML:\s*(.*?)\s*Código universal:\s*(.*?)\s*SKU:\s*(.*?)\n(.*)'
                    match = re.search(pattern, product_info, re.DOTALL)

                    if match:
                        codigo_ml = match.group(1).strip()
                        codigo_universal = match.group(2).strip()
                        sku = match.group(3).strip()
                        descricao = match.group(4).replace("\n", " ").strip()

                        extracted_data.append({
                            'Código ML': codigo_ml,
                            'Código universal': codigo_universal,
                            'SKU': sku,
                            'Descrição': descricao,
                            'Unidade': unit.strip(),
                            'Etiquetagem': etiquetagem.strip(),
                            'Instruções de Preparação': preparation_instructions.strip() if preparation_instructions else ''
                        })

    df = pd.DataFrame(extracted_data)
    df.to_csv(csv_path, index=False, encoding='utf-8')

def process_pdf():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if not pdf_path:
        return

    csv_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if not csv_path:
        return

    loading_label.grid(row=2, column=0, pady=10, sticky="nsew")
    root.update_idletasks()

    try:
        extract_table_data(pdf_path, csv_path)
        messagebox.showinfo("Sucesso", f"Dados extraídos e salvos em {csv_path}")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
    finally:
        loading_label.grid_remove()

root = tk.Tk()
root.title("Conversor de PDF para CSV")

window_width, window_height = 400, 200
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_cordinate = int((screen_width/2) - (window_width/2))
y_cordinate = int((screen_height/2) - (window_height/2))
root.geometry(f"{window_width}x{window_height}+{x_cordinate}+{y_cordinate}")

root.grid_columnconfigure(0, weight=1)
root.grid_rowconfigure(0, weight=1)
root.grid_rowconfigure(1, weight=1)
root.grid_rowconfigure(2, weight=1)

process_button = tk.Button(root, text="Selecionar PDF e Extrair para CSV", command=process_pdf, 
                           font=("Arial", 10, "bold"), bg="green", fg="black", 
                           width=15, height=1,  # Diminuindo largura e altura
                           relief="solid", borderwidth=1, highlightbackground="black", highlightthickness=2)
process_button.grid(row=1, column=0, pady=10, sticky="nsew")

loading_label = ttk.Label(root, text="Processando...", font=("Arial", 12))
loading_label.grid_remove()

root.mainloop()
