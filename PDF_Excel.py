import tkinter as tk
from tkinter import filedialog, ttk
import tabula
import pandas as pd
from threading import Thread

def pdf_to_excel(pdf_path):
    # Extrai tabelas do PDF
    dfs = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
    # Define o caminho do arquivo de saída com base no caminho do PDF
    output_excel_path = pdf_path.replace('.pdf', '_converted.xlsx')
    
    # Criar um escritor pandas Excel
    writer = pd.ExcelWriter(output_excel_path, engine='openpyxl')
    
    # Escrever cada DataFrame em uma aba separada e atualizar a barra de progresso
    total = len(dfs)
    for i, df in enumerate(dfs):
        df.to_excel(writer, sheet_name=f'Sheet{i+1}')
        # Atualiza a barra de progresso
        progress_var.set((i + 1) / total * 100)
        root.update_idletasks()
    writer.save()
    label.config(text=f"Arquivo convertido e salvo como: {output_excel_path}")
    progress_var.set(0)  # Resetar a barra de progresso

def open_file():
    filepath = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if filepath:
        # Roda a conversão em uma thread separada para evitar que a GUI congele
        Thread(target=pdf_to_excel, args=(filepath,)).start()

# Configuração da GUI
root = tk.Tk()
root.title("Conversor PDF para Excel")

frame = tk.Frame(root)
frame.pack(pady=20)

# Botão para selecionar o PDF
button = tk.Button(frame, text="Selecionar PDF", command=open_file)
button.pack(side=tk.LEFT, padx=10)

# Barra de progresso
progress_var = tk.DoubleVar()  # aqui a variável que controla a barra
progress_bar = ttk.Progressbar(frame, variable=progress_var, maximum=100)
progress_bar.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

# Label para mostrar o status da conversão
label = tk.Label(frame, text="")
label.pack(side=tk.LEFT, padx=10)

root.mainloop()