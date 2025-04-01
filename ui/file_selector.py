import tkinter as tk
from tkinter import filedialog

def select_dxf_file():
    """Abre uma janela para selecionar um arquivo DXF e retorna o caminho."""
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal do Tkinter

    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione um arquivo DXF",
        filetypes=[("Arquivos DXF", "*.dxf")],
    )

    return caminho_arquivo if caminho_arquivo else None
