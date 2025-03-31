import tkinter as tk
from tkinter import filedialog
from src.gui import launch_gui
import sys

def abrir_seletor_dxf():
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecionar Arquivo DXF",
        filetypes=[("Arquivos DXF", "*.dxf")]
    )
    if caminho_arquivo:
        print(f"Arquivo selecionado: {caminho_arquivo}")
        launch_gui(caminho_arquivo)
    else:
        print("Nenhum arquivo selecionado.")

def iniciar_interface():
    janela = tk.Tk()
    janela.title("DXF CEVASA")
    janela.geometry("400x200")
    janela.configure(bg="#F0F0F0")
    janela.resizable(False, False)

    titulo = tk.Label(janela, text="DXF CEVASA", font=("Arial", 18, "bold"), fg="#4CAF50", bg="#F0F0F0")
    titulo.pack(pady=20)

    botao = tk.Button(
        janela,
        text="Selecionar DXF",
        command=abrir_seletor_dxf,
        font=("Arial", 14),
        bg="#4CAF50",
        fg="white",
        activebackground="#45a049",
        relief="flat",
        padx=20,
        pady=10
    )
    botao.pack()

    janela.mainloop()

    # Criação automática do atalho ao iniciar o app compilado
    if getattr(sys, 'frozen', False):
        try:
            from create_shortcut import criar_atalho
            criar_atalho()
        except Exception as e:
            print(f"❌ Erro ao criar atalho: {e}")

if __name__ == '__main__':
    iniciar_interface()
