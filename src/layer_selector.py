import tkinter as tk

def selecionar_layers(layers_disponiveis, titulo):
    """
    Abre uma janela para o usuário selecionar os layers desejados.
    Por padrão, todos os layers estarão marcados.
    Retorna uma lista com os layers selecionados.
    """
    def confirmar_selecao():
        nonlocal selecionados
        selecionados = [layer for layer, var in zip(layers_disponiveis, variaveis) if var.get()]
        janela.destroy()

    janela = tk.Toplevel()
    janela.title(titulo)
    janela.geometry("400x400")

    tk.Label(janela, text="Selecione os layers desejados:").pack(pady=10)

    variaveis = []
    for layer in layers_disponiveis:
        var = tk.BooleanVar(value=True)  # Todos selecionados inicialmente
        chk = tk.Checkbutton(janela, text=layer, variable=var)
        chk.pack(anchor='w')
        variaveis.append(var)

    selecionados = []
    tk.Button(janela, text="Confirmar", command=confirmar_selecao).pack(pady=10)

    janela.grab_set()
    janela.wait_window()

    return selecionados