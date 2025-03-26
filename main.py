import os
from src.file_selector import select_dxf_file
from src.gui import launch_gui

if __name__ == '__main__':
    # Permite que o usuário selecione o arquivo DXF
    dxf_file_path = select_dxf_file()
    
    if dxf_file_path:
        launch_gui(dxf_file_path)
    else:
        print("Nenhum arquivo DXF foi selecionado!")