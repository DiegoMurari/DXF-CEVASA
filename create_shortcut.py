import os
import sys
import win32com.client

def criar_atalho(nome_exe='DXF-CEVASA.exe', nome_atalho='DXF CEVASA', icone='icon.ico'):
    desktop = os.path.join(os.environ["USERPROFILE"], "Desktop")
    caminho_atalho = os.path.join(desktop, f"{nome_atalho}.lnk")
    
    # Verifica se o atalho já existe
    if os.path.exists(caminho_atalho):
        return

    shell = win32com.client.Dispatch("WScript.Shell")
    atalho = shell.CreateShortCut(caminho_atalho)
    atalho.TargetPath = os.path.join(os.getcwd(), nome_exe)
    atalho.WorkingDirectory = os.getcwd()
    atalho.IconLocation = os.path.join(os.getcwd(), icone)
    atalho.save()
    print(f"✅ Atalho criado: {caminho_atalho}")

# Executa se for chamado diretamente
if __name__ == '__main__':
    criar_atalho()
