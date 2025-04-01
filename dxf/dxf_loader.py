# dxf_loader.py
import ezdxf

def load_dxf(file_path):
    """
    Carrega o arquivo DXF e retorna o objeto do documento.
    """
    try:
        doc = ezdxf.readfile(file_path)
        return doc
    except IOError:
        print("Erro: não foi possível ler o arquivo DXF.")
    except ezdxf.DXFStructureError:
        print("Erro: estrutura do arquivo DXF inválida ou corrompida.")
    return None
