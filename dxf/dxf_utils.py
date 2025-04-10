import ezdxf
from ezdxf.colors import aci2rgb

def get_entity_color(entity, doc):
    """
    Obtém a cor RGB normalizada (0..1) de uma entidade DXF.
    
    :param entity: Entidade DXF
    :param doc: Documento DXF
    :return: Tupla (r, g, b) com valores entre 0 e 1
    """
    try:
        color_aci = entity.dxf.color
        if color_aci is None or color_aci in [0, 256]:
            layer = doc.layers.get(entity.dxf.layer)
            color_aci = layer.color
        r, g, b = aci2rgb(color_aci)
        return (r / 255.0, g / 255.0, b / 255.0)
    except Exception:
        return (0, 0, 0)  # Preto como fallback
