import math
import ezdxf
from collections import defaultdict
from .dxf_utils import get_entity_color  # função que retorna (r, g, b) normalizado
import re

def get_entity_color_original(entity, doc):
    """
    Fallback que converte o ACI em RGB normalizado, mantendo compatibilidade.
    """
    color_aci = entity.dxf.color
    if color_aci is None or color_aci in (0, 256):
        layer = doc.layers.get(entity.dxf.layer)
        color_aci = layer.color
    r, g, b = ezdxf.colors.aci2rgb(color_aci)
    return (r / 255.0, g / 255.0, b / 255.0)

def parse_entity(entity, doc):
    """
    Converte cada entidade DXF em um dict padronizado com:
      - type: str ('LINE', 'CIRCLE', 'POLYLINE', 'SOLID', 'TEXT', etc.)
      - atributos específicos (start, end, center, points, text, etc.)
      - layer: nome da camada
      - color: tupla (r, g, b)
    """
    etype = entity.dxftype()
    layer = entity.dxf.layer
    color = get_entity_color(entity, doc)

    # Block INSERT: explodir e herdar layer/cor
    if etype == 'INSERT':
        result = []
        try:
            for sub in entity.virtual_entities():
                sub.dxf.layer = layer
                result.extend(parse_entity(sub, doc))
            for attrib in entity.attribs:
                result.append({
                    'type': 'TEXT',
                    'text': attrib.dxf.text,
                    'position': tuple(attrib.dxf.insert),
                    'rotation': getattr(attrib.dxf, 'rotation', 0),
                    'height': getattr(attrib.dxf, 'height', 12),
                    'layer': layer,
                    'color': color,
                })
        except Exception as e:
            print(f"[AVISO] Erro ao explodir INSERT: {e}")
        return result

    # Textos
    if etype in ('TEXT', 'MTEXT', 'ATTRIB', 'ATTDEF'):
        return [{
            'type': 'TEXT',
            'text': entity.dxf.text if hasattr(entity.dxf, 'text') else entity.text,
            'position': tuple(entity.dxf.insert),
            'rotation': getattr(entity.dxf, 'rotation', 0),
            'height': getattr(entity.dxf, 'height', 12),
            'layer': layer,
            'color': color,
        }]

    # Linhas
    if etype == 'LINE':
        start = tuple(entity.dxf.start)
        end = tuple(entity.dxf.end)
        return [{
            'type': 'LINE',
            'start': start,
            'end': end,
            'layer': layer,
            'color': color,
            'length': math.dist(start, end)
        }]

    # Polilinhas
    if etype in ('LWPOLYLINE', 'POLYLINE'):
        pts = [tuple(pt[:2]) for pt in entity.get_points()] \
              if hasattr(entity, 'get_points') \
              else [tuple(v.dxf.location) for v in entity.vertices()]
        total_length = sum(math.dist(pts[i], pts[i+1]) for i in range(len(pts)-1))
        return [{
            'type': 'POLYLINE',
            'points': pts,
            'layer': layer,
            'color': color,
            'length': total_length
        }]

    # Círculos
    if etype == 'CIRCLE':
        return [{
            'type': 'CIRCLE',
            'center': tuple(entity.dxf.center),
            'radius': entity.dxf.radius,
            'layer': layer,
            'color': color
        }]

    # Solids (triângulos, polígonos 2D preenchidos)
    if etype == 'SOLID':
        pts = [tuple(p) for p in entity.dxf.points]
        return [{
            'type': 'SOLID',
            'points': pts,
            'layer': layer,
            'color': color
        }]

    # 3DFACE (faces 3D com 4 vértices)
    if etype == '3DFACE':
        pts = [tuple(entity.dxf.get_dxf_attrib(f'vtx{i}')) for i in range(4)]
        return [{
            'type': '3DFACE',
            'points': pts,
            'layer': layer,
            'color': color
        }]

    # Arcos
    if etype == 'ARC':
        return [{
            'type': 'ARC',
            'center': tuple(entity.dxf.center),
            'radius': entity.dxf.radius,
            'start_angle': entity.dxf.start_angle,
            'end_angle': entity.dxf.end_angle,
            'layer': layer,
            'color': color
        }]

    # Elipses
    if etype == 'ELLIPSE':
        return [{
            'type': 'ELLIPSE',
            'center': tuple(entity.dxf.center),
            'major_axis': tuple(entity.dxf.major_axis),
            'ratio': entity.dxf.ratio,
            'layer': layer,
            'color': color
        }]

    # Splines
    if etype == 'SPLINE':
        pts = [tuple(p) for p in entity.control_points]
        return [{
            'type': 'SPLINE',
            'points': pts,
            'layer': layer,
            'color': color
        }]

    # Hatches
    if etype == 'HATCH':
        return [{
            'type': 'HATCH',
            'pattern': entity.dxf.pattern_name,
            'layer': layer,
            'color': color
        }]

    # Pontos
    if etype == 'POINT':
        return [{
            'type': 'POINT',
            'position': tuple(entity.dxf.location),
            'layer': layer,
            'color': color
        }]

    # Leaders, dimensões, etc.
    if etype in ('MLEADER', 'LEADER', 'DIMENSION'):
        out = []
        for sub in entity.virtual_entities():
            out.extend(parse_entity(sub, doc))
        return out

    # Qualquer outro tipo
    return [{
        'type': etype,
        'layer': layer,
        'color': color,
        'raw': str(entity)
    }]

def parse_dxf(doc):
    """
    Percorre o modelspace e retorna a lista de todas as entidades processadas.
    Também imprime um diagnóstico dos textos encontrados.
    """
    all_entities = []
    msp = doc.modelspace()
    for ent in msp:
        all_entities.extend(parse_entity(ent, doc))

    print("=== DIAGNÓSTICO DOS TEXTOS ===")
    for e in all_entities:
        if e.get('type') == 'TEXT':
            print(f"Texto encontrado: '{e['text']}' | Posição: {e['position']} | Layer: {e['layer']}")
    print("=== FIM DO DIAGNÓSTICO ===")

    return all_entities

def calcular_tabelas(dxf_entities):
    """
    Retorna dois dicts:
      - layer_data: {layer: {'qtd': n_entidades, 'total': soma_comprimentos}}
      - talhoes_data: {layer: {'area_ha': soma_áreas_lidas_em_TEXT}}
    """
    layer_data = defaultdict(lambda: {'qtd': 0, 'total': 0.0})
    for e in dxf_entities:
        if 'length' in e:
            layer_data[e['layer']]['qtd'] += 1
            layer_data[e['layer']]['total'] += e['length']

    talhoes_data = defaultdict(lambda: {'area_ha': 0.0})
    for e in dxf_entities:
        if e.get('type') == 'TEXT':
            txt = e['text'].replace('ha', '').strip()
            try:
                val = float(txt)
                talhoes_data[e['layer']]['area_ha'] += val
            except ValueError:
                continue

    return layer_data, talhoes_data
