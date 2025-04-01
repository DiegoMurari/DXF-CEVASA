
import math
import ezdxf
from collections import defaultdict
from ui.layout_generator import gerar_layout_final  # antes estava errado
from .dxf_utils import get_entity_color  # NOVO
import re

# Mantido para compatibilidade, se necessário
def get_entity_color_original(entity, doc):
    color_aci = entity.dxf.color
    if color_aci is None or color_aci == 256:
        layer = doc.layers.get(entity.dxf.layer)
        color_aci = layer.color
    elif color_aci == 0:
        layer = doc.layers.get(entity.dxf.layer)
        color_aci = layer.color
    r, g, b = ezdxf.colors.aci2rgb(color_aci)
    return (r/255.0, g/255.0, b/255.0)

def parse_entity(entity, doc):
    etype = entity.dxftype()
    layer = entity.dxf.layer
    color = get_entity_color(entity, doc)

    if etype == 'INSERT':
        result = []
        try:
            block_entities = list(entity.virtual_entities())
            for be in block_entities:
                result.extend(parse_entity(be, doc))

            for attrib in entity.attribs:
                result.append({
                    'type': 'TEXT',
                    'text': attrib.dxf.text,
                    'position': tuple(attrib.dxf.insert),
                    'rotation': getattr(attrib.dxf, 'rotation', 0),
                    'height': getattr(attrib.dxf, 'height', 12),
                    'layer': entity.dxf.layer,
                    'color': get_entity_color(entity, doc),
                })
        except Exception as e:
            print(f"[AVISO] Erro ao explodir INSERT: {e}")
        return result

    if etype == 'DIMENSION':
        result = []
        try:
            for sub in entity.virtual_entities():
                result.extend(parse_entity(sub, doc))
        except Exception:
            pass
        return result

    if etype in ['MLEADER', 'LEADER']:
        result = []
        try:
            for sub in entity.virtual_entities():
                result.extend(parse_entity(sub, doc))
        except Exception:
            pass
        return result

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

    if etype == 'LINE':
        start = tuple(entity.dxf.start)
        end = tuple(entity.dxf.end)
        length = math.sqrt((end[0] - start[0])**2 + (end[1] - start[1])**2)
        return [{
            'type': 'LINE', 'start': start, 'end': end,
            'layer': layer, 'color': color, 'length': length
        }]

    if etype == 'CIRCLE':
        return [{
            'type': 'CIRCLE', 'center': tuple(entity.dxf.center), 'radius': entity.dxf.radius,
            'layer': layer, 'color': color
        }]

    if etype == 'ARC':
        return [{
            'type': 'ARC', 'center': tuple(entity.dxf.center), 'radius': entity.dxf.radius,
            'start_angle': entity.dxf.start_angle, 'end_angle': entity.dxf.end_angle,
            'layer': layer, 'color': color
        }]

    if etype in ('LWPOLYLINE', 'POLYLINE'):
        pts = [tuple(pt[:2]) for pt in entity.get_points()] if hasattr(entity, 'get_points') else [tuple(v.dxf.location) for v in entity.vertices()]
        total_length = sum(math.dist(pts[i], pts[i + 1]) for i in range(len(pts) - 1))
        return [{'type': 'POLYLINE', 'points': pts, 'layer': layer, 'color': color, 'length': total_length}]

    if etype == 'ELLIPSE':
        major_axis = entity.dxf.major_axis
        angle = math.degrees(math.atan2(major_axis[1], major_axis[0]))
        major_len = math.hypot(major_axis[0], major_axis[1])
        return [{
            'type': 'ELLIPSE', 'center': tuple(entity.dxf.center),
            'width': major_len * 2, 'height': major_len * entity.dxf.ratio * 2,
            'angle': angle, 'layer': layer, 'color': color
        }]

    if etype == 'SPLINE':
        pts = [tuple(p) for p in entity.control_points]
        return [{'type': 'SPLINE', 'points': pts, 'layer': layer, 'color': color}]

    if etype == 'HATCH':
        return [{'type': 'HATCH', 'pattern': entity.dxf.pattern_name, 'layer': layer, 'color': color}]

    if etype == 'SOLID':
        pts = [tuple(p) for p in entity.dxf.points]
        return [{'type': 'SOLID', 'points': pts, 'layer': layer, 'color': color}]

    if etype == '3DFACE':
        pts = [tuple(entity.dxf.get_dxf_attrib(f'vtx{i}')) for i in range(4)]
        return [{'type': '3DFACE', 'points': pts, 'layer': layer, 'color': color}]

    if etype == 'IMAGE':
        return [{'type': 'IMAGE', 'layer': layer, 'color': color, 'info': str(entity)}]

    if etype == 'POINT':
        return [{'type': 'POINT', 'position': tuple(entity.dxf.location), 'layer': layer, 'color': color}]

    if etype == 'XLINE':
        return [{'type': 'XLINE', 'start': tuple(entity.dxf.start), 'unit_dir': tuple(entity.dxf.unit_dir), 'layer': layer, 'color': color}]

    if etype == 'RAY':
        return [{'type': 'RAY', 'start': tuple(entity.dxf.start), 'unit_dir': tuple(entity.dxf.unit_dir), 'layer': layer, 'color': color}]
    
    if etype == 'BLOCK':
        return [{'type': 'BLOCK', 'name': getattr(entity.dxf, "name", "Unnamed"), 'layer': layer, 'color': color, 'raw': str(entity)}]

    if etype == 'TRACE':
        points = [tuple(getattr(entity.dxf, f'vtx{i}', (0, 0))) for i in range(4)]
        return [{'type': 'TRACE', 'points': points, 'layer': layer, 'color': color}]

    if etype in ['MESH', 'REGION', 'SURFACE', '3DSOLID']:
        return [{'type': etype, 'layer': layer, 'color': color, 'raw': str(entity)}]

    return [{'type': etype, 'layer': layer, 'color': color, 'raw': str(entity)}]

def parse_dxf(doc):
    all_entities = []
    msp = doc.modelspace()
    for e in msp:
        all_entities.extend(parse_entity(e, doc))

    print("=== DIAGNÓSTICO DOS TEXTOS ===")
    for entity in all_entities:
        if entity.get("type") in ["TEXT", "MTEXT", "ATTRIB", "ATTDEF"]:
            print(f"Texto encontrado: '{entity.get('text')}' | Posição: {entity.get('position')} | Layer: {entity.get('layer')}")
    print("=== FIM DO DIAGNÓSTICO ===")

    return all_entities

def calcular_tabelas(dxf_entities):
    layer_data = defaultdict(lambda: {'qtd': 0, 'total': 0.0})
    for entity in dxf_entities:
        if 'length' in entity:
            layer_data[entity['layer']]['qtd'] += 1
            layer_data[entity['layer']]['total'] += entity['length']

    talhoes_data = defaultdict(lambda: {'area_ha': 0.0})
    for entity in dxf_entities:
        if entity.get('type') == 'TEXT':
            try:
                area_value = float(entity['text'].replace('ha', '').strip())
                talhoes_data[entity['layer']]['area_ha'] += area_value
            except ValueError:
                pass

    return layer_data, talhoes_data
