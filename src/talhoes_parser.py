# talhoes_parser.py
import re
import math

def extrair_talhoes_por_proximidade(entities, distance_threshold=100.0, debug=False):
    """
    Lê uma lista de entidades de texto do DXF, identifica qual é 'número do talhão'
    e qual é 'área em hectares' (ex.: '7.38 ha', '7.38', '7,38 ha'), e agrupa-as por proximidade.
    
    Retorna um dicionário: {"03": 7.38, "07": 22.51, ...}
    
    :param entities: lista de dicionários, cada um com pelo menos:
        {
          "type": "TEXT" ou "MTEXT",
          "text": "...",
          "position": (x, y, [z])
        }
    :param distance_threshold: distância máxima para considerar que a área pertence ao número
    :param debug: se True, imprime informações de depuração
    """
    # Regex para identificar números (apenas dígitos)
    numero_pattern = re.compile(r"^\d+$")
    
    # Regex para área: permite ponto ou vírgula, pode ter ou não "ha" no final
    area_pattern = re.compile(r"^(\d+(?:[.,]\d+))(?:\s*ha)?$", re.IGNORECASE)

    numero_list = []
    area_list = []

    for ent in entities:
        if ent.get("type") in ["TEXT", "MTEXT"]:
            txt = ent.get("text", "").strip()
            pos = ent.get("position", (0, 0, 0))
            x, y = pos[:2]

            if numero_pattern.match(txt):
                numero_list.append({"numero": txt, "pos": (x, y)})
                if debug:
                    print(f"[DEBUG] Número detectado: '{txt}' em pos=({x:.2f}, {y:.2f})")
            else:
                match_area = area_pattern.match(txt)
                if match_area:
                    area_str = match_area.group(1).replace(',', '.')
                    try:
                        area_val = float(area_str)
                    except ValueError:
                        area_val = 0.0
                    area_list.append({"area": area_val, "pos": (x, y)})
                    if debug:
                        print(f"[DEBUG] Área detectada: '{txt}' => {area_val} ha em pos=({x:.2f}, {y:.2f})")
    
    if debug:
        print(f"[DEBUG] Total de números detectados: {len(numero_list)}")
        print(f"[DEBUG] Total de áreas detectadas: {len(area_list)}")

    talhoes_dict = {}
    for n in numero_list:
        nx, ny = n["pos"]
        best_distance = float("inf")
        best_area_val = None
        
        if debug:
            print(f"[DEBUG] Associando número '{n['numero']}' em ({nx:.2f}, {ny:.2f})")
        
        for a in area_list:
            ax, ay = a["pos"]
            dist = math.hypot(ax - nx, ay - ny)
            if dist < best_distance and dist <= distance_threshold:
                best_distance = dist
                best_area_val = a["area"]
        
        if best_area_val is not None:
            talhoes_dict[n["numero"]] = best_area_val
            if debug:
                print(f"   -> Associado à área: {best_area_val} (dist={best_distance:.2f})")
        else:
            if debug:
                print(f"   -> Nenhuma área encontrada dentro de {distance_threshold} unidades")
    
    return talhoes_dict

def extrair_legenda_layers(dxf_entities):
    """
    Retorna um dicionário {layer_name: {"color": (r, g, b)}},
    baseado nas entidades extraídas do DXF.
    """
    legenda = {}
    for ent in dxf_entities:
        layer_name = ent.get("layer", "Sem Layer")
        color = ent.get("color", (0, 0, 0))  # (r, g, b) em floats, por exemplo
        if layer_name not in legenda:
            legenda[layer_name] = {"color": color}
    return legenda