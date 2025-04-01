import matplotlib
import sys
matplotlib.use('TkAgg')  # Força o backend interativo
import matplotlib.pyplot as plt
from matplotlib.widgets import Button, CheckButtons
from matplotlib.patches import Arc, Ellipse
import matplotlib.gridspec as gridspec
import math
import re
import os
import tkinter as tk
from datetime import datetime
from dxf.dxf_loader import load_dxf
from dxf.dxf_parser import parse_dxf, calcular_tabelas
from ui.layout_generator import gerar_layout_final
from ui.talhoes_parser import extrair_talhoes_por_proximidade, extrair_legenda_layers
from .layout_generator import gerar_layout_final
from .talhoes_parser import extrair_talhoes_por_proximidade, extrair_legenda_layers
from matplotlib.patches import FancyBboxPatch
import tkinter as tk
# Variáveis globais para medição e controle
measurement_mode = False
measurement_points = []
current_doc = None
visible_layers = None
viewport_ax = None
fig = None
measure_button = None
dxf_file_path = None  # Variável global para armazenar o caminho do DXF

def desenhar_botao_arredondado(ax, label, callback):
    # Cria um patch arredondado que ocupará todo o eixo
    bbox = FancyBboxPatch(
        (0, 0), 1, 1,
        boxstyle="round,pad=0.02,rounding_size=0.1",  # Ajuste rounding_size conforme necessário
        transform=ax.transAxes,
        facecolor='#4CAF50',
        edgecolor='none'
    )
    # Adiciona o patch no fundo do eixo
    ax.add_patch(bbox)
    # Cria o botão com cor 'none' para que o fundo seja visto
    btn = Button(ax, label, color='none', hovercolor='#45a049')
    btn.label.set_color("white")
    btn.label.set_fontsize(10)
    btn.on_clicked(callback)
    return btn

def get_output_dir():
    """Retorna o caminho correto da pasta 'output' na raiz do projeto, mesmo quando chamado de dentro do src/."""
    if getattr(sys, 'frozen', False):
        # Empacotado (PyInstaller)
        base_path = sys._MEIPASS
    else:
        # Caminho normal (ex: C:/Users/Usuario/DXF-CEVASA/src)
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

    output_path = os.path.join(base_path, 'output')
    return output_path

def setup_plot(ax):
    ax.set_facecolor('white')
    ax.grid(True, linestyle='--', color='gray', alpha=0.3)
    ax.set_aspect('equal', adjustable='box')

def draw_dxf(ax, dxf_entities, visible_layers=None):
    ax.cla()
    setup_plot(ax)
    for entity in dxf_entities:
        if visible_layers and entity.get("layer") not in visible_layers:
            continue
        etype = entity.get("type")
        color = entity.get("color", (0, 0, 0))
        if etype == "LINE":
            x1, y1 = entity["start"]
            x2, y2 = entity["end"]
            ax.plot([x1, x2], [y1, y2], color=color, linewidth=1)
        elif etype == "CIRCLE":
            center = entity["center"]
            radius = entity["radius"]
            circle = plt.Circle(center, radius, edgecolor=color, facecolor='none', linewidth=1)
            ax.add_patch(circle)
        elif etype == "ARC":
            center = entity["center"]
            radius = entity["radius"]
            start_angle = entity["start_angle"]
            end_angle = entity["end_angle"]
            arc = Arc(center, 2 * radius, 2 * radius, theta1=start_angle, theta2=end_angle, edgecolor=color, linewidth=1)
            ax.add_patch(arc)
        elif etype == "POLYLINE":
            pts = entity["points"]
            if len(pts) > 1:
                xs, ys = zip(*pts)
                ax.plot(xs, ys, color=color, linewidth=1)
        elif etype == "ELLIPSE":
            center = entity["center"]
            width = entity["width"]
            height = entity["height"]
            angle = entity["angle"]
            ellipse = Ellipse(center, width, height, angle=angle, edgecolor=color, facecolor='none', linewidth=1)
            ax.add_patch(ellipse)
        elif etype == "LEADER":
            pts = entity["points"]
            if len(pts) > 1:
                xs, ys = zip(*pts)
                ax.plot(xs, ys, color=color, linewidth=1, linestyle='--')
        elif etype == "DIMENSION":
            pts = entity["points"]
            if len(pts) > 1:
                xs, ys = zip(*pts)
                ax.plot(xs, ys, color=color, linewidth=1, linestyle='--')
        elif etype == "SPLINE":
            pts = entity["points"]
            if len(pts) > 1:
                xs, ys = zip(*pts)
                ax.plot(xs, ys, color=color, linewidth=1, linestyle='--')
        elif etype in ["TEXT", "MTEXT", "ATTRIB", "ATTDEF"]:
            pos = entity.get("position", (0, 0))
            txt = entity.get("text", "")
            rot = entity.get("rotation", 0)
            text_color = color if color != (1, 1, 1) else (0, 0, 0)
            font_size = entity.get("height", 12) * 0.5 if re.match(r'^\d+(\.\d+)?(\s*ha)?$', txt.strip()) else entity.get("height", 12)
            ax.text(pos[0], pos[1], txt, color=text_color, rotation=rot, fontsize=font_size)
        elif etype == "HATCH":
            pattern = entity.get("pattern", "")
            ax.text(0, 0, f"HATCH: {pattern}", color=color, fontsize=8)
    ax.autoscale(enable=True, axis='both', tight=True)
    plt.draw()

def get_unique_layers(dxf_entities):
    layers = set()
    for entity in dxf_entities:
        layers.add(entity.get("layer", "undefined"))
    return sorted(list(layers))

def reset_view(event, ax, f):
    print("Reset View acionado!")
    ax.relim()
    ax.autoscale_view()
    ax.set_aspect('equal', adjustable='box')
    f.canvas.draw_idle()

def toggle_measurement_mode(event):
    global measurement_mode, measurement_points, measure_button
    measurement_mode = not measurement_mode
    measurement_points = []
    measure_button.label.set_text("Medindo..." if measurement_mode else "Medir Distância")
    fig.canvas.draw_idle()

def on_save_button_clicked(event):
    global dxf_file_path
    print("Botão Salvar Figura pressionado!")

    if not dxf_file_path:
        print("❌ Erro: Arquivo DXF não selecionado.")
        return

    doc = load_dxf(dxf_file_path)
    if doc is None:
        print("❌ Erro ao carregar o DXF.")
        return

    # parse_dxf deve retornar todas as entidades, incluindo TEXT
    entities = parse_dxf(doc)

    # 1) Extrair dicionário { "03": 7.38, "07": 22.51, ... } por proximidade
    talhoes_dict = extrair_talhoes_por_proximidade(entities, distance_threshold=50.0)

    # 2) Calcular tabelas de comprimentos (layer_data) ou outras coisas se precisar
    layer_data, _ = calcular_tabelas(entities)

    # 3) Salvar o mapa como imagem (mapa.png)
    salvar_mapa_como_png()

    # 4) Extrair a legenda dos layers
    legenda_layers = extrair_legenda_layers(entities)

    # 5) Gerar planilha final, passando os 4 parâmetros: dxf_file_path, layer_data, talhoes_dict e legenda_layers
    gerar_layout_final(dxf_file_path, layer_data, talhoes_dict, legenda_layers)


def on_click_measurement(event):
    global measurement_mode, measurement_points, viewport_ax
    if not measurement_mode or event.inaxes != viewport_ax:
        return
    measurement_points.append((event.xdata, event.ydata))
    if len(measurement_points) == 2:
        x1, y1 = measurement_points[0]
        x2, y2 = measurement_points[1]
        distance = math.sqrt((x2 - x1)**2 + (y2 - y1)**2)
        viewport_ax.plot([x1, x2], [y1, y2], color='yellow', linewidth=2)
        mid_x = (x1 + x2) / 2
        mid_y = (y1 + y2) / 2
        viewport_ax.annotate(f"{distance:.2f}", (mid_x, mid_y),
                             color='yellow', fontsize=10,
                             bbox=dict(boxstyle="round", fc="black", ec="yellow", alpha=0.5))
        measurement_points.clear()
        fig.canvas.draw_idle()

def salvar_mapa_como_png():
    global fig, viewport_ax, current_doc, visible_layers
    try:
        print("Iniciando a geração do mapa...")
        if fig and viewport_ax:
            print("Visualização identificada. Gerando 'mapa.png'...")
            if current_doc:
                dxf_data = parse_dxf(current_doc)
                draw_dxf(viewport_ax, dxf_data, visible_layers)
                
                all_x = []
                all_y = []
                for entity in dxf_data:
                    if entity['type'] in ['LINE', 'POLYLINE']:
                        for point in entity.get('points', []):
                            all_x.append(point[0])
                            all_y.append(point[1])
                    elif entity['type'] == 'CIRCLE':
                        all_x.append(entity['center'][0])
                        all_y.append(entity['center'][1])
                
                if not all_x or not all_y:
                    print("❌ Erro: Nenhum ponto detectado para definir os limites do mapa.")
                    return

                # Calcular os limites dos dados
                data_x_min = min(all_x)
                data_x_max = max(all_x)
                data_y_min = min(all_y)
                data_y_max = max(all_y)

                # Obter os limites atuais do eixo
                current_xlim = viewport_ax.get_xlim()
                current_ylim = viewport_ax.get_ylim()

                # Definir um valor de threshold (limite crítico) para margem em unidades de dados
                threshold = 50

                # Verificar se algum lado está muito "apertado"
                if (data_x_min - current_xlim[0] < threshold) or (current_xlim[1] - data_x_max < threshold) \
                   or (data_y_min - current_ylim[0] < threshold) or (current_ylim[1] - data_y_max < threshold):
                    zoom_out_factor = 1.2  # Aumenta os limites em 20%
                else:
                    zoom_out_factor = 1.0

                # Calcular o centro e o range dos dados
                x_center = (data_x_min + data_x_max) / 2
                y_center = (data_y_min + data_y_max) / 2
                x_range = (data_x_max - data_x_min) / 2 * zoom_out_factor
                y_range = (data_y_max - data_y_min) / 2 * zoom_out_factor

                new_xlim = (x_center - x_range, x_center + x_range)
                new_ylim = (y_center - y_range, y_center + y_range)

                # Aplicar os novos limites no eixo
                viewport_ax.set_xlim(new_xlim)
                viewport_ax.set_ylim(new_ylim)
                viewport_ax.set_aspect('equal')

            # Obter a extensão da área do eixo para salvar a figura
            extent = viewport_ax.get_window_extent().transformed(fig.dpi_scale_trans.inverted())
            output_dir = get_output_dir()
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
                print(f"Pasta 'output' criada em: {output_dir}")

            output_path = os.path.join(output_dir, "mapa.png")
            fig.savefig(output_path, dpi=150, bbox_inches=extent, pad_inches=0.1)
            print(f"✅ Mapa salvo com sucesso em: {output_path}")
        else:
            print("❌ Erro: Visualização do mapa não identificada.")
    except Exception as e:
        print(f"❌ Erro ao salvar o mapa como imagem: {e}")

def launch_gui(file_path):
    import tkinter as tk
    global dxf_file_path, current_doc, visible_layers, fig, viewport_ax
    dxf_file_path = file_path
    doc = load_dxf(file_path)
    if doc is None:
        print("❌ Erro ao carregar o DXF.")
        return
    current_doc = doc
    dxf_data = parse_dxf(doc)
    unique_layers = get_unique_layers(dxf_data)
    visible_layers = unique_layers.copy()

    # Obter resolução da tela via Tkinter
    root = tk.Tk()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    root.destroy()

    # Resolução base do design
    base_width = 1280
    base_height = 720

    # Usar no máximo 90% da resolução disponível
    max_width = int(screen_width * 0.9)
    max_height = int(screen_height * 0.9)

    # A janela terá a dimensão base se a tela for maior; senão, será redimensionada para caber
    window_width_px = min(base_width, max_width)
    window_height_px = min(base_height, max_height)

    # Fator de escala: se a tela for menor, o layout é reduzido para caber
    scale_factor = min(window_width_px / base_width, window_height_px / base_height)
    print(f"[DEBUG] Tela: {screen_width}x{screen_height} - Scale factor: {scale_factor:.2f}")

    dpi = 100  # DPI base
    fig = plt.figure(figsize=(window_width_px / dpi, window_height_px / dpi), dpi=dpi)
    manager = plt.get_current_fig_manager()
    try:
        manager.window.wm_geometry(f"{window_width_px}x{window_height_px}+50+50")
    except Exception:
        pass

    # Função auxiliar para escalar valores (fontes, etc.)
    def escalar(valor_base):
        return int(valor_base * scale_factor)

    # Ajustar título e layout
    fig.subplots_adjust(top=0.92)

    # Layout com GridSpec: painel do mapa e painel de controle
    gs = gridspec.GridSpec(ncols=2, nrows=1, width_ratios=[4, 1], wspace=0.1)
    viewport_ax = fig.add_subplot(gs[0, 0])
    setup_plot(viewport_ax)
    draw_dxf(viewport_ax, dxf_data, visible_layers)

    control_ax = fig.add_subplot(gs[0, 1])
    control_ax.axis('off')
    control_pos = control_ax.get_position()

    # Tamanhos dos botões e posicionamento (valores relativos não precisam ser escalados, apenas as fontes)
    btn_width = control_pos.width * 0.8
    btn_height = control_pos.height * 0.08
    btn_x = control_pos.x0 + 0.1 * control_pos.width

    # Botão: Redefinir Visualização
    reset_ax = fig.add_axes([btn_x, control_pos.y0 + 0.8 * control_pos.height, btn_width, btn_height])
    reset_button = desenhar_botao_arredondado(reset_ax, "Redefinir Visualização", lambda event: reset_view(event, viewport_ax, fig))
    reset_button.label.set_fontsize(escalar(11))

    # Botão: Medir Distância
    measure_ax = fig.add_axes([btn_x, control_pos.y0 + 0.68 * control_pos.height, btn_width, btn_height])
    measure_button = desenhar_botao_arredondado(measure_ax, "Medir Distância", toggle_measurement_mode)
    measure_button.label.set_fontsize(escalar(11))

    # Botão: Salvar Figura
    save_ax = fig.add_axes([btn_x, control_pos.y0 + 0.56 * control_pos.height, btn_width, btn_height])
    save_button = desenhar_botao_arredondado(save_ax, "Salvar Figura", on_save_button_clicked)
    save_button.label.set_fontsize(escalar(11))

    # Checkboxes para seleção de layers
    check_ax = fig.add_axes([
        control_pos.x0 - 0.01,
        control_pos.y0 + 0.08,
        control_pos.width * 1.7,
        control_pos.height * 0.4
    ])
    check = CheckButtons(check_ax, unique_layers, [True] * len(unique_layers))
    check_ax.set_facecolor("#333333")
    check_ax.set_alpha(1.0)
    try:
        check._spacing = 0.15
    except Exception:
        pass

    for lbl in check.labels:
        lbl.set_fontsize(escalar(9))
        lbl.set_color("white")
        lbl.set_fontweight("bold")

    def update_layers(label):
        if label in visible_layers:
            visible_layers.remove(label)
        else:
            visible_layers.append(label)
        draw_dxf(viewport_ax, dxf_data, visible_layers)
        fig.canvas.draw_idle()

    check.on_clicked(update_layers)
    fig.canvas.mpl_connect('button_press_event', on_click_measurement)
    fig.canvas.draw()
    plt.show()