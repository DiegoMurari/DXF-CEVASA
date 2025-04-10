import matplotlib
import sys
matplotlib.use('TkAgg')
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
from matplotlib.patches import FancyBboxPatch
from PySide6.QtWidgets import QLineEdit, QFileDialog, QMessageBox, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QCheckBox, QListWidget, QListWidgetItem, QApplication
from PySide6.QtGui import QPixmap
from PySide6.QtCore import Qt
from ui.imagem_utils import salvar_mapa_como_png

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
    ax.grid(False, linestyle='--', color='gray', alpha=0.3)
    ax.set_aspect('equal', adjustable='box')


def draw_dxf(ax, dxf_entities, visible_layers=None):
    ax.cla()
    setup_plot(ax)  # Certifique-se que essa função está definida no seu código

    for entity in dxf_entities:
        if visible_layers and entity.get("layer") not in visible_layers:
            continue

        etype = entity.get("type")
        color = entity.get("color", (0, 0, 0))

        if etype == "LINE":
            x1, y1 = entity["start"][:2]
            x2, y2 = entity["end"][:2]
            ax.plot([x1, x2], [y1, y2], color=color, linewidth=1)

        elif etype == "CIRCLE":
            center = entity["center"][:2]
            radius = entity["radius"]
            circle = plt.Circle(center, radius, edgecolor=color, facecolor='none', linewidth=1)
            ax.add_patch(circle)

        elif etype == "ARC":
            center = entity["center"][:2]
            radius = entity["radius"]
            start_angle = entity["start_angle"]
            end_angle = entity["end_angle"]
            arc = Arc(center, 2 * radius, 2 * radius, theta1=start_angle, theta2=end_angle,
                      edgecolor=color, linewidth=1)
            ax.add_patch(arc)

        elif etype == "POLYLINE":
            pts = [tuple(pt[:2]) for pt in entity.get("points", []) if isinstance(pt, (list, tuple)) and len(pt) >= 2]
            if len(pts) > 1:
                xs, ys = zip(*pts)
                ax.plot(xs, ys, color=color, linewidth=1)

        elif etype == "ELLIPSE":
            center = tuple(entity.get("center", (0, 0)))[:2]
            width = entity.get("width", 1)
            height = entity.get("height", 1)
            angle = entity.get("angle", 0)
            ellipse = Ellipse(center, width, height, angle=angle,
                              edgecolor=color, facecolor='none', linewidth=1)
            ax.add_patch(ellipse)


        elif etype in ["LEADER", "DIMENSION", "SPLINE"]:
            pts = [tuple(pt[:2]) for pt in entity.get("points", []) if isinstance(pt, (list, tuple)) and len(pt) >= 2]
            if len(pts) > 1:
                xs, ys = zip(*pts)
                ax.plot(xs, ys, color=color, linewidth=1, linestyle='--')

        elif etype in ["TEXT", "MTEXT", "ATTRIB", "ATTDEF"]:
            pos = entity.get("position", (0, 0))[:2]
            txt = entity.get("text", "")
            rot = entity.get("rotation", 0)
            font_size = entity.get("height", 12)
            is_area = re.match(r'^\d+(\.\d+)?(\s*ha)?$', txt.strip())
            font_size = font_size * 0.5 if is_area else font_size
            text_color = color if color != (1, 1, 1) else (0, 0, 0)
            ax.text(pos[0], pos[1], txt, color=text_color, rotation=rot, fontsize=font_size)

        elif etype == "HATCH":
            ax.text(0, 0, f"HATCH: {entity.get('pattern', '')}", color=color, fontsize=8)

    ax.autoscale(enable=True, axis='both', tight=True)
    plt.draw()

