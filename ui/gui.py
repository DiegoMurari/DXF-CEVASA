import matplotlib
import sys
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from matplotlib.widgets import Button, CheckButtons
from matplotlib.patches import Arc, Ellipse, Polygon, Circle
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


import re
import math
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon, Circle, Arc, Ellipse

def draw_dxf(ax, dxf_entities, visible_layers=None):
    """
    Desenha entidades DXF em ax, sem criar linhas extras entre pontos distantes.
    - Layer “XLEGENDA SISTEMATIZAÇÃO”: triângulo preenchido.
    - Layers contendo “LOMB”: círculos preenchidos.
    - Outras POLYLINE/SOLID: só segmentos consecutivos (nunca fecha).
    - TEXT/MTEXT/ATTRIB/ATTDEF: exatamente como no primeiro código.
    """
    ax.cla()
    setup_plot(ax)

    for ent in dxf_entities:
        layer   = ent.get("layer", "")
        if visible_layers and layer not in visible_layers:
            continue

        etype    = ent.get("type")
        color    = ent.get("color", (0, 0, 0))
        layer_up = layer.upper()
        is_lomb  = "LOMB" in layer_up
        is_xleg  = layer_up == "XLEGENDA SISTEMATIZAÇÃO"

        # LINE
        if etype == "LINE":
            x1,y1 = ent["start"][:2]
            x2,y2 = ent["end"][:2]
            ax.plot([x1,x2],[y1,y2], color=color, linewidth=1)

        # POLYLINE: só segmentos consecutivos, ou triângulo preenchido no XLEGENDA
        elif etype == "POLYLINE":
            pts = [tuple(p[:2]) for p in ent.get("points",[])]
            if is_xleg and len(pts) >= 3:
                tri = Polygon(
                    pts[:3], closed=True,
                    facecolor=color, edgecolor=color, linewidth=1
                )
                ax.add_patch(tri)
            else:
                for i in range(len(pts)-1):
                    x1,y1 = pts[i]
                    x2,y2 = pts[i+1]
                    ax.plot([x1,x2],[y1,y2], color=color, linewidth=1)

        # SOLID: idem POLYLINE, mas respeita fill em XLEGENDA/LOMB
        elif etype == "SOLID":
            pts = [tuple(p) for p in ent.get("points",[])]
            if is_xleg and len(pts)>=3:
                poly = Polygon(pts[:3], closed=True,
                               facecolor=color, edgecolor=color, linewidth=1)
                ax.add_patch(poly)
            else:
                for i in range(len(pts)-1):
                    x1,y1 = pts[i]
                    x2,y2 = pts[i+1]
                    ax.plot([x1,x2],[y1,y2], color=color, linewidth=1)

        # CIRCLE
        elif etype == "CIRCLE":
            cx,cy = ent["center"][:2]
            r     = ent["radius"]
            circ = Circle(
                (cx,cy), r,
                facecolor=color if (is_lomb or is_xleg) else "none",
                edgecolor=color, linewidth=1
            )
            ax.add_patch(circ)

        # ARC
        elif etype == "ARC":
            cx,cy = ent["center"][:2]
            r     = ent["radius"]
            a0,a1 = ent["start_angle"], ent["end_angle"]
            arc = Arc((cx,cy),2*r,2*r, theta1=a0, theta2=a1,
                      edgecolor=color, linewidth=1)
            ax.add_patch(arc)

        # ELLIPSE
        elif etype == "ELLIPSE":
            cx,cy = ent.get("center",(0,0))[:2]
            w     = ent.get("width",1)
            h     = ent.get("height",1)
            ang   = ent.get("angle",0)
            ell = Ellipse((cx,cy),w,h,angle=ang,
                          facecolor=color if (is_lomb or is_xleg) else "none",
                          edgecolor=color, linewidth=1)
            ax.add_patch(ell)

        # SPLINE / LEADER / DIMENSION: só segmentos
        elif etype in ("SPLINE","LEADER","DIMENSION"):
            pts = [tuple(p[:2]) for p in ent.get("points",[])]
            for i in range(len(pts)-1):
                x1,y1 = pts[i]
                x2,y2 = pts[i+1]
                ax.plot([x1,x2],[y1,y2], linestyle="--", color=color, linewidth=1)

        # TEXT / MTEXT / ATTRIB / ATTDEF — **como no seu primeiro código**
        elif etype in ("TEXT","MTEXT","ATTRIB","ATTDEF"):
            pos = ent.get("position",(0,0))[:2]
            txt = ent.get("text","")
            rot = ent.get("rotation",0)
            font_size = ent.get("height",12)
            # reduz fonte para “ha”
            if re.match(r"^\d+(\.\d+)?(\s*ha)?$", txt.strip(), re.IGNORECASE):
                font_size *= 0.5
            # texto branco vira preto
            text_color = color if color != (1,1,1) else (0,0,0)
            ax.text(
                pos[0], pos[1], txt,
                color=text_color,
                rotation=rot,
                fontsize=font_size
            )

        # HATCH (fallback)
        elif etype == "HATCH":
            ax.text(0,0,f"HATCH:{ent.get('pattern','')}",
                    color=color, fontsize=6)

    ax.autoscale(enable=True, axis="both", tight=True)
    plt.draw()