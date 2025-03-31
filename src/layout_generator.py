import os
import tkinter as tk
from tkinter import simpledialog, messagebox
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from PIL import Image
from datetime import datetime
from openpyxl.styles import Font
from openpyxl.worksheet.table import Table, TableStyleInfo
from .layer_selector import selecionar_layers
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from src.talhoes_parser import extrair_talhoes_por_proximidade,extrair_legenda_layers
import win32com.client as win32

MAX_DESENHISTA = 24  # Limite máximo para o nome do DESENHISTA

def ask_limited_string(prompt, limit=MAX_DESENHISTA):
    def validate_input(P):
        return len(P) <= limit

    dlg = tk.Toplevel()
    dlg.title(prompt)
    dlg.resizable(False, False)
    dlg.geometry("350x150+400+250")

    tk.Label(dlg, text=f"{prompt} (máx {limit} caracteres)").pack(padx=10, pady=5)

    var = tk.StringVar()
    entry = tk.Entry(dlg, textvariable=var, validate='key',
                     validatecommand=(dlg.register(validate_input), '%P'))
    entry.pack(padx=10, pady=5)
    entry.focus()

    result = None

    def on_ok():
        nonlocal result
        result = var.get().strip().upper()
        if not result:
            messagebox.showerror("Erro", "Campo obrigatório.", parent=dlg)
        else:
            dlg.destroy()

    tk.Button(dlg, text="OK", command=on_ok).pack(pady=10)
    dlg.grab_set()
    dlg.wait_window()

    return result
def excel_to_pdf(excel_path, pdf_path):
    """
    Converte um arquivo Excel para PDF utilizando o Microsoft Excel via COM.
    Requer que o Excel esteja instalado no sistema (funciona apenas no Windows).
    """
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        # Abre o workbook
        workbook = excel.Workbooks.Open(os.path.abspath(excel_path))
        # Exporta como PDF (0 indica PDF)
        workbook.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
        workbook.Close(False)
        print(f"✅ Planilha convertida para PDF: {pdf_path}")
    except Exception as e:
        print(f"❌ Erro ao converter Excel para PDF: {e}")
    finally:
        excel.Quit()

def set_cell_value(ws, cell_coord, value):
    for merged_range in ws.merged_cells.ranges:
        if cell_coord in merged_range:
            anchor = merged_range.start_cell.coordinate
            ws[anchor].value = value
            return
    ws[cell_coord].value = value

def adicionar_legenda_layers(ws, legenda_layers, start_row=1, start_col=1):
    """
    Gera uma legenda baseada no dicionário {layer_name: {"color": (r, g, b)}}.
    Cada linha terá:
      - 1 célula com cor de fundo (PatternFill)
      - 1 célula com o nome do layer
    """
    # Exemplo de título "PROJETO DE SISTEMATIZAÇÃO"
    titulo = "PROJETO DE SISTEMATIZAÇÃO"
    titulo_cell = ws.cell(row=start_row, column=start_col, value=titulo)
    titulo_cell.font = Font(bold=True, size=14)
    ws.merge_cells(
        start_row=start_row, start_column=start_col,
        end_row=start_row, end_column=start_col + 2
    )
    titulo_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Deixe uma linha em branco após o título
    row = start_row + 2

    for layer_name, info in legenda_layers.items():
        color_floats = info["color"]  # Ex.: (0.0, 1.0, 0.0) para verde
        r_float, g_float, b_float = color_floats
        # Converter floats [0..1] em RGB hex (ex.: "FF00FF")
        r = int(r_float * 255)
        g = int(g_float * 255)
        b = int(b_float * 255)
        color_hex = f"{r:02X}{g:02X}{b:02X}"

        # Celula colorida
        color_cell = ws.cell(row=row, column=start_col)
        fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
        color_cell.fill = fill

        # Celula com o nome do layer ao lado
        name_cell = ws.cell(row=row, column=start_col + 1, value=layer_name)
        name_cell.alignment = Alignment(horizontal="left", vertical="center")

        row += 1

def redimensionar_imagem(imagem_path, largura, altura):
    try:
        with Image.open(imagem_path) as img:
            resized_img = img.resize((largura, altura), Image.LANCZOS)
            resized_img.save(imagem_path)
            print("✅ Imagem redimensionada para:", resized_img.size)
    except Exception as e:
        print(f"❌ Erro ao redimensionar imagem: {e}")

def centralizar_imagem_na_planilha(ws, imagem_path, cell_coord="E20"):
    if not os.path.exists(imagem_path):
        print("❌ Imagem do mapa não encontrada.")
        return
    try:
        img = XLImage(imagem_path)
        cell = ws[cell_coord]
        col_letter = get_column_letter(cell.column)
        row_num = cell.row
        img.anchor = f"{col_letter}{row_num}"
        ws.add_image(img)
        print("✅ Imagem inserida na planilha na célula", cell_coord)
    except Exception as e:
        print(f"❌ Erro ao inserir imagem na planilha: {e}")

def adicionar_tabela_comprimentos_custom(ws, layer_data, start_row=1, start_col=1):
    """
    Cria uma tabela de "COMPRIMENTOS POR LAYER" com estilo e bordas finas,
    semelhante ao modelo fornecido, e garante que a célula onde os nomes dos layers 
    são inseridos não altere seu tamanho.
    
    :param ws: Worksheet onde a tabela será criada
    :param layer_data: Dicionário {layer: {"qtd": int, "total": float}}
    :param start_row: Linha inicial da tabela
    :param start_col: Coluna inicial da tabela
    """
    # --------------------------
    #   Configurações de estilo
    # --------------------------
    title_font = Font(bold=True, size=12)
    header_font = Font(bold=True, size=10)
    cell_font = Font(size=10)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=False)
    left_alignment = Alignment(horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=False)
    
    # --------------------------
    #   Montar título mesclado
    # --------------------------
    title_cell = ws.cell(row=start_row, column=start_col, value="COMPRIMENTOS POR LAYER")
    title_cell.font = title_font
    title_cell.alignment = center_alignment
    ws.merge_cells(
        start_row=start_row, start_column=start_col,
        end_row=start_row, end_column=start_col + 3
    )
    
    # --------------------------
    #   Cabeçalho
    # --------------------------
    headers = ["NOME DO LAYER", "QTD", "TOTAL (m)", "MÉDIA (m)"]
    header_row = start_row + 1
    for i, header_text in enumerate(headers):
        cell = ws.cell(row=header_row, column=start_col + i, value=header_text)
        cell.font = header_font
        cell.alignment = center_alignment
        cell.border = thin_border
    
    # --------------------------
    #   Inserir dados
    # --------------------------
    data_start_row = header_row + 1
    current_row = data_start_row
    
    for layer, data in layer_data.items():
        qtd = data.get("qtd", 0)
        total_m = data.get("total", 0.0)
        media_m = total_m / qtd if qtd > 0 else 0.0
        
        # Coluna 1: NOME DO LAYER
        cell = ws.cell(row=current_row, column=start_col, value=layer)
        cell.font = cell_font
        # Aqui, usamos left_alignment com wrap_text=False para manter o tamanho fixo
        cell.alignment = left_alignment
        cell.border = thin_border
        
        # Coluna 2: QTD
        cell = ws.cell(row=current_row, column=start_col + 1, value=qtd)
        cell.font = cell_font
        cell.alignment = center_alignment
        cell.border = thin_border
        
        # Coluna 3: TOTAL (m)
        cell = ws.cell(row=current_row, column=start_col + 2, value=round(total_m, 2))
        cell.font = cell_font
        cell.alignment = center_alignment
        cell.border = thin_border
        
        # Coluna 4: MÉDIA (m)
        cell = ws.cell(row=current_row, column=start_col + 3, value=round(media_m, 2))
        cell.font = cell_font
        cell.alignment = center_alignment
        cell.border = thin_border
        
        current_row += 1

def parse_talhao_layer_name(layer_name):
    """
    Recebe algo como '06.11.14' e retorna ('06', 11.14).
    Se não houver ponto ou não for possível converter a área, 
    retorna (layer_name, 0.0).
    """
    parts = layer_name.split('.', 1)  # Divide em 2 partes no primeiro ponto
    if len(parts) == 2:
        numero_str, area_str = parts
        numero_str = numero_str.strip()
        try:
            area_ha = float(area_str)
        except ValueError:
            area_ha = 0.0
    else:
        # Se não houver ponto, ou não der para converter, 
        # assume o layer_name inteiro como número e área = 0
        numero_str = layer_name
        area_ha = 0.0

    return numero_str, area_ha

def adicionar_tabela_talhoes_custom(ws, talhoes_dict, start_row=1, start_col=1):
    """
    Cria uma tabela "TALHÕES" sem a coluna de %,
    exibindo:
      - Número
      - Área (ha)
      - Área (alq)*

    E uma linha TOTAL em vermelho.
    """
    from openpyxl.styles import Alignment, Font, Border, Side
    from openpyxl.utils import get_column_letter

    title_font = Font(bold=True, size=12)
    header_font = Font(bold=True, size=10)
    cell_font = Font(size=10)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)

    # Título
    ws.cell(row=start_row, column=start_col, value="TALHÕES").font = title_font
    ws.merge_cells(
        start_row=start_row, start_column=start_col,
        end_row=start_row, end_column=start_col + 2
    )
    ws.cell(row=start_row, column=start_col).alignment = center_alignment

    # Cabeçalho
    headers = ["Número", "Área (ha)", "Área (alq)*"]
    for i, header_text in enumerate(headers):
        c = ws.cell(row=start_row+1, column=start_col + i, value=header_text)
        c.font = header_font
        c.alignment = center_alignment
        c.border = thin_border

    # Inserir dados
    total_ha = 0.0
    total_alq = 0.0
    row_data_start = start_row + 2

    current_row = row_data_start
    for numero, area_ha in talhoes_dict.items():
        area_alq = area_ha / 2.42  # Ajuste se quiser outro fator

        # Número
        c = ws.cell(row=current_row, column=start_col, value=numero)
        c.font = cell_font
        c.alignment = center_alignment
        c.border = thin_border

        # Área (ha)
        c = ws.cell(row=current_row, column=start_col+1, value=round(area_ha, 2))
        c.font = cell_font
        c.alignment = center_alignment
        c.border = thin_border

        # Área (alq)*
        c = ws.cell(row=current_row, column=start_col+2, value=round(area_alq, 2))
        c.font = cell_font
        c.alignment = center_alignment
        c.border = thin_border

        total_ha += area_ha
        total_alq += area_alq
        current_row += 1

    # Linha TOTAL
    c = ws.cell(row=current_row, column=start_col, value="TOTAL")
    c.font = Font(bold=True, color="FF0000")  # vermelho
    c.alignment = center_alignment
    c.border = thin_border

    c = ws.cell(row=current_row, column=start_col+1, value=round(total_ha, 2))
    c.font = Font(bold=True)
    c.alignment = center_alignment
    c.border = thin_border

    c = ws.cell(row=current_row, column=start_col+2, value=round(total_alq, 2))
    c.font = Font(bold=True)
    c.alignment = center_alignment
    c.border = thin_border

    # Observação "*Alqueires Paulistas"
    ws.cell(row=current_row+1, column=start_col+2, value="*Alqueires Paulistas").alignment = Alignment(horizontal="right")

    # Ajuste de largura (remova se quiser manter o template)
    ws.column_dimensions[get_column_letter(start_col)].width = 10
    ws.column_dimensions[get_column_letter(start_col+1)].width = 12
    ws.column_dimensions[get_column_letter(start_col+2)].width = 12

def gerar_layout_final(dxf_file_path, layer_data, talhoes_dict, legenda_layers):
    """
    Gera a planilha final com:
      - Mapa na aba 'Pagina1'
      - Tabelas (comprimentos e talhões) na aba 'Pagina2'
      - Informações do rodapé em ambas as abas
      - Permite que o usuário selecione quais layers incluir nas tabelas
    """
    # Solicita entradas do usuário
    desenhista = ask_limited_string("Informe o nome do DESENHISTA")
    if desenhista is None:
        print("Operação cancelada.")
        return
    escala = simpledialog.askstring("Input", "Informe a ESCALA:").strip().upper()
    distancia = simpledialog.askstring("Input", "Informe a DISTÂNCIA:").strip().upper()
    area_cana = simpledialog.askstring("Input", "Informe a ÁREA CANA (ha):").strip().upper()
    prev = simpledialog.askstring("Input", "Informe a VERSÃO ANTERIOR (ou deixe em branco para 0.0):")
    mun_est = simpledialog.askstring("Input", "Informe MUN. EST:").strip().upper()
    parc = simpledialog.askstring("Input", "Informe PARC OUTORGANTE (opcional):")
    parc = parc.strip().upper() if parc else ""
    
    try:
        prev_version = float(prev) if prev and prev.strip() != "" else 0.0
    except ValueError:
        prev_version = 0.0
    nova_versao = round(prev_version + 0.1, 1)
    
    data_atual = datetime.now().strftime("%d/%m/%Y")
    propriedade = os.path.splitext(os.path.basename(dxf_file_path))[0].upper()
    
    # Obter as entidades do DXF usando load_dxf e parse_dxf
    from .dxf_loader import load_dxf
    from .dxf_parser import parse_dxf
    doc = load_dxf(dxf_file_path)
    entities = parse_dxf(doc)
    
    # Permitir que o usuário selecione os layers para a tabela de comprimentos
    layers_disponiveis = list(layer_data.keys())
    layers_comprimentos = selecionar_layers(layers_disponiveis, "Selecione os layers para a Tabela de Comprimentos")
    
    # Para a tabela de talhões, extraímos automaticamente os dados a partir das entidades do DXF
    talhoes_dict = extrair_talhoes_por_proximidade(entities, distance_threshold=150.0, debug=False)
    
    # Carregar a planilha template
    template_file = os.path.join(os.path.dirname(__file__), '..', 'resources', 'excel', 'Planilha_template.xlsx')
    output_file = os.path.join(os.path.dirname(__file__), '..','resources', 'excel', 'Planilha_Final.xlsx')
    wb = openpyxl.load_workbook(template_file)
    
    if "Pagina1" not in wb.sheetnames or "Pagina2" not in wb.sheetnames:
        print("❌ As abas 'Pagina1' ou 'Pagina2' não foram encontradas no template.")
        return
    
    ws_pagina1 = wb["Pagina1"]  # Aba para inserir o mapa e o rodapé
    ws_pagina2 = wb["Pagina2"]  # Aba para inserir as tabelas e o rodapé
    
    # Inserir informações do rodapé na aba Pagina1
    set_cell_value(ws_pagina1, "I28", desenhista)
    set_cell_value(ws_pagina1, "J29", data_atual)
    set_cell_value(ws_pagina1, "I30", distancia)
    set_cell_value(ws_pagina1, "I31", area_cana)
    set_cell_value(ws_pagina1, "J31", nova_versao)
    set_cell_value(ws_pagina1, "I29", escala)
    set_cell_value(ws_pagina1, "B33", propriedade)
    set_cell_value(ws_pagina1, "E33", mun_est)
    set_cell_value(ws_pagina1, "H33", parc)
    
    # Inserir informações do rodapé na aba Pagina2
    set_cell_value(ws_pagina2, "G28", desenhista)
    set_cell_value(ws_pagina2, "H29", data_atual)
    set_cell_value(ws_pagina2, "G30", distancia)
    set_cell_value(ws_pagina2, "G31", area_cana)
    set_cell_value(ws_pagina2, "H31", nova_versao)
    set_cell_value(ws_pagina2, "G29", escala)
    set_cell_value(ws_pagina2, "B33", propriedade)
    set_cell_value(ws_pagina2, "C33", mun_est)
    set_cell_value(ws_pagina2, "F33", parc)
    
    # Inserir imagem (mapa) somente na aba Pagina1
    if os.path.exists(os.path.join("output", "mapa.png")):
        # use o caminho completo na chamada das funções:
        image_path = os.path.join("output", "mapa.png")
        try:
            redimensionar_imagem(image_path, 800, 575)
            centralizar_imagem_na_planilha(ws_pagina1, image_path, "A02")
            print("✅ Imagem 'mapa.png' adicionada na aba 'Pagina1'.")
        except Exception as e:
            print(f"❌ Erro ao inserir imagem 'mapa.png': {e}")
    else:
        print("❌ Imagem do mapa não foi gerada.")
    

    # Inserir as tabelas na aba Pagina2
    adicionar_tabela_comprimentos_custom(ws_pagina2, layer_data, start_row=2, start_col=2)
    adicionar_tabela_talhoes_custom(ws_pagina2, talhoes_dict, start_row=2, start_col=7)
    adicionar_legenda_layers(ws_pagina1, legenda_layers, start_row=1, start_col=9)

    wb.save(output_file)
    print(f"✅ Planilha salva como '{output_file}'.")