import os
from PIL import Image as PILImage
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import column_index_from_string
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from matplotlib.transforms import Bbox
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
try:
    from openpyxl.drawing.geometry import Ext
except ImportError:
    # Se a importação falhar, define uma classe Ext simples
    class Ext:
        def __init__(self, cx, cy):
            self.cx = cx
            self.cy = cy
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.utils import get_column_letter


def redimensionar_imagem(imagem_path, largura, altura):
    """
    Redimensiona a imagem original para o tamanho desejado (substitui a original).
    """
    try:
        with PILImage.open(imagem_path) as img:
            resized_img = img.resize((largura, altura), PILImage.LANCZOS)
            resized_img.save(imagem_path)
            print(f"✅ Imagem redimensionada para: {resized_img.size}")
    except Exception as e:
        print(f"❌ Erro ao redimensionar imagem: {e}")

def centralizar_imagem_na_planilha(ws, imagem_path, cell_coord="E20"):
    """
    Insere a imagem centralizada na célula definida pela coordenada (ex.: "E20").
    
    A função utiliza valores padrão para a largura e a altura da célula:
      - Largura padrão da célula: 64 pixels
      - Altura padrão da célula: 20 pixels
      
    Esses valores podem ser ajustados conforme a real formatação da sua planilha.
    
    A imagem é posicionada via OneCellAnchor, centralizando-a dentro da célula.
    """
    # Verifica se a imagem existe
    if not os.path.exists(imagem_path):
        print("❌ Imagem não encontrada")
        return

    try:
        # Carrega a imagem do Excel
        img = XLImage(imagem_path)
        # Obtém as dimensões da imagem (em pixels)
        img_width, img_height = img.width, img.height

        # Defina os tamanhos padrão da célula (em pixels)
        # Se você tiver outra forma de obter os valores reais, substitua aqui.
        cell_width = 64  
        cell_height = 20

        # Calcula os offsets (em pixels) para centralizar a imagem na célula
        offset_x = max((cell_width - img_width) / 2, 0)
        offset_y = max((cell_height - img_height) / 2, 0)

        # Converter pixels para EMUs (1 pixel ≃ 9525 EMUs)
        EMU_PER_PIXEL = 9525

        # Cria o marcador de âncora do tipo OneCellAnchor
        marker = AnchorMarker()
        cell = ws[cell_coord]
        # openpyxl trata as colunas como 0-indexed para o marcador, enquanto cell.col_idx é 1-indexed.
        marker.col = cell.col_idx - 1  
        marker.row = cell.row - 1
        marker.colOff = int(offset_x * EMU_PER_PIXEL)
        marker.rowOff = int(offset_y * EMU_PER_PIXEL)

        # Define a extensão da imagem (tamanho em EMU)
        ext = Ext(cx=int(img_width * EMU_PER_PIXEL), cy=int(img_height * EMU_PER_PIXEL))

        # Cria o objeto OneCellAnchor e associa à imagem
        anchor = OneCellAnchor(_from=marker, ext=ext)
        img.anchor = anchor

        ws.add_image(img)
        print(f"✅ Imagem inserida centralizada na célula {cell_coord}")
    except Exception as e:
        print(f"❌ Erro ao inserir imagem centralizada: {e}")

def gerar_imagem_centrada(imagem_path, nova_largura, nova_altura, output_path):
    """
    Gera imagem centralizada com transparência, com correção visual para o Excel.
    """
    try:
        img = PILImage.open(imagem_path).convert("RGBA")
        bg = PILImage.new("RGBA", (nova_largura, nova_altura), (255, 255, 255, 0))

        offset_x = (nova_largura - img.width) // 2
        offset_y = max(0, (nova_altura - img.height) // 2 - 2)  # ← ajuste fino

        bg.paste(img, (offset_x, offset_y), mask=img)
        bg.save(output_path, format="PNG")
        print(f"✅ Imagem com padding e transparência salva: {output_path}")
    except Exception as e:
        print(f"❌ Erro ao gerar imagem centralizada: {e}")

def inserir_imagem(ws, imagem_path, cell_anchor):
    """
    Insere uma imagem do Excel com ancoragem em uma célula (ex: 'K28').
    """
    try:
        img = XLImage(imagem_path)
        img.anchor = cell_anchor
        ws.add_image(img)
        print(f"✅ Imagem inserida na célula: {cell_anchor}")
    except Exception as e:
        print(f"❌ Erro ao inserir imagem: {e}")

def salvar_mapa_como_png(fig, viewport_ax, output_path="output/mapa.png", dpi=150, padding_factor=0.1):
    try:
        # Guarda os limites originais
        xlim = viewport_ax.get_xlim()
        ylim = viewport_ax.get_ylim()

        # Calcula novo zoom out (aplicando padding)
        x_range = xlim[1] - xlim[0]
        y_range = ylim[1] - ylim[0]
        new_xlim = (xlim[0] - x_range * padding_factor, xlim[1] + x_range * padding_factor)
        new_ylim = (ylim[0] - y_range * padding_factor, ylim[1] + y_range * padding_factor)
        viewport_ax.set_xlim(new_xlim)
        viewport_ax.set_ylim(new_ylim)
        fig.canvas.draw()

        # (Opcional) Desativa spines e grid para o salvamento
        for spine in viewport_ax.spines.values():
            spine.set_visible(False)
        viewport_ax.grid(False)
        fig.canvas.draw()

        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        # Calcula o extent original (em polegadas)
        from matplotlib.transforms import Bbox
        extent = viewport_ax.get_window_extent().transformed(fig.dpi_scale_trans.inverted())
        left, bottom, right, top = extent.extents

        # Converte 4 pixels para polegadas
        offset_in_inches = 4 / dpi

        # Cria um novo bbox removendo 4 pixels de cada lado
        extent_adjusted = Bbox.from_extents(left + offset_in_inches, bottom + offset_in_inches,
                                            right - offset_in_inches, top - offset_in_inches)

        # Salva a figura usando o extent ajustado
        fig.savefig(output_path, dpi=dpi, bbox_inches=extent_adjusted, pad_inches=0.0,
                    facecolor='white', edgecolor='none')
        print(f"✅ Mapa salvo com sucesso em: {output_path}")

        # Restaura os limites originais e reativa o grid se necessário
        viewport_ax.set_xlim(xlim)
        viewport_ax.set_ylim(ylim)
        fig.canvas.draw()
        for spine in viewport_ax.spines.values():
            spine.set_visible(True)
        viewport_ax.grid(True, linestyle='--', color='gray', alpha=0.3)
        fig.canvas.draw()
    except Exception as e:
        print(f"❌ Erro ao salvar o mapa visível: {e}")