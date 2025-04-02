import os
from PIL import Image as PILImage
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import column_index_from_string
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor

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