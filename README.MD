# 📐 DXF - CEVASA

Sistema completo para visualização, medição e geração automática de planilhas e relatórios a partir de arquivos DXF, voltado para projetos de sistematização agrícola (cana-de-açúcar).

---

## ✅ Funcionalidades

- 📁 Interface intuitiva para seleção de arquivos `.dxf`
- 🗺️ Visualização interativa do mapa DXF com filtros de layers
- 📏 Ferramenta de medição direta no mapa
- 🧾 Geração automática de:
  - Tabela de **comprimentos por layer**
  - Tabela de **talhões**
  - Legenda com cores e nomes dos layers
- 🖼️ Inserção automática da imagem do mapa
- 📄 Geração de planilha Excel + PDF
- 🧩 Ícone personalizado
- 🔒 Interface limpa, sem terminal aberto na versão empacotada

---

## 🖼️ Imagem do Sistema

<p align="center">
  <img src="docs/interface_preview.png" alt="Preview do Sistema" width="600">
</p>

---

## 🚀 Como Executar (Versão Empacotada)

1. Clique no executável: `DXF-CEVASA.exe`
2. Selecione o arquivo `.dxf`
3. Interaja com o mapa, clique em "Salvar Figura"
4. Preencha os dados solicitados
5. Sua planilha será gerada em `resources/excel/Planilha_Final.xlsx`

---

## 🔧 Tecnologias Utilizadas

- Python 3.13
- `ezdxf`, `matplotlib`, `tkinter`, `openpyxl`, `Pillow`, `win32com`, `PyInstaller`

---

## 📁 Estrutura do Projeto

DXF-CEVASA/ 
├── main.py 
├── src/ 
│ ├── app_interface.py 
│ ├── gui.py 
│ ├── dxf_loader.py 
│ ├── dxf_parser.py 
│ ├── layout_generator.py 
│ ├── talhoes_parser.py 
│ └── layer_selector.py 
├── resources/ 
│ └── excel/ 
│ ├── Planilha_template.xlsx 
│ └── Planilha_Final.xlsx 
├── output/ 
│ └── mapa.png 
└── README.md

---

## 👨‍💻 Autor

**Diego Murari**  
🔗 [linkedin.com/in/diegomurari](https://linkedin.com/in/diegomurari)
