# DXF DIEGO Project

Projeto para visualização, processamento e conversão de arquivos DXF.  
O sistema permite carregar um arquivo DXF, visualizar seu conteúdo, selecionar layers, medir distâncias e gerar layouts finais (incluindo planilhas e conversão para PDF).

## Funcionalidades

- **Carregamento e Parsing de DXF:**  
  Importa e interpreta arquivos DXF, extraindo entidades gráficas.

- **Visualização:**  
  Exibe o desenho com interface gráfica interativa, permitindo medições e seleções.

- **Geração de Layout:**  
  Processa os dados extraídos e gera planilhas (Excel) e imagens com o layout final.

- **Conversão para PDF:**  
  (Se aplicável) Converte o resultado para formato PDF.

## Estrutura do Projeto

- `main.py`: Ponto de entrada do programa.
- `src/`: Contém os módulos-fonte do projeto.
- `resources/`: Armazena imagens, arquivos de exemplo e planilhas.
- `output/`: Pasta para salvar as saídas geradas pelo programa.
