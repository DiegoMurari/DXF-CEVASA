# 📐 DXF - CEVASA

O sistema é uma solução especializada desenvolvida para simplificar e automatizar o processamento de arquivos DXF (formato padrão utilizado por softwares de desenho técnico, como AutoCAD), com o objetivo de gerar relatórios e planilhas detalhadas e precisas para projetos de sistematização agrícola, especialmente para o plantio.

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

<h3 align="center">🎬 Demonstração do Projeto</h3>

<p align="center">
  <a href="https://drive.google.com/file/d/1XvG21EYv-gb0cMijzg9xmujGQz_Yirt5/view?usp=sharing">
    <img src="https://img.icons8.com/fluency/240/play-button-circled.png" alt="Clique para assistir à demonstração" />
  </a>
</p>

## 🖥️ Interface Inicial

<p align="center">
  <img src="docs/layout_excel_mapa.png" alt="Interface Inicial" width="300">
</p>

Tela de entrada do sistema. Ao clicar em **“Selecionar DXF”**, você escolhe o arquivo topográfico a ser processado.

---

## 🗺️ Visualização Interativa do DXF

<p align="center">
  <img src="docs/visualizacao_dxf.png" alt="Visualização DXF" width="800">
</p>

O sistema renderiza automaticamente o mapa extraído do DXF, preservando as cores, textos e geometrias.  
Na lateral direita, você tem botões para redefinir visualização, medir distâncias e salvar a figura final.

---

## ✅ Preencher excel

<p align="center">
  <img src="docs/informações_planilha.png" alt="Preencher excel" width="400">
</p>

Apos confirmar os filtros, o sistema apresenta essa janela **para inserir as informações na planilha**.

O campo desenhista sempre vai salvar o ultimo nome inserido para agilizar o preenchimento.

---

## ✅ Filtros por Layer

<p align="center">
  <img src="docs/selecionar_layers.png" alt="Seleção de Layers" width="400">
</p>

Antes de gerar a planilha, o sistema permite selecionar **quais layers devem ser incluídos nos cálculos**.

---

## 📄 Mapa e Legenda Gerados na Planilha

<p align="center">
  <img src="docs/interface_inicial.png" alt="Mapa na Planilha" width="700">
</p>

A planilha gerada inclui:

- Mapa do DXF renderizado
- Legenda automática baseada nos layers utilizados
- Dados preenchidos automaticamente:
  - **Data atual**
  - **Versão (incremental)**
  - Nome da propriedade (arquivo DXF)
  - Área Cana, Escala, Distância (preenchidos pelo usuário)

---

## 📊 Tabelas de Comprimentos e Talhões

<p align="center">
  <img src="docs/planilha_tabelas.png" alt="Tabelas" width="700">
</p>

Com base nas linhas do DXF, o sistema calcula:

- **Comprimento por layer** (quantidade, total e média)
- **Área por talhão**, em hectares e alqueires
- **Soma total da área cultivável**

> 🔄 As tabelas são montadas automaticamente com base nos layers visíveis e textos próximos às geometrias.

---
