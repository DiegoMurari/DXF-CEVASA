
# 📐 DXF - CEVASA

A complete system for visualizing, measuring, and automatically generating spreadsheets and reports from DXF files, focused on agricultural systematization projects (specifically sugarcane).

---

## ✅ Features

- 📁 User-friendly interface to select `.dxf` files
- 🗺️ Interactive DXF map viewer with layer filters
- 📏 On-map distance measurement tool
- 🧾 Automatic generation of:
  - **Line length by layer** table
  - **Field (talhão)** area table
  - Legend with colors and layer names
- 🖼️ Automatic insertion of the map image
- 📄 Excel + PDF report generation
- 🧩 Custom executable icon
- 🔒 Clean interface (no terminal on packaged version)

---

## 🖥️ Initial Interface

<p align="center">
  <img src="docs/interface_inicial.png" alt="Initial Interface" width="300">
</p>

Start screen of the system. Click **“Selecionar DXF”** to choose the DXF file to be processed.

---

## 🗺️ DXF Map Viewer

<p align="center">
  <img src="docs/visualizacao_dxf.png" alt="DXF Viewer" width="800">
</p>

The system automatically renders the map extracted from the DXF file, preserving colors, texts, and geometry.  
On the right panel, you have buttons to reset the view, measure distances, and save the final map.

---

## ✅ Layer Selection

<p align="center">
  <img src="docs/selecionar_layers.png" alt="Layer Selection" width="400">
</p>

Before generating the spreadsheet, the system allows you to select **which layers will be used for the length calculation**.

---

## 📄 Map and Legend in the Spreadsheet

<p align="center">
  <img src="docs/layout_excel_mapa.png" alt="Excel Layout with Map" width="700">
</p>

The generated spreadsheet includes:

- DXF map with geospatial layout
- Automatically generated legend based on layer colors
- Auto-filled fields:
  - **Current date**
  - **Version number** (automatically incremented)
  - Property name (from DXF filename)
  - Cane Area, Scale, and Distance (entered by user)

---

## 📊 Lengths and Field Area Tables

<p align="center">
  <img src="docs/planilha_tabelas.png" alt="Tables" width="700">
</p>

Based on the lines and texts from the DXF, the system calculates:

- **Line lengths per layer** (count, total and average)
- **Field (talhão) areas**, in hectares and alqueires
- **Total cultivated area**

> 🔄 Tables are generated automatically based on visible layers and proximity of text annotations.

---
