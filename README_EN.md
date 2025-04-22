📐 DXF - CEVASA
The system is a specialized solution developed to simplify and automate the processing of DXF files (a standard format used by technical drawing software such as AutoCAD), aiming to generate detailed and accurate reports and spreadsheets for agricultural systematization projects, especially for planting.

✅ Features
📁 Intuitive interface for selecting .dxf files

🗺️ Interactive DXF map viewer with layer filters

📏 Direct measurement tool on the map

🧾 Automatic generation of:

Length by layer table

Plot division table

Legend with layer names and colors

🖼️ Automatic insertion of the map image

📄 Excel spreadsheet + PDF generation

🧩 Custom icon

🔒 Clean interface, no terminal window in the packaged version

<h3 align="center">🎬 Project Demo</h3> <p align="center"> <a href="https://drive.google.com/file/d/1XvG21EYv-gb0cMijzg9xmujGQz_Yirt5/view?usp=sharing"> <img src="https://img.icons8.com/fluency/240/play-button-circled.png" alt="Click to watch the demo" /> </a> </p>
🖥️ Initial Interface
<p align="center"> <img src="docs/Tela_inicial.png" alt="Initial Screen" width="600"> </p> <p align="center"> <img src="docs/Tela_inicial_arrastado.png" alt="Initial Screen - Dragged" width="600"> </p>
🗺️ Interactive DXF Viewer
<p align="center"> <img src="docs/Tela_inicial_dxf_aberto.png" alt="DXF Viewer" width="800"> </p>
The system automatically renders the map extracted from the DXF, preserving colors, texts, and geometries.
On the right panel, you have buttons to reset view, measure distances, and save the final image.

✅ Spreadsheet Entry and Layer Filters
<p align="center"> <img src="docs/Janela_layout.png" alt="Layer Selection" width="400"> </p>
The "Designer" field always saves the last entered name to speed up the process.

Before generating the spreadsheet, the system allows you to choose which layers should be included in the calculations.

📄 Map and Legend Generated in the Spreadsheet
<p align="center"> <img src="docs/Pdf_frente.png" alt="Spreadsheet Map - Front" width="700"> </p> <p align="center"> <img src="docs/Pdf_costas.png" alt="Spreadsheet Map - Back" width="700"> </p>
The spreadsheet includes:

Rendered DXF map

Automatically generated legend based on used layers

Automatically filled fields:

Current date

Version (incremental)

Property name (from the DXF file)

Cane area, scale, distance (entered by user)

📊 Length and Plot Tables
Based on the lines in the DXF, the system calculates:

Length per layer (count, total, and average)

Area per plot, in hectares and alqueires

Total cultivable area

🔄 The tables are generated automatically based on visible layers and text near the geometries.

