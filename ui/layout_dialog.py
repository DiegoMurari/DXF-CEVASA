import os
from datetime import datetime
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QMessageBox, QListWidget, QListWidgetItem, QScrollArea, QWidget, QComboBox
)
from PySide6.QtCore import Qt

MAX_DESENHISTA = 60
LAST_DESENHISTA_FILE = "data/last_desenhista.txt"

class ExtendedLayoutInfoDialog(QDialog):
    def __init__(self, dxf_path=None, available_layers=None, parent=None):
        super().__init__(parent)
        self.dxf_path = dxf_path
        self.available_layers = available_layers or []
        self.result = {}
        self.init_ui()
        # Aplica a folha de estilo para um visual minimalista e integrado
        self.setStyleSheet("""
            QDialog {
                background-color: #2d2d30;
                color: #ffffff;
                font-family: "Segoe UI", Verdana, sans-serif;
                font-size: 12px;
            }
            QLabel {
                color: #ffffff;
            }
            QLineEdit {
                background-color: #3e3e42;
                border: 1px solid #555555;
                padding: 4px;
                border-radius: 4px;
                color: #ffffff;
            }
            QComboBox {
                background-color: #3e3e42;
                padding: 4px;
                border: 1px solid #555555;
                border-radius: 4px;
                color: #ffffff;
            }
            QPushButton {
                background-color: #007ACC;
                border: none;
                padding: 6px;
                border-radius: 4px;
                color: #ffffff;
            }
            QPushButton:hover {
                background-color: #005999;
            }
            QListWidget {
                background-color: #3e3e42;
                border: 1px solid #555555;
                color: #ffffff;
            }
            QListWidget::item {
                padding: 4px;
            }
            QListWidget::item:selected, QListWidget::item:selected:hover {
                background-color: #005999;
                color: #ffffff;
            }
            QListWidget::item:hover {
                background-color: #464646;
                color: #ffffff;
            }
            QScrollArea {
                border: none;
                background-color: #2d2d30;
            }
        """)

    def init_ui(self):
        self.setWindowTitle("Informações do Layout")
        self.setMinimumWidth(400)
        layout = QVBoxLayout(self)
        
        # Campo: Desenhista – pré-popula se houver valor salvo
        self.desenhista_edit = QLineEdit(self)
        if os.path.exists(LAST_DESENHISTA_FILE):
            try:
                with open(LAST_DESENHISTA_FILE, "r", encoding="utf-8") as f:
                    ultimo_nome = f.read().strip()
                    self.desenhista_edit.setText(ultimo_nome)
            except Exception as e:
                print(f"Erro ao ler {LAST_DESENHISTA_FILE}: {e}")
        layout.addLayout(self._create_row("Desenhista:", self.desenhista_edit))
        
        # Campo: Escala – usando QComboBox editável com valores pré-definidos
        self.escala_combo = QComboBox(self)
        self.escala_combo.setEditable(True)
        # Preenche com escalas de 1:1000 até 1:10000 (pode ajustar conforme necessário)
        for escala in range(1000, 10001, 1000):
            # Exibe no formato "1:1000", "1:2000", etc.
            self.escala_combo.addItem(f"1:{escala}")
        layout.addLayout(self._create_row("Escala:", self.escala_combo))
        
        # Campo: Distância – QLineEdit
        self.distancia_edit = QLineEdit(self)
        layout.addLayout(self._create_row("Distância:", self.distancia_edit))
        
        # Campo: Área Cana (ha)
        self.area_cana_edit = QLineEdit(self)
        layout.addLayout(self._create_row("Área Cana (ha):", self.area_cana_edit))
        
        # Campo: Mun/Estado
        self.mun_estado_edit = QLineEdit(self)
        layout.addLayout(self._create_row("Mun/Estado:", self.mun_estado_edit))
        
        # Campo: Parc Outorgante
        self.parc_outorgante_edit = QLineEdit(self)
        layout.addLayout(self._create_row("Parc Outorgante:", self.parc_outorgante_edit))
        
        # Se available_layers for fornecida, cria a área para seleção de layers com botões
        if self.available_layers:
            label_layers = QLabel("Selecione os layers para legendas e tabelas:", self)
            layout.addWidget(label_layers)
            
            btn_layout = QHBoxLayout()
            self.btn_select_all = QPushButton("Selecionar Tudo", self)
            self.btn_deselect_all = QPushButton("Desmarcar Tudo", self)
            self.btn_select_all.setStyleSheet("padding: 4px;")
            self.btn_deselect_all.setStyleSheet("padding: 4px;")
            self.btn_select_all.clicked.connect(self.selecionar_todos)
            self.btn_deselect_all.clicked.connect(self.desmarcar_todos)
            btn_layout.addWidget(self.btn_select_all)
            btn_layout.addWidget(self.btn_deselect_all)
            layout.addLayout(btn_layout)
            
            self.layer_list = QListWidget(self)
            self.layer_list.setSelectionMode(QListWidget.MultiSelection)
            for layer in self.available_layers:
                item = QListWidgetItem(layer, self.layer_list)
                item.setSelected(True)
            scroll_area = QScrollArea(self)
            scroll_area.setWidgetResizable(True)
            container = QWidget(self)
            container_layout = QVBoxLayout(container)
            container_layout.addWidget(self.layer_list)
            scroll_area.setWidget(container)
            layout.addWidget(scroll_area)
        else:
            self.layer_list = None
        
        # Campos automáticos: Data, Versão e Propriedade – agora editáveis
        self.data_edit = QLineEdit(self)
        current_date = datetime.now().strftime("%d/%m/%Y")
        self.data_edit.setText(current_date)
        layout.addLayout(self._create_row("Data:", self.data_edit))
        
        self.versao_edit = QLineEdit(self)
        self.versao_edit.setText("0.1")
        layout.addLayout(self._create_row("Versão:", self.versao_edit))
        
        self.propriedade_edit = QLineEdit(self)
        if self.dxf_path:
            base_name = os.path.splitext(os.path.basename(self.dxf_path))[0]
            self.propriedade_edit.setText(base_name)
        layout.addLayout(self._create_row("Propriedade:", self.propriedade_edit))
        
        self.confirm_button = QPushButton("Confirmar", self)
        self.confirm_button.clicked.connect(self.confirm_info)
        layout.addWidget(self.confirm_button)
    
    def _create_row(self, label_text, widget):
        row = QHBoxLayout()
        label = QLabel(label_text, self)
        label.setMinimumWidth(150)
        row.addWidget(label)
        row.addWidget(widget)
        return row
    
    def selecionar_todos(self):
        if self.layer_list:
            for index in range(self.layer_list.count()):
                item = self.layer_list.item(index)
                item.setSelected(True)
    
    def desmarcar_todos(self):
        if self.layer_list:
            for index in range(self.layer_list.count()):
                item = self.layer_list.item(index)
                item.setSelected(False)
    
    def confirm_info(self):
        desenhista = self.desenhista_edit.text().strip().upper()
        if not desenhista:
            QMessageBox.critical(self, "Erro", "O campo 'Desenhista' é obrigatório.")
            return
        if len(desenhista) > MAX_DESENHISTA:
            QMessageBox.critical(self, "Erro", f"O nome do Desenhista deve ter no máximo {MAX_DESENHISTA} caracteres.")
            return
        
        try:
            os.makedirs(os.path.dirname(LAST_DESENHISTA_FILE), exist_ok=True)
            with open(LAST_DESENHISTA_FILE, "w", encoding="utf-8") as f:
                f.write(desenhista)
        except Exception as e:
            print(f"Erro ao salvar o nome do desenhista: {e}")
        
        self.result['desenhista'] = desenhista
        self.result['escala'] = self.escala_combo.currentText().strip()
        self.result['distancia'] = self.distancia_edit.text().strip()
        self.result['area_cana'] = self.area_cana_edit.text().strip()
        self.result['mun_est'] = self.mun_estado_edit.text().strip()
        self.result['parc'] = self.parc_outorgante_edit.text().strip()
        self.result['data_atual'] = self.data_edit.text().strip()
        self.result['nova_versao'] = self.versao_edit.text().strip()
        self.result['propriedade'] = self.propriedade_edit.text().strip()
        
        if self.layer_list:
            selected = [self.layer_list.item(i).text() for i in range(self.layer_list.count())
                        if self.layer_list.item(i).isSelected()]
            self.result['selected_layers'] = selected
        
        self.accept()
    
    def get_result(self):
        return self.result