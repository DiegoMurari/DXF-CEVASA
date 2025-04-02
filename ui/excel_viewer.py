import os
from PySide6.QtWidgets import QWidget, QVBoxLayout
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtCore import QUrl

class PDFViewer(QWidget):
    def __init__(self, pdf_path):
        super().__init__()
        self.setWindowTitle("Visualização da Planilha Final")
        self.resize(1000, 700)

        layout = QVBoxLayout()
        self.browser = QWebEngineView()
        layout.addWidget(self.browser)
        self.setLayout(layout)

        self.setStyleSheet("""
            QWidget {
                background-color: #1e1e1e;
            }
        """)

        self.browser.load(QUrl.fromLocalFile(os.path.abspath(pdf_path)))
