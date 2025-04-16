import sys
import os
import math
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QPainter, QColor, QPixmap, QRegion, QPainterPath, QPen
from PySide6.QtWidgets import QApplication, QLabel, QWidget


class SplashScreen(QWidget):
    def __init__(self, image_path: str):
        super().__init__()

        self.setWindowFlags(Qt.FramelessWindowHint | Qt.SplashScreen)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFixedSize(600, 400)

        # Borda arredondada real
        path = QPainterPath()
        path.addRoundedRect(self.rect(), 20, 20)
        self.setMask(QRegion(path.toFillPolygon().toPolygon()))

        # Imagem de fundo
        pixmap = QPixmap(image_path)
        if pixmap.isNull():
            print(f"❌ Falha ao carregar imagem: {image_path}")
        else:
            print("✅ Imagem carregada com sucesso")

        self.bg = QLabel(self)
        self.bg.setPixmap(pixmap.scaled(self.size(), Qt.KeepAspectRatioByExpanding, Qt.SmoothTransformation))
        self.bg.setGeometry(0, 0, self.width(), self.height())

        # Círculo animado
        self.angle = 0
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_spinner)
        self.timer.start(16)

        # Círculo por cima
        self.spinner_overlay = QWidget(self)
        self.spinner_overlay.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.spinner_overlay.setGeometry(0, 0, self.width(), self.height())
        self.spinner_overlay.paintEvent = self.paint_spinner

    def update_spinner(self):
        self.angle = (self.angle + 4) % 360
        self.spinner_overlay.repaint()

    def paint_spinner(self, event):
        painter = QPainter(self.spinner_overlay)
        painter.setRenderHint(QPainter.Antialiasing)

        center_x = self.width() // 2
        center_y = self.height() - 60
        radius = 28
        thickness = 5

        # Anel branco translúcido
        pen_bg = QColor(255, 255, 255, 40)
        painter.setPen(QPen(pen_bg, thickness))
        painter.setBrush(Qt.NoBrush)
        painter.drawEllipse(center_x - radius, center_y - radius, radius * 2, radius * 2)

        # Arco giratório branco opaco
        pen_arc = QColor(255, 255, 255, 220)
        painter.setPen(QPen(pen_arc, thickness))
        painter.save()
        painter.translate(center_x, center_y)
        painter.rotate(self.angle)
        painter.drawArc(-radius, -radius, radius * 2, radius * 2, 0, 60 * 16)  # 60 graus
        painter.restore()


def show_splash_and_launch(main_window_callback, delay=2000):
    app = QApplication.instance() or QApplication(sys.argv)

    # Caminho compatível com execução empacotada (PyInstaller)
    base_dir = os.path.dirname(os.path.abspath(__file__))
    splash_path = os.path.join(base_dir, "..", "resources", "images", "splash.png")
    splash_path = os.path.normpath(splash_path)

    splash = SplashScreen(splash_path)
    splash.show()

    def launch():
        splash.close()
        window = main_window_callback()
        window.show()

    QTimer.singleShot(delay, launch)
    app.exec()
