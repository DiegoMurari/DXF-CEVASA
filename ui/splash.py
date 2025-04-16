import sys
import os
from PySide6.QtCore import Qt, QTimer, QEventLoop
from PySide6.QtGui import QPainter, QColor, QPixmap, QRegion, QPainterPath, QPen
from PySide6.QtWidgets import QApplication, QLabel, QWidget


class SplashScreen(QWidget):
    def __init__(self, image_path: str):
        super().__init__()

        self.setWindowFlags(Qt.FramelessWindowHint | Qt.SplashScreen)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFixedSize(600, 400)

        # Arredondar bordas reais
        path = QPainterPath()
        path.addRoundedRect(self.rect(), 20, 20)
        self.setMask(QRegion(path.toFillPolygon().toPolygon()))

        # Carregar e exibir imagem de fundo
        self.pixmap = QPixmap(image_path)
        if self.pixmap.isNull():
            print(f"❌ Falha ao carregar imagem: {image_path}")
        else:
            print("✅ Imagem carregada com sucesso")

        # Ângulo da animação
        self.angle = 0
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.animate)
        self.timer.start(16)  # ~60 FPS

    def animate(self):
        self.angle = (self.angle + 4) % 360
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # Fundo com imagem
        if not self.pixmap.isNull():
            scaled = self.pixmap.scaled(self.size(), Qt.KeepAspectRatioByExpanding, Qt.SmoothTransformation)
            offset_x = (scaled.width() - self.width()) // 2
            offset_y = (scaled.height() - self.height()) // 2
            source_rect = scaled.copy(offset_x, offset_y, self.width(), self.height())
            painter.drawPixmap(0, 0, source_rect)

        # Spinner (círculo)
        center_x = self.width() // 2
        center_y = self.height() - 60
        radius = 28
        thickness = 5

        # Anel circular
        painter.setPen(QPen(QColor(255, 255, 255, 40), thickness))
        painter.setBrush(Qt.NoBrush)
        painter.drawEllipse(center_x - radius, center_y - radius, radius * 2, radius * 2)

        # Arco girando
        painter.setPen(QPen(QColor(255, 255, 255, 220), thickness))
        painter.save()
        painter.translate(center_x, center_y)
        painter.rotate(self.angle)
        painter.drawArc(-radius, -radius, radius * 2, radius * 2, 0, 60 * 16)
        painter.restore()


def show_splash_and_launch(main_window_callback):
    app = QApplication.instance() or QApplication(sys.argv)

    # Caminho da imagem
    base_dir = os.path.dirname(os.path.abspath(__file__))
    splash_path = os.path.normpath(os.path.join(base_dir, "..", "resources", "images", "splash.png"))

    splash = SplashScreen(splash_path)
    splash.show()

    # Garante que a splash fique animando por 2 segundos
    app.processEvents()
    loop = QEventLoop()
    QTimer.singleShot(2000, loop.quit)
    loop.exec()

    # Carrega e mostra interface principal
    window = main_window_callback()
    splash.close()
    window.show()

    app.exec()
