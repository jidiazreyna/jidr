# sentencia_window.py
from PySide6.QtWidgets import QMainWindow
from tramsent import confirm_and_quit
from tramsent import SentenciaWidget
from PySide6.QtGui import QGuiApplication, QScreen
from PySide6.QtGui import QIcon


class SentenciaWindow(QMainWindow):
    skip_confirm: bool = False          # ← bandera

    def __init__(self, data, parent: QMainWindow | None = None):
        super().__init__(parent)
        self.setWindowTitle("Sentencia")
                # Heredamos el ícono de la ventana principal, si existe
        # Heredamos el ícono de la ventana principal:
        if parent is not None:
            self.setWindowIcon(parent.windowIcon())

        self.setCentralWidget(SentenciaWidget(data, self))

        # ------------------------------------------------------------------
        # 1) Igualamos el TAMAÑO al de la ventana de Trámites (si existe);
        #    caso contrario, fijamos una resolución razonable.
        # ------------------------------------------------------------------
        if parent is not None:
            self.resize(parent.size())          # mismo ancho/alto
        else:
            self.resize(1200, 700)              # fallback

        # ------------------------------------------------------------------
        # 2) CENTRAMOS la ventana en la pantalla disponible.
        # ------------------------------------------------------------------
        geo = self.frameGeometry()                                   # rect actual
        center_point = QGuiApplication.primaryScreen()\
                                     .availableGeometry().center()   # centro pantalla
        geo.moveCenter(center_point)
        self.move(geo.topLeft())                                     # la llevamos allí

    def closeEvent(self, ev):
        if self.skip_confirm:
            ev.accept()
            # si tenés main_win, lo mostramos
            if hasattr(self, 'main_win'):
                self.main_win.show()
            else:
                # fallback a la lógica original
                self.parent().show()
        else:
            confirm_and_quit(self)
            ev.ignore()
