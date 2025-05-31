from PySide6.QtWidgets import QApplication
from core_data import CausaData
from main import MainWindow
import sys
from PySide6.QtGui import QIcon
from main import resource_path

if __name__ == "__main__":
    app   = QApplication(sys.argv)
    app.setWindowIcon(QIcon(resource_path("icono5.ico")))
    model = CausaData.instance()         # la única copia
    win   = MainWindow(model)
    win.show()
    sys.exit(app.exec())
