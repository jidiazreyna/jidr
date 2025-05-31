from PySide6.QtWidgets import QComboBox, QSpinBox

class NoWheelComboBox(QComboBox):
    """QComboBox que ignora la rueda del rat\u00f3n cuando el desplegable est\u00e1 cerrado."""
    def wheelEvent(self, event):
        if self.view().isVisible():
            super().wheelEvent(event)
        else:
            event.ignore()


class NoWheelSpinBox(QSpinBox):
    """QSpinBox que ignora la rueda del ratón para evitar cambios accidentales."""

    def wheelEvent(self, event):
        event.ignore()
