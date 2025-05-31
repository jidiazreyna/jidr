from __future__ import annotations
from dataclasses import dataclass, field
from typing import List, Dict, Any

from PySide6.QtCore import QSignalBlocker
import json, dataclasses, pathlib
from PySide6.QtWidgets import (
    QFileDialog, QMessageBox, QLineEdit, QComboBox, QCheckBox
)
from pathlib import Path

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from main import MainWindow
    from tramsent import SentenciaWidget

# ---------------------------------------------------------------------------
@dataclass
class CausaData:

    # ─────────── Generales (ambas pantallas) ───────────
    caratula: str = ""
    articulo: str = ""           # Cámara / Juzgado (combo en main)
    tribunal: str = ""
    sala: str = ""
    fecha_audiencia: str = ""     # «01/02/2025» o en letras
    hora_audiencia: str = ""
    localidad: str = "Córdoba"

    funcionario: str = ""        # quien firma pedidos / oficios
    fiscal_nombre: str = ""
    fiscal_sexo: str = "M"

    sentencia_num: str = ""      # «123/2025»
    resuelvo: str = ""
    firmantes: str = ""
    renuncia: bool = False

    # ─────────── Atributos propios de SentenciaWidget ───────────
    juez_nombre: str = ""
    juez_sexo: str = "M"
    juez_cargo: str = "juez"      # «juez» / «vocal»

    n_imputados: int = 1

    # Datos auxiliares de sentencia (ya estaban)
    sujeto_eventual: str = ""
    manifestacion_sujeto: str = ""
    victima: str = ""
    victima_plural: bool = False
    manifestacion_victima: str = ""
    pruebas: str = ""
    pruebas_relevantes: str = ""
    alegato_fiscal: str = ""
    alegato_defensa: str = ""
    calif_legal: str = "Correcta"
    calif_correccion: str = ""
    usa_potenciales: bool = False
    decomiso_si: bool = False
    decomiso_texto: str = ""
    restriccion_si: bool = False
    restriccion_texto: str = ""
    caso_vf: str = "No"

    # Listas
    imputados: List[Dict[str, Any]] = field(default_factory=list)
    hechos: List[Dict[str, Any]] = field(default_factory=list)

    # ---------------------------------------------------------------------
    #  MÉTODOS  – SYNC CON WIDGETS
    # ---------------------------------------------------------------------
    #  Nota: usamos "hasattr" para no romper si el widget aún no existe.

        # ------------- MainWindow ↔ modelo ------------------
    def from_main(self, win: "MainWindow") -> None:
        # print("[DEBUG from_main] Modelo antes:", self.imputados)
        """Lee TODOS los widgets de MainWindow y actualiza este objeto."""
        # Generales
        self.caratula        = getattr(win, "entry_caratula",     None).text() if hasattr(win, "entry_caratula") else self.caratula
        if hasattr(win, "combo_articulo"):
            self.articulo = win.combo_articulo.currentText()
            # Sincronizamos el cargo del juez con el tipo de tribunal
            self.juez_cargo = "vocal" if self.articulo.startswith("Cámara") else "juez"
        else:
            self.articulo = self.articulo
        self.tribunal        = getattr(win, "entry_tribunal",     None).currentText() if hasattr(win, "entry_tribunal") else self.tribunal
        self.sala            = getattr(win, "combo_sala",         None).currentText() if hasattr(win, "combo_sala") else self.sala
        self.fecha_audiencia = getattr(win, "entry_fecha",        None).text() if hasattr(win, "entry_fecha") else self.fecha_audiencia
        self.hora_audiencia  = getattr(win, "combo_hora",         None).currentText() if hasattr(win, "combo_hora") else self.hora_audiencia
        self.funcionario     = getattr(win, "entry_funcionario",  None).text() if hasattr(win, "entry_funcionario") else self.funcionario
        self.fiscal_nombre   = getattr(win, "entry_fiscal",       None).text() if hasattr(win, "entry_fiscal") else self.fiscal_nombre
        self.sentencia_num   = getattr(win, "entry_sentencia",    None).text() if hasattr(win, "entry_sentencia") else self.sentencia_num
        if hasattr(win, "entry_resuelvo"):
            # Cuando el widget no tiene HTML almacenado en la property,
            # ``toHtml()`` devuelve la plantilla vacía de Qt (con DOCTYPE y
            # meta etiquetas). Guardar eso provoca que aparezcan caracteres
            # extraños al reconstruir la sentencia o generar el DOCX.
            html_full = win.entry_resuelvo.property("html") or ""

            self.resuelvo_html = html_full
            from PySide6.QtGui import QTextDocument
            doc = QTextDocument(); doc.setHtml(html_full)
            self.resuelvo = doc.toPlainText().replace("\n", " ")
        self.firmantes       = getattr(win, "entry_firmantes",    None).text() if hasattr(win, "entry_firmantes") else self.firmantes
        self.renuncia        = getattr(win, "combo_renuncia",     None).currentText() == "Sí" if hasattr(win, "combo_renuncia") else self.renuncia
        self.n_imputados     = int(getattr(win, "combo_n",        None).currentText()) if hasattr(win, "combo_n") else self.n_imputados

        # Imputados — copiamos todos los widgets actuales
        self.imputados.clear()
        if hasattr(win, "imputados_widgets"):
            for w in win.imputados_widgets:
                self.imputados.append({
                    key: (
                        widget.text()        if isinstance(widget, QLineEdit) else
                        widget.currentText() if isinstance(widget, QComboBox) else
                        widget.isChecked()   if isinstance(widget, QCheckBox) else None
                    )
                    for key, widget in w.items()
                })
        # print("[DEBUG from_main] Modelo después:", self.imputados)

    def apply_to_main(self, win: "MainWindow") -> None:
        """Carga los widgets de MainWindow con los valores guardados."""
        if not hasattr(win, "entry_caratula"):
            return  # aún no está construida la UI
        # (usamos "setText" / "setCurrentText" solo cuando el valor difiere para evitar señales infinitas)
        _set = lambda w, val: w.setText(val)         if w.text() != val else None
        _setc= lambda w, val: w.setCurrentText(val)  if w.currentText() != val else None

        _set(win.entry_caratula, self.caratula)
        _setc(win.combo_articulo, self.articulo)
        # Tribunal: si es combo editable, usamos setCurrentText; si no, volvemos a setText
        trib_widget = getattr(win, "entry_tribunal", None)
        if trib_widget is not None:
            if hasattr(trib_widget, "setCurrentText"):
                if trib_widget.currentText() != self.tribunal:
                    trib_widget.setCurrentText(self.tribunal)
            else:
                # sigue usando _set para QLineEdit u otros
                _set(trib_widget, self.tribunal)

        _setc(win.combo_sala, self.sala)
        _set(win.entry_fecha, self.fecha_audiencia)
        _setc(win.combo_hora, self.hora_audiencia)
        _set(win.entry_funcionario, self.funcionario)
        _set(win.entry_fiscal, self.fiscal_nombre)
        _set(win.entry_sentencia, self.sentencia_num)
        if hasattr(win, "entry_resuelvo"):
            blocker = QSignalBlocker(win.entry_resuelvo)
            html_full = getattr(self, "resuelvo_html", self.resuelvo)
            win.entry_resuelvo.setProperty("html", html_full)
            win.entry_resuelvo.setHtml(html_full)
        _set(win.entry_firmantes, self.firmantes)
        _setc(win.combo_renuncia, "Sí" if self.renuncia else "No")
        # Evito que al cambiar combo_n se dispare update_template() → data.from_main()
        was_blocked = win.combo_n.blockSignals(True)
        win.combo_n.setCurrentText(str(self.n_imputados))
        win.combo_n.blockSignals(was_blocked)
        # reconstruir pestañas y volcar datos imputados
        win.rebuild_imputados()

        # bucle de copia
        for idx, w in enumerate(win.imputados_widgets):
            if idx >= len(self.imputados):
                break
            dato = self.imputados[idx]
            for k, widget in w.items():
                if k not in dato:
                    continue
                if isinstance(widget, QLineEdit):
                    widget.setText(dato[k])
                elif isinstance(widget, QComboBox):
                    widget.setCurrentText(dato[k])
                elif isinstance(widget, QCheckBox):
                    widget.setChecked(bool(dato[k]))

        win._refresh_imp_names_in_selector()

        # (Quedaba un segundo bucle duplicado; si realmente lo necesitas, pégalo aquí,
        # pero normalmente basta con uno solo)


    # ------------- SentenciaWidget ↔ modelo ----------------
# ------------- SentenciaWidget ↔ modelo ----------------
    def from_sentencia(self, sw: "SentenciaWidget") -> None:
        # print("[DEBUG from_sentencia] Modelo antes:", self.imputados)
        # ── datos generales ────────────────────────────────────────────────
        
        self.localidad       = sw.var_localidad.text().strip()
        self.caratula        = sw.var_caratula.text().strip()
        self.tribunal        = sw.var_tribunal.currentText().strip()
        self.sala            = sw.var_sala.currentText().strip()
        self.juez_nombre     = sw.var_juez.text().strip()
        self.juez_sexo       = "F" if sw.rb_juez_f.isChecked() else "M"
        self.juez_cargo      = sw.boton_cargo_juez.text().lower()
        # Mantener sincronizado el tipo de órgano con el cargo del juez.
        # Cuando el usuario cambia el cargo en ``SentenciaWidget`` entre
        # ``juez`` y ``vocal`` también debemos actualizar el campo
        # ``articulo`` que se muestra en la pantalla principal.
        self.articulo = (
            "Cámara en lo Criminal y Correccional"
            if self.juez_cargo == "vocal"
            else "Juzgado de Control"
        )
        self.fiscal_nombre   = sw.var_fiscal.text().strip()
        self.fiscal_sexo     = "F" if sw.rb_fiscal_f.isChecked() else "M"
        self.fecha_audiencia = sw.var_dia_audiencia.text().strip()
        self.n_imputados     = sw.var_num_imputados.value()
        html_full = sw.var_resuelvo.property("html") or ""
        self.resuelvo_html = html_full
        from PySide6.QtGui import QTextDocument
        doc = QTextDocument(); doc.setHtml(html_full)
        self.resuelvo = doc.toPlainText().replace("\n", " ")
        self.alegato_fiscal  = sw.var_alegato_fiscal
        self.alegato_defensa = sw.var_alegato_defensa

        old = list(self.imputados)
        self.imputados.clear()

        # Usamos enumerate normal (idx empieza en 0)
        for idx, imp_w in enumerate(sw.imputados):
            nuevos = {
                "nombre"      : imp_w["nombre"].text().strip(),
                "sexo"        : "F" if imp_w["sexo_rb"][1].isChecked() else "M",
                "datos"       : imp_w["datos"].text().strip(),
                "defensa"     : imp_w["defensor"].text().strip(),
                "tipo"        : imp_w["tipo_def"].currentText().strip(),
                "delitos"     : imp_w["delitos"].text().strip(),
                "condena"     : imp_w["condena"].text().strip(),
                "condiciones" : imp_w["condiciones"].text().strip(),
                "anteced_no"  : imp_w["antecedentes_opcion"][0].isChecked(),
                "anteced"     : imp_w["antecedentes"].text().strip(),
                "confesion"   : imp_w["confesion"].text().strip(),
                "ultima"      : imp_w["ultima"].text().strip(),
                "pautas"      : imp_w["pautas"].text().strip(),
            }

            if idx < len(old):
                base = old[idx].copy()
                base.update(nuevos)
            else:
                base = nuevos

            # print(f"[from_sentencia] imputado #{idx+1} fused →", base)
            self.imputados.append(base)

        # print("[DEBUG from_sentencia] Modelo después:", self.imputados)


    # (si también necesitás los ‘hechos’, añadí un bucle similar sobre sw.hechos)

        #  ... (añadí algunos; podés extender según necesites) ...

    def apply_to_sentencia(self, sw: "SentenciaWidget") -> None:
        """Vuelca los datos almacenados en el modelo a SentenciaWidget."""
        sw.var_localidad.setText(self.localidad)
        sw.var_caratula.setText(self.caratula)
        sw.var_tribunal.setCurrentText(self.tribunal)
        sw.var_sala.setCurrentText(self.sala)
        sw.var_juez.setText(self.juez_nombre)
        (sw.rb_juez_f if self.juez_sexo == 'F' else sw.rb_juez_m).setChecked(True)
        sw.boton_cargo_juez.setText(self.juez_cargo)
        sw.var_fiscal.setText(self.fiscal_nombre)
        (sw.rb_fiscal_f if self.fiscal_sexo == 'F' else sw.rb_fiscal_m).setChecked(True)
        sw.var_dia_audiencia.setText(self.fecha_audiencia)
        sw.var_num_imputados.setValue(self.n_imputados)
        html_full = getattr(self, "resuelvo_html", self.resuelvo)
        sw.var_resuelvo.setProperty("html", html_full)
        if hasattr(sw.var_resuelvo, "setHtml"):
            sw.var_resuelvo.setHtml(html_full)
        else:
            from PySide6.QtGui import QTextDocument
            doc = QTextDocument(); doc.setHtml(html_full)
            sw.var_resuelvo.setText(doc.toPlainText().replace("\n", " "))
        sw.var_alegato_fiscal  = self.alegato_fiscal
        sw.var_alegato_defensa = self.alegato_defensa

        # ── asegurémonos de que las pestañas de imputados existen ────────
        sw.update_imputados_section()

        # 2) Luego volcamos LOS DATOS de imputados (como te propuse antes)
        for idx, datos_imp in enumerate(self.imputados):
            if idx >= len(sw.imputados):
                break
            w = sw.imputados[idx]
            w["nombre"].setText( datos_imp.get("nombre","") )
            if datos_imp.get("sexo","M") == "F":
                w["sexo_rb"][1].setChecked(True)
            else:
                w["sexo_rb"][0].setChecked(True)
            w["datos"].setText( datos_imp.get("datos","") )
            w["defensor"].setText( datos_imp.get("defensa","") )
            w["tipo_def"].setCurrentText( datos_imp.get("tipo","") )
            w["delitos"].setText( datos_imp.get("delitos","") )
            w["condena"].setText( datos_imp.get("condena","") )
            w["condiciones"].setText( datos_imp.get("condiciones","") )
            # antecedentes (QRadioButton + QLineEdit)
            w["antecedentes_opcion"][0].setChecked( datos_imp.get("anteced_no",True) )
            w["antecedentes_opcion"][1].setChecked(not datos_imp.get("anteced_no",True))
            w["antecedentes"].setText( datos_imp.get("anteced","") )
            w["confesion"].setText( datos_imp.get("confesion","") )
            w["ultima"].setText( datos_imp.get("ultima","") )
            w["pautas"].setText( datos_imp.get("pautas","") )

        # 3) Ahora sincronizamos los hechos:
        count_hechos = len(self.hechos)
        sw.var_num_hechos.setValue(count_hechos or 1)
        sw.update_hechos_section()


        sw.actualizar_plantilla()

    def to_json(self, path: str | pathlib.Path) -> None:
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(dataclasses.asdict(self), fh, ensure_ascii=False, indent=2)

    @classmethod

    def from_json(cls, path: str | pathlib.Path) -> "CausaData":
        with open(path, "r", encoding="utf-8") as fh:
            raw = json.load(fh)
        return cls(**raw)

    # Guardar causa
    def guardar_causa(self):
        CAUSAS_DIR = Path("causas_guardadas")
        CAUSAS_DIR.mkdir(exist_ok=True)
        self.data.from_main(self)                           # sincroniza
        path, _ = QFileDialog.getSaveFileName(
            self, "Guardar causa", str(CAUSAS_DIR), "JSON (*.json)")
        if path:
            self.data.to_json(path)
            QMessageBox.information(self, "OK", "Causa guardada.")

    # Cargar causa
    def cargar_causa(self):
        CAUSAS_DIR = Path("causas_guardadas")
        CAUSAS_DIR.mkdir(exist_ok=True)
        path, _ = QFileDialog.getOpenFileName(
            self, "Cargar causa", str(CAUSAS_DIR), "JSON (*.json)")
        if path:
            self.data = CausaData.from_json(path)           # ¡nueva instancia!
            self.data.apply_to_main(self)                   # refresca widgets

# Eliminar causa  (sin cambios relevantes)
    def eliminar_causa(self):
        CAUSAS_DIR = Path("causas_guardadas")
        CAUSAS_DIR.mkdir(exist_ok=True)
        path, _ = QFileDialog.getOpenFileName(self, "Eliminar causa",
                                              str(CAUSAS_DIR), "JSON (*.json)")
        if path and QMessageBox.question(
            self, "Confirmar", f"¿Eliminar {Path(path).name}?"
        ) == QMessageBox.Yes:
            Path(path).unlink(missing_ok=True)



    # ------------------------------------------------------------------
    #  Factory / singleton (opcional)
    # ------------------------------------------------------------------
    _singleton: "CausaData | None" = None

    @classmethod
    def instance(cls) -> "CausaData":
        if cls._singleton is None:
            cls._singleton = cls()
        return cls._singleton