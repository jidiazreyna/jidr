#!/usr/bin/env python3
"""
Generador GUI – Sobreseimiento por Prescripción
==============================================
Versión revisada (bug-fix 1)
----------------------------
• Corrige `AttributeError: HechoWidget` → ahora guarda `self.idx` y lo usa en `texto()`.
• Resto idéntico a la versión anterior (fecha automática, tipo de tribunal, sexo M/F, n-imputados/hechos, zoom, copiar, exportar DOCX).

Dependencias rápidas
--------------------
    pip install PySide6 python-docx

Ejecución
---------
    python prescripcion_gui.py
"""
from __future__ import annotations

import sys, re, datetime
from typing import List, Dict
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QTextEdit, QTextBrowser,
    QSpinBox, QComboBox, QRadioButton, QButtonGroup, QPushButton,
    QVBoxLayout, QHBoxLayout, QGridLayout, QScrollArea, QFileDialog,
    QMessageBox, QSizePolicy
)

# ─────────────────────────────────────────────────────────── Helpers plantilla ──

def _pronombres_imputados(sexos: List[str]) -> Dict[str, str]:
    n = len(sexos)
    cant_m = sexos.count("M"); cant_f = n - cant_m
    if n == 1:
        art = "el imputado" if sexos[0] == "M" else "la imputada"
        asistido = "asistido" if sexos[0] == "M" else "asistida"
        acusado = "acusado" if sexos[0] == "M" else "acusada"
        le_les = "le"
    else:
        if cant_f == n:
            art, asistido, acusado = "las imputadas", "asistidas", "acusadas"
        elif cant_m == n:
            art, asistido, acusado = "los imputados", "asistidos", "acusados"
        else:
            art, asistido, acusado = "los imputados", "asistidos", "acusados"
        le_les = "les"
    return {
        "imputado_articulo": art,
        "le_les": le_les,
        "asistido_label": asistido,
        "acusado_label": acusado,
    }

# Fecha en letras
MESES = ("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre")
NUMS_UNI = ("cero","uno","dos","tres","cuatro","cinco","seis","siete","ocho","nueve","diez","once","doce","trece","catorce","quince","dieciséis","diecisiete","dieciocho","diecinueve","veinte","veintiuno","veintidós","veintitrés","veinticuatro","veinticinco","veintiséis","veintisiete","veintiocho","veintinueve")
DECENAS = ("treinta","cuarenta")

def _num_letras(n:int)->str:
    if n<30: return NUMS_UNI[n]
    dec, uni = divmod(n,10); base = DECENAS[dec-3]
    return base if uni==0 else f"{base} y {NUMS_UNI[uni]}"

def fecha_hoy_letras()->str:
    hoy = datetime.date.today(); dia = _num_letras(hoy.day); mes = MESES[hoy.month-1]
    return f"{dia} de {mes} de {hoy.year}"

TEMPLATE = """<p align='justify'>Córdoba, {fecha_letras}.</p>
<p align='justify'>VISTA: la presente causa caratulada {caratula}, venida a {este_esta} {tribunal} a los efectos de resolver la situación procesal de {nombre_apellido}, {datos_personales}</p>
<p align='justify'>DE LA QUE RESULTA: Que {imputado_articulo} {nombre_apellido} se {le_les} atribuye {hechos_label}:</p>
<p align='justify'><i>{hecho}</i></p>
<p align='justify'><b>Y CONSIDERANDO:</b></p>
<p align='justify'>I) Durante la instrucción se colectaron los siguientes elementos probatorios: {prueba}</p>
<p align='justify'>II) El fiscal {fiscal} requiere el sobreseimiento total por prescripción… (demo).</p>
<p align='justify'><b>RESUELVO:</b></p>
<p align='justify'>I. Sobreseer totalmente a {nombre_apellido}…</p>"""

def render_prescripcion(*,sexos_imputados:List[str],num_hechos:int,**campos)->str:
    auto = _pronombres_imputados(sexos_imputados)
    auto["hechos_label"] = "el siguiente hecho" if num_hechos==1 else "los siguientes hechos"
    return TEMPLATE.format(**auto,**campos)

# ───────────────────────────────────────────────────────────── Widgets ──
class ImputadoWidget(QWidget):
    def __init__(self, idx:int):
        super().__init__(); self.idx=idx
        lay = QHBoxLayout(self)
        self.edit_nombre = QLineEdit(); self.edit_nombre.setPlaceholderText(f"Nombre imputado #{idx+1}")
        self.rb_m = QRadioButton("M"); self.rb_f = QRadioButton("F"); self.rb_m.setChecked(True)
        grp=QButtonGroup(self); grp.addButton(self.rb_m); grp.addButton(self.rb_f)
        for w in (self.edit_nombre,self.rb_m,self.rb_f): lay.addWidget(w)
        lay.addStretch()
    def nombre(self)->str: return self.edit_nombre.text().strip() or f"Imputado#{self.idx+1}"
    def sexo(self)->str: return "M" if self.rb_m.isChecked() else "F"

class HechoWidget(QWidget):
    def __init__(self, idx:int):
        super().__init__(); self.idx = idx
        lay = QVBoxLayout(self)
        self.txt = QTextEdit(); self.txt.setPlaceholderText(f"Descripción hecho #{idx+1}"); self.txt.setFixedHeight(50)
        lay.addWidget(self.txt)
    def texto(self)->str: return self.txt.toPlainText().strip() or f"[hecho {self.idx+1}]"

class ZoomBrowser(QTextBrowser):
    def wheelEvent(self,ev):
        if ev.modifiers() & Qt.ControlModifier:
            self.zoomIn(1) if ev.angleDelta().y()>0 else self.zoomOut(1); ev.accept()
        else: super().wheelEvent(ev)

class PrescripcionGUI(QWidget):
    def __init__(self):
        super().__init__(); self.setWindowTitle("Prescripción – Generador rápido"); self.resize(1100,600)
        main = QHBoxLayout(self)
        # formulario con scroll
        scroll=QScrollArea(); scroll.setWidgetResizable(True); form_holder=QWidget(); self.form_layout=QVBoxLayout(form_holder); self.form_layout.setAlignment(Qt.AlignTop); scroll.setWidget(form_holder)
        main.addWidget(scroll,2)
        # preview
        self.preview=ZoomBrowser(); self.preview.document().setDefaultFont(QFont("Times New Roman",12)); main.addWidget(self.preview,3)
        # campos fijos
        grid=QGridLayout(); row=0; self.form_layout.addLayout(grid)
        grid.addWidget(QLabel("Carátula:"),row,0); self.ed_caratula=QLineEdit(); grid.addWidget(self.ed_caratula,row,1); row+=1
        grid.addWidget(QLabel("Tribunal:"),row,0); self.ed_tribunal=QLineEdit(); grid.addWidget(self.ed_tribunal,row,1); row+=1
        grid.addWidget(QLabel("Tipo tribunal:"),row,0); self.cb_tipo=QComboBox(); self.cb_tipo.addItems(["Juzgado","Cámara"]); grid.addWidget(self.cb_tipo,row,1); row+=1
        grid.addWidget(QLabel("Prueba:"),row,0); self.ed_prueba=QLineEdit(); grid.addWidget(self.ed_prueba,row,1); row+=1
        grid.addWidget(QLabel("Fiscal:"),row,0); self.ed_fiscal=QLineEdit(); grid.addWidget(self.ed_fiscal,row,1); row+=1
        grid.addWidget(QLabel("N° imputados:"),row,0); self.spin_imp=QSpinBox(minimum=1,maximum=8,value=1); grid.addWidget(self.spin_imp,row,1); row+=1
        grid.addWidget(QLabel("N° hechos:"),row,0); self.spin_hec=QSpinBox(minimum=1,maximum=8,value=1); grid.addWidget(self.spin_hec,row,1); row+=1
        # contenedores dinámicos
        self.box_imputados=QVBoxLayout(); self.form_layout.addLayout(self.box_imputados)
        self.box_hechos=QVBoxLayout(); self.form_layout.addLayout(self.box_hechos)
        # botones
        btn_row=QHBoxLayout(); self.form_layout.addLayout(btn_row)
        btn_copy=QPushButton("Copiar"); btn_docx=QPushButton("Exportar DOCX"); btn_row.addWidget(btn_copy); btn_row.addWidget(btn_docx); btn_row.addStretch()
        # señales
        for w in (self.ed_caratula,self.ed_tribunal,self.ed_prueba,self.ed_fiscal): w.textChanged.connect(self.update_preview)
        self.cb_tipo.currentTextChanged.connect(self.update_preview)
        self.spin_imp.valueChanged.connect(self.refresh_imputados)
        self.spin_hec.valueChanged.connect(self.refresh_hechos)
        btn_copy.clicked.connect(self.copy_clip)
        btn_docx.clicked.connect(self.save_docx)
        # dinámicos
        self.imputados_widgets:List[ImputadoWidget]=[]; self.hechos_widgets:List[HechoWidget]=[]
        self.refresh_imputados(); self.refresh_hechos(); self.update_preview()

    def refresh_imputados(self):
        target=self.spin_imp.value()
        while len(self.imputados_widgets)<target:
            w=ImputadoWidget(len(self.imputados_widgets)); self.imputados_widgets.append(w); self.box_imputados.addWidget(w)
            w.edit_nombre.textChanged.connect(self.update_preview); w.rb_m.toggled.connect(self.update_preview); w.rb_f.toggled.connect(self.update_preview)
        while len(self.imputados_widgets)>target:
            w=self.imputados_widgets.pop(); w.setParent(None)
        self.update_preview()

    def refresh_hechos(self):
        target=self.spin_hec.value()
        while len(self.hechos_widgets)<target:
            w=HechoWidget(len(self.hechos_widgets)); self.hechos_widgets.append(w); self.box_hechos.addWidget(w); w.txt.textChanged.connect(self.update_preview)
        while len(self.hechos_widgets)>target:
            w=self.hechos_widgets.pop(); w.setParent(None)
        self.update_preview()

    def update_preview(self):
        sexos=[w.sexo() for w in self.imputados_widgets]
        hechos_desc="\n".join(f"{i+1}. {w.texto()}" for i,w in enumerate(self.hechos_widgets))
        campos={
            "fecha_letras":fecha_hoy_letras(),
            "caratula": self.ed_caratula.text().strip() or "[carátula]",
            "este_esta": "este" if self.cb_tipo.currentText()=="Juzgado" else "esta",
            "tribunal": self.ed_tribunal.text().strip() or "[tribunal]",
            "nombre_apellido": ", ".join(w.nombre() for w in self.imputados_widgets),
            "datos_personales": "[datos]",
            "hecho": hechos_desc or "[hechos]",
            "prueba": self.ed_prueba.text().strip() or "[prueba]",
            "fiscal": self.ed_fiscal.text().strip() or "[fiscal]",
        }
        html=render_prescripcion(sexos_imputados=sexos,num_hechos=self.spin_hec.value(),**campos)
        self.preview.setHtml(html)

    def copy_clip(self):
        QApplication.clipboard().setText(self.preview.toPlainText())
        QMessageBox.information(self,"Copiado","Texto copiado al portapapeles")

    def save_docx(self):
        try:
            from docx import Document
            from docx.shared import Pt
            from docx.enum.text import WD_ALIGN_PARAGRAPH
        except ImportError:
            QMessageBox.warning(self,"python-docx faltante","Instale con: pip install python-docx"); return
        path,_=QFileDialog.getSaveFileName(self,"Guardar DOCX","","Word (*.docx)")
        if not path: return
        doc=Document(); doc._body.clear_content()
        for par in self.preview.toPlainText().split('\n'):
            p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
            run=p.add_run(par); run.font.name="Times New Roman"; run.font.size=Pt(12)
        doc.save(path); QMessageBox.information(self,"Guardado",f"Archivo guardado en:\n{path}")

# ───────────────────────────────────────────────────────── main ──
if __name__=="__main__":
    app=QApplication(sys.argv)
    gui=PrescripcionGUI(); gui.show()
    sys.exit(app.exec())
