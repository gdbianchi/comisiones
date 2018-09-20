import sys

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.uic import loadUi

from comisiones import *


class winPrincipal(QMainWindow):
    def __init__(self):
        super(winPrincipal, self).__init__()
        loadUi('comisiones.ui', self)
        self.setWindowTitle('Comisiones Argos v1.0')
        self.spinAno.setMaximum(self.calendario.yearShown())
        self.spinAno.setValue(self.calendario.yearShown())
        self.spinMes.setValue(self.calendario.monthShown())
        self.btnComisionar.clicked.connect(self.on_push_comisiones)
        self.calendario.clicked.connect(self.calendar_change)
        self.FIniA = self.calendario.selectedDate()
        self.rbIniA.setText("Inicio Negro: " + self.FIniA.toString("yyyy-MM-dd"))
        self.FFinA = self.calendario.selectedDate()
        self.rbFinA.setText("Final Negro: " + self.FFinA.toString("yyyy-MM-dd"))
        self.FIniB = self.calendario.selectedDate()
        self.rbIniB.setText("Inicio Blanco: " + self.FIniB.toString("yyyy-MM-dd"))
        self.FFinB = self.calendario.selectedDate()
        self.rbFinB.setText("Final Blanco: " + self.FFinB.toString("yyyy-MM-dd"))

    @pyqtSlot()
    def on_push_comisiones(self):

        progressBar = self.pgBar
        taskInfo = self.lblTarea

        meses = "ENERO", "FEBRERO", "MARZO", "ABRIL", "MARZO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
        print("Comisionar")
        print("BLANCO: "+self.FIniB.toString("yyyy-MM-dd")+" "+self.FFinB.toString("yyyy-MM-dd"))
        print("NEGRO: "+self.FIniA.toString("yyyy-MM-dd")+" "+self.FFinA.toString("yyyy-MM-dd"))
        comisionar(self.spinAno.value(), meses[self.spinMes.value()-1], self.FIniA.toString("yyyy-MM-dd"), self.FIniB.toString("yyyy-MM-dd"), self.FFinA.toString("yyyy-MM-dd"), self.FFinB.toString("yyyy-MM-dd"), progressBar, taskInfo)

    @pyqtSlot()
    def calendar_change(self):
        if self.rbIniA.isChecked():
            self.FIniA = self.calendario.selectedDate()
            self.rbIniA.setText("Inicio Negro: " + self.FIniA.toString("yyyy-MM-dd"))
        if self.rbIniB.isChecked():
            self.FIniB = self.calendario.selectedDate()
            self.rbIniB.setText("Inicio Blanco: " + self.FIniB.toString("yyyy-MM-dd"))
        if self.rbFinA.isChecked():
            self.FFinA = self.calendario.selectedDate()
            self.rbFinA.setText("Final Negro: " + self.FFinA.toString("yyyy-MM-dd"))
        if self.rbFinB.isChecked():
            self.FFinB = self.calendario.selectedDate()
            self.rbFinB.setText("Final Blanco: " + self.FFinB.toString("yyyy-MM-dd"))

        self.btnComisionar.setEnabled((self.FIniA.toJulianDay()<self.FFinA.toJulianDay())and(self.FIniB.toJulianDay()<self.FFinB.toJulianDay()))

app = QApplication(sys.argv)
widget = winPrincipal()
widget.show()
sys.exit(app.exec_())