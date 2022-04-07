#@Autor: Gustavo Armitano
#App para el corralón Armitano
#Permite seleccionar un archivo pdf o xlsx, ordenarlo, limpiarlo y setearlo de acuerdo a un formato preestablecido
# y luego crear un xlsx con el resultado final.
# Los archivos son listas de precios. 

import sys
import pandas as pd
import numpy as np
import os
from datetime import date
from PyQt5.QtWidgets import QApplication, QMainWindow, QGraphicsDropShadowEffect, QFileDialog, QVBoxLayout
from PyQt5.QtCore import QPropertyAnimation, QEasingCurve
from PyQt5.QtCore import *
from PyQt5 import QtCore, QtWidgets
from PyQt5.uic import loadUi

#Define clase principal como hija de QMainWindows (osea, hereda todos sus métodos)
class Ventana_principal(QMainWindow):
    #Constructor de la clase hija
    def __init__(self):
        #Constructor de la clase padre
        super(Ventana_principal, self).__init__()
        #Carga el archivo con el diseño de pantalla QtDesign
        loadUi('C:/Users/Usuario/Mi unidad/PYTHON/PyQt/app_listas.ui', self)
     
        #Botones de acción
        self.btn_buscar.clicked.connect(self.procesa_archivo)
        self.btn_salir.clicked.connect(self.close)

        #Borra la barra de título de la ventana principal
        self.setWindowFlag(QtCore.Qt.FramelessWindowHint)
        self.setWindowOpacity(1)

    #Procesa el archivo seleccionado, segun extención y nombre, para limpiar la lista y exportar el definitivo
    def procesa_archivo(self):

         #Abre la ventana de Explorar Archivos y carga el valor en path
        path_inicial = QFileDialog.getOpenFileName(parent=self, caption='Selecccione una lista (xlsx)', directory='C:/Users/Usuario/Desktop/', filter='Archivo Excel (*.xlsx)')
    
        #Tupla que contiene [0]=ruta+nombre y [1]=extensión
        nombre = os.path.splitext(path_inicial[0])
        
        #Determinar la lista y procesar
        if 'malano' in str.casefold(nombre[0]):

            self.lista_malano(path_inicial)
        else: 
            self.mostrar_mensaje(0)        
        return

    def mostrar_mensaje(self, texto):  
        mensaje = texto
        #Tipo: 
        #      0-Lista no corresponde
        #      1-Procesando archivo...
        #      2-Lista valida
        #      3-Lista procesada con 'éxito.
        lista_mensajes=['Lista no corresponde', 'Procesando archivo...', 
                        'Lista válida', 'Lista procesada con éxito.']          
        self.lbl_cartel.setText(lista_mensajes[texto])
        return

##################################################
#DEFINICION DE METODOS PARA CADA LISTA DE PRECIOS#
##################################################

    #MALANO
    def lista_malano(self, path_inicial):
        self.mostrar_mensaje(1)
        #ABRIR ARCHIVO DE LISTA DE PRECIOS Y DESECHO LAS COLUMNAS Y FILAS SIN DATOS IMPORTANTES
        lista_malano = pd.read_excel(path_inicial[0], header=None , usecols="A:C", skiprows=13)

        #LISTA CON LOS NOMBRES DE LAS COLUMNAS
        columnas=['cod articulo', 'Codigo Provedor', 'Codigo articulo proveedor',
            'descripcion', 'stock', 'rubro', 'subrubro', 'Precio Lista Proveedor',
            'IVA', 'Dtos. Blanco', 'costo S/IVA Blanco', 'Utilidad Blanco',
            'Dtos. Rosa', 'costo S/IVA ROSA', 'Utilidad ROSA']

        #CAMBIAR NOMBRE DE COLUMNAS ORIGINALES
        lista_malano.columns = ['cod articulo', 'descripcion', 'Precio Lista Proveedor']

        #AGREGAR COLUMNAS PARA COMPLETAR LA DEFINITIVA
        for var in [1,2,4,5,6,8,9,10,11,12,13,14]:
            lista_malano.insert(var, columnas[var], '')
        
        #DEFINIR NOMBRE DE LISTA FINAL
        fecha = date.today()
  
        path_final = os.path.dirname(path_inicial[0]) + "/LISTA MALANO " + fecha.strftime('%d-%m-%Y') + '.xlsx'

        #EXPORTA LA LISTA FINAL A EXCEL
        lista_malano.to_excel(path_final, index=False)
        self.mostrar_mensaje(3)
        return


#Esto se ejecuta si es llamada la app, si se instacia la clase Ventana_principal no se ejecuta
if __name__=="__main__":
    app = QApplication(sys.argv)
    #Instancia de la clase y se muestra
    app_corralon = Ventana_principal()
    app_corralon.show()
    #Permite que la salida de la app sea general
    sys.exit(app.exec_())

