from Automatizador import Automatizador

automatizador=Automatizador()
# automatizador.empezarTarea('codigosCSV.csv') #COLOCAR NOMBRE DE ARCHIVO CSV CABECERA = 'REGISTRO SANITARIO'

import tkinter as tk
from tkinter import *
from tkinter import filedialog

ventana=tk.Tk()
ventana.title("WebScraping-Digemid")
ventana.geometry('380x300')
ventana.configure(background='green yellow')

def abrirArchivo():
    archivo=filedialog.askopenfilename(title='Escoja el Csv',filetypes=(('ARCHIVOS CSV','*.csv'),))
    au=Automatizador()
    au.empezarTarea(archivo)
    exit()

Button(ventana,text='ABRIR CSV',command=abrirArchivo).pack()

ventana.mainloop()