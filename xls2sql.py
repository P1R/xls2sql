#!/usr/bin/env python2

__author__ = 'p1r0'

import sys
import xlrd
import os

Help="ERROR - malos parametros:(\n\t la forma de uso es: xls2sql.py archivo.xlsx"

if __name__ == "__main__":
    if(len(sys.argv) == 2):
#almacenamos el nombre del Fichero
        FileName=sys.argv[1]
#creamos archivo de salida con nombre generico
        OuFile = open("first", 'w')
#Abrmos la hoja de calculo
        book = xlrd.open_workbook(sys.argv[1])
        print "El numero de hojas es:", book.nsheets
        print "nombre(s) de las hojas:", book.sheet_names()
#Tomamos la primer Hoja
        sh = book.sheet_by_index(0)
#Imprimimos  nombre de hoja, filas y columnas
        print sh.name, sh.nrows, sh.ncols
        OuFile.writelines("first try")
        for i in range(1,sh.nrows):
            print "la fila "+str(i)+" tiene:"+sh.cell_value(rowx=i, colx=0),
            Fecha = xlrd.xldate.xldate_as_tuple(sh.cell_value(rowx=i,colx=1),0)
            Anio, Mes, Dia = Fecha[0], Fecha[1], Fecha[2]
            print str(Dia)+"-"+str(Mes)+"-"+str(Anio),
            Hora = xlrd.xldate.xldate_as_tuple(sh.cell_value(rowx=i,colx=2),0) 
            Horas, Minutos, Segundos =  "%02d"%Hora[3], "%02d"%Hora[4], "%02d"%Hora[5]
            print str(Horas)+":"+str(Minutos)+":"+str(Segundos),
            print sh.cell_value(rowx=i,colx=3)

#rescribimos el nombre del archivo agregando anio+mes+sql
        os.rename('first',(sys.argv[1].split(".")[0])+str(Anio)+"-"+str(Mes)+".sql")
    else:
        print Help
