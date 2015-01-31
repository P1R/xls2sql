#!/usr/bin/env python2
import sys
import xlrd
from os import rename

__author__ = 'p1r0'


Help="ERROR - malos parametros:(\n\t la forma de uso es: xls2sql.py archivo.xlsx"

if __name__ == "__main__":
    if(len(sys.argv) == 2):
#almacenamos el nombre del Fichero
        FileName=sys.argv[1]
#creamos archivo de salida con nombre generico
        OuFile = open("first", 'w')
#Abrmos la hoja de calculo
        book = xlrd.open_workbook(sys.argv[1])
        #print "El numero de hojas es:", book.nsheets
        #print "nombre(s) de las hojas:", book.sheet_names()
#Tomamos la primer Hoja
        sh = book.sheet_by_index(0)
#Imprimimos  nombre de hoja, filas y columnas
        print sh.name, sh.nrows, sh.ncols  
        for i in range(1,sh.nrows):
        #imprime de linea procesado
            print "Procesando fila:"+str(i)

        #manipulacion con formato de fecha xlrd y separando en variables la tupla para fecha
            Fecha = xlrd.xldate.xldate_as_tuple(sh.cell_value(rowx=i,colx=1),0)
            Anio, Mes, Dia = Fecha[0], "%02d"%Fecha[1], "%02d"%Fecha[2]
            
        #Agregamos comando sql y nombre de tabla al inicio
            if i == 1:
                OuFile.writelines("Insert Into medi"+str(Anio)+str(Mes)
                        +"\n(eq_id,med_fecha,med_kwh) Values\n") 
        #manipulacion con formato de fecha xlrd y separado en variable la tupla para hora
            Hora = xlrd.xldate.xldate_as_tuple(sh.cell_value(rowx=i,colx=2),0) 
            Horas, Minutos, Segundos =  "%02d"%Hora[3], "%02d"%Hora[4], "%02d"%Hora[5]
        #agregamos todas las celdas como filas en sql excepto la ultima
            if i < (sh.nrows-1):
                OuFile.writelines("('"+str(sh.cell_value(rowx=i,colx=0))+"', '"+str(Anio)+"-"+str(Mes)
                        +"-"+str(Dia)+" "+str(Horas)+":"+str(Minutos)+":"
                        +str(Segundos)+"', '"+str(sh.cell_value(rowx=i,colx=3))+"'),\n")
        #agregamos la ultima fila con terminacion ; para finalizar comando sql
            else:
                OuFile.writelines("('"+str(sh.cell_value(rowx=i,colx=0))+"', '"+str(Anio)+"-"+str(Mes)
                        +"-"+str(Dia)+" "+str(Horas)+":"+str(Minutos)+":"
                        +str(Segundos)+"', '"+str(sh.cell_value(rowx=i,colx=3))+"');\n")

        print "hecho! ;)"
#cerramos el archivo de salida
        OuFile.close()
#rescribimos el nombre del archivo agregando anio+mes+sql
        rename('first',(sys.argv[1].split(".")[0])+str(Anio)+"-"+str(Mes)+".sql")
    else:
        print Help
