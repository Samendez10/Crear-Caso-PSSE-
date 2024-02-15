# -*- coding: cp1252 -*-
import os,sys
import win32com.client.dynamic
#import numpy as np   ### Este módulo es importante para hacer cuentas con python y que no haga cagada

class Excel:

    """A utility to make it easier to get at Excel. Remembering
    to save the data is your problem, as is error handling.
    Operates on one workbook at a time."""

    def __init__(self, filename=None, carpeta=None):
        from win32com.client import Dispatch
        import win32com.client.dynamic 
        self.xlApp = win32com.client.Dispatch('Excel.Application')
        if filename:
            self.filename = filename
            self.xlBook = self.xlApp.Workbooks.Open(carpeta+"\\"+filename)
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ""

    def save(self, newfilename=None):
        if newfilename:
            self.filename = newfilename
            self.xlBook.SaveAs(newfilename)
        else:
            self.xlBook.Save()

    def close(self):
        self.xlBook.Close(SaveChanges=0)
        del self.xlApp

    def getCell(self, sheet, row, col):
        "Get value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Cells(row, col).Value

    def setCell(self, sheet, row, col, value):
        "set value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Value = value

    def getRange(self, sheet, row1, col1, row2, col2):
        "return a 2d array (i.e. tuple of tuples)"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value

    def setRange(self, sheet, leftCol, topRow, data):
        """insert a 2d array starting at given location.
        Works out the size needed for itself"""

        bottomRow = topRow + len(data) - 1

        rightCol = leftCol + len(data[0]) - 1
        sht = self.xlBook.Worksheets(sheet)
        sht.Range(
        sht.Cells(topRow, leftCol),
        sht.Cells(bottomRow, rightCol)
        ).Value = data

    def Listaencolumna (self, sheet, filaprimercelda,columnaprimercelda,lista):
        "Inserta los elementos de una lista en una columna"
        N=len(lista)
        for i in range (0,N):
            sht = self.xlBook.Worksheets(sheet)
            sht.Cells(filaprimercelda+i, columnaprimercelda).Value = lista[i]   

    def fixStringsAndDates(self, aMatrix):
        # converts all unicode strings and times
        newmatrix = []
        for row in aMatrix:
            newrow = []
            for cell in row:
                if type(cell) is UnicodeType:
                    newrow.append(str(cell))
                elif type(cell) is TimeType:
                    newrow.append(int(cell))
                else:
                    newrow.append(cell)
            newmatrix.append(tuple(newrow))
        return newmatrix

    def Visible(self):
        # Hace visible el documento en el que se esta trabajando
        self.xlApp .Visible =1

    def Definorango(self,sheet,filainicial,colinicial,filafinal,colfinal):
        Hoja=self.xlApp.Workbooks(self.filename).Sheets(sheet)
        Rango=Hoja.Range(Hoja.Cells(filainicial,colinicial), Hoja.Cells(filafinal,colfinal))
        return Rango

    def Eliminarfilas(self,sheet,filas="fila1:fila1,fila2:fila2"):
        Hoja=self.xlApp.Workbooks(self.filename).Sheets(sheet)
        Filas=Hoja.Range(filas)
        Filas.Select()
        Filas.Delete()
        
    def Columnaenlista(self,sheet,Fila1,Columna1,Cantdatos):
        Lista=list()
        for i in range(0,Cantdatos):
            a=Excel.getCell(sheet,Fila1+i,Columna1)
            Lista.append(a)
        return Lista        

##sys.path.append(PSSE_LOCATION)
##os.environ['PATH'] = os.environ['PATH'] + ';' + PSSE_LOCATION
import redirect
import psspy
import pssarrays
from psspy import _i, _f, _c

raiz=os.getcwd()

##############################################################################################################


Excel=Excel("CREA PAT_PRUEBA-OCT23",raiz)
Excel.Visible()

finic=3
colinic=4
BANDA=1 ##Banda de asjute para P y Q en MW o MVAr
PASOP=1 ##Paso inicial para escalar P de la demanda en MW
PASOQ=1 ##Paso inicial para escalar Q de la demanda en MVAr

m=finic
while str(Excel.getCell('REACTORES_SIP',m,3))<>'None':
    bus=int(Excel.getCell('REACTORES_SIP',m,3))
    Binit=Excel.getCell('REACTORES_SIP',m,5)
    psspy.switched_shunt_chng_4(bus,[_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,Binit,_f],_s)
    m=m+1
    
## Ajuste de la demanda del nodo ESP
## barras del subsistema ESP
bus=[115,118,556,721]
## Se crea un subsistema para las demandas de ESP
psspy.bsys(0,0,[0.0, 525.],0,[],len(bus),bus,0,[],0,[])
## Se busca la consigna de P y Q el trafo de ESP
Pfin=Excel.getCell('RED PATAGONIA',62,3)
Qfin=Excel.getCell('RED PATAGONIA',63,3)
## bucle de escala de demanda en P
m=1
(ierr, S)=psspy.wnddt2(19,117,545,'1',"FLOW") #se lee el estado inicial del flujo por el trafo
print S
P=S.real
Q=S.imag
while (abs(P-Pfin)>BANDA or abs(Q-Qfin)>BANDA) and m<20:#se compara el estado del flujo con el flujo consigna. Se admiten hasta 30 iteraciones
    PASOP=max(abs(P-Pfin),1)
    PASOQ=max(abs(Q-Qfin),1)
    (ierr,totals,moto)=psspy.scal_2(0,0,1,[0,0,0,0,0],[0.0,0.0,0.0,0.0,0.0,0.0,0.0])# se leen los totales del subsistema
    #escala P
    print('escala P')
    if P<Pfin:#si el fluo está bajo se sube la demanda
        (ierr,totals,moto)=psspy.scal_2(0,1,2,[_i,1,0,2,0],[totals[1]+PASOP,0.0,0.0,-.0,0.0,-.0,totals[0]])
    else:#si el flujo está alto se baja la demanda
        (ierr,totals,moto)=psspy.scal_2(0,1,2,[_i,1,0,2,0],[totals[1]-PASOP,0.0,0.0,-.0,0.0,-.0,totals[0]])
    #escala Q
    print('escala Q')
    (ierr,totals,moto)=psspy.scal_2(0,0,1,[0,0,0,0,0],[0.0,0.0,0.0,0.0,0.0,0.0,0.0])# se leen los totales del subsistema
    if Q<Qfin:#si el fluo está bajo se sube la demanda
        (ierr,totals,moto)=psspy.scal_2(0,1,2,[_i,1,0,2,0],[totals[1],0.0,0.0,-.0,0.0,-.0,totals[0]+PASOQ])
    else:#si el flujo está alto se baja la demanda
        (ierr,totals,moto)=psspy.scal_2(0,1,2,[_i,1,0,2,0],[totals[1],0.0,0.0,-.0,0.0,-.0,totals[0]-PASOQ])
    psspy.fnsl([0,0,0,1,0,0,99,0])#se corre el flujo: todo Lock, not flat start
    (ierr, S)=psspy.wnddt2(19,117,545,'1',"FLOW")#se vuelve a leer el estado del trafo por el trafo
    print S
    P=S.real
    Q=S.imag
    m=m+1
print('m='+str(m))
## Se elimina el subsistema
psspy.bsys(0,0,[0.0, 525.],0,[],0,[],0,[],0,[])


Excel.close()
