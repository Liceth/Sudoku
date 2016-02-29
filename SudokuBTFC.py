# -*- coding: cp1252 -*-
'''
Created on 14/08/2011

@author: liceth
'''

import xlrd

from pyExcelerator import *

#Esta sera la direccion en donde se encontrara el archivo de Excel del Sudoku Sin Resolver
directorio = "/home/liceth/Escritorio/"

tablero = list()
posibilidades = list()
jugadas = list()

#Leer el Archivo de Excel
try:
    #Se debe Especificar el Nombre del Archivo que contiene el Sudoku sin resolver, es este caso prueba.xls
    book = xlrd.open_workbook(directorio+"prueba.xls")
    sheet = book.sheet_by_index(0)
    for i in range(9):
        tablero.append(list())
        for j in range(9):
            valor = sheet.cell(i,j).value
            if valor == '':
                valor = 0
            tablero[i].append(int(valor))
except IOError:
    print "El Archivo No Existe"


def valores_restrictos_posibilidades(tablero):
    #Busca las Posibilidades por Filas
    #Recibe por parametro el tablero y devulve lista con Posibilidades X Fila
    posibilidades = list()
    for i in range(len(tablero)):
        posibilidadesfila = list()
        for j in range(len(tablero[i])):
            fila = list()
            columna = list()
            dominio = list()
            if tablero[i][j] == 0:
                #Hallar Elementos de la Fila
                for l in range(len(tablero[i])):
                    if tablero[i][l] != 0:
                        fila.append(tablero[i][l])
                #Hallar Elementos de la Columna
                for l in range(len(tablero)):
                    for m in range(len(tablero[l])):
                        if m == j:
                            if tablero[l][m] != 0:
                                columna.append(tablero[l][m])
                #Busca en que Cuadrante se encuentra la Coordena i , j
                #Para luego buscar sus posibilidades respectivas
                if i <= 2:
                    l = 0
                    m = 3
                    if j <= 2:
                        n = 0
                        o = 3
                    elif j <= 5:
                        n = 3
                        o = 6
                    elif j <= 8:
                        n = 6
                        o = 9
                elif i <= 5:
                    l = 3
                    m = 6
                    if j<= 2:
                        n = 0
                        o = 3
                    elif j <= 5:
                        n = 3
                        o = 6
                    elif j <= 8:
                        n = 6
                        o = 9
                elif i <= 8:
                    l = 6
                    m = 9
                    if j<= 2:
                        n = 0
                        o = 3
                    elif j <= 5:
                        n = 3
                        o = 6
                    elif j <= 8:
                        n = 6
                        o = 9
                for p in range(l,m):
                    for q in range(n,o):
                        if p != i and j != q and tablero[p][q] != 0:
                            dominio.append(tablero[p][q])
                posibilidadescelda = list()
                for p in range(1,10):
                    if fila.count(p) == 0 and columna.count(p) == 0 and dominio.count(p) == 0:
                        posibilidadescelda.append(p)
                posibilidadesfila.append(posibilidadescelda)
        #print "Posibilidades Fila" , i+1
        #print posibilidadesfila
        posibilidades.append(posibilidadesfila)
    return posibilidades

def valores_p(tablero,i,j,posibilidades):
    posibles_valores=[]
    cont=0
    k=0
    # primero a buscar la fila de la posicion q esta vacia dentro de la matris de posibilidades
    while cont<j or k<len(tablero):
        if tablero[i][k]==0:
            cont=cont+1
        k=k+1
    # ya con la fila encontrada saco el vector de probabilidad de la casilla(variable)
    print cont
    posibles_valores=posibilidades[i][cont]

    k=0
    # ya con los valores encontrados paso a calcular los valores de la probabilidades para cada posible valor
    for k in range(len(posibles_valores)):
        valor=posibles_valores[k]
        probvalor=[]
        for z in range(len(tablero)):
            if tablero[z][j] == 0:
                    l=0
                    cont2=0
                    while cont2!=j or l<len(tablero):
                            if tablero[z][l]==0:
                                cont2=cont2+1
                                l=l+1
                    p_valores=posibilidades(z,cont2)

                    if p_valores.__contains__(k):
                        probvalor[k]=probvalor[k]+len(p_valores)-1
                    else:
                        probvalor[k]=probvalor[k]+len(p_valores)
    print "salio!!!!!"
    return posibles_valores,probvalor

def cambiar_n_posibilidad(tablero,posibilidades,n):
    for i in range(len(tablero)):
        cont = 0
        for j in range(len(tablero[i])):
            #print len(posibilidades[i][cont])
            if tablero[i][j] == 0:
                if len(posibilidades[i][cont]) == n:
                    posibilidades_celda = posibilidades[i][cont]
                    for i in posibilidades_celda:
                        for k in range(i+1):
                            if tablero[k][j] == 0:
                                cont1 = 0
                                for l in range(j+1):
                                    if tablero[k][l] == 0:
                                        cont1 = cont1 + 1


                cont = cont + 1

def no_lleno(tablero):
    for i in tablero:
        for j in i:
            if j == 0:
                return True
    return False

def calcular_jugadas_totales(tablero, posibilidades):
    jugadas = []
    for i in range(len(tablero)):
        cont = 0
        for j in range(len(tablero[i])):
            if tablero[i][j] == 0:
                jugadas.append(len(posibilidades[i][cont]))
                cont = cont + 1
    return jugadas

def valores_menos_restrictos(tablero,posibilidades,i,j):
    #Busca los valores menos restrictores
    #Recibe el tablero , las posibilidades para este mismo y la respectiva
    #Posicion i , j sobre la cual se desean buscar los valores
    posibles_valores=list()
    cont=0
    k=0
    #Primero a buscar la fila de la posicion q esta vacia dentro de la matris de posibilidades
    while k<=j :
        if tablero[i][k]==0:
            cont=cont+1
        k=k+1
    #Ya con la fila encontrada saco el vector de probabilidad de la casilla(variable)
    cont=cont-1;
    posibles_valores=posibilidades[i][cont]


    # ya con los valores encontrados paso a calcular los valores de la probabilidades para cada posible valor
    probvalor=list()
    #Creo la lista con 0
    k=0
    for k in range(len(posibles_valores)):
        probvalor.append(0)
    k=0
    for k in range(len(posibles_valores)):
        valor=posibles_valores[k]

        for z in range(len(tablero)):
            #Primero busco los vacios en su columna
            if tablero[z][j] == 0 and z!=i:
                    l=0
                    cont2=0
                    #Ahora q encontre uno vacio busco su columna en probabilidades
                    while l<=j :
                            if tablero[z][l]==0 :
                                cont2=cont2+1
                            l=l+1
                    cont2=cont2-1;
                    p_valores=posibilidades[z][cont2]


                    #Ahora q encontre la columna y posteriormente en vector de posivilidades pregunto si el valor esta entre las posibilidades de esa casilla
                    #Voy sumando, las posbilidades
                    if p_valores.__contains__(valor):
                        probvalor[k] = probvalor[k] + 1
                    else:
                        probvalor[k] = probvalor[k] + 0
            if tablero[i][z] == 0 and z!=j:
                    l=0
                    cont2=0
                    #Ahora q encontre uno vacio busco su columna en probabilidades
                    while l<=z :
                            if tablero[i][l]==0 :
                                cont2=cont2+1
                            l=l+1
                    cont2=cont2-1;
                    p_valores=posibilidades[i][cont2]


                    #Ahora q encontre la columna y posteriormente en vector de posivilidades pregunto si el valor esta en entre las posibilidades del esa casilla
                    # y se van sumando las posibilidades quitandoles el numero escogido hipoteticamente
                    if p_valores.__contains__(valor):
                        probvalor[k]=probvalor[k]+1
                    else:
                        probvalor[k]=probvalor[k]+0
    #Ahora busco en su cuadrado respectivo
        fil=0
        col=0
        if i<=2:
            fil=0
        if i<=5 and i>2:
            fil=3
        if i<=8 and i>5:
            fil=6

        if j<=2:
            col=0
        elif j<=5 and j>2:
            col=3
        elif j<=8 and j>5:
            col=6

        for ii in range(fil,fil+3):
            for jj in range(col,col+3):

                if tablero[ii][jj] == 0 and  ii!=i and jj!=j :
                        print
                        l=0
                        cont2=0
                        #Ahora q encontre uno vacio busco su columna en probabilidades
                        while l<=jj :
                                if tablero[ii][l]==0 :
                                    cont2=cont2+1
                                l=l+1
                        cont2=cont2-1;

                        p_valores=posibilidades[ii][cont2]

                        #Ahora q encontre la columna y posteriormente en vector de posivilidades pregunto si el valor esta
                        #En entre las posibilidades del esa casilla
                        #Y se van sumando las posibilidades quitandoles el numero escogido hipoteticamente
                        if p_valores.__contains__(valor):
                            probvalor[k]=probvalor[k]+1
                        else:
                            probvalor[k]=probvalor[k]+0

    #Ordenar Vectores del menos restrictos al mayor para efectos del Backtrack
    for g in range(len(probvalor)):
        for m in range(len(probvalor)):
            if(probvalor[g]<probvalor[m]):
                temp1=probvalor[g]
                probvalor[g]=probvalor[m]
                probvalor[m]=temp1

                temp2=posibles_valores[g]
                posibles_valores[g]=posibles_valores[m]
                posibles_valores[m]=temp2
    return posibles_valores


posibilidades = valores_restrictos_posibilidades(tablero)
#print posibilidades

sw = True
contx = 0
#Cambiar el respectivo directorio sobre el cual se colocara el log de las acciones que se realizan sobre
#El tablero
arch = open(directorio+"log.txt","w")
arch.writelines("--------------------------------------SUDOKU--------------------------------\n")
arch.writelines("--Tablero--\n")
for i in tablero:
    arch.writelines(str(i)+"\n")
arch.writelines("--Fin Tablero Sudoku--\n")

while sw:
    #print contx
    posibilidades = valores_restrictos_posibilidades(tablero)
    ultima_jugada = 0
    i = 0
    j = 0
    while i < len(tablero):
        cont = 0
        while j < len(tablero[i]):
            cont1 = 0
            if tablero[i][j] == 0:
                #print "Posibilidades: " , posibilidades[i]
                #arch.writelines("Fila " + str(i) + "\n")
                #arch.writelines("Posibilidades: " + str(posibilidades[i]) + "\n")
                #arch.writelines("Contador: "+ str(cont)+"\n")
                #arch.writelines("Tablero: "+str(tablero)+"\n")
                #print "Contador: ", cont
                #print "Tablero: ",tablero
                #print "i : ",i , " j: " ,j

                if len(posibilidades[i][cont]) == min(calcular_jugadas_totales(tablero, posibilidades)):
                    arch.writelines("Se ha seleccionado la casilla con fila "+str(i+1)+" y columna "+str(j)+"\n")
                    valores_restrictos = valores_menos_restrictos(tablero,posibilidades,i,j)
                    arch.writelines("Para esta casilla tenemos los siguientes valores ordenado del menos restricto al mas restricto:\n")
                    arch.writelines(str(valores_restrictos)+"\n")
                    temp = []
                    jugar = 0
                    for k in range(len(valores_restrictos)):
                        if valores_restrictos[k] != ultima_jugada:
                            jugar = valores_restrictos[k]
                            ultima_jugada = 0
                            cont1 = k
                            break
                    temp.append(jugar)
                    temp.append(i)
                    temp.append(j)
                    temp.append(valores_restrictos)
                    jugadas.append(temp)
                    tablero[i][j] = jugar
                    contx = contx + 1
                    arch.writelines("----------------------------------------------------------------------------\n")
                    arch.writelines("--Tablero--\n")
                    for h in tablero:
                        arch.writelines(str(h)+"\n")
                    arch.writelines("--Fin Tablero--\n")
                    arch.writelines("Asignacion en la fila "+str(i+1)+" Columna: "+str(j+1)+" Valor: "+str(jugar))
                    arch.writelines(" Nro. de Jugada: "+str(contx)+"\n")
                    arch.writelines("----------------------------------------------------------------------------\n")

                    posibilidades = valores_restrictos_posibilidades(tablero)
                    cont = - 1
                    i = 0
                    j = -1
                    #Ojo El Sudoku Es Valido
                    if len(calcular_jugadas_totales(tablero, posibilidades)) != 0:
                        while min(calcular_jugadas_totales(tablero, posibilidades)) == 0:
                            if (cont1 + 1) < len(valores_restrictos):
                                cont1 = cont1 + 1
                                temp = list()
                                temp.append(valores_restrictos[cont1])
                                temp.append(i)
                                temp.append(j)
                                temp.append(valores_restrictos)
                                jugadas[len(jugadas) - 1] = temp
                                tablero[i][j] = 0
                                arch.writelines("---------------------BackTracking--------------------------\n")
                                arch.writelines("Se ha reversado la posicion i: "+str(i+1)+" , j: "+str(i+1)+" con el Valor: "+str(temp[0])+"\n")
                                arch.writelines("---------------------Fin BackTracking--------------------------\n")
                                arch.writelines("--Tablero--\n")
                                for h in tablero:
                                    arch.writelines(str(h)+"\n")
                                arch.writelines("--Fin Tablero--\n")
                                tablero[i][j] = valores_restrictos[cont1]
                                contx = contx + 1
                                arch.writelines("----------------------------------------------------------------------------\n")
                                arch.writelines("--Tablero--\n")
                                for h in tablero:
                                    arch.writelines(str(h)+"\n")
                                arch.writelines("--Fin Tablero--\n")
                                arch.writelines("Asignacion en la fila "+str(i+1)+" Columna: "+str(j+1)+" Valor: "+str(valores_restrictos[cont1]))
                                arch.writelines(" Nro. de Jugada: "+str(contx)+"\n")
                                arch.writelines("----------------------------------------------------------------------------\n")
                                posibilidades = valores_restrictos_posibilidades(tablero)
                            else:
                                temp = jugadas.pop()
                                tablero[temp[1]][temp[2]] = 0
                                arch.writelines("---------------------BackTracking--------------------------\n")
                                arch.writelines("Se ha reversado la posicion i: "+str(temp[1]+1)+" , j: "+str(temp[2]+1)+" con el Valor: "+str(temp[0])+"\n")
                                arch.writelines("---------------------Fin BackTracking--------------------------\n")
                                arch.writelines("--Tablero--\n")
                                for h in tablero:
                                    arch.writelines(str(h)+"\n")
                                arch.writelines("--Fin Tablero--\n")
                                temp = jugadas.pop()
                                arch.writelines("---------------------BackTracking--------------------------\n")
                                arch.writelines("Se ha reversado la posicion i: "+str(temp[1]+1)+" , j: "+str(temp[2]+1)+" con el Valor: "+str(temp[0])+"\n")
                                arch.writelines("---------------------Fin BackTracking--------------------------\n")
                                arch.writelines("--Tablero--\n")
                                for h in tablero:
                                    arch.writelines(str(h)+"\n")
                                arch.writelines("--Fin Tablero--\n")
                                ultima_jugada = temp[0]
                                i = temp[1]
                                j = temp[2]
                                tablero[i][j] = 0
                                while len(temp[3]) == 1:
                                    temp = jugadas.pop()
                                    arch.writelines("---------------------BackTracking--------------------------\n")
                                    arch.writelines("Se ha reversado la posicion i: "+str(temp[1]+1)+" , j: "+str(temp[2]+1)+" con el Valor: "+str(temp[0])+"\n")
                                    arch.writelines("---------------------Fin BackTracking--------------------------\n")
                                    arch.writelines("--Tablero--\n")
                                    for h in tablero:
                                        arch.writelines(str(h)+"\n")
                                    arch.writelines("--Fin Tablero--\n")
                                    tablero[temp[1]][temp[2]] = 0
                                    temp = jugadas.pop()
                                    arch.writelines("---------------------BackTracking--------------------------\n")
                                    arch.writelines("Se ha reversado la posicion i: "+str(temp[1]+1)+" , j: "+str(temp[2]+1)+" con el Valor: "+str(temp[0])+"\n")
                                    arch.writelines("---------------------Fin BackTracking--------------------------\n")
                                    arch.writelines("--Tablero--\n")
                                    for h in tablero:
                                        arch.writelines(str(h)+"\n")
                                    arch.writelines("--Fin Tablero--\n")
                                    tablero[temp[1]][temp[2]] = 0
                                    ultima_jugada = temp[0]
                                i = 0
                                j = - 1
                                cont = - 1
                                posibilidades = valores_restrictos_posibilidades(tablero)
                cont = cont + 1
            j = j + 1
        j = 0
        i = i + 1
    contx= contx + 1
    sw = no_lleno(tablero)
contx = contx - 1
arch.writelines("Nro. Total de Asignaciones: "+str(contx))
wb = Workbook()
ws0 = wb.add_sheet('0')
for i in range(len(tablero)):
    for j in range(len(tablero[i])):
        ws0.write(i+1, j+1, tablero[i][j])
#Escribir el Directorio sobre el cual se colocara la respuesta
wb.save(directorio+"Salida.xls")
#print tablero
#print contx
print "Procedimiento Resuelto en sus Respectivos Archivos!!!!!"
print "Observar el archivo log.txt, detalles de la Salida"
print "Sudoku Resuelto Salida.xls"
arch.close()




