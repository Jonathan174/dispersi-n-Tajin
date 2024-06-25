import tkinter as tk
from tkinter import filedialog
import shutil
import os
from openpyxl import Workbook, load_workbook


def sumarListas(lista1, lista2, inicio_sku):
    listaSumada=[0] * len(lista2)
    for i in range(inicio_sku, len(lista2)):
        if (lista2[i] is None or lista1[i] == "") and (lista1 is None or lista1[i] == ""):
            listaSumada[i]=0
        elif (lista2[i] is None or lista2[i]=="")and lista1 is not None:
            listaSumada[i] = lista1[i]
        elif lista2[i] is not None and (lista1 is None or lista1[i]==""):
            listaSumada[i] = lista2[i]
        else:
        #elif lista2[i] and lista1[i]:
            listaSumada[i]=int(lista1[i]) + int(lista2[i])
    return listaSumada

def seleccionar_archivo():
    # Configurar el filtro para mostrar todos los archivos
    tipos_archivo = [("Archivos de Excel", "*.xlsx")]
    ruta_archivo = filedialog.askopenfilename(title="Selecciona un archivo", filetypes=tipos_archivo)
    
    # Verificar si se seleccionó un archivo
    if ruta_archivo:
        # Ruta del archivo con el nombre "archivo_incentivos"
        ruta_destino = os.path.join(os.path.dirname(os.path.abspath(__file__)), "archivo_incentivos.xlsx")

        # Copiar el archivo seleccionado con el nombre "archivo_incentivos"
        shutil.copy(ruta_archivo, ruta_destino)
        #print("El archivo seleccionado se ha guardado como 'archivo_incentivos.xlsx'.")
    

def filtrar_datos(hoja_activa):
    # Obtener los datos de la hoja de trabajo
    datos = []
    indices = []
    inicio_sku = 0
    bandera_inicio_sku = True
    encabezados = []
    filas_encabezado = True
    for row in hoja_activa.iter_rows(values_only=True):
        if row[0] == "Zona Clave" and filas_encabezado:
            filas_encabezado = False
            datos.append(row)
            for j in range(len(row)):
                if row[j]:
                    indices.append(j)
                    #print(row[j], j)
                    #len(row[j]) es igual a 5, porque se toma la palabra cuota
                    if len(row[j]) == 5 and bandera_inicio_sku:
                        bandera_inicio_sku = False
                    elif len(row[j]) != 5 and bandera_inicio_sku:
                        inicio_sku+=1                                       #Indica en que indice de columna comenzarán los productos
        
        elif filas_encabezado:
            encabezados.append(row)
        
        else:
            if row[0] == None:
                datos.append(row)
            elif row[0].isdigit() and row[0] != "8010":
                datos.append(row)
            elif row[0] == "Z501" or row[0] == "Z505" or row[0].lower() == "total general":
                datos.append(row)
            


    datos = encabezados[len(encabezados)-3:] + datos

    filtro = []

    referencia_Colaborador = 11             #Es 11 porque empieza por la zona 1100, por lo que tomamos los primeros 2
    referencia_Zona = 1                     #Son 4 zonas, empezando por la 1
    
    suma_Colaboradores=[0]*1000
    suma_Zona=[0]*1000
    suma_totales = [0]*1000

    for k in range(len(datos)):
        fila_filtro =[]
        iteracion = 0
        for j in range(len(datos[k])):
            if j == indices[iteracion]:
                fila_filtro.append(datos[k][j])
                
                if iteracion >= len(indices):
                    iteracion = len(indices)-1
                else:
                    iteracion+=1
        
        if k >= 4:
            if fila_filtro[0]:
                if fila_filtro[0].isdigit():
                    #Bloque de código para sumar los valores de las zonas
                    if int(int(fila_filtro[0])/1000) == referencia_Zona:
                        
                        #Bloque de código para sumar los valores de los vendedores
                        if int(int(fila_filtro[0])/100) == referencia_Colaborador:
                            suma_Colaboradores = sumarListas(suma_Colaboradores, fila_filtro, inicio_sku)
                        else:
                            suma_Colaboradores[0]="NOMBRE"
                            filtro.append(suma_Colaboradores)
                            suma_Zona=sumarListas(suma_Zona, suma_Colaboradores, inicio_sku)
                            suma_Colaboradores=[0]*1000
                            suma_Colaboradores=sumarListas(suma_Colaboradores, fila_filtro, inicio_sku)
                            referencia_Colaborador = int(int(fila_filtro[0])/100)
                    
                    else:
                        suma_Colaboradores[0]="NOMBRE"
                        filtro.append(suma_Colaboradores)
                        suma_Zona=sumarListas(suma_Zona, suma_Colaboradores, inicio_sku)
                        suma_Colaboradores=[0]*1000
                        suma_Colaboradores=sumarListas(suma_Colaboradores, fila_filtro, inicio_sku)
                        referencia_Colaborador = int(int(fila_filtro[0])/100)
                        if referencia_Zona == 1:
                            suma_Zona[0]="TOTAL PACIFICO - OCCIDENTE"
                        if referencia_Zona == 2:
                            suma_Zona[0]="TOTAL NORTE"
                        if referencia_Zona == 3:
                            suma_Zona[0]="TOTAL CENTRO-VDM"
                        if referencia_Zona == 4:
                            suma_Zona[0]="TOTAL SUR"
                        filtro.append(suma_Zona)
                        suma_Zona=[0]*1000
                        #suma_Zona=sumarListas(suma_Zona, fila_filtro)
                        referencia_Zona = int(int(fila_filtro[0])/1000) 
                    
                #if referencia_Zona
                elif not fila_filtro[0].isdigit() and referencia_Zona == 4:
                    suma_Colaboradores[0]="NOMBRE"
                    filtro.append(suma_Colaboradores)
                    suma_Zona=sumarListas(suma_Zona, suma_Colaboradores, inicio_sku)
                    suma_Colaboradores=[0]*1000
                    suma_Colaboradores=sumarListas(suma_Colaboradores, fila_filtro, inicio_sku)
                    referencia_Colaborador = 11
                    if referencia_Zona == 1:
                        suma_Zona[0]="TOTAL PACIFICO - OCCIDENTE"
                    if referencia_Zona == 2:
                        suma_Zona[0]="TOTAL NORTE"
                    if referencia_Zona == 3:
                        suma_Zona[0]="TOTAL CENTRO-VDM"
                    if referencia_Zona == 4:
                        suma_Zona[0]="TOTAL SUR"
                    filtro.append(suma_Zona)
                    suma_Zona=[0]*1000
                    #suma_Zona=sumarListas(suma_Zona, fila_filtro)
                    
                    referencia_Zona = 5


                if fila_filtro[0].lower() != "total general":
                    suma_totales = sumarListas(suma_totales,fila_filtro,+ inicio_sku)
                
                if fila_filtro[0].lower() == "total general":
                    for i in range(1, len(fila_filtro)):
                        fila_filtro[i]=suma_totales[i]

        filtro.append(fila_filtro)

    #print(suma_totales)    
    return filtro, inicio_sku


def calcular_porcentaje(cuota, venta):
    # Verificar si cuota o venta son cadenas vacías, None o cero
    if (not cuota and venta):
        if venta>0:
            return 0, venta, 1
        else:
            return 0, venta, 0
    elif (not venta and cuota):
        return cuota, 0, 0
    elif not cuota and not venta:
        return cuota, venta, cuota  # Manejar el caso cuando uno de los valores es None
    else:
        # Calcular el porcentaje como (venta / cuota) * 100
        #return cuota, venta, round(((venta / cuota) * 100), 2)
        return cuota, venta, round(venta / cuota)

# Crear una ventana de tkinter
ventana = tk.Tk()

# Ocultar la ventana principal
ventana.withdraw()
