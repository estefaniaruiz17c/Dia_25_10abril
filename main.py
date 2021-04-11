# Series con pandas utilizando archivos de excel
print("Series con pandas utilizando archivos de excel")

# Importar las librerías necesarias para el ejercicio
import openpyxl
import pandas as pd

# Crearemos la variable para leer el excel 1, contiene números enteros solamente
num_1 = pd.read_excel("pandas1.xlsx",sheet_name=0)
print(num_1,"\n")

# Esta variable recorre la fila 2 del excel
array1 = pd.Series(num_1.iloc[0,:])
#arr = nmp.array(array1)
# Esta variable recorre la fila 3 del excel
array2 = pd.Series(num_1.iloc[1,:])
print(array1,"\n")
print(array2,"\n")

print()
print()

# Procedermos a crear el archivo donde guardaremos las operaciones 

# Crear un archivo de Excel
opera_doc_1 = openpyxl.Workbook()

# Aquí asignamos una hoja de cálculo en blanco en el archivo creado en el paso anterior
hojacalculo = opera_doc_1.active

# Guardaremos lo que llevamos en el archivo con el nombre: 'opera1_excel_1.xlsx'
opera_doc_1.save('opera1_excel_1.xlsx')

# Ya creado, continuaremos diseñando algunas operaciones
hojacalculo['A1'] = ("Operaciones realizadas con la librería pandas de Pyhton")
hojacalculo['A3'] = ("Ejercicios parte 1")
hojacalculo['A4'] = 1
hojacalculo['A5'] = 2
hojacalculo['A6'] = 3
hojacalculo['A7'] = 4

# Con esto, haremos algunos ejercicios

# Ejercicio 1: size - devuelve el número de elementos de la serie array1
hojacalculo['B4'] = ("Número de elementos de la fila 2 del archivo 'pandas1.xlsx")
print("size")

# Creación de la operación
size1_array1 = array1.size
print(size1_array1)

# Almacenar la respuesta en una celda de 'opera1_excel_1.xlsx'
hojacalculo['C4'] = size1_array1


# Ejercicio 2: sdt - desviación típica de los datos de la serie (datos de tipo numérico)
hojacalculo['B5'] = ("Desviación típica de los datos de la fila 3 de 'pandas1.xlsx':")
print("std()")

# Creación da la operación
std1_array2 = array2.std()
print(std1_array2)

# Almacenar la respuesta en una celda de 'opera1_excel_1.xlsx'
hojacalculo['C5'] = std1_array2


# Ejercicio 3: min - devuelve el menor de los datos de la serie
hojacalculo['B6'] = ("El menor de los datos de la fila 2 de 'pandas1.xlsx':")
print("min")

# Creación da la operación
min1_array1 = array1.min()
print(min1_array1)

# Almacenar la respuesta en una celda de 'opera1_excel_1.xlsx'
hojacalculo['C6'] = min1_array1

# Ejercicio 4: max - devuelve el mayor de los datos de la serie
hojacalculo['B7'] = ("El mayor de los datos de la fila 3 de 'pandas1.xlsx':")
print("max")

# Creación da la operación
max1_array2 = array2.max()
print(max1_array2)

# Almacenar la respuesta en una celda de 'opera1_excel_1.xlsx'
hojacalculo['C7'] = max1_array2

print()
git = array1[1:3]
print((git))

# Guadaremos los ejercicios realizados
opera_doc_1.save('opera1_excel_1.xlsx')

print("---------"*8)

# Crearemos la variable para leer el excel 2, contiene números decimales y letras
num_2 = pd.read_excel("pandas2.xlsx",sheet_name=0)
print(num_2,"\n")

# Esta variable recorre la fila 2 del excel
array1_num2 = pd.Series(num_2.iloc[0,:])

# Esta variable recorre la fila 3 del excel
array2_num2 = pd.Series(num_2.iloc[1,:])
print(array1_num2,"\n")
print(array2_num2,"\n")

# Procedermos a crear el archivo donde guardaremos las operaciones 

# Crear un archivo de Excel
opera_doc_2 = openpyxl.Workbook()

# Aquí asignamos una hoja de cálculo en blanco en el archivo creado en el paso anterior
hojacalculo2 = opera_doc_2.active

# Guardaremos lo que llevamos en el archivo con el nombre: 'opera2_excel_2.xlsx'
opera_doc_2.save('opera2_excel_2.xlsx')

# Ya creado, continuaremos diseñando algunas operaciones
hojacalculo2['A1'] = ("Operaciones realizadas con la librería pandas de Pyhton")
hojacalculo2['A3'] = ("Ejercicios parte 2")
hojacalculo2['A4'] = 1
hojacalculo2['A5'] = 2
hojacalculo2['A6'] = 3
hojacalculo2['A7'] = 4

# Ejercicio 1: mean() - devuelve la media de los datos de la serie
hojacalculo2['B4'] = ("Media de los datos de la fila 2 del archivo 'pandas2.xlsx':")
print("mean")

# Creación da la operación
mean2_array1_num2 = array1_num2.mean()
print(mean2_array1_num2 )

# Almacenar la respuesta en una celda de 'opera2_excel_2.xlsx'
hojacalculo2['C4'] = mean2_array1_num2 


# Ejercicio 2: suma() - suma los datos si son númericos, si son string, los concatena
hojacalculo2['B5'] = ("Suma de los datos de la fila 2 del archivo 'pandas2.xlsx':")
print("suma1")

# Creación da la operación
suma1_2_array1_num2 = array1_num2.sum()
print(suma1_2_array1_num2)

# Almacenar la respuesta en una celda de 'opera2_excel_2.xlsx'
hojacalculo2['C5'] = suma1_2_array1_num2


# Ejercicio 3: sum() - suma los datos si son númericos, si son string, los concatena
hojacalculo2['B6'] = ("Concatenación de los datos de la fila 3 del archivo 'pandas2.xlsx':")
print("suma2")

# Creación da la operación
suma2_array2_num2 = array2_num2.sum()
print(suma2_array2_num2)

# Almacenar la respuesta en una celda de 'opera2_excel_2.xlsx'
hojacalculo2['C6'] = str(suma2_array2_num2 )


# Ejercicio 4: count() - devuelve el número de elementos que no son nulos
hojacalculo2['B7'] = ("Número de datos no nulos de la fila 3 del archivo 'pandas2.xlsx':")
print("count")

# Creación da la operación
count2_array2_num2 = array2_num2.count()
print(count2_array2_num2)

# Almacenar la respuesta en una celda de 'opera2_excel_2.xlsx'
hojacalculo2['C7'] = (count2_array2_num2 )


# Guadaremos los ejercicios realizados
opera_doc_2.save('opera2_excel_2.xlsx')