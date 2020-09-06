Attribute VB_Name = "Módulo1"
Option Explicit
Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim Dato As String
Dim Resul As Boolean
Dim Cantidad As Integer
Dim Posicion As Integer
Dim inver As Variant
Dim Fila As Integer
Sub CuentaFilas()
'*************************************************
'Determinación de la primera fila vacía / Contador de filas ocupadas (Cantidad). Cuidado: incluye la fila de título
'**************************************************
    Fila = 2
    While Worksheets("Hoja1").Cells(Fila, 1) <> ""
        Fila = Fila + 1
    Wend
    Cantidad = Fila - 1
End Sub
'***************************
'Botón LIMPIAR
'***************************
Sub LIMPIAR()
'Para cada elemento del arreglo.
' limpio la celda de Hoja1
    Worksheets("Hoja1").Range("a2:e100").Clear
End Sub
'***************************
'Botón BUSCAR
'***************************
Sub Buscar()
'Leo el nombre del estudiante que voy a buscar. Uso InputBox para simplificar la programación.
    Dato = InputBox("Nombre del estudiante:")
'Para la búsqueda voy a aplicar una estrategia que me ayude a determinar si encuentro o no el dato en el arreglo.
'Inicializo la variable Resul en False para establecer de entrada que el dato no está en el arreglo. Si en el
'proceso de búsqueda encuentro el dato cambiaré el valor de la variable Resul a True. Así sabré si el dato está
'o no está con una estrategia simple.
'Inicializo respuesta
    Resul = False
'Busco el nombre en cada registro del arreglo. Aunque el arreglo es de 2 dimensiones, en este caso estoy buscando
'por nombre (primer elemento). Podría buscar por cualquier otro elemento.
    Call CuentaFilas    'Cuento las filas ocupadas
    For I = 2 To Cantidad
        If Sheets("Hoja1").Cells(I, 1) = Dato Then  'Si el contenido de la celda es igual al nombre que busco
                Resul = True                        'Lo encontré, está en el arreglo
                Posicion = I                        'Fila donde está
        End If
    Next I
'Si lo encontré reporto en cuál fila, si no lo encontré reporto que no aparece
    If Resul Then
        MsgBox ("Se encuentra en la fila " & Posicion)
    Else
        MsgBox ("No aparece")
    End If
End Sub
'************************
'Boton ORDENAR**
'************************
Sub ORDENAR()
'********************************************************
'Algoritmo para ordenar los registros del arreglo (Método de la burbuja)
'********************************************************
'Este algoritmo ordena en sentido ascendente usando como criterio el elemento que está en la columna 1, es
'decir, ordena por nombre ascendente.
    Call CuentaFilas
    With Sheets("Hoja1")
        For I = 2 To (Cantidad - 1)                     'Desde el primer registro hasta el penúltimo
            For J = (I + 1) To Cantidad                 'Desde el segundo hasta el último
                If .Cells(I, 1) > .Cells(J, 1) Then     'Se compara cada uno con todos los que están debajo
                    Call INVERTIR                       'Si se cumple la condición, se invierten los registros
                End If                                  'Si no se cumple, no se hace nada
            Next J
        Next I
    End With
End Sub
'***************************

Sub INVERTIR()
'************************************************************************
'Algoritmo para invertir los registros que no están en el orden correcto
'************************************************************************
    With Sheets("Hoja1")
'Debo invertir los registros completos, es decir, debo invertir los datos que están en todas las columnas
'del registro o descompondré los datos. Son 5 elementos (5 columnas).
        For K = 1 To 5                  'Uso K para controlar las columnas e invierto desde la 1 hasta la 5.
            inver = .Cells(I, K)        'Para cada columna: Guardo en una variable el contenido de la celda en I
            .Cells(I, K) = .Cells(J, K) 'Para cada columna: Copio en la celda en I el contenido de la celda en J
            .Cells(J, K) = inver        'Para cada columna: Escribo en la celda en J el contenido de la variable
        Next K                          'Próxima columna
    End With
End Sub
'************************************
'Boton AGREGAR EN UNA POSICIÓN FIJA**
'************************************
Sub AGREGAR_POSICION()
'*******************************************************************************************************
'Sub para ingresar un dato en una posición seleccionada por el usuario
'*******************************************************************************************************
    Call CuentaFilas
    If Cantidad < 2 Then
        MsgBox ("No se puede agregar en una posición, el arreglo está vacío. Agregue al final")
    Else
'Leo la posición donde voy a agregar el elemento
        Posicion = 0    'La inicializo en cero
        While Posicion < 1 Or Posicion > Cantidad   'Me aseguro de que el usuario introduzca una posición válida
            Posicion = InputBox("Posición donde se va a agregar:")
        Wend
'Bajo una fila todos los registros que están por debajo de la posición para abrir una línea en la que
'insertar el nuevo registro. Tengo que bajar primero la última fila (la copio en la siguiente), luego bajo
'la penúltima fila (la copio en la siguiente) y así sucesivamente.
        For I = Cantidad To Posicion Step -1
'Debo bajar los registros completos, es decir, debo bajar los datos que están en todas las columnas
'del registro o descompondré los datos. Son 5 elementos (5 columnas).
            For K = 1 To 5      'Uso K para controlar las columnas
                Sheets("Hoja1").Cells(I + 1, K) = Sheets("Hoja1").Cells(I, K)
            Next K
        Next I
'limpio la posición (todos los elementos del registro, son 5)
        For K = 1 To 5
            Sheets("Hoja1").Cells(Posicion, K).Clear
'Pido el dato que se debe colocar en cada elemento de esa posición y lo escribo
            Sheets("Hoja1").Cells(Posicion, K) = InputBox("Valor del elemento:")
        Next K
    End If
End Sub

'************************
'Boton AGREGAR AL FINAL**
'************************
Sub AGREGAR_FINAL()
'*******************************************************************************************************
'Sub para ingresar un dato al final del arreglo
'*******************************************************************************************************
    Call CuentaFilas
'Pido el nuevo dato (cada uno de los 5 elementos) y lo escribo en la primera fila vacía (ahora la última del arreglo)
    For K = 1 To 5
        Sheets("Hoja1").Cells(Fila, K) = InputBox("Valor del elemento:")
    Next K
End Sub

'************************
'Boton AGREGAR EN ORDEN**
'************************
Sub AGREGAR_ORDEN()
'*******************************************************************************************************
'Sub para ingresar un dato en donde corresponde por orden ascendente
'*******************************************************************************************************
'Primero ordeno para asegurarme que está ordenado. Voy a ordenar por nombre puesto que el Sub ORDENAR
'utiliza como criterio de comparación la columna 1. Si quisiera ordenar por la Nota 4 tendría que utilizar
'como criterio para ordenar la columna 5. Si queremos que el usuario escoja por cuál columna quiere ordenar
'el número de columna correspondiente al criterio podría pasarse como parámetro al Sub ORDENAR.
    Call ORDENAR
'Pido el nombre del estudiante que voy a agregar
    Dato = InputBox("Nombre del estudiante:")
'Uso la variable Resul para determinar si por el orden al registro le corresponde estar al final de la tabla
    Resul = False
    Call CuentaFilas
'Busco la posición donde va
    For I = 2 To Cantidad
        If Sheets("Hoja1").Cells(I, 1) > Dato And Resul = False Then
'Coloco todos los registros siguientes una fila más abajo
            For J = Cantidad To I Step -1
                For K = 1 To 5
                    Sheets("Hoja1").Cells(J + 1, K) = Sheets("Hoja1").Cells(J, K)
                Next K
            Next J
'Ya tengo disponible la fila (es la I). Ahora pido y escribo cada elemento del registro.
            Sheets("Hoja1").Cells(I, 1) = Dato  'El nombre ya lo había leído, lo escribo en la columna 1
            Sheets("Hoja1").Cells(I, 2) = InputBox("Nota 1:")
            Sheets("Hoja1").Cells(I, 3) = InputBox("Nota 2:")
            Sheets("Hoja1").Cells(I, 4) = InputBox("Nota 3:")
            Sheets("Hoja1").Cells(I, 5) = InputBox("Nota 4:")
            Resul = True
         End If
    Next I
'Si entra en este if es porque por el orden el dato va al final
    If Resul = False Then
        Sheets("Hoja1").Cells(Cantidad + 1, 1) = Dato
        Sheets("Hoja1").Cells(I, 2) = InputBox("Nota 1:")
        Sheets("Hoja1").Cells(I, 3) = InputBox("Nota 2:")
        Sheets("Hoja1").Cells(I, 4) = InputBox("Nota 3:")
        Sheets("Hoja1").Cells(I, 5) = InputBox("Nota 4:")
    End If
End Sub

'******************************
'Boton ELIMINAR DE UNA POSICION
'******************************
Sub ELIMINAR_POSICION()
    Call CuentaFilas
'Leo la posición donde está el registro que voy a eliminar
    Posicion = 0
    While Posicion < 1 Or Posicion > Cantidad
        Posicion = InputBox("Posición donde se va a eliminar:")
    Wend
'Al eliminar un registro, todos los demás deben subir para llenar la fila que desaparece
'Subo todos los registros siguientes, copiándolos una fila más arriba
    For I = Posicion To Cantidad - 1
        For K = 1 To 5  'Debo subir todos los elementos de cada registro, son 5
            Sheets("Hoja1").Cells(I, K) = Sheets("Hoja1").Cells(I + 1, K)
        Next K
    Next I
'Cuando termino la última fila está duplicada. Debo limpiarla (los 5 elementos)
    For K = 1 To 5
        Sheets("Hoja1").Cells(Cantidad, K) = ""
    Next K
End Sub

'******************************
'Boton ELIMINAR UN VALOR
'******************************
Sub ELIMINAR_VALOR()
    Call CuentaFilas
'Pido el elemento a eliminar
    Dato = InputBox("Nombre del estudiante que se va a eliminar:")
    For I = Cantidad To 1 Step -1
        If Sheets("Hoja1").Cells(I, 1) = Dato Then
'Al eliminar un registro, todos los demás deben subir para llenar la fila que desaparece
'Subo todos los registros siguientes, copiándolos una fila más arriba
            For J = I To Cantidad - 1
                For K = 1 To 5  'Debo subir todos los elementos de cada registro, son 5
                    Sheets("Hoja1").Cells(J, K) = Sheets("Hoja1").Cells(J + 1, K)
                Next K
            Next J
'Cuando termino la última fila está duplicada. Debo limpiarla (los 5 elementos)
            For K = 1 To 5
                Sheets("Hoja1").Cells(Cantidad, K) = ""
            Next K
        End If
    Next I
End Sub


'******************************
'Boton MODIFICAR DE UNA POSICION
'******************************
Sub MODIFICAR_POSICION()
    Call CuentaFilas
'Leo la posición donde voy a modificar el elemento validando que el usuario introduzca una posición válida
    Posicion = 0
    While Posicion < 1 Or Posicion > Cantidad
        Posicion = InputBox("Posición donde se va a modificar:")
    Wend
'En este caso se asume que el usuario va a modificar los 5 elementos y por lo tanto se piden todos.
'Generalmente se utilizan Forms, el usuario cambia los datos que quiere y luego se escriben todos. Para
'los modificados se escribirá el nuevo valor y para los no modificados se escribirá el mismo valor
'que ya tenían
    For K = 1 To 5
        Sheets("Hoja1").Cells(Posicion, K) = InputBox("Valor del elemento:")
    Next K
End Sub

'******************************
'Boton MODIFICAR UN VALOR
'******************************
Sub MODIFICAR_VALOR()
    Call CuentaFilas
'Uso la estrategia ya conocida para determinar si encuentro o no el elemento
    Resul = False
'Pido el elemento a modificar
    Dato = InputBox("Nombre del estudiante cuyos datos se van a modificar:")
'Busco el elemento en todo el arreglo
    For I = 1 To Cantidad
        If Sheets("Hoja1").Cells(I, 1) = Dato Then
                Resul = True
                For K = 1 To 5  'Pido los nuevos valores para los 3 elementos
                    Sheets("Hoja1").Cells(I, K) = InputBox("Nuevo Valor:")
                Next K
        End If
    Next I
    If Resul = False Then
        MsgBox ("No aparece")
    End If
End Sub


