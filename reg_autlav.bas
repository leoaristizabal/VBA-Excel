Attribute VB_Name = "Módulo1"
Option Explicit
Dim tipo As String
Dim uso As String
Dim placa As String
Dim hay As Boolean
Dim Fila As Integer
Dim FilaAux As Integer
Dim FilaDatos As Integer
Dim FilaResp As Integer
Dim I As Integer
Dim K As Integer
Dim J As Integer
Dim Fil As Integer
Dim Col As Integer
Dim Nombre As String
Dim Mayor As Integer
Dim Respuesta As String
Dim NombreHoja As String
Dim Posicion As Integer
Dim Criterio As Integer
Dim Inver As String
Sub Autos()
'**********************************************************************************************
'Sub para agregar autos al inventario
'**********************************************************************************************
    Call CuentaFilas
    form_datos.txt_fecha = Format(Date, "d/mmm/yyyy")
    form_datos.Show
End Sub
Sub CuentaFilas()
'********************************************************************************
'Sub que encuentra la primera fila vacía en la hoja de inventario
'********************************************************************************
    Fila = 2
    While Worksheets("Inventario de Autos").Cells(Fila, 1) <> ""
        Fila = Fila + 1
    Wend
End Sub
Sub EscRep()
'*********************************************************************************
'Sub para escribir los datos de los autos en la hoja de inventario
'*********************************************************************************
    With Sheets("Inventario de Autos")
        .Cells(Fila, 1) = form_datos.txt_placa
        .Cells(Fila, 2) = form_datos.cmb_tipo
        .Cells(Fila, 3) = form_datos.txt_color
        If (form_datos.rbtn_particular) Then
            .Cells(Fila, 4) = "Particular"
        Else
            .Cells(Fila, 4) = "Colectivo"
        End If
        .Cells(Fila, 5) = 0
        .Cells(Fila, 6) = "Disponible"
        .Cells(Fila, 7) = "Ok"
        .Cells(Fila, 8) = form_datos.txt_fecha
    End With
    ' actualiza el numero de filas del inventario
    Fila = Fila + 1
End Sub
Sub LimpiarForm()
'******************************************************************************
'Sub que limpia el form de inventario
'******************************************************************************
    With form_datos
        .txt_fecha = ""
        .txt_placa = ""
        .txt_color = ""
        .cmb_tipo = ""
        .rbtn_colectivo = False
        .rbtn_particular = True
        .txt_fecha = Format(Date, "d/mmm/yyyy")
    End With
End Sub
Sub LimpiarHoja()
'********************************************************************************
'Sub para limpiar la hoja de inventario
'********************************************************************************
    Call CuentaFilas
    With Sheets("Inventario de Autos")
        For I = 2 To Fila - 1
            For K = 1 To 8
                .Cells(I, K).Clear
            Next
        Next
    End With
End Sub
Sub alquilar()
'*********************************************************************************
'Sub para el proceso de alquilar un auto
'*********************************************************************************
'Como estrategia inicializo la variable "hay" indicando que no hay autos con las características solicitadas.
'Si después encuentro alguno, cambiaré el valor a True
hay = False
form_alquiler.Show
' Leer el tipo y uso del vehiculo a alquilar desde el form
tipo = form_alquiler.cmb_tipo
If (form_alquiler.rbtn_particular) Then
    uso = "Particular"
Else
    uso = "Colectivo"
End If
' Crear la hoja con los carros posibles a alquilar
Call crearDatosAlqui(tipo, uso, hay)
' Si encontré algún auto, pido la placa del vehiculo deseado de acuerdo a los datos que muestra la hoja Auxiliar
If hay = True Then
    placa = InputBox("Ingrese la placa del carro que prefiere alquilar:")

' Ya no necesitaré los datos que están en la hoja Auxiliar, la limpio
    With Sheets("Auxiliar")
        .Range(.Cells(2, 1), .Cells(50, 4)).Clear   'Otra forma de especificar un rango en VBA
    End With
' Copiar los datos del auto seleccionado en la hoja Datos de Alquiler
    Call copiarDatos(placa)
Else
    Sheets("Inventario de Autos").Select
    MsgBox ("No hay carros disponibles con esas caracteristicas")
End If
    
End Sub
Sub crearDatosAlqui(t As String, u As String, h As Boolean)
'*******************************************************************************************
'Sub para crear la hoja de datos auxiliar con los datos de los autos disponibles
'*******************************************************************************************
' inicializo las variables
FilaAux = 2 'Contador para controlar las filas de la hoja Auxiliar
Call CuentaFilas
' buscar los registros qeu coincidan con los datos pedidos
For I = 2 To Fila - 1
    If Sheets("Inventario de Autos").Cells(I, 2) = t Then                       'Si coincide el tipo
        If Sheets("Inventario de Autos").Cells(I, 4) = u Then                   'Si coincide el uso
            If Sheets("Inventario de Autos").Cells(I, 6) = "Disponible" Then    'Si está disponible
                For K = 1 To 4  'Paso a la hoja Auxiliar los datos que están en las primeras cuatro columnas del inventario
                    Sheets("Auxiliar").Cells(FilaAux, K) = Sheets("Inventario de Autos").Cells(I, K)
                Next K
                FilaAux = FilaAux + 1   'Ya usé esta fila de la hoja Auxiliar, paso a la siguiente
                h = True    'Indico que encontré algún auto con las características solicitadas
            End If
        End If
    End If
Next I
' Activo la hoja Auxiliar para que se vea el listado de carros que tienen las características
Sheets("Auxiliar").Select
End Sub
Sub copiarDatos(p As String)
'********************************************************************************************************
' Sub para pasar los datos del vehiculo desde el inventario al control de alquilados
'********************************************************************************************************
Call CuentaFilas    'Contador de las filas ocupadas en la hoja de inventario
FilaDatos = 2       'Procedimiento similar para buscar la primera fila vacía en la hoja Datos de Alquiler
While Worksheets("Datos de Alquiler").Cells(FilaDatos, 1) <> ""
    FilaDatos = FilaDatos + 1
Wend
For I = 2 To Fila - 1
    If Sheets("Inventario de Autos").Cells(I, 1) = p Then   'Si la placa coincide
        For K = 1 To 4  'Copio los valores en las cuatro primeras columnas de la hoja Datos de Alquiler
            Sheets("Datos de Alquiler").Cells(FilaDatos, K) = Sheets("Inventario de Autos").Cells(I, K)
        Next K
        ' Pido el nombre del responsable del vehículo y lo escribo en la columna 5 de la hoja Datos de Alquiler
        Sheets("Datos de Alquiler").Select
        Nombre = InputBox("Ingrese el nombre del responsable del vehículo:")
        Sheets("Datos de Alquiler").Cells(FilaDatos, 5) = Nombre
        ' Cambio el status del vehiculo en la hoja de inventario
        Sheets("Inventario de Autos").Cells(I, 6) = "Alquilado"
    End If
Next I
' Actualiza la base de datos de responsables
Call responsables(Nombre)
End Sub
Sub responsables(n As String)
' Actualizo la base de datos de responsables
Sheets("BaseDatos").Select
' Calculo el número de registros de esta hoja
FilaResp = 2
While Worksheets("BaseDatos").Cells(FilaResp, 1) <> ""
    FilaResp = FilaResp + 1
Wend
' Recorro la base de datos para buscar el nombre
Posicion = 0
With Sheets("BaseDatos")
    For I = 2 To FilaResp - 1
        If .Cells(I, 1) = n Then    'Si lo encuentro
            Posicion = I            'Guardo la posición donde está
        End If
    Next
    If Posicion > 0 Then    'Si lo encontré
        .Cells(Posicion, 2) = .Cells(Posicion, 2) + 1 'Incremento las veces que sus carros se han alquilado
    Else     'si no lo encontré
        .Cells(I, 1) = n    'Escribo su nombre al final de la tabla
        .Cells(I, 2) = 1    'Escribo que ha alquilado su primer carro
    End If
End With
Sheets("Inventario de Autos").Select    'Activo la hoja de inventario
End Sub
Sub MayorResp()
'***************************************************************************************************
'Concurso: Premio a quien haya alquilado más vehículos
'***************************************************************************************************
' Calculo el número de registros de esta hoja
FilaResp = 2
While Worksheets("BaseDatos").Cells(FilaResp, 1) <> ""
    FilaResp = FilaResp + 1
Wend
'Recorro todos los registros para determinar el mayor número de autos alquilado por algún responsable
'Comienzo asumiendo que el mayor está en el primer registro
Mayor = Sheets("BaseDatos").Cells(2, 2)
For I = 2 To FilaResp
    If Sheets("BaseDatos").Cells(I, 2) > Mayor Then
        Mayor = Sheets("BaseDatos").Cells(I, 2)
    End If
Next I
' reporto los ganadores en la hoja
K = 2
For I = 2 To FilaResp
    If Sheets("BaseDatos").Cells(I, 2) = Mayor Then
        Sheets("BaseDatos").Cells(K, 4) = Sheets("BaseDatos").Cells(I, 1)
        Sheets("BaseDatos").Cells(K, 5) = Mayor
        K = K + 1
    End If
Next I
End Sub
Sub devolver()
' Proceso de devolución de un vehículo
FilaDatos = 2
While Worksheets("Datos de Alquiler").Cells(FilaDatos, 1) <> ""
    FilaDatos = FilaDatos + 1
Wend
' Muestro la hoja Datos de Alquiler
Sheets("Datos de Alquiler").Select
'Los nombres de las hojas son cadenas de texto, puedo guardarlos en variables y usar esa variable después
'En la siguiente instrucción guardo el nombre de la hoja activa
NombreHoja = ActiveSheet.Name
' Para devolver un auto al inventario
placa = InputBox("Indique la placa del auto a devolver")
'busco la posición del auto que tiene esa placa
Posicion = 0
With Sheets(NombreHoja)
    For I = 2 To FilaDatos - 1
        If placa = .Cells(I, 1) Then
            Posicion = I
        End If
    Next
End With
' Si lo encontré procedo a eliminar el registro en la hoja Datos de Alquiler.
' Si no lo encontré emito un mensaje
If Posicion > 0 Then
    ' Eliminar el registro del auto en alquileres
    For I = Posicion To FilaDatos - 1
        For K = 1 To 5 'Los 5 elementos del registro
            Sheets(NombreHoja).Cells(I, K) = Sheets(NombreHoja).Cells(I + 1, K)
        Next K
    Next I
    ' Limpio la última fila
    For K = 1 To 5
        Sheets(NombreHoja).Cells(I, K).Clear
    Next K
Else
    MsgBox ("Error: La placa no existe")
End If
' Cambio el status en la hoja de inventario
Call CuentaFilas
With Sheets("Inventario de Autos")
    For I = 2 To Fila - 1
        If .Cells(I, 1) = placa Then    'Si la placa coincide
            .Cells(I, 5) = .Cells(I, 5) + 1
            .Cells(I, 6) = "Disponible"
            If .Cells(I, 5) > 10 Then   'Si se ha alquilado más de 10 veces
                .Cells(I, 6) = "En servicio"
                .Cells(I, 7) = "A revisión"
            ElseIf .Cells(I, 5) > 6 Then    'Si se ha alquilado entre 6 y 10 veces
                .Cells(I, 7) = "Pronto a revisión"
            Else        'En cualquier otro caso
                .Cells(I, 7) = "OK"
            End If
        End If
    Next I
    .Select
End With
End Sub
Sub ORDENAR()
'Sub para ordenar los datos del inventario o de los alquileres.
NombreHoja = ActiveSheet.Name
If NombreHoja = "Inventario de Autos" Then
    Call CuentaFilas
    Fil = Fila - 1
    Col = 8
Else
    FilaDatos = 2
    While Worksheets("Datos de Alquiler").Cells(FilaDatos, 1) <> ""
        FilaDatos = FilaDatos + 1
    Wend
    Fil = FilaDatos - 1
    Col = 5
End If
' Pido la columna por la que se quiere ordenar
Criterio = InputBox("Indique el número de la columna por la que desea ordenar:")
With Sheets(NombreHoja)
    For I = 2 To Fil - 1
        For J = (I + 1) To Fil
            If .Cells(I, Criterio) > .Cells(J, Criterio) Then
                For K = 1 To Col
                    Inver = .Cells(I, K)
                    .Cells(I, K) = .Cells(J, K)
                    .Cells(J, K) = Inver
                Next K
            End If
        Next J
    Next I
End With
End Sub
Sub reactivar()
' Proceso de reactivación de un vehículo
Call CuentaFilas
' Muestro la hoja de inventario
Sheets("Inventario de Autos").Select
' Guardo el nombre de la hoja activa
NombreHoja = ActiveSheet.Name
' Para reactivar un auto en el inventario pido la placa
placa = InputBox("Indique la placa del auto a devolver")
'busco la posición del auto que tiene esa placa
Posicion = 0
With Sheets(NombreHoja)
    For I = 2 To FilaDatos - 1
        If placa = .Cells(I, 1) Then
            Posicion = I
        End If
    Next
End With
' Si lo encontré procedo a modificar los datos en la hoja de inventario.
' Si no lo encontré emito un mensaje
If Posicion > 0 Then
    With Sheets("Inventario de Autos")
        .Cells(Posicion, 6) = "Disponible"
        .Cells(Posicion, 5) = 0
        .Cells(Posicion, 7) = "Ok"
    End With
Else
    MsgBox ("Error: La placa no existe")
End If
End Sub
