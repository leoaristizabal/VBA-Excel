Attribute VB_Name = "Módulo1"
Option Explicit
Dim fila As Integer
Dim contador As Integer
' declarar las variables que  faltan paar el reporte

Sub inicio()
'limpia el form y lo muestra
Call limpiar_form
form_datos.Show
End Sub
' procesar al data que viene del formulario y colocarla en al hoja
Sub procesar()
' cuenta las filas llenas
Call contar_filas
' paso los datos del formulario a la hoja
With Sheets("Datos")
' numero de secuencia
    .Cells(fila, 1) = contador
' sexo
    If form_datos.opt_m = True Then
        .Cells(fila, 2) = "M"
    Else
        .Cells(fila, 2) = "F"
    End If
'categoria
    .Cells(fila, 3) = form_datos.cmb_cat
' edad
    If form_datos.opt_1415 = True Then
        .Cells(fila, 4) = "14-15"
    Else
        .Cells(fila, 4) = "16 o mas"
    End If
' nivel de dificultad
    .Cells(fila, 5) = form_datos.cmb_niv
End With
End Sub
' limpiar el form
Sub limpiar_form()
With form_datos
    .opt_m = False
    .opt_f = False
    .opt_1415 = False
    .opt_16 = False
End With
End Sub
' contar filas
Sub contar_filas()
fila = 3
While Sheets("Datos").Cells(fila, 1) <> ""
    fila = fila + 1
Wend
If fila = 3 Then
    contador = 1
Else
    contador = Sheets("Datos").Cells(fila - 1, 1) + 1
End If
End Sub

' calcular y escribir el reporte pedido
' construye los sub requeridos siguiendo al secuencia
' desarrolla el codigo faltante en el for
' para calcular lo pedido

Sub reporte()
' contar filas
Call contar_filas
' inicializar variables
Call inicio
' ciclo de repeticion para recorrer la hoja
For i = 3 To fila - 1
    ' leer los datos de la hoja
    Call lectura
    ' Cantidad de niños participantes de cada categoría
    
    'Cantidad de niños de 3er. Año en cada nivel de dificultad
    
    ' Porcentaje de niños participantes de cada sexo
    
    'Rango de edad en el cual hay más participantes
Next
' calculo de los porcentajes

' reportar los resultados
Call escritura
End Sub
