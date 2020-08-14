Attribute VB_Name = "Módulo1"
Option Explicit
' entrada
Dim valor As Integer
' salida
Dim prom As Double
Dim c05 As Integer
Dim c615 As Integer
Dim cm15 As Integer
Dim mayor As Integer
Dim menor As Integer
Dim porcm15 As Double
Dim total As Integer
' auxiliares
Dim fila As Integer
Dim i As Integer


' contar filas
Call contar_filas
' inicializacion de contadores y acumuladores
Call inicio
' ciclo de repeticion
For i = 2 To fila - 1
    valor = Sheets("Datos").Cells(i, 1)
    ' calcular el total
    total = total + valor
    'contar en rangos dados
    Select Case valor
        Case Is < 0
            MsgBox ("valor negativo")
        Case Is <= 5
            c05 = c05 + 1
        Case Is <= 15
            c615 = c615 + 1
        Case Else
            cm15 = cm15 + 1
    End Select
    ' determino el mayor
    If valor > mayor Then
        mayor = valor
    End If
    ' determino el menor
    If valor < menor Then
        menor = valor
    End If
Next i
' calculo del promedio
If (fila - 2) > 0 Then
    prom = total / (fila - 2)
Else
    prom = 0
End If
' calculo del porcentaje de valores mayores a 15
If (fila - 2) > 0 Then
    porcm15 = (cm15 / (fila - 2)) * 100
Else
    porcm15 = 0
End If
' escribir los resultados
Call escribir



' cuenta cuantas filas están llenas
fila = 2
While Sheets("Datos").Cells(fila, 1) <> ""
    fila = fila + 1
Wend



' inicializacion de contadores y acumuladores
c05 = 0
c615 = 0
cm15 = 0
mayor = Sheets("Datos").Cells(2, 1)
menor = Sheets("Datos").Cells(2, 1)
total = 0


' resultados obtenidos
With Sheets("Reporte")
    .Cells(2, 2) = prom
    .Cells(3, 2) = c05
    .Cells(4, 2) = c615
    .Cells(5, 2) = cm15
    .Cells(6, 2) = mayor
    .Cells(7, 2) = menor
    .Cells(8, 2) = porcm15
    .Cells(9, 2) = total
End With
Sheets("Reporte").Select

