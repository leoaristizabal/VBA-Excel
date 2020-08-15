Attribute VB_Name = "Módulo1"
Option Explicit
Dim fila As Integer
Dim uno As Integer
Dim dos As Integer
Dim nova As Integer
Dim prim As Integer
Dim secu As Integer
Dim ambas As Integer



Sub notas()
fila = 4
uno = 0
dos = 0
nova = 0
prim = 0
secu = 0
ambas = 0

With Sheets("Datos ")
    While .Cells(fila, 1) <> ""
        .Cells(fila, 4) = (.Cells(fila, 2) + .Cells(fila, 3)) / 2
   '    >= 18  Semestre DOS
    '   >= 12 Semestre UNO
   '   < 12 RECHAZADO
Select Case .Cells(fila, 4)
    Case Is >= 18
        .Cells(fila, 5) = "semestre DOS"
        dos = dos + 1
    Case 12 To 17.9
        .Cells(fila, 5) = "semestre UNO"
        uno = uno + 1
    Case Else
        .Cells(fila, 5) = "Rechazado"
        nova = nova + 1
End Select
' critero especial
If .Cells(fila, 2) > .Cells(fila, 3) Then
    .Cells(fila, 6) = "primaria"
    prim = prim + 1
Else
    If .Cells(fila, 2) < .Cells(fila, 3) Then
       .Cells(fila, 6) = "bachillerato"
       secu = secu + 1
     Else
        .Cells(fila, 6) = "ambos"
        ambas = ambas + 1
    End If
End If

fila = fila + 1

Wend
' impresion de las estadisticas
.Cells(5, 9) = uno
.Cells(6, 9) = dos
.Cells(7, 9) = nova
.Cells(8, 9) = uno / (uno + dos + nova)
.Cells(9, 9) = dos / (uno + dos + nova)
.Cells(10, 9) = nova / (uno + dos + nova)
.Cells(14, 9) = prim
.Cells(15, 9) = secu
.Cells(16, 9) = ambas


End With
Sheets("Datos ").Select
End Sub
