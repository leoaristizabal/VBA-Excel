Attribute VB_Name = "Módulo1"
Option Explicit
Dim nombre As String
Dim cedula As Long
Dim enfermedad As String
'contadores y acumuladores
Dim total_pagar As Integer
Dim c As Integer
Dim r As Integer
Dim d As Integer

Dim monto As Integer
Dim monto_rad As Integer
Dim fila As Integer

Sub principal()

fila = 3
monto = 0
monto_rad = 0



With Sheets("datos")
   nombre = .Cells(fila, 1)
While nombre <> ""
    monto = monto + 60
    
    cedula = .Cells(fila, 2)
    enfermedad = .Cells(fila, 3)
    
    
        If enfermedad = "R" Then
         .Cells(fila, 4) = "2"
         monto_rad = (monto_rad) + (2 * 30)
        Else
            If enfermedad = "C" Then
             .Cells(fila, 4) = "1"
             monto_rad = monto_rad + 30
            End If
        End If
        
        If enfermedad = "D" Then
             .Cells(fila, 4) = "3"
             monto_rad = (monto_rad) + (3 * 30)
        End If
    
    total_pagar = monto_rad + monto
    Sheets("datos").Cells(fila, 5) = total_pagar

    fila = fila + 1
    nombre = .Cells(fila, 1)
    

Wend
End With
    
'impresiones total a pagar


    

End Sub
