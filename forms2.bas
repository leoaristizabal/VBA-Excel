Attribute VB_Name = "Módulo1"
Option Explicit
Dim fila As Integer
Sub inicio()
' limpia el form y lo abre para trabajar
Call limpiar_form
form_datos.Show
End Sub
Sub contar_filas()
' cuenta las filas ocupadas en la hoja
fila = 3
While Sheets("Datos").Cells(fila, 1) <> ""
    fila = fila + 1
Wend
End Sub
Sub procesar()
' coloca los datos desde el form a la hoja
Call contar_filas
With Sheets("Datos")
    .Cells(fila, 1) = form_datos.txt_carnet
    .Cells(fila, 2) = form_datos.cmb_carrera
    .Cells(fila, 3) = form_datos.spin_cred.Value
    If form_datos.opt_M1 = True Then
        .Cells(fila, 4) = 1
    ElseIf form_datos.opt_m2 = True Then
        .Cells(fila, 4) = 2
    Else
        .Cells(fila, 4) = 3
    End If
End With
End Sub
Sub limpiar_form()
' limpiar los campos del formulario
With form_datos
    .txt_carnet = ""
    .cmb_carrera = ""
    .spin_cred = 3
    .opt_M1 = False
    .opt_m2 = False
    .opt_m3 = False
End With
End Sub
Sub calculos()
Dim materias As Integer
Dim i As Integer
Dim creditos As Integer
Dim desc As Double
Dim neto As Double
Dim pago As Integer
' calcula los pagos
Call contar_filas
' ciclo de recorrido
For i = 3 To fila - 1
    ' calcular monto básico
    creditos = Sheets("Datos").Cells(i, 3)
    
    Select Case creditos
        Case 3
            pago = 3 * 350
        Case 6
            pago = 6 * 300
        Case 9
            pago = 9 * 200
    End Select
    ' calculo el descuento
    materias = Sheets("Datos").Cells(i, 4)
    If materias = 2 Then
        desc = pago * 0.1
    ElseIf materias = 3 Then
        desc = pago * 0.15
    End If
    ' calculo del neto a pagar
    neto = pago - desc
    ' reporte de los resultados
    Sheets("Datos").Cells(i, 5) = desc
    Sheets("Datos").Cells(i, 6) = neto
Next i
            
End Sub
