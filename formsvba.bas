Attribute VB_Name = "Módulo1"
Option Explicit
Dim TotOrig As Integer
Dim TotQuem As Integer
Dim PromDur As Double
Dim AcuDur As Integer
Dim c1 As Integer
Dim c2 As Integer
Dim c3 As Integer
Dim c4 As Integer
Dim Fila As Integer

'***********************************************
' Se activa al hacer click en el botón Iniciar
'***********************************************
Sub Inicio()
    Call filas
    Sheets("Inventario").Select
    frm_ejemplo.Show
End Sub

Sub Procesar()
'***********************************************
'Paso de los datos a la hoja
'***********************************************
    With Sheets("Inventario")
        .Cells(Fila, 1) = frm_ejemplo.txt_nombre
        If frm_ejemplo.opt_original Then
            .Cells(Fila, 2) = "Original"
        Else
            .Cells(Fila, 2) = "Quemado"
        End If
            .Cells(Fila, 3) = frm_ejemplo.txt_duracion
            .Cells(Fila, 4) = frm_ejemplo.cmb_tipo
    End With

'************************************************
'Actualizo el contador de filas
'************************************************
    Fila = Fila + 1
End Sub
Sub cierre()
'************************************************
'Actualizo los acumuladores y contadores
'************************************************
    Call Actualizar
    Call reporte
    Sheets("Inventario").Select
End Sub
'************************************************
'Reinicializo la forma
'************************************************
Sub limpiar_form()
    With frm_ejemplo
        .txt_nombre = ""
        .opt_original = False
        .opt_quemado = False
        .txt_duracion = 0
        .cmb_tipo = ""
    End With
End Sub

Sub Actualizar()
Dim i As Integer
Dim tipo As String
TotOrig = 0
TotQuem = 0
c1 = 0
c2 = 0
c3 = 0
c4 = 0

' determino cuantas filas de la hoja inventario estan llenas
Call filas
' recorro la hoja de inventario
For i = 2 To Fila - 1
' actualizo contadores
    If Sheets("Inventario").Cells(i, 2) = "Original" Then
            TotOrig = TotOrig + 1
        Else
            TotQuem = TotQuem + 1
    End If
' actualizo acumulador
    AcuDur = AcuDur + Sheets("Inventario").Cells(i, 3)
' cuanto cuantos hay de cada tipo de disco
    tipo = Sheets("Inventario").Cells(i, 4)
    Select Case tipo
    Case "CDROM"
        c1 = c1 + 1
    Case "CDRW"
        c2 = c2 + 1
    Case "DVD"
        c3 = c3 + 1
    Case "DVDRW"
        c4 = c4 + 1
    End Select
Next i
End Sub
Sub Finalizar()
    MsgBox ("Cerrar Inventario")
    frm_ejemplo.Hide
End Sub
Sub reporte()
' escribo el reporte en la misma hoja inventario al finalizar la data
Call filas
If Fila > 2 Then
With Sheets("Inventario")
        .Cells(Fila + 3, 5) = "Originales"
        .Cells(Fila + 3, 6) = TotOrig
        .Cells(Fila + 4, 5) = "Quemados"
        .Cells(Fila + 4, 6) = TotQuem
        .Cells(Fila + 5, 5) = "Prom. Duracion"
        .Cells(Fila + 5, 6) = AcuDur / (TotOrig + TotQuem)
        .Cells(Fila + 6, 5) = "Porcentaje de CD"
        .Cells(Fila + 6, 6) = (c1 / (c1 + c2 + c3 + c4)) * 100
        .Cells(Fila + 7, 5) = "Porcentaje de CDRW"
        .Cells(Fila + 7, 6) = (c2 / (c1 + c2 + c3 + c4)) * 100
        .Cells(Fila + 8, 5) = "Porcentaje de DVD"
        .Cells(Fila + 8, 6) = (c3 / (c1 + c2 + c3 + c4)) * 100
        .Cells(Fila + 9, 5) = "Porcentaje de DVDRW"
        .Cells(Fila + 9, 6) = (c4 / (c1 + c2 + c3 + c4)) * 100
    End With
Else
    MsgBox ("No hay datos para generar el reporte, debe cargar al menos un elemento")
End If
End Sub
Sub filas()
Fila = 2
While Sheets("Inventario").Cells(Fila, 1) <> ""
    Fila = Fila + 1
Wend
End Sub
Sub limpiar_data()
' lipiar la hoja de Inventario
Dim i As Integer
Dim j As Integer
For i = 2 To 200
    For j = 1 To 8
        Sheets("Inventario").Cells(i, j) = ""
    Next j
Next i
Sheets("Inventario").Select
End Sub
