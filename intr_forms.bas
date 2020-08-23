Attribute VB_Name = "Módulo1"
Option Explicit
Dim fila As Integer
Sub inicio()
' abrir el formulario para comenzar a cargar datos
form_datos.Show
End Sub
Sub procesar()
Call contar_filas
With Sheets("Datos")
    ' leemos el campo texto cedula
    .Cells(fila, 1) = form_datos.txt_cedula
    ' leemos el sexo
    If form_datos.opt_fem = True Then
        .Cells(fila, 3) = "F"
    Else
        .Cells(fila, 3) = "M"
    End If
    ' leemos la categoria
    If form_datos.opt_prof = True Then
        .Cells(fila, 2) = "P"
    ElseIf form_datos.opt_est = True Then
        .Cells(fila, 2) = "E"
    Else
        .Cells(fila, 2) = "X"
    End If
    ' leemos el libro preferido
    If form_datos.chk_novela = True Then
        .Cells(fila, 4) = "X"
    End If
    If form_datos.chk_Ciencia = True Then
        .Cells(fila, 5) = "X"
    End If
    If form_datos.chk_poesia = True Then
        .Cells(fila, 6) = "X"
    End If
    If form_datos.chk_otro = True Then
        .Cells(fila, 7) = "X"
    End If
End With
End Sub
Sub contar_filas()
' cuenta las fials llenas
fila = 3
While Sheets("Datos").Cells(fila, 1) <> ""
    fila = fila + 1
Wend
End Sub
Sub limpiar_form()
' limpia los campos del form
With form_datos
    .txt_cedula = ""
    .opt_fem = False
    .opt_masc = False
    .opt_prof = False
    .opt_est = False
    .opt_otro = False
    .chk_novela = False
    .chk_Ciencia = False
    .chk_poesia = False
    .chk_otro = False
End With
End Sub
