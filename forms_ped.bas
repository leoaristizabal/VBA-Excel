Attribute VB_Name = "Módulo1"
Option Explicit
Dim ContMas As Integer
Dim ContFem As Integer
Dim AcuAltMas As Integer
Dim AcuAltFem As Integer
Dim AcuPesoMas As Integer
Dim AcuPesoFem As Integer
Dim PromAltMas As Double
Dim PromAltFem As Double
Dim PromPesoMas As Double
Dim PromPesoFem As Double
Dim ContSint As Integer 'contador de sintomas
Dim Sintomas As String 'variable para mostrar en la hoja de excel
Dim ContSanos As Integer
Dim ContObs As Integer
Dim ContUrg As Integer
Dim Fila As Integer

Sub Inicio_var()
    Fila = 2
    ContMas = 0
    ContFem = 0
    ContSint = 0
    ContSanos = 0
    ContObs = 0
    ContUrg = 0
    AcuAltMas = 0
    AcuAltFem = 0
    AcuPesoMas = 0
    AcuPesoFem = 0
    Sintomas = ""
    
End Sub
Sub inicio()
    Sheets("Datos").Select
    for_pediatra.Show
End Sub
Sub contar_filas()
'cuenta las filas llenas
Fila = 2
While Sheets("Datos").Cells(Fila, 1) <> ""
    Fila = Fila + 1
Wend
End Sub
Sub Procesar()
Call contar_filas
    With Sheets("Datos")
        .Cells(Fila, 1) = for_pediatra.txt_nombre
        If for_pediatra.opt_femenino Then
            .Cells(Fila, 2) = "Niña"
        Else
            .Cells(Fila, 2) = "Niño"
        End If
        .Cells(Fila, 3) = CInt(for_pediatra.txt_edad) 'EL OPERADOR CINT CONVIERTE EL TEXTO A NUMERO ENTERO
        .Cells(Fila, 4) = CDbl(for_pediatra.txt_peso) ' EL OPERADOR CDBL CONVIERTE EL TEXTO A UNA VALOR DOUBLE
        .Cells(Fila, 5) = CDbl(for_pediatra.txt_altura)
        .Cells(Fila, 6) = for_pediatra.cmb_grupo
'concatenación de las vacunas que se le aplicaron al niño
'se mostraran en una sola celda
        .Cells(Fila, 7) = for_pediatra.cmb_vacuna1 & " " & for_pediatra.cmb_vacuna2 & " " & for_pediatra.cmb_vacuna3
 'se revisan los sintomas y se MARCAN EN LA HOJA CON UNA X
        If for_pediatra.chk_fiebre Then
            .Cells(Fila, 8) = "X"
        End If
        If for_pediatra.chk_vomito Then
            .Cells(Fila, 9) = "X"
        End If
        If for_pediatra.chk_diarrea Then
            .Cells(Fila, 10) = "X"
        End If
        If for_pediatra.chk_erupcion Then
            .Cells(Fila, 11) = "X"
        End If
    End With
End Sub
Sub Actualizar()
' SE REVISA LA HOJA DE DATOS Y SE CUENTAN LOS SINTOMAS PARA COLOCAR EL DIAGNOSTICO EN LA ULTIMA COLUMNA
Dim i As Integer
Dim j As Integer

Dim diag As String
Call Inicio_var ' se inicializan los contadores y acumuladores
Call contar_filas ' se cuentan cuantas filas llenas hay en la tabla
' se reccore la tabla para contar los sintomas y rellenar la ultima columna
For i = 2 To Fila - 1
    ContSint = 0
    With Sheets("Datos")
        For j = 8 To 11
            If .Cells(i, j) = "X" Then
                ContSint = ContSint + 1
            End If
        Next j
        ' se determina el diagnostico y se escribe
        Select Case ContSint
        Case 0
            .Cells(i, 12) = "Sano"
        Case 1, 2
                .Cells(i, 12) = "En Observación"
        Case Is >= 3
                .Cells(i, 12) = "Caso Urgente"
        End Select
    End With
Next i
        
'se recorre  de nuevo la tabla y se actualizan todos los contadores y acumuladores pedidos

For i = 2 To Fila - 1
    With Sheets("Datos")
        If .Cells(i, 2) = "Niña" Then
            ContFem = ContFem + 1
            AcuPesoFem = AcuPesoFem + .Cells(i, 4)
            AcuAltFem = AcuAltFem + .Cells(i, 5)
        Else
            ContMas = ContMas + 1
            AcuPesoMas = AcuPesoMas + .Cells(i, 4)
            AcuAltMas = AcuAltMas + .Cells(i, 5)
        End If
        diag = .Cells(i, 12)
        Select Case diag
        Case "Sano"
            ContSanos = ContSanos + 1
        Case "En Observación"
            ContObs = ContObs + 1
        Case "Caso Urgente"
            ContUrg = ContUrg + 1
        End Select
        
    End With
Next i
End Sub
Sub Limpiar_form()
    With for_pediatra
        .txt_nombre = ""
        .txt_altura = ""
        .txt_peso = ""
        .txt_edad = ""
        .opt_femenino = False
        .opt_masculino = False
        .cmb_grupo = ""
        .cmb_vacuna1 = ""
        .cmb_vacuna2 = ""
        .cmb_vacuna3 = ""
        .chk_diarrea = False
        .chk_erupcion = False
        .chk_vomito = False
        .chk_fiebre = False
    End With
    Sintomas = ""
    ContSint = 0
End Sub

Sub Finalizar()
    MsgBox ("Final del día")
    for_pediatra.Hide
End Sub

Sub Reporte()
Dim proAF As Double
Dim proAM As Double
Dim proPF As Double
Dim proPM As Double
' limpiar el area de reporte
Call limpiar_reporte
' actualizar contadores y acumuladores
Call Actualizar
Fila = 0
' calculo de promedios
If ContFem > 0 Then
    proAF = AcuAltFem / ContFem
    proAM = AcuAltMas / ContMas
Else
    proAF = 0
    proAM = 0
End If
If ContMas > 0 Then
    proPF = AcuPesoFem / ContFem
    proPM = AcuPesoMas / ContMas
Else
    proPF = 0
    proPM = 0
End If

    With Sheets("Reporte")
        .Cells(Fila + 3, 3) = "Niñas"
        .Cells(Fila + 3, 4) = "Niños"
        .Cells(Fila + 4, 2) = "Atendidos"
        .Cells(Fila + 4, 3) = ContFem
        .Cells(Fila + 4, 4) = ContMas
        .Cells(Fila + 5, 2) = "Promedio altura"
        .Cells(Fila + 5, 3) = proAF
        .Cells(Fila + 5, 4) = proAM
        .Cells(Fila + 6, 2) = "Promedio Peso"
        .Cells(Fila + 6, 3) = proPF
        .Cells(Fila + 6, 4) = proPM
        .Cells(Fila + 8, 2) = "Casos Sanos"
        .Cells(Fila + 8, 3) = ContSanos
        .Cells(Fila + 9, 2) = "Casos en observación"
        .Cells(Fila + 9, 3) = ContObs
        .Cells(Fila + 10, 2) = "Casos Urgentes"
        .Cells(Fila + 10, 3) = ContUrg
    End With
End Sub
Sub limpiar_reporte()
Dim i As Integer
Dim j As Integer
With Sheets("Reporte")
    For i = 4 To 11
        For j = 3 To 4
            .Cells(i, j) = ""
        Next j
    Next i
End With
            
End Sub
