Attribute VB_Name = "Módulo2"
Option Explicit
'CONTADORES Y ACUMULADORES
Dim cant_encues As Integer
Dim si As Integer
Dim no As Integer
Dim bolsita As Integer

Dim edad As Integer
Dim nombre As String
Dim sexo As String
Dim calidad_juguete As String
Dim llevarlo As String
Dim premio As String
Dim fila As Integer

Sub principal()

fila = 2
cant_encues = 0
si = 0
no = 0
bolsita = 0

edad = InputBox("INGRESA TU EDAD: ", "EDAD")

If edad <= 10 Then

    While edad <= 10
cant_encues = cant_encues + 1
        nombre = InputBox("INGRESA TU NOMBRE: ", "NOMBRE")
        
        sexo = InputBox("INGRESA TU TIPOS DE SEXO (M,F): ", "SEXO")

        calidad_juguete = InputBox("INGRESA LA CALIDAD DEL JUGUETE (B,R,M): ", "CALIDAD DEL JUGUETE")

        llevarlo = InputBox("¿DESEA UD COMO PADRE LLAVARLO? (SI O NO)")

    With Sheets("resultados")
        .Cells(fila, 1) = nombre
        .Cells(fila, 2) = edad
        .Cells(fila, 3) = sexo
        .Cells(fila, 4) = calidad_juguete
        .Cells(fila, 5) = llevarlo
    

    If llevarlo = "SI" Then
    si = si + 1
        If calidad_juguete = "B" Then
            .Cells(fila, 6) = "BOLSITA FELIZ"
            bolsita = bolsita + 1
        Else
            .Cells(fila, 6) = "CARAMELO PEPPA"
        End If
    End If
    
    If llevarlo = "NO" Then
    no = no + 1
        If calidad_juguete = "B" Then
            .Cells(fila, 6) = "PAPAS FRITAS"
        Else
            .Cells(fila, 6) = "CALCOMANIA"
        End If
    End If
    

    fila = fila + 1
    edad = InputBox("INGRESA TU EDAD: ", "EDAD")
    End With
    Wend

Else
    MsgBox ("LO SENTIMOS, ESTA ENCUESTA ES PARA NIÑOS MENORES DE 10 AÑOS")
End If

'IMPRESIONES CONTADORES
With Sheets("resultados")
    .Cells(2, 10) = cant_encues
    .Cells(3, 10) = si / (si + no)
    .Cells(4, 10) = bolsita
End With
End Sub


Sub limpiar()

With Sheets("resultados")
    .Range("A2:F12").Value = " "
    .Range("J2:J4").Value = " "
End With
End Sub
