Attribute VB_Name = "Módulo2"
Option Explicit
' entradas
Dim j1 As String
Dim j2 As String
' salida
Dim resultado As Integer
' auxiliares
Dim linea As Integer
Dim filas As Integer
Dim resp As String
' lectura de jugadas
Sub lectura()
linea = 3
resp = MsgBox("Hay mas jugadas por reportar?", vbYesNo, "Jugadas")
' ciclo de repeticion
While resp = vbYes
    ' leo la jugada de jugador1
    j1 = InputBox("Jugador1: ")
    Select Case j1
        Case "Piedra", "piedra", "PIEDRA"
            Sheets("datos").Cells(linea, 1) = "P"
        Case "Papel", "papel", "PAPEL"
            Sheets("datos").Cells(linea, 1) = "X"
        Case "Tijera", "tijera", "TIJERA"
            Sheets("datos").Cells(linea, 1) = "T"
    End Select
    ' leo la jugada de jugador 2
    j2 = InputBox("Jugador 2: ")
    Select Case j2
        Case "Piedra", "piedra", "PIEDRA"
            Sheets("datos").Cells(linea, 2) = "P"
        Case "Papel", "papel", "PAPEL"
            Sheets("datos").Cells(linea, 2) = "X"
        Case "Tijera", "tijera", "TIJERA"
            Sheets("datos").Cells(linea, 2) = "T"
    End Select
    ' solicitar de nuevo la variable de control
    resp = MsgBox("Hay mas jugadas por reportar?", vbYesNo, "Jugadas")
    ' actualizo la linea
    linea = linea + 1
Wend
End Sub


