Attribute VB_Name = "Módulo2"
Option Explicit
' entradas
Dim NomAlu As String
Dim Nota1 As Integer
Dim Nota2 As Integer
' salidas
Dim ProNot As Single
Dim ConAlu As Integer
Dim AcuPro As Single
Dim ProSec As Single
Dim ConApro As Integer
Dim ConRepro As Integer
Dim ConMas15 As Integer
Dim Con10a15 As Integer
Dim Con5a10 As Integer
Dim Con0a5 As Integer
Dim Lineas As Integer
Dim MinPro As Single
Dim MinNom As String
Dim MaxPro As Single
Dim MaxNom As String


    Call inicio
    Call repeticion
    Call reporte_final


' Estructura de repetición
    While NomAlu <> ""  'Mientras el valor de NomAlu sea diferente de vacío
        Call lectura
        Call proceso
        Call reporte_interno
        ' Actualización del número de la fila en la hoja
        Lineas = Lineas + 1
        ' Lectura de la variable que controla el ciclo de repetición
        NomAlu = Sheets("Datos").Cells(Lineas, 1)
    Wend


' Inicialización de variables
    Sheets("Reporte").Select ' Activación de la hoja de cálculo
    Lineas = 2
    ConAlu = 0
    AcuPro = 0
    ConApro = 0
    ConRepro = 0
    ConMas15 = 0
    Con10a15 = 0
    Con5a10 = 0
    Con0a5 = 0
    MinPro = 21
    MaxPro = -1
' Lectura de la variable que controla el ciclo de repetición
    NomAlu = Sheets("Datos").Cells(Lineas, 1)


' Lectura desde la hoja de cálculo
    With Sheets("Datos")
        Nota1 = .Cells(Lineas, 2)
        Nota2 = .Cells(Lineas, 3)
    End With


' Cálculos
        ProNot = (Nota1 + Nota2) / 2
        If ProNot < MinPro Then
            MinPro = ProNot
            MinNom = NomAlu
        End If
        If ProNot > MaxPro Then
            MaxPro = ProNot
            MaxNom = NomAlu
        End If
' Actualizaciones
        ConAlu = ConAlu + 1
        AcuPro = AcuPro + ProNot
' Contar los alumnos
        If ProNot >= 10 Then
            ConApro = ConApro + 1
            If ProNot >= 15 Then
                ConMas15 = ConMas15 + 1
            Else
                Con10a15 = Con10a15 + 1
            End If
        Else
            ConRepro = ConRepro + 1
            If ProNot >= 5 Then
                Con5a10 = Con5a10 + 1
            Else
                Con0a5 = Con0a5 + 1
            End If
        End If


' Escritura del reporte interno (para cada alumno)
        With Sheets("Reporte")
            .Cells(Lineas, 1) = NomAlu
            .Cells(Lineas, 2) = Nota1
            .Cells(Lineas, 3) = Nota2
            .Cells(Lineas, 4) = ProNot
        End With


' Cálculo del promedio
    If ConAlu <> 0 Then
        ProSec = AcuPro / ConAlu
    Else
        ProSec = 0
    End If
' Escritura del reporte final (para toda la sección)
    With Sheets("Reporte")
        .Cells(1, 7) = ConAlu
        .Cells(2, 7) = ProSec
        .Cells(3, 7) = ConApro
        .Cells(4, 7) = ConRepro
        .Cells(5, 7) = ConMas15
        .Cells(6, 7) = Con10a15
        .Cells(7, 7) = Con5a10
        .Cells(8, 7) = Con0a5
        .Cells(9, 7) = MinPro
        .Cells(9, 8) = MinNom
        .Cells(10, 7) = MaxPro
        .Cells(10, 8) = MaxNom
    End With






