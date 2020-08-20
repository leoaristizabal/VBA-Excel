Attribute VB_Name = "Módulo2"
Option Explicit
'declaracion
Dim formato As String
Dim especie As String
'contadores
Dim cont_f As Integer
Dim cont_v As Integer
Dim cont_perro As Integer
Dim cont_total As Integer

Dim fila As Integer
Dim fila2 As Integer

Sub principal()
'iniciadores de filas
fila = 5 'bases de datos
fila2 = 5 ' base de datos en Propuesta

'inicializadores de contadores
cont_f = 0
cont_v = 0
cont_perro = 0
cont_total = 0


'entradas de formato y especie
formato = InputBox("INGRESE EL FORMATO QUE DESEE (V,F): ", "FORMATO")
especie = InputBox("INGRESE EL ANIMAL A BUSCAR: ", "CLASE DE ANIMAL")

'Repetición While
While Sheets("Datos").Cells(fila, 2) <> ""
    If (especie = "elefante" Or especie = "hipopotamo" Or especie = "perro" Or especie = "gato" Or especie = "loro" Or especie = "tiburon" Or especie = "ballena" Or especie = "ganso" Or especie = "serpiente" Or especie = "tortuga") And (formato = "V" Or formato = "F") Then
        If formato = "F" Then
            cont_f = cont_f + 1
            Sheets("Estadisticas").Cells(5, 3) = cont_f
        End If
        If especie = "perro" Then
            cont_perro = cont_perro + 1
            Sheets("Estadisticas").Cells(9, 3) = cont_perro
        
        If formato = Sheets("Datos").Cells(fila, 2) And especie = Sheets("Datos").Cells(fila, 3) Then
           
            Sheets("Propuesta").Cells(fila2, 2) = Sheets("Datos").Cells(fila, 2)
            Sheets("Propuesta").Cells(fila2, 3) = Sheets("Datos").Cells(fila, 3)
            Sheets("Propuesta").Cells(fila2, 4) = Sheets("Datos").Cells(fila, 4)
            Sheets("Propuesta").Cells(fila2, 5) = Sheets("Datos").Cells(fila, 5)
        
            'contador de veces consutlada
            Sheets("Datos").Cells(fila, 6) = Sheets("Datos").Cells(fila, 6) + 1
            
            fila2 = fila2 + 1
        
        End If
    End If
            
fila = fila + 1

Wend

End Sub

Sub limpiar_boton()

'boton limpiar Hoja Propuesta y Estadistica a la vez
'NOTA: de querer limpiar cada hoja por separado se sepran en diferentes sub y se crea un boton para cada uno


Sheets("Propuesta").Range("B5:E34").Value = ""

Sheets("Estadisticas").Range("C5:C9").Value = ""

End Sub

