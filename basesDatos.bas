Attribute VB_Name = "Módulo2"
Option Explicit
'Entrada
Dim fila As Integer
Dim fila2 As Integer
Dim color As String
Dim tipo As String


Sub programa_principal()
fila = 5 ' donde comienza la base de datos y podamos comparar con los datos ingresados
fila2 = 5

    tipo = InputBox("INGRESE EL TIPO DE VEHICULO(S,C): ", "TIPO DE VEHICULO")
    color = InputBox("INGRESE EL COLOR DEL VEHICULO (N,A,V,R): ", "COLOR VEHICULO")
   
   While Sheets("Datos").Cells(fila, 2) <> "" 'REPETICION BUSQUEDA ENLA BASE DE DATOS
    'SEGUN EL COLOR Y TIPO DE VEHICULO INGRESADO, SE IMPRIMEN LOS DATOS EN RESULTADOS
    If (color = "N" Or color = "V" Or color = "A" Or color = "R") And (tipo = "S" Or tipo = "C") Then
        'SABIENDO COLOR Y TIPO DE DATOS, QUEREMOS IMPRIMIR LOS OTROS DATOS EN RESULTADOS
        If color = Sheets("Datos").Cells(fila, 3) And tipo = Sheets("Datos").Cells(fila, 4) Then
            Sheets("Resultados").Cells(fila2, 2) = Sheets("Datos").Cells(fila, 2)
            Sheets("Resultados").Cells(fila2, 3) = Sheets("Datos").Cells(fila, 5)
            Sheets("Resultados").Cells(fila2, 4) = Sheets("Datos").Cells(fila, 6)
            Sheets("Resultados").Cells(fila2, 5) = Sheets("Datos").Cells(fila, 7)
             'SUMA DE VECES QUE SE CONSULTA CON CADA NUEVO INGRESO
             Sheets("Datos").Cells(fila, 8) = Sheets("Datos").Cells(fila, 8)
             
   fila2 = fila2 + 1
        End If
   End If
   fila = fila + 1
   Wend
   
    
End Sub
