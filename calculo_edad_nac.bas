Attribute VB_Name = "M�dulo2"
Option Explicit
'Semana 1 - Tarea 2. Parte 1

'Declaraci�n de variables de entrada (ambas partes)

Dim perimetro As Single
Dim ancho As Single
Dim ano_nac As Integer
Dim mes_nac As Integer
'Declaracion variables de calculo

Dim largo As Single
Dim edad As Integer

'Declaraci�n variables de salida

Dim edad_total As Single
Dim area_patio As Single

perimetro = InputBox("Introduce el valor del perimetro del patio: ", "Perimetro patio")
ancho = InputBox("Introduce el valor del ancho del patio: ", " Ancho del patio")

'El area de un rect�ngulo viene dada por la multiplicacion de sus lados, al tener el ancho y perimetro debemos despejar quedando:

largo = ancho - perimetro / 2

area_patio = ancho * largo

MsgBox ("El �rea del patio cuyo largo es " & largo & ",  es de: " & area_patio)



ano_nac = Sheets("Hoja1").Cells(5, 2)

edad = 2020 - ano_nac

mes_nac = Sheets("Hoja1").Cells(5, 4)

edad_total = (8 - mes_nac + edad * 12) / 12 '8 representa el mes acutal de agosto, 12 los meses del a�o para pasar todos los a�os a meses y luego devolver el total

Sheets("Hoja1").Cells(5, 6) = edad_total 'impresion de la edad total, incluidos los meses, en a�os.

'Fin del programa

'NOTA: Para completar el programa podriamos escribir una estructura de decision compuesta para especificar a traves de un mensaje si la persona ya cumpli� a�os o no

If 8 - mes_nac >= 0 Then
    MsgBox ("Ya cumpli� a�os")
Else
    MsgBox ("No ha cumplido a�os a�n")
End If

'Puede existir un error para las personas que cumplen a�os en agosto, ya que tomamos el mes completo"

End Sub

