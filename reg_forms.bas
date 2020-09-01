Attribute VB_Name = "Módulo1"
Option Explicit
Dim Linea As Integer
Dim LineaConsultas As Integer
Dim Nombre As String
Dim Autor As String
Dim Area As String
Dim Buscar As String
Dim NombreBuscar As String
Dim AutorBuscar As String
Dim I As Integer

Sub Principal()
    Call Inicio
    For_Ingresar.txt_nombre.SetFocus
    For_Ingresar.Show
End Sub
Sub Iniciar()
    Sheets("Datos").Range("A2:C30").Clear
    Sheets("Consultas").Range("A2:C30").Clear
    Sheets("Ocultos").Cells(1, 7) = 2
End Sub
Sub Inicio()
'En la hoja "Ocultos", celda G1 tengo guardado el valor de la primera fila disponible
'Si esa celda está vacía, significa que no hay datos y la primera fila vacía es la 2.
'Si esa celda está llena leo el valor que contiene.
    If Sheets("Ocultos").Cells(1, 7) = "" Then
        Linea = 2
    Else
        Linea = Sheets("Ocultos").Cells(1, 7)
    End If
End Sub
Sub Procesar()
    Call Leer
    Call PasarExcel
    Call LimpiarForm
End Sub
Sub Leer()
    Nombre = For_Ingresar.txt_nombre
    Autor = For_Ingresar.txt_autor
    Area = For_Ingresar.com_area.Value
End Sub
Sub PasarExcel()
    Sheets("Datos").Cells(Linea, 1) = Nombre
    Sheets("Datos").Cells(Linea, 2) = Autor
    Sheets("Datos").Cells(Linea, 3) = Area
    Linea = Linea + 1
End Sub
Sub LimpiarForm()
    For_Ingresar.txt_nombre = ""
    For_Ingresar.txt_autor = ""
    For_Ingresar.com_area.Value = ""
End Sub
Sub Final()
    Sheets("Ocultos").Cells(1, 7) = Linea 'Guardo el valor para poder continuar después
                                          'De esta forma no perdemos los datos anteriores
    For_Ingresar.Hide
End Sub
Sub Consultar()
    Sheets("Consultas").Select
    for_consultas.lst_areas.Visible = False
    for_consultas.txt_nombre.Visible = False
    for_consultas.txt_autor.Visible = False
    for_consultas.Show
End Sub
Sub Consultas()
    Sheets("Consultas").Range("A2:C30").Clear
    Linea = Sheets("Ocultos").Cells(1, 7)
    for_consultas.Hide
    If for_consultas.opt_area Then
        for_consultas.opt_area = False
        Call PorArea
    Else
        If for_consultas.opt_nombre Then
            for_consultas.opt_nombre = False
            Call PorNombre
        Else
            for_consultas.opt_autor = False
            Call PorAutor
        End If
    End If
End Sub
Sub PorNombre()
    Buscar = for_consultas.txt_nombre
    LineaConsultas = 2
    For I = 2 To Linea
        If Sheets("Datos").Cells(I, 1) = Buscar Then
            Sheets("Consultas").Cells(LineaConsultas, 1) = Sheets("Datos").Cells(I, 1)
            Sheets("Consultas").Cells(LineaConsultas, 2) = Sheets("Datos").Cells(I, 2)
            Sheets("Consultas").Cells(LineaConsultas, 3) = Sheets("Datos").Cells(I, 3)
            LineaConsultas = LineaConsultas + 1
        End If
    Next I
End Sub
Sub PorAutor()
    Buscar = for_consultas.txt_autor
    LineaConsultas = 2
    For I = 2 To Linea
        If Sheets("Datos").Cells(I, 2) = Buscar Then
            Sheets("Consultas").Cells(LineaConsultas, 1) = Sheets("Datos").Cells(I, 1)
            Sheets("Consultas").Cells(LineaConsultas, 2) = Sheets("Datos").Cells(I, 2)
            Sheets("Consultas").Cells(LineaConsultas, 3) = Sheets("Datos").Cells(I, 3)
            LineaConsultas = LineaConsultas + 1
        End If
    Next I
End Sub
Sub PorArea()
    Buscar = for_consultas.lst_areas.Value
    LineaConsultas = 2
    For I = 2 To Linea
        If Sheets("Datos").Cells(I, 3) = Buscar Then
            Sheets("Consultas").Cells(LineaConsultas, 1) = Sheets("Datos").Cells(I, 1)
            Sheets("Consultas").Cells(LineaConsultas, 2) = Sheets("Datos").Cells(I, 2)
            Sheets("Consultas").Cells(LineaConsultas, 3) = Sheets("Datos").Cells(I, 3)
            LineaConsultas = LineaConsultas + 1
        End If
    Next I
End Sub

