Attribute VB_Name = "Módulo1"
Option Explicit
Dim fila As Integer
Dim amedp As Integer
Dim tloc As Integer
Dim totl As Integer ' total de locales
Dim totmp As Integer ' mmedicos privados
Dim hosps As Integer ' hospitales con subsidios
Dim tots As Integer ' tota; con subsidio
Dim pub As Integer
Dim priv As Integer
Dim promp As Double
Dim proms As Double


' muestra el menu
Sub menu()
form_menu.Show
End Sub
' inicio limpia el form y lo muestra
Sub inicio()
Call limpiar_form
form_datos.Show
End Sub
' contar alñs filas llenas
Sub contar_filas()
fila = 3
While Sheets("Datos").Cells(fila, 1) <> ""
    fila = fila + 1
Wend
End Sub
' pasar lso datos del form a la hoja
Sub procesar()
' cuento las filas
Call contar_filas
' paso los datos
With Sheets("Datos")
    ' nombre
    .Cells(fila, 1) = form_datos.txt_nom
    ' categoria
    If form_datos.tgl_cat = True Then
        .Cells(fila, 2) = "PUBLICO"
    Else
        .Cells(fila, 2) = "PRIVADO"
    End If
    ' subsidio
    If form_datos.opt_si = True Then
        .Cells(fila, 3) = form_datos.cmb_subsidio
    Else
    
        .Cells(fila, 3) = "NO"
    End If
    ' cant medicos
    .Cells(fila, 4) = form_datos.spin_cmed
    ' tipo de local
    If form_datos.cmb_tipo = "OTRO" Then
        .Cells(fila, 5) = form_datos.txt_otro
    Else
        .Cells(fila, 5) = form_datos.cmb_tipo
    End If
End With
End Sub
' limpiar el form
Sub limpiar_form()
With form_datos
    .txt_nom = ""
    .txt_cmed = ""
    .txt_otro = ""
    .opt_no = False
    .opt_si = False
    .tgl_cat = False
    .spin_cmed = 1
    .txt_otro.Visible = False
    .cmb_subsidio.Visible = False
    .cmb_tipo = ""
    .cmb_tipo.RowSource = "PRIVADO"
End With
End Sub
Sub calculos()
Dim cat As String
Dim subs As String
Dim cantm As Integer
Dim tipo As String
Dim i As Integer
' calcula y escribe las estadisticas
Call contar_filas
' inicializo los acumuladores y contadores
Call inicializar
' ciclo de recorrido de la tabla
For i = 3 To fila - 1
    With Sheets("Datos")
        ' lectura de valores
        cat = .Cells(i, 2)
        subs = .Cells(i, 3)
        cantm = .Cells(i, 4)
        tipo = .Cells(i, 5)
        '  total de locales privados
        totl = totl + 1
        ' medicos en el sector privado
        If cat = "PRIVADO" Then
            ' medicos en el sector privado
            totmp = totmp + cantm
            '  total de locales privados
            totl = totl + 1
        End If
        ' locales con subsidio
        If subs <> "NO" Then
            tots = tots + 1
        ' hospitales con subsidio
            If tipo = "HOSPITAL" Then
                hosps = hosps + 1
            End If
        End If
        ' locales de cada categoria
        Select Case cat
        Case "PUBLICO"
            pub = pub + 1
        Case "PRIVADO"
            priv = priv + 1
        End Select
    End With
Next
End Sub
Sub reporte()
'  calcula los promediso y escribe resultados
If totl > 0 Then
    promp = totmp / totl
Else
    promp = 0
End If
If tots > 0 Then
    proms = hosps / tots
Else
    proms = 0
End If
'  reporte
With Sheets("Reporte")
    .Cells(3, 4) = promp
    .Cells(4, 4) = proms
    .Cells(6, 4) = pub
    .Cells(6, 5) = priv
    .Select
End With
End Sub

Sub inicializar()
' inicializar variables
totl = 0
totmp = 0
hosps = 0
tots = 0
pub = 0
priv = 0
End Sub
