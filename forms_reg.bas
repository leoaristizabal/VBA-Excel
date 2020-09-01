Attribute VB_Name = "Módulo1"
Option Explicit
Dim fila As Integer
Dim pos As Integer
Dim i As Integer
Dim fila_encontro As Integer
Dim encontro_reg As Boolean
Sub menu()
form_menu.Show
End Sub


Sub procesar()

fila = 4
While Sheets("Datos").Cells(fila, 1) <> ""
    fila = fila + 1
Wend
With Sheets("Datos")
    
    .Cells(fila, 1) = form_datos.txt_codigo 'Impresion Codigo
    .Cells(fila, 2) = form_datos.txt_nombre
    .Cells(fila, 3) = form_datos.txt_usu
    .Cells(fila, 4) = form_datos.txt_cont
    .Cells(fila, 5) = form_datos.cmb_edociv 'Impresion cuadro combinado Estado Civil
    .Cells(fila, 6) = form_datos.spin_edad
    .Cells(fila, 7) = form_datos.spin_antig

End With
End Sub

Sub limpiar_form_datos()
With form_datos
    .txt_nombre = ""
    .txt_usu = ""
    .txt_cont = ""
    .txt_edad = ""
    .txt_antig = ""
    .spin_edad = 18
    .spin_antig = 0
End With
End Sub
Sub eliminar_f()
'Para eliminar un registro debemos buscarlo primero por lo que el programa Buscar se repite

encontro_reg = False
fila_encontro = 0
fila = 4
With Sheets("Datos")
    While .Cells(fila, 2) <> ""
        If form_eliminar.txt_cod_eli = .Cells(fila, 1) Then
            encontro_reg = True
            fila_encontro = fila
        End If
    fila = fila + 1
    Wend
    
    If encontro_reg = True Then
        For i = fila_encontro To fila - 1 ' subiendo registros restantes para no dejar filas em blanco
            .Cells(i, 1) = .Cells(i + 1, 1)
            .Cells(i, 2) = .Cells(i + 1, 2)
            .Cells(i, 3) = .Cells(i + 1, 3)
            .Cells(i, 4) = .Cells(i + 1, 4)
            .Cells(i, 5) = .Cells(i + 1, 5)
            .Cells(i, 6) = .Cells(i + 1, 6)
            .Cells(i, 7) = .Cells(i + 1, 7)
        Next
        form_eliminar.txtx_mensaje_elim = "¡Registro Eliminado Exitosamente!"
    Else
        form_eliminar.txtx_mensaje_elim = "Registro NO Encontrado!"
    End If
End With
End Sub

Sub buscar_f()
encontro_reg = False
fila_encontro = 0
fila = 4
With Sheets("Datos")
    While .Cells(fila, 2) <> ""
        If form_buscar.txt_codigo_b = .Cells(fila, 1) Then
            encontro_reg = True
            fila_encontro = fila
        End If
    fila = fila + 1
    Wend
        
    If encontro_reg = True Then
        form_buscar.txt_mensaje_b = "¡Registro Encontrado!"
        
        form_buscar.txt_nomb_bus = .Cells(fila_encontro, 2)
        form_buscar.txt_usu_busc = .Cells(fila_encontro, 3)
        form_buscar.txt_edad_busc = .Cells(fila_encontro, 6)
    Else
        form_buscar.txt_mensaje_b = "¡Registro NO Encontrado!"
    End If
End With

End Sub

Sub actualizar_buscar() 'ACTUALIZAR BASE CON CAMPOS DEL FORM ENCONTRADOS (Continuacion Buscar)
With Sheets("Datos")
    .Cells(fila_encontro, 2) = form_buscar.txt_nomb_bus
    .Cells(fila_encontro, 3) = form_buscar.txt_usu_busc
    .Cells(fila_encontro, 6) = form_buscar.txt_edad_busc
End With

End Sub
Sub modificar_f()

'Antes de modificar debemos BUSCAR el registro a trabajar
encontro_reg = False
fila_encontro = 0
fila = 4
With Sheets("Datos")
    While .Cells(fila, 2) <> ""
        If form_modificar.txt_codigo_mod = .Cells(fila, 1) Then
            encontro_reg = True
            fila_encontro = fila
        End If
    fila = fila + 1
    Wend

    If encontro_reg = True Then
        
    form_modificar.txt_mensaje_mod = "¡Registro Encontrado!"
        
        form_modificar.txt_nomb_mod = .Cells(fila_encontro, 2)
        form_modificar.txt_usu_mod = .Cells(fila_encontro, 3)
        form_modificar.txt_edad_mod = .Cells(fila_encontro, 6)
    Else
        form_modificar.txt_mensaje_mod = "¡Registro NO Encontrado!"
    End If
End With
    
End Sub
    
Sub modificar_act() 'Continuacion Sub modificar_f
'Actualizar Datos

With Sheets("Datos")
    .Cells(fila_encontro, 2) = form_modificar.txt_nomb_mod
    .Cells(fila_encontro, 3) = form_modificar.txt_usu_mod
    .Cells(fila_encontro, 6) = form_modificar.txt_edad_mod
End With
   
End Sub


