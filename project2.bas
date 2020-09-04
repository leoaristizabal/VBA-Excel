Attribute VB_Name = "Módulo1"
Option Explicit
Dim fila As Integer

Dim fila_encontro As Integer
Dim encontro_reg As Boolean

'INGRESAR PROVEEDOR
Dim servicios As String
'ELIMINAR
Dim i As Integer
'SELECCIONAR
Dim fila2 As Integer

Sub ing_prov()

fila = 5
servicios = ""
With Sheets("proveedores")

While .Cells(fila, 2) <> "" 'buscar primera fila vacia
    fila = fila + 1
Wend

encontro_reg = False

    While .Cells(fila, 2) <> ""
        If form_ingresar.txt_ced_ing = .Cells(fila, 2) Then
            
            encontro_reg = True
    End If
    fila = fila + 1
    Wend
    
    If encontro_reg = True Then
        form_ingresar.txt_result_ing = "Error: ¡Proveedor Existente!"
    
    Else
        form_ingresar.txt_result_ing = "¡Proveedor Ingresado Exitosamente!"
            'Entradas por cuadros de texto en form
            .Cells(fila, 2) = form_ingresar.txt_ced_ing
            .Cells(fila, 3) = form_ingresar.txt_nomb_ing
            .Cells(fila, 4) = form_ingresar.txt_tlf_ing
    
            'Ubicacion/Direccion por cuadro combinado
            .Cells(fila, 5) = form_ingresar.cmb_dir_ing
    
            'Servicio. Casillas de Verificacion
    
            If form_ingresar.chk_asce_ing = True Then
                 servicios = servicios & "Mantenimiento de Ascensores,"
            End If
    
            If form_ingresar.chk_bombas_ing Then
                servicios = servicios & " Mantenimiento de Bombas, "
            End If
    
            If form_ingresar.chk_elect_ing Then
                servicios = servicios & "Servicio Electrico, "
            End If
    
            If form_ingresar.chk_limpieza_ing Then
                servicios = servicios & "Servicio de Limpieza, "
            End If
    
            If form_ingresar.chk_otro_ing Then
                servicios = servicios & form_ingresar.txt_chkotro_ing
            End If
    
            'Impresion de datos en casilla de verificacion
            .Cells(fila, 6) = servicios
    
            'Sedes usando boton de opciones
            If form_ingresar.opt_sucur_ing = True Then
                .Cells(fila, 7) = "Sucursal"
            Else
                .Cells(fila, 7) = "Unica"
            End If

    End If
End With

    
End Sub

Sub busq_mod_f()

encontro_reg = False
fila_encontro = 0
fila = 5

With Sheets("proveedores")
    While .Cells(fila, 2) <> ""
        If form_modificar.txt_ced_mod = .Cells(fila, 2) Then
            encontro_reg = True
            fila_encontro = fila 'Registro buscado guardado en variable fila_encontro
        End If
    fila = fila + 1
    Wend
    
    If encontro_reg = True Then
        form_modificar.txt_resul_mod = "¡Proveedor Encontrado!"
        
        form_modificar_2.txt_ced_mod_mod = .Cells(fila_encontro, 2) 'impresiones en form modificar2
        
        form_modificar_2.txt_nomb_mod_mod = .Cells(fila_encontro, 3)
        form_modificar_2.txt_tlf_mod_mod = .Cells(fila_encontro, 4)
        form_modificar_2.cmb_dir_ing = .Cells(fila_encontro, 5)
        form_modificar_2.cmb_ser_mod_mod = .Cells(fila_encontro, 7)
    Else
        form_modificar.txt_resul_mod = "Proveedor NO Encontrado"
    End If
End With

End Sub

Sub modif_2_f()

With Sheets("proveedores")
    .Cells(fila_encontro, 2) = form_modificar_2.txt_ced_mod_mod
    .Cells(fila_encontro, 3) = form_modificar_2.txt_nomb_mod_mod
    .Cells(fila_encontro, 4) = form_modificar_2.txt_tlf_mod_mod
    .Cells(fila_encontro, 5) = form_modificar_2.cmb_dir_ing
    .Cells(fila_encontro, 7) = form_modificar_2.cmb_ser_mod_mod

    'Servicio. Casillas de Verificacion
    
    If form_modificar_2.chk_asce_mod_mod = True Then
        servicios = servicios & "Mantenimiento de Ascensores"
    End If
    
    If form_modificar_2.chk_bombas_mod_mod Then
        servicios = servicios & "Mantenimiento de Bombas"
    End If
    
    If form_modificar_2.chk_elect_mod_mod Then
        servicios = servicios & " Servicio Electrico"
    End If
    
    If form_modificar_2.chk_limpieza_mod_mod Then
        servicios = servicios & " Servicio de Limpieza"
    End If
    
    If form_modificar_2.chk_otro_mod_mod Then
        servicios = servicios & form_modificar_2.txt_chkotro_mod_mod
    End If
    
    'Impresion de datos actualizados en casilla de verificacion
    .Cells(fila_encontro, 6) = servicios

    form_modificar_2.txt_result_mod_mod = "¡Proveedor Actualizado con Exito!"
End With


End Sub

Sub eli_f()

'Definimos programa buscar que venimos usando
fila = 5
encontro_reg = False
fila_encontro = 0

With Sheets("proveedores")
    While .Cells(fila, 2) <> ""
        If form_eliminar.txt_ced_eli = .Cells(fila, 2) Then
            encontro_reg = True
            fila_encontro = fila
        End If
    fila = fila + 1
    Wend
    
    If encontro_reg = True Then
        For i = fila_encontro To fila - 1 'Subir el restante de registros la fila eliminada
            'Reescribiendo registro en la fila eliminada
            .Cells(i, 2) = .Cells(i + 1, 2)
            .Cells(i, 3) = .Cells(i + 1, 3)
            .Cells(i, 4) = .Cells(i + 1, 4)
            .Cells(i, 5) = .Cells(i + 1, 5)
            .Cells(i, 6) = .Cells(i + 1, 6)
            .Cells(i, 7) = .Cells(i + 1, 7)
            
        Next
        
        form_eliminar.txt_resul_eli = "Registro Eliminado Exitosamente"
    Else
        form_eliminar.txt_resul_eli = "Proveedor NO existe en base de datos"
    End If
End With

End Sub

Sub seleccionar_f()

encontro_reg = False
fila = 5
fila2 = 5

With Sheets("proveedores")
    While .Cells(fila, 2) <> ""
        If form_seleccionar.cmb_dir_sel = .Cells(fila, 5) Then
            Sheets("ubicacion").Cells(fila2, 2) = .Cells(fila, 2)
            Sheets("ubicacion").Cells(fila2, 3) = .Cells(fila, 3)
            Sheets("ubicacion").Cells(fila2, 4) = .Cells(fila, 4)
            Sheets("ubicacion").Cells(fila2, 5) = .Cells(fila, 5)
            Sheets("ubicacion").Cells(fila2, 6) = .Cells(fila, 6)
            Sheets("ubicacion").Cells(fila2, 7) = .Cells(fila, 7)
            fila2 = fila2 + 1
            encontro_reg = True
        End If
    fila = fila + 1
    Wend
    If encontro_reg = True Then
        form_seleccionar.txt_result_sel = "Reporte emitido exitosamente, ir a la hoja UBICACION"
    Else
        form_seleccionar.txt_result_sel = "No existen proveedores en esta zona"
    End If
End With

End Sub

Sub limpiar_BD()

Sheets("proveedores").Range("B5:G20").Value = ""

End Sub
