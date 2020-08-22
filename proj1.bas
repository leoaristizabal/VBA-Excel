Attribute VB_Name = "Módulo2"
Option Explicit

'CONSUMIBLES
Dim cod_prod As String
Dim nom_cons As String
Dim marca As String
Dim tipo_med As String
Dim cantidad As Integer
Dim fecha As Date
Dim prec_uni As Double

Dim cant_buscar As Integer
Dim cod_buscar As String
Dim accion As String
Dim fecha1 As Date
Dim dif_fechas As Date

'PREVENTIVOS
Dim fecha_i As Date
Dim fecha_f As Date
Dim fila2 As Integer
Dim frecuencia As Integer


'PROVEEDORES
Dim fila As Integer
Dim mensaje As String
Dim rif As String
Dim nombre As String
Dim tlf As String
Dim email As String
Dim red_social As String
Dim servicios As String

Sub programa_proveedores()
fila = 5 'FILA DONDE INICIA LA BASE DE DATOS
    While Sheets("Prov").Cells(fila, 2) <> "" 'BUSCAR PRIMERA FILA DISPONIBLE EN BASE DE DATOS
        fila = fila + 1
    Wend
    
mensaje = MsgBox("¿DESEA INGRESAR MAS DATOS?", vbYesNo, "INGRESAR DATOS") 'CONDICION DE PARADA DEL WHILE
    
    While mensaje = vbYes
    
inicio:
        
        rif = InputBox("INGRESA RIF O CI: ", "RIF/CI")
        nombre = InputBox("INGRESA NOMBRE DE LA EMPRESA O PERSONA: ", "NOMBRE")
        
        If rif = "" Or nombre = "" Then
        
             MsgBox ("ERROR: USTED NO HA INGRESADO EL RIF O NOMBRE") 'PRIMER REQUERIMIENTO DEL PROGRAMA, SI SE DEJA EL RIF O NOMBRE EN BLANCO DA ERROR Y VUELVE AL INICIO
             GoTo inicio
        End If
        
        tlf = InputBox("INGRESE TELEFONO DE CONTACTO: ", "TELEFONO")
        email = InputBox("INGRESE EMAIL DE CONTACTO: ", "EMAIL")
        red_social = InputBox("INGRESE USER DE RED SOCIAL: ", "RED SOCIAL")
        servicios = InputBox("INGRESE EL TIPO DE SERVICIO QUE OFRECE: ", "SERVICIO")
    
    'IMPRESION DE DATOS INGRESADOS
    
    With Sheets("Prov")
        .Cells(fila, 2) = rif
        .Cells(fila, 3) = nombre
        .Cells(fila, 4) = tlf
        .Cells(fila, 5) = email
        .Cells(fila, 6) = red_social
        .Cells(fila, 7) = servicios
    End With
    
    fila = fila + 1 'ACTUALIZACION DE FILA
    mensaje = MsgBox("¿Hay mas datos?", vbYesNo, "Ingresar Datos") 'PARADA/SALIDA DEL WHILE EN CASO DE VBNO
    Wend

    End Sub
Sub programa_incluir_consumibles()
fila = 5 'FILA DONDE INICIA LA BASE DE DATOS
    While Sheets("Consumible").Cells(fila, 2) <> "" 'BUSCAR PRIMERA FILA DISPONIBLE EN BASE DE DATOS
        fila = fila + 1
    Wend

mensaje = MsgBox("¿DESEA INGRESAR NUEVOS PRODUCTOS?", vbYesNo, "INGRESAR PRODUCTOS") 'CONDICION DE PARADA DEL WHILE
    
    While mensaje = vbYes
    
    'ENTRADAS DE DATOS CONSUMIBLES

        cod_prod = InputBox("INGRESA EL CODIGO DEL PRODUCTO: ", "CODIGO DE PRODUCTO")
        nom_cons = InputBox("INGRESA EL NOMBRE DEL PRODUCTO: ", "NOMBRE DEL PRODUCTO")
        marca = InputBox("INGRESA LA MARCA DEL PRODUCTO: ", "MARCA DEL PRODUCTO")
        tipo_med = InputBox("INGRESA EL TIPO DE MEDIDA: ", " TIPO DE MEDIDA")
        cantidad = InputBox("INGRESA LA CANTIDAD EN EXISTENCIA: ", "CANTIDAD EN EXISTENCIA")
        fecha = InputBox("INGRESA LA FECHA DE LA ULTIMA COMPRA DEL PRODUCTO: ", "FECHA ULTIMA COMPRA")
        prec_uni = InputBox("INGRESA EL PRECIO UNITARIO DEL PRODUCTO: ", "PRECIO UNITARIO")
    
    'IMPRESION DE DATOS INGRESADOS
    
      With Sheets("Consumible")
        .Cells(fila, 2) = cod_prod
        .Cells(fila, 3) = nom_cons
        .Cells(fila, 4) = marca
        .Cells(fila, 5) = tipo_med
        .Cells(fila, 6) = cantidad
        .Cells(fila, 7) = fecha
        .Cells(fila, 8) = prec_uni
        End With
        
fila = fila + 1
mensaje = MsgBox("¿DESEA INGRESAR NUEVOS PRODUCTOS?", vbYesNo, "INGRESAR PRODUCTOS") 'CONDICION DE PARADA DEL WHILE
Wend
End Sub

Sub modificar_consumibles()

fila = 5 'inicializador de filas

    cod_buscar = InputBox("INGRESA EL CODIGO DEL PRODUCTO A BUSCAR: ", "CODIGO DE PRODUCTO A BUSCAR")
    
    While Sheets("Consumible").Cells(fila, 2) <> "" 'Mientras la fila no esté vacía

        If cod_buscar = Sheets("Consumible").Cells(fila, 2) Then
    
            cant_buscar = InputBox("INGRESE LA CANTIDAD A AÑADIR O RETIRAR: ", "CANTIDAD")
            fecha1 = InputBox("INGRESE LA FECHA DE LA TRANSACCION", "FECHA")
            accion = InputBox("INGRESE A PARA AÑADIR PRODUCTOCHr(13) INGRESE R PARA RETIRAR PRODUCTO", "AÑADIR O RETIRAR PRODUCTO")
  
                If accion = "A" Then
                    Sheets("Consumible").Cells(fila, 6) = Sheets("Consumible").Cells(fila, 6) + cant_buscar
                    Sheets("Consumible").Cells(fila, 7) = fecha1
                End If
            
                If accion = "R" Then
                    Sheets("Consumible").Cells(fila, 6) = Sheets("Consumible").Cells(fila, 6) - cant_buscar
                    Sheets("Consumible").Cells(fila, 7) = fecha1
                End If
        
        End If
    fila = fila + 1 'actualizar fila
    Wend
    
End Sub
Sub preventivo()

fila = 5 'inicio fila Mtoequipos
fila2 = 5 'inicio fila Seleccion

mensaje = MsgBox("¿DESEA VER EL MANTENIMIENTO DE LOS PRODUCTOS?", vbYesNo, "MANTENIMIENTO DE EQUIPOS") 'condicion de parada dandole NO

While mensaje = vbYes

fecha_i = InputBox("INGRESA LA FECHA DE MANTENIMIENTO A EXAMINAR: ", "FECHA MANTENIMIENTO") 'ingresar fecha por cuadro de dialogo

    While Sheets("Consumible").Cells(fila, 6) <> "" 'condicoin parada mientras no sea vacio

        fecha_f = Sheets("Mtoequipos").Cells(fila, 6) 'asignando rango de valores
        
        dif_fechas = fecha_i - fecha_f 'diferencia entre fechas
        frecuencia = dif_fechas / 30 'diferencia en meses
        
        
        If frecuencia >= Sheets("Mtoequipos").Cells(fila, 5) Then ' comprobando si los meses de la diferencia son mayores a la frecuencia de meses

            'impresiones
            
            Sheets("Seleccion").Cells(fila2, 2) = Sheets("Mtoequipos").Cells(fila, 2)
            Sheets("Seleccion").Cells(fila2, 3) = Sheets("Mtoequipos").Cells(fila, 3)
            Sheets("Seleccion").Cells(fila2, 4) = Sheets("Mtoequipos").Cells(fila, 4)
            Sheets("Seleccion").Cells(fila2, 5) = Sheets("Mtoequipos").Cells(fila, 5)
            Sheets("Seleccion").Cells(fila2, 6) = Sheets("Mtoequipos").Cells(fila, 6)
            Sheets("Seleccion").Cells(fila2, 7) = -Sheets("Mtoequipos").Cells(fila, 5) + frecuencia
            
        fila2 = fila2 + 1 'actualizar fila de Seleccion
        End If
            

    fila = fila + 1 ''Actualizar filas Mtoequipos
    Wend

mensaje = MsgBox("¿Hay mas datos?", vbYesNo, "Ingresar Datos") 'parada o inicio
Wend
End Sub


Sub limpiar_provedores()
Sheets("Prov").Range("B15:G34").Value = ""

End Sub

