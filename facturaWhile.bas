Attribute VB_Name = "Módulo1"
Option Explicit
'Declaracion de variables
Dim fila As Integer
Dim mensaje As String

Dim iva As Double
Dim monto_total As Double
Dim cod_comp As String
Dim linea_prod As String
Dim modo_pago As String
Dim descuento As Double
Dim monto_compra As Double


cod_comp = InputBox("INTRODUCE EL CÓDIGO DE COMPRADOR: ", "CODIGO COMPRADOR")

linea_prod = InputBox("INTRODUCE EL CÓDIGO SEGÚN EL TIPO DE PRODUCTO(Q,F,H): ", "TIPO DE LÍNEA DE PRODUCTO")

modo_pago = InputBox("INTRODUCE EL CÓDIGO SEGÚN EL MÉTODO DE PAGO (Cbs,  CRbs,Z): ", "MODALIDAD DE PAGO")

With Sheets("factura")

    .Cells(fila, 2) = cod_comp
    .Cells(fila, 3) = linea_prod
    .Cells(fila, 4) = modo_pago
End With


If linea_prod = "Q" Then 'estructura condicional anidada para productos Quimicos, segun el modo de pago
    
    If modo_pago = "Z" Then
        monto_compra = InputBox("INTRODUCE EL MONTO TOTAL DE LA COMPRA ($): ", "MONTO DE COMPRA")
        Sheets("factura").Cells(fila, 5) = monto_compra
    
        iva = monto_compra * 0.16 'Impuesto del 16% para los productos Químicos
        Sheets("factura").Cells(fila, 7) = iva
        
        monto_total = monto_compra + iva
        Sheets("factura").Cells(fila, 8) = monto_total
        
        
            If modo_pago = "Z" And monto_compra >= 10000 Then 'estructura condicional simple para descuento si el pago es con zelle y mayor a 10000$
    
              descuento = monto_compra * 0.1
              Sheets("factura").Cells(fila, 6) = descuento
              MsgBox ("UD TIENE UN DESCUENTO DE: " & descuento)
              monto_total = monto_compra + iva - descuento
                Sheets("factura").Cells(fila, 8) = monto_total
            End If
    
        
    Else
        
        monto_compra = InputBox("INTRODUCE EL MONTO TOTAL DE LA COMPRA (Bs): ", "MONTO DE COMPRA") 'Pago en bs sin descuento
        Sheets("factura").Cells(fila, 5) = monto_compra
        
        iva = monto_compra * 0.16 'Impuesto del 16% para los productos Químicos
        Sheets("factura").Cells(fila, 7) = iva
        
        monto_total = monto_compra + iva
        Sheets("factura").Cells(fila, 8) = monto_total
   
   End If


Else
   If linea_prod = "H" Then
    
        If modo_pago = "Z" Then
        monto_compra = InputBox("INTRODUCE EL MONTO TOTAL DE LA COMPRA ($): ", "MONTO DE COMPRA")
        Sheets("factura").Cells(fila, 5) = monto_compra
    
        iva = monto_compra * 0.08 'Impuesto del 8% para los productos Hidrocarburos
        Sheets("factura").Cells(fila, 7) = iva
        monto_total = monto_compra + iva
        Sheets("factura").Cells(fila, 8) = monto_total
        
            If modo_pago = "Z" And monto_compra >= 10000 Then 'estructura condicional simple para descuento si el pago es con zelle y mayor a 10000$
    
              descuento = monto_compra * 0.1
              Sheets("factura").Cells(fila, 6) = descuento
              MsgBox ("UD TIENE UN DESCUENTO DE: " & descuento)
              monto_total = monto_compra + iva - descuento
                Sheets("factura").Cells(fila, 8) = monto_total
            End If
        
        
        
        Else
        
        monto_compra = InputBox("INTRODUCE EL MONTO TOTAL DE LA COMPRA (Bs): ", "MONTO DE COMPRA") 'Pago en bs sin descuento
        Sheets("factura").Cells(fila, 5) = monto_compra
        
        iva = monto_compra * 0.08 'Impuesto del 8% para los productos Hidrocarburos
        Sheets("factura").Cells(fila, 7) = iva
        
        monto_total = monto_compra + iva
        Sheets("factura").Cells(fila, 8) = monto_total
   
        End If
End If

    
If linea_prod = "F" Then
     
        MsgBox ("UD ESTÁ EXONERADO DE IMPUESTO")
        
        If modo_pago = "Z" Then
        monto_compra = InputBox("INTRODUCE EL MONTO TOTAL DE LA COMPRA ($): ", "MONTO DE COMPRA")
        Sheets("factura").Cells(fila, 5) = monto_compra
        
        iva = monto_compra * 0 'Impuesto del 0% para los productos farnaceúticos (HACEMOS ESTO PARA ASIGNARLE EL VALOR CERO A LA CASILLA IVA Y NO QUEDE VACÍA)
        Sheets("factura").Cells(fila, 7) = iva
        
        monto_total = monto_compra + iva
        Sheets("factura").Cells(fila, 8) = monto_total
        
            If modo_pago = "Z" And monto_compra >= 10000 Then 'estructura condicional simple para descuento si el pago es con zelle y mayor a 10000$
    
              descuento = monto_compra * 0.1
              Sheets("factura").Cells(fila, 6) = descuento
              MsgBox ("UD TIENE UN DESCUENTO DE: " & descuento)
              monto_total = monto_compra + iva - descuento
                Sheets("factura").Cells(fila, 8) = monto_total
            End If
        
        
        Else
        
        monto_compra = InputBox("INTRODUCE EL MONTO TOTAL DE LA COMPRA (Bs): ", "MONTO DE COMPRA") 'Pago en bs sin descuento
        Sheets("factura").Cells(fila, 5) = monto_compra
        
        iva = monto_compra * 0 'Impuesto del 0% para los productos farnaceúticos (HACEMOS ESTO PARA ASIGNARLE EL VALOR CERO A LA CASILLA IVA Y NO QUEDE VACÍA)
        Sheets("factura").Cells(fila, 7) = iva
        
        monto_total = monto_compra + iva
        Sheets("factura").Cells(fila, 8) = monto_total
   
        End If
     
End If

End If



With Sheets("factura")
.Range("B5:H12").Value = " "
.Range("L5:L12").Value = " "

End With

'NOTA: SIGO TENIENDO PROBLEMAS CON LOS BOTONES, A VECES APARECEN Y OTRAS NO. LEÍ EN INTERNET Y ES UN PROBLEMA QUE A VECES SUCEDE EN OFFICE 2010 EN LAS OTRAS VERSIONES DEBERIAN SALIR
'IGUALMENTE, ESTAN CONFIGURADOS 4 BOTONES, UNO PARA EJECUTAR LAS ENTRADAS DE DATOS, OTRAS PARA PROCESAR LA FACTURA Y EL ÚLTIMO PARA
'LIMPIAR LAS CELDAS. ADICIONALMENTE, EN MODO DE PRUEBA, AÑADÍ OTRO BOTON PARA VER SI APARECIAN LOS MODULOS DE LOS OTROS, ESTE SE LLAMÓ PRUEBA
'EN CASO DE NO APARECER LOS BOTONOS PUEDE VERIFICAR SUS MODULOS EN LA PESTAÑA PROGRAMADOR, MODO DISEÑO Y LUEGO VER COFIGO




fila = 5 'inicializador de la posicion en la hoja de calculo
mensaje = MsgBox("¿Hay mas datos?", vbYesNo, "Ingresar Datos")

While mensaje = vbYes
    Call entradas_datos
    Call condiciones

mensaje = MsgBox("¿Hay mas datos?", vbYesNo, "Ingresar Datos")
fila = fila + 1
Wend



