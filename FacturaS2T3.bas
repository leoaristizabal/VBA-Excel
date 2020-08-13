Attribute VB_Name = "Módulo2"
Option Explicit
'Declaracion de variables de entrada
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

Sheets("factura").Cells(5, 2) = cod_comp
Sheets("factura").Cells(5, 3) = linea_prod
Sheets("factura").Cells(5, 4) = modo_pago


If linea_prod = "Q" Then 'estructura condicional anidada para productos Quimicos, segun el modo de pago
    
    If modo_pago = "Z" Then
        monto_compra = InputBox("INTRODUCE EL MONTO TOTAL DE LA COMPRA ($): ", "MONTO DE COMPRA")
        Sheets("factura").Cells(5, 5) = monto_compra
    
        iva = monto_compra * 0.16 'Impuesto del 16% para los productos Químicos
        Sheets("factura").Cells(5, 9) = iva
        
        monto_total = monto_compra + iva
        Sheets("factura").Cells(5, 7) = monto_total
        
        
            If modo_pago = "Z" And monto_compra >= 10000 Then 'estructura condicional simple para descuento si el pago es con zelle y mayor a 10000$
    
              descuento = monto_compra * 0.1
              Sheets("factura").Cells(5, 6) = descuento
              MsgBox ("UD TIENE UN DESCUENTO DE: " & descuento)
              monto_total = monto_compra + iva - descuento
                Sheets("factura").Cells(5, 7) = monto_total
            End If
    
        
    Else
        
        monto_compra = InputBox("INTRODUCE EL MONTO TOTAL DE LA COMPRA (Bs): ", "MONTO DE COMPRA") 'Pago en bs sin descuento
        Sheets("factura").Cells(5, 5) = monto_compra
        
        iva = monto_compra * 0.16 'Impuesto del 16% para los productos Químicos
        Sheets("factura").Cells(5, 9) = iva
        
        monto_total = monto_compra + iva
        Sheets("factura").Cells(5, 7) = monto_total
   
   End If


Else
   If linea_prod = "H" Then
    
        If modo_pago = "Z" Then
        monto_compra = InputBox("INTRODUCE EL MONTO TOTAL DE LA COMPRA ($): ", "MONTO DE COMPRA")
        Sheets("factura").Cells(5, 5) = monto_compra
    
        iva = monto_compra * 0.08 'Impuesto del 8% para los productos Hidrocarburos
        Sheets("factura").Cells(5, 9) = iva
        monto_total = monto_compra + iva
        Sheets("factura").Cells(5, 7) = monto_total
        
            If modo_pago = "Z" And monto_compra >= 10000 Then 'estructura condicional simple para descuento si el pago es con zelle y mayor a 10000$
    
              descuento = monto_compra * 0.1
              Sheets("factura").Cells(5, 6) = descuento
              MsgBox ("UD TIENE UN DESCUENTO DE: " & descuento)
              monto_total = monto_compra + iva - descuento
                Sheets("factura").Cells(5, 7) = monto_total
            End If
        
        
        
        Else
        
        monto_compra = InputBox("INTRODUCE EL MONTO TOTAL DE LA COMPRA (Bs): ", "MONTO DE COMPRA") 'Pago en bs sin descuento
        Sheets("factura").Cells(5, 5) = monto_compra
        
        iva = monto_compra * 0.08 'Impuesto del 8% para los productos Hidrocarburos
        Sheets("factura").Cells(5, 9) = iva
        
        monto_total = monto_compra + iva
        Sheets("factura").Cells(5, 7) = monto_total
   
        End If
End If

    
If linea_prod = "F" Then
     
        MsgBox ("UD ESTÁ EXONERADO DE IMPUESTO")
        
        If modo_pago = "Z" Then
        monto_compra = InputBox("INTRODUCE EL MONTO TOTAL DE LA COMPRA ($): ", "MONTO DE COMPRA")
        Sheets("factura").Cells(5, 5) = monto_compra
        
        iva = monto_compra * 0 'Impuesto del 0% para los productos farnaceúticos (HACEMOS ESTO PARA ASIGNARLE EL VALOR CERO A LA CASILLA IVA Y NO QUEDE VACÍA)
        Sheets("factura").Cells(5, 9) = iva
        
        monto_total = monto_compra + iva
        Sheets("factura").Cells(5, 7) = monto_total
        
            If modo_pago = "Z" And monto_compra >= 10000 Then 'estructura condicional simple para descuento si el pago es con zelle y mayor a 10000$
    
              descuento = monto_compra * 0.1
              Sheets("factura").Cells(5, 6) = descuento
              MsgBox ("UD TIENE UN DESCUENTO DE: " & descuento)
              monto_total = monto_compra + iva - descuento
                Sheets("factura").Cells(5, 7) = monto_total
            End If
        
        
        Else
        
        monto_compra = InputBox("INTRODUCE EL MONTO TOTAL DE LA COMPRA (Bs): ", "MONTO DE COMPRA") 'Pago en bs sin descuento
        Sheets("factura").Cells(5, 5) = monto_compra
        
        iva = monto_compra * 0 'Impuesto del 0% para los productos farnaceúticos (HACEMOS ESTO PARA ASIGNARLE EL VALOR CERO A LA CASILLA IVA Y NO QUEDE VACÍA)
        Sheets("factura").Cells(5, 9) = iva
        
        monto_total = monto_compra + iva
        Sheets("factura").Cells(5, 7) = monto_total
   
        End If
     
End If

End If



Sheets("factura").Range("B5:I5").Value = " "




End Sub
