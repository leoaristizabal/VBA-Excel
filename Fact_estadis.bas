Attribute VB_Name = "Módulo2"
Option Explicit
'Declaracion de variables nuevas
Dim fila As Integer
Dim mensaje As String

'declaracion de variables estadisticas
Dim cant_f As Integer
Dim cant_q As Integer
Dim cant_h As Integer
Dim cant_z As Integer

Dim acu_impuesto As Double
Dim acu_descuento As Double
Dim acu_total As Double


Dim iva As Double
Dim monto_total As Double
Dim cod_comp As String
Dim linea_prod As String
Dim modo_pago As String
Dim descuento As Double
Dim monto_compra As Double
Sub programa_principal() ' ejecutar programa completo

    Call repeticion
    
End Sub

Sub entradas_datos()

cod_comp = InputBox("INTRODUCE EL CÓDIGO DE COMPRADOR: ", "CODIGO COMPRADOR")

linea_prod = InputBox("INTRODUCE EL CÓDIGO SEGÚN EL TIPO DE PRODUCTO(Q,F,H): ", "TIPO DE LÍNEA DE PRODUCTO")

modo_pago = InputBox("INTRODUCE EL CÓDIGO SEGÚN EL MÉTODO DE PAGO (Cbs,  CRbs,Z): ", "MODALIDAD DE PAGO")

With Sheets("factura")

    .Cells(fila, 2) = cod_comp
    .Cells(fila, 3) = linea_prod
    .Cells(fila, 4) = modo_pago
End With

If linea_prod = "Q" Then 'estructura condicional anidada para productos Quimicos, segun el modo de pago
    cant_q = cant_q + 1
    
    If modo_pago = "Z" Then
        monto_compra = InputBox("INTRODUCE EL MONTO TOTAL DE LA COMPRA ($): ", "MONTO DE COMPRA")
        Sheets("factura").Cells(fila, 5) = monto_compra
        cant_z = cant_z + 1
    
        iva = monto_compra * 0.16 'Impuesto del 16% para los productos Químicos
        Sheets("factura").Cells(fila, 7) = iva
        acu_impuesto = iva + acu_impuesto
        
        monto_total = monto_compra + iva
        Sheets("factura").Cells(fila, 8) = monto_total
        acu_total = monto_total + acu_total
        
        
            If modo_pago = "Z" And monto_compra >= 10000 Then 'estructura condicional simple para descuento si el pago es con zelle y mayor a 10000$
    
              descuento = monto_compra * 0.1
              Sheets("factura").Cells(fila, 6) = descuento
              acu_descuento = acu_descuento + descuento
              
              MsgBox ("UD TIENE UN DESCUENTO DE: " & descuento)
              monto_total = monto_compra + iva - descuento
                Sheets("factura").Cells(fila, 8) = monto_total
                acu_total = acu_total + monto_total
                
            End If
    
        
    Else
        
        monto_compra = InputBox("INTRODUCE EL MONTO TOTAL DE LA COMPRA (Bs): ", "MONTO DE COMPRA") 'Pago en bs sin descuento
        Sheets("factura").Cells(fila, 5) = monto_compra
        
        
        iva = monto_compra * 0.16 'Impuesto del 16% para los productos Químicos
        Sheets("factura").Cells(fila, 7) = iva
        acu_impuesto = iva + acu_impuesto
        
        monto_total = monto_compra + iva
        Sheets("factura").Cells(fila, 8) = monto_total
        acu_total = acu_total + monto_total
   End If


Else
   If linea_prod = "H" Then
        cant_h = cant_h + 1
        
        If modo_pago = "Z" Then
        monto_compra = InputBox("INTRODUCE EL MONTO TOTAL DE LA COMPRA ($): ", "MONTO DE COMPRA")
        Sheets("factura").Cells(fila, 5) = monto_compra
        cant_z = cant_z + 1
    
        iva = monto_compra * 0.08 'Impuesto del 8% para los productos Hidrocarburos
        Sheets("factura").Cells(fila, 7) = iva
        acu_impuesto = acu_impuesto + iva
        
        monto_total = monto_compra + iva
        Sheets("factura").Cells(fila, 8) = monto_total
        acu_total = acu_total + monto_total
        
            If modo_pago = "Z" And monto_compra >= 10000 Then 'estructura condicional simple para descuento si el pago es con zelle y mayor a 10000$
        
               
              descuento = monto_compra * 0.1
              Sheets("factura").Cells(fila, 6) = descuento
             acu_descuento = acu_descuento + descuento
             
              MsgBox ("UD TIENE UN DESCUENTO DE: " & descuento)
              monto_total = monto_compra + iva - descuento
                Sheets("factura").Cells(fila, 8) = monto_total
                acu_total = acu_total + monto_total
                
            End If
        
        
        
        Else
        
        monto_compra = InputBox("INTRODUCE EL MONTO TOTAL DE LA COMPRA (Bs): ", "MONTO DE COMPRA") 'Pago en bs sin descuento
        Sheets("factura").Cells(fila, 5) = monto_compra
        
        iva = monto_compra * 0.08 'Impuesto del 8% para los productos Hidrocarburos
        Sheets("factura").Cells(fila, 7) = iva
        acu_impuesto = acu_impuesto + iva
        
        monto_total = monto_compra + iva
        Sheets("factura").Cells(fila, 8) = monto_total
        acu_total = acu_total + monto_total
        
        End If
End If

    
If linea_prod = "F" Then
        
        cant_f = cant_f + 1
        
        MsgBox ("UD ESTÁ EXONERADO DE IMPUESTO")
        
        If modo_pago = "Z" Then
        monto_compra = InputBox("INTRODUCE EL MONTO TOTAL DE LA COMPRA ($): ", "MONTO DE COMPRA")
        Sheets("factura").Cells(fila, 5) = monto_compra
        cant_z = cant_z + 1
        
        iva = monto_compra * 0 'Impuesto del 0% para los productos farnaceúticos (HACEMOS ESTO PARA ASIGNARLE EL VALOR CERO A LA CASILLA IVA Y NO QUEDE VACÍA)
        Sheets("factura").Cells(fila, 7) = iva
        acu_impuesto = acu_impuesto + iva
        
        monto_total = monto_compra + iva
        Sheets("factura").Cells(fila, 8) = monto_total
        acu_total = acu_total + monto_total
        
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
        acu_impuesto = acu_impuesto + iva
        
        monto_total = monto_compra + iva
        Sheets("factura").Cells(fila, 8) = monto_total
        acu_total = acu_total + monto_total
        
        End If
     
End If

End If

End Sub

Sub limpiar_celdas()

With Sheets("factura")
.Range("B5:H12").Value = " "
.Range("L5:L12").Value = " "

End With


End Sub

Sub repeticion()
fila = 5
mensaje = MsgBox("¿Hay mas datos?", vbYesNo, "Ingresar Datos")

'iniciadores contadores estadísticos

cant_f = 0
cant_q = 0
cant_h = 0
cant_z = 0

acu_impuesto = 0
acu_descuento = 0
acu_total = 0


While mensaje = vbYes
    Call entradas_datos

fila = fila + 1
mensaje = MsgBox("¿Hay mas datos?", vbYesNo, "Ingresar Datos")
Wend
Call impresiones_estadisticas

End Sub

Sub impresiones_estadisticas()

With Sheets("factura")
    .Cells(5, 12) = cant_f
    .Cells(6, 12) = cant_q
    .Cells(7, 12) = cant_h
    .Cells(8, 12) = cant_z
    .Cells(10, 12) = acu_impuesto
    .Cells(11, 12) = acu_descuento
    .Cells(12, 12) = acu_total
End With

End Sub

