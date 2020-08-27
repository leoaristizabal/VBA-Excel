Attribute VB_Name = "M�dulo2"
Option Explicit
Dim fila As Integer
Dim monto As Integer
'contadores
Dim cont_carro As Integer
Dim cont_moto As Integer
Dim cont_nat As Integer
Dim cont_jur As Integer

Sub programa_f()

fila = 4
While Sheets("diario").Cells(fila, 2) <> "" 'buscar primera fila vacia
    fila = fila + 1
Wend

With Sheets("diario")
'Entradas por cuadros de texto
    .Cells(fila, 3) = form_agregar.txt_cedula
    .Cells(fila, 4) = form_agregar.txt_nombre
    .Cells(fila, 8) = form_agregar.txt_entregas
    
    'Boton de opciones Tipo Cliente
    
    If form_agregar.opt_natural = True Then
        .Cells(fila, 2) = "Natural"
    Else
        .Cells(fila, 2) = "Juridico"
    End If
    
    'Boton de Opciones Vehiculo
    
    If form_agregar.opt_carro = True Then
        .Cells(fila, 5) = "Carro"
    Else
        .Cells(fila, 5) = "Moto"
    End If
    
    'Boton de opciones Recolecta
    
    If form_agregar.opt_recobaruta = True Then
        .Cells(fila, 6) = "Mun Baruta"
    End If

    If form_agregar.opt_recohatillo Then 'no es necesario colocar le true
        .Cells(fila, 6) = "Mun Hatillo"
    End If

    If form_agregar.opt_recosucre Then
        .Cells(fila, 6) = "Mun Sucre"
    End If
    
    If form_agregar.opt_recochacao Then
        .Cells(fila, 6) = "Mun Chacao"
    End If
    
    If form_agregar.opt_recolibertador Then
        .Cells(fila, 6) = "Mun Libertador"
    End If
    
    'Boton de opciones Entrega
     If form_agregar.opt_entbaruta = True Then
        .Cells(fila, 7) = "Mun Baruta"
    End If

    If form_agregar.opt_enthatillo Then 'no es necesario colocar le true
        .Cells(fila, 7) = "Mun Hatillo"
    End If

    If form_agregar.opt_entsucre Then
        .Cells(fila, 7) = "Mun Sucre"
    End If
    
    If form_agregar.opt_entchacao Then
        .Cells(fila, 7) = "Mun Chacao"
    End If
    
    If form_agregar.opt_entlibertador Then
        .Cells(fila, 7) = "Mun Libertador"
    End If
    
    'Condiciones sobre el Monto
If .Cells(fila, 5) = "Moto" Then

    If .Cells(fila, 6) = .Cells(fila, 7) Then
        monto = 5
    Else
        monto = 8
    End If
Else
       
    If .Cells(fila, 6) = .Cells(fila, 7) Then
        monto = 10
    Else
        monto = 12
    End If
End If
    
    If .Cells(fila, 8) > 1 Then
        monto = monto + .Cells(fila, 8) * 2
    End If
    
    .Cells(fila, 9) = monto
    

End With

End Sub

Sub estad_f()

'Estadisticas
fila = 4
cont_nat = 0
cont_jur = 0
cont_carro = 0
cont_moto = 0

While Sheets("diario").Cells(fila, 2) <> "" 'buscar primera fila vacia

    With Sheets("diario")

    If .Cells(fila, 2) = "Natural" Then
        cont_nat = cont_nat + 1
    Else
        cont_jur = cont_jur + 1
    End If

    If .Cells(fila, 5) = "Moto" Then
        cont_moto = cont_moto + 1
    Else
        cont_carro = cont_carro + 1
    End If

fila = fila + 1
    End With
Wend


'impresiones estadisticas
With Sheets("Estadisticas")
    .Cells(3, 3) = cont_nat / (cont_nat + cont_jur)
    .Cells(4, 3) = cont_jur / (cont_nat + cont_jur)
    .Cells(5, 3) = Sheets("diario").Cells(fila, 8)
    .Cells(6, 3) = cont_moto
    .Cells(7, 3) = cont_carro

End With
End Sub
Sub limpiar_f()
Sheets("diario").Range("B4:I49").Value = "" ' base de datos del form
Sheets("Estadisticas").Range("C3:C7").Value = "" ' limpiar estadisticas

End Sub

