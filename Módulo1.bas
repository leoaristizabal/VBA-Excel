Attribute VB_Name = "Módulo1"
Option Explicit
Dim tipo_cli As String
Dim cedula As String
Dim nombre As String
Dim vehiculo As String
Dim zona_reco As String
Dim zona_ent As String
Dim cant_ent As Integer
Dim cont_ent As Integer
Dim fila As Integer
Dim monto As Integer
Dim entregas As Integer

Sub ingres_for()


For cont_ent = 4 To 8 'for iniciando en 4 por las filas en la hoja diario, hasta 8 porque se cuentan los extremos completando cinco ciclos
    
    'ingreso de datos
    
    tipo_cli = InputBox("INGRESA (N) PARA PERSONA NATURAL Y (J) JURÍDICA", "TIPO CLIENTE")
    cedula = InputBox("INGRESE CEDULO O RIF DEL CLIENTE", "CEDULA O RIF")
    nombre = InputBox("INGRESE EL NOMBRE DEL CLIENTE", "NOMBRE CLIENTE")
    vehiculo = InputBox("INGRESA EL TIPO DE VEHICULO (CARRO, MOTO)", "TIPO DE VEHICULO")
    zona_reco = InputBox("INGRESA LA ZONA DE RECOLECTA SEGUN TABLA", "ZONA DE RECOLECTA")
    zona_ent = InputBox("INGRESA LA ZONA DE ENTREGA", "ZONA ENTREGA")
    entregas = InputBox("INGRESE EL NUMERO DE ENTREGAS A REALIZAR", "NUMERO DE ENTREGAS")
    
    If vehiculo = "MOTO" Then
        If zona_reco = zona_ent Then
            monto = 5
        Else
            monto = 8
        End If
    Else
        If zona_reco = zona_ent Then
            monto = 10
        Else
            monto = 12
        End If
    End If
    
    If entregas > 1 Then
        monto = monto + entregas * 2
    
    End If
    
    With Sheets("diario") 'impresiones
    
        .Cells(cont_ent, 2) = tipo_cli
        .Cells(cont_ent, 3) = cedula
        .Cells(cont_ent, 4) = nombre
        .Cells(cont_ent, 5) = zona_reco
        .Cells(cont_ent, 6) = zona_ent
        .Cells(cont_ent, 8) = monto
        .Cells(cont_ent, 7) = entregas
    End With
     
Next cont_ent 'aumenta el contador en 1 y sube


End Sub

Sub limp()

Sheets("diario").Range("B4:H10").Value = ""
End Sub
