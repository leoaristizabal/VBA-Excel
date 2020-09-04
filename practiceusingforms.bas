Attribute VB_Name = "Módulo1"
Option Explicit
Dim fila As Integer
Dim encontre As Boolean ' variable que permite dar valores true y false
Dim i As Integer
Dim cod As String
Dim cant As Integer
Dim existencia As Integer
Sub inicio()
' inicia abriendo el menu
Sheets("inventario").Select ' se posiciona en la hoja inventario
form_menu.Show ' abre el formulario menu
End Sub
Sub contar_filas()
' cuenta las filas llenas en el inventario
fila = 5
While Sheets("inventario").Cells(fila, 1) <> ""
    fila = fila + 1
Wend
End Sub

Sub validar_entrega()
' este sub permite agregar elmentos al inventario o modificar las existencias
' de algun producto que  ya esxiste en el inventario
' se llama desde el primer boton del form codigo

' inicializar algunas variables de trabajo
' esta variable trabaja como una bandera, la inicializo en false
'indicando que voy a buscar el codigo, por eso esta apagada, cuando lo encuentre si es asi le
' asigno true para encenderal e indicar que el producto ya existe
encontre = False
' cuento las filas llenas en el inventario
Call contar_filas
' indico cual es al primera fila de la base de datos con uan variable i
i = 5
' extraigo el codigo del producto entregado desde el formulario
cod = form_codigo.txt_cod
cant = form_codigo.spin_cant
'hago un ciclo de repeticion para verificar si el producto entregado ya existe.
With Sheets("inventario") ' para trabajar abreviadamente con la hoja inventario
    While (encontre = False) And (i < fila) ' no he encontrado el producto y no he llegado al final de los datos
        If .Cells(i, 1) = cod Then ' si encontre el codigo
            .Cells(i, 3) = .Cells(i, 3) + cant ' agrego la cantidad
            encontre = True ' indico que lo encontre
        Else ' no he encontrado el producto
            i = i + 1 ' avanzo a la siguiente fila para seguir verificando
        End If
    Wend ' termino el ciclo
    ' ahora debo proceder a verificar si el codigo fue encontrado
    If encontre = False Then ' significa que el codigo no existia en la hoja inventario
        form_codigo.Hide ' oculto el form codigo
        form_datos.Show ' muestro el form_datos
    Else
        Call limpiar_codigo ' limpio el form_codigo para que pida mas datos
    End If
End With
End Sub
Sub cargar_datos()
With Sheets("inventario")
' procedo a cargar los datos en al hoja
        .Cells(i, 1) = cod
        .Cells(i, 2) = form_datos.cmb_tipo
        .Cells(i, 3) = cant
        .Cells(i, 4) = form_datos.cmb_marca
        If .Cells(i, 4) <> "COLITA" Then ' si la marca no es colita
            If form_datos.opt_light = True Then ' indico la clase de refresco
                .Cells(i, 5) = "X"
            Else
                .Cells(i, 6) = "X"
            End If
        End If ' no necesito el else porque si es colita no hago nada
End With ' finalizo el with
End Sub
Sub limpiar_entregar()
' limpia el form_datos para dejarlo listo para usar
With form_datos
    .cmb_marca = ""
    .cmb_tipo = ""
    .opt_light = False ' para apagar el boton se coloca false no ""
    .opt_reg = False ' para apagar el boton se coloca false no ""
    .frm_clase.Visible = False ' oculto el marco con los botones de clase
End With
End Sub
Sub limpiar_codigo()
' limpio el form_codigo
With form_codigo
    .txt_cant = ""
    .txt_cod = ""
    .spin_cant = 1
End With
End Sub
Sub solicitar()
'este sub procesa una solicitud, si hay suficiente cantidad disminuye la cantidad disponible
' restando la cantidad solictada de la existente, si no hay suficiente cantidad
' emite un mensaje en el mismo form indicando que cantidad hay disponible para que
' se vuelva a hacer el pedido si asi lo desea.

' inicializar algunas variables de trabajo
' esta variable trabaja como una bandera, la inicializo en false
'indicando que voy a buscar el codigo, por eso esta apagada, cuando lo encuentre si es asi le
' asigno true para encenderla e indicar que el producto ya existe
encontre = False
' cuento las filas llenas en el inventario
Call contar_filas
' indico cual es al primera fila de la base de datos con uan variable i
i = 5
' extraigo el codigo del producto entregado desde el formulario
cod = form_solicitar.txt_cods
cant = form_solicitar.spin_canti
'hago un ciclo de repeticion para buscar el producto y ver que cantidad hay en inventario
With Sheets("inventario") ' para trabajar abreviadamente con la hoja inventario
    While (encontre = False) And (i < fila) ' no he encontrado el producto y no he llegado al final de los datos
        If .Cells(i, 1) = cod Then ' si encontre el codigo
            encontre = True ' indico que lo encontre
            existencia = .Cells(i, 3) ' veo que cantidad hay en existencia
            If existencia > 0 Then
                ' si hay existencia
                If existencia >= cant Then ' valido si es suficiente
                    .Cells(i, 3) = .Cells(i, 3) - cant ' disminuyo el inventario en al cantidad pedida
                    Call limpiar_solicitar ' limpio el formulario
                Else ' si no hay suficiente
                    ' escribo el mensaje adecuado en la etiqueta
                    form_solicitar.lbl_candis = "si usted desea pedir una cantidad igual o menor a esta por favor modifique su pedido y validelo de nuevo"
                    form_solicitar.txt_candis = existencia ' coloco el valor de la existencia en el campo texto para mostrarlo
                    form_solicitar.txt_candis.Visible = True
                    form_solicitar.frm_candis.Visible = True ' muestro el marco con el mensaje y la cantidad en existencia
                End If
            Else ' si no hay nada en existencia
                form_solicitar.lbl_candis = "ese producto no hay existencia por favor solicite otro codigo" ' cambio el mensaje en la etiqueta
                form_solicitar.frm_candis.Visible = True
                form_solicitar.txt_candis.Visible = False
            End If
        Else ' no he encontrado el producto
            i = i + 1 ' avanzo a la siguiente fila para seguir verificando
        End If
    Wend ' termino el ciclo
    ' ahora debo proceder a verificar si el codigo fue encontrado
    If encontre = False Then ' significa que el codigo no existia en la hoja inventario
        MsgBox ("el codigo indicado no existe por favor indique los datos nuevamente")
    End If
End With
End Sub
Sub limpiar_solicitar()
' limpia el form_solicitar
With form_solicitar
    .txt_canti = ""
    .spin_canti = 1
    .txt_cods = ""
    .lbl_candis = ""
    .frm_candis.Visible = False ' oculto el marco con el mensaje de cantidad disponible
End With
End Sub
Sub reporte()
' declaro las variables que voy a usar en este sub solamente
Dim cp As Integer
Dim cch As Integer
Dim ccol As Integer
Dim cl As Integer
Dim total As Integer
Dim por As Double
Dim menor As Integer
Dim codm As String
Dim tipom As String
Dim marcam As String
Dim tipo  As String

' obtener el reporte de estadisticas
' debo hacer un recorrido por toda la base de datos para calcular lo que me piden
' para ello primero cuento las filas llenas
Call contar_filas
With Sheets("inventario")
' inicializo los contadores y acumuladores que necesito
cp = 0
cch = 0
ccol = 0
cl = 0
total = 0
menor = .Cells(5, 3) ' tomo como patron el primer valor de la hoja
codm = .Cells(5, 1)
tipom = .Cells(5, 2)
marcam = .Cells(5, 4)
For i = 5 To fila - 1 ' ciclo para recorrer la hoja inventario
    existencia = .Cells(i, 3) ' OBTENGO LA CANTIDAD EN EXISTENCIA
    tipo = .Cells(i, 4) ' obtengo el valor del tipo de refresco
    Select Case tipo ' cuento de acuerdo al valor obtenido en cada fila
    Case "PEPSI"
        cp = cp + .Cells(i, 3)
    Case "CHINOTTO"
        cch = cch + .Cells(i, 3)
    Case "COLITA"
        ccol = ccol + .Cells(i, 3)
    End Select
    ' hallo el producto de menor existencia
    
    If existencia < menor Then
        menor = existencia
        codm = .Cells(i, 1)
        tipom = .Cells(i, 2)
        marcam = .Cells(i, 4)
    End If
    ' cuento cuantos refrescos light hay en el deposito
    If .Cells(i, 5) = "X" Then
        cl = cl + existencia
    End If
    ' cuento el total de unidades en existencia para calcular el porcentaje
    total = total + existencia
    ' AQUI DEBERIA AGREGAR LA ESTADISTICA ADICIONAL QUE LE PIDEN
    
Next i ' cierro el ciclo
End With
' reporto lo calculado
With Sheets("reporte")
    ' refrescos de cada tipo
    .Cells(4, 4) = cp
    .Cells(5, 4) = cch
    .Cells(6, 4) = ccol
    ' datos del menor
    .Cells(9, 4) = codm
    .Cells(10, 4) = marcam
    .Cells(11, 4) = tipom
    .Cells(12, 4) = menor
    ' porcentaje de refrescos ligth
    If total > 0 Then
        .Cells(16, 4) = (cl / total) * 100
    Else
        .Cells(16, 4) = "N/a"
    End If
    ' AQUI VA EL REPORTE DE LA ESTADISTICA ADICIONAL QUE LE PIDEN
End With
End Sub
Sub BOTON_ADICIONAL()
' AQUI DEBE IR EL CODIGO PARA EL BOTON DE ACCION ADICIONAL
' DEBE SER UNA DE LAS OPERACIONES DE REGISTRO TALES COMO ELIMINAR O
'ORDENAR YA QUE LAS DE AGREGAR Y MODIFICAR YA SE ESTAN HACIENDO AQUI.
End Sub
