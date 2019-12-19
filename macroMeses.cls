VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'Macro usado para conocer las cedulas repetidas en meses anteriores
'los comentarios con ### funcionan como parametros.

Sub Cedulas()

mesActual = "G" '### mes analizar
inicio = 1 '### mes inicio
fin = 6 '### mes final
colorMes = 8 'Color de tabla de colores


numFilas = ActiveSheet.UsedRange.Rows.Count 'numero de filas del documento

For i = 2 To numFilas 'fila en la que se incia el recorrido

    mesIter = inicio 'mes actual sobre el cual se esta iterando

    cedulaActual = Cells(i, mesActual).Value 'cedula del mes actual que se compara con meses anteriores

     'En caso de que la columna no tenga mas cedulas se da por terminadola busqueda
    If IsEmpty(cedulaActual) Then
        Exit For
    End If

    For j = 2 To numFilas 'filas de meses anteriores a recorrer

        cComparar = Cells(j, mesIter).Value ' cedula de mes anterior para comparar

        'si la cedula es la misma entra aqui
        If cComparar = cedulaActual Then
            Cells(i, mesActual).Interior.ColorIndex = colorMes 'colorea cedula del mes actual
            Cells(j, mesIter).Interior.ColorIndex = colorMes ''colorea cedula del mes pasado
            Cells(19, "K") = Cells(19, "K") + 1  ' ### casilla donde se sumara las cedulas viejas
            Exit For
        End If

        'si al terminar de recorrer la columna del mes no encuentra la cedula del mes actual y el mes es diferente del final
        'se reinicia la cuenta de la fila y se continua con la siguiente columna que corresponde al siguiente mes
        If j = numFilas And mesIter <> fin Then
        j = 1
        mesIter = mesIter + 1
        End If

    Next j

    'Si al terminar de recorrer todos los meses y el recorrido de filas es mayor a su tama�o,
    'entonces la cedula no se encuentra en ningunos de los meses y se sumarar en la celda seleccionada.
    If mesIter = fin And j >= numFilas Then
        Cells(19, "L") = Cells(19, "L") + 1 '### casilla donde se sumara las cedulas nuevas
    End If

Next i

End Sub
