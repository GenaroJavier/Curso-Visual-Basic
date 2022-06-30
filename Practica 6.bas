Sub buscarRango()

Dim RangoBuscar As Range
Dim RangoEncontrado As String

Set RangoBuscar = Range("A2:A31")
RangoEncontrado = RangoBuscar.Find(What:="Marzo", LookAt:=xlPart, MatchCase:=False).Address
Sheets("Practica 5").Range(RangoEncontrado).Interior.Color = RGB(40, 55, 71)
Sheets("Practica 5").Range(RangoEncontrado).Font.Color = RGB(253, 254, 254)

End Sub
Sub buscarRangoPractica()

Dim RangoBuscar As Range
Dim RangoTotalBuscar As Range
Dim totalDeFilas As Integer

Set RangoBuscar = Sheets("Practica 5").Range("A1").CurrentRegion
Set RangoTotalBuscar = RangoBuscar.Offset(1, 0).Resize(RangoBuscar.Rows.Count - 1, 1)
totalDeFilas = RangoTotalBuscar.Rows.Count

For J = 2 To totalDeFilas
    If Range("A" & J).Value = "Agosto" Then
        Sheets("Practica 5").Range("A" & J).Interior.Color = RGB(40, 55, 71)
        Sheets("Practica 5").Range("A" & J).Font.Color = RGB(253, 254, 254)
    End If
Next

End Sub
Sub SimpleSET()
'Declaramos una nueva variable de tipo worksheet
Dim nuevaHoja As Worksheet

'Asignamos la variable y creamos una nueva hoja haciendo uso de sus metodos
Set nuevaHoja = Application.Sheets.Add

'Asignamos nombre a la nueva hoja a crear
nuevaHoja.Name = "Hoja prueba"
End Sub

Sub variables_objeto()

Dim hoja As Worksheet
Set hoja = Application.ActiveSheet
MsgBox hoja.Name

End Sub

        'Crear lista de validacion
Sub Validacion()
Dim hojaLista As Worksheet
Dim rangoLista As Range
Dim miRango As Range

Set hojaLista = ThisWorkbook.Worksheets("Practica 5")
Set rangoLista = hojaLista.Range("A2:A13")
Set miRango = hojaLista.Range("D2")

miRango.Validation.Delete
miRango.Validation.Add xlValidateList, Formula1:="='" & hojaLista.Name & "'!" & rangoLista.Address
End Sub
