Option Explicit

Sub cambiarNombreHojas()
Dim contador As Integer
Dim hoja As Worksheet
Dim titulo As String

contador = 1
titulo = InputBox("Por favor ingresa un titulo en común para las hojas", "Cambiar nombre a las hojas")

For Each hoja In Application.Worksheets
    hoja.Name = titulo & " " & contador
    contador = contador + 1
Next hoja

End Sub

Sub elegirRango()
Dim rango As Range

Set rango = Application.InputBox("Selecciona el rango que desees", "Cambiar formato", Type:=8)

rango.Style = "comma"

End Sub


Sub obtenerPorcentaje()
Dim valor As Integer
Dim porcentaje As Integer

valor = Range("G5").Value
porcentaje = Range("G6").Value

If IsNumeric(valor) = False Or IsNumeric(porcentaje) = False Then
    MsgBox "Por favor introduce el valor o porcentaje, en la celda correspondiente"
    Else
    Range("G7").Value = (valor * (porcentaje / 100))
End If
End Sub
