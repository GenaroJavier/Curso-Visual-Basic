Option Explicit

Sub usando_foreach()
    
    Dim hoja As Worksheet
    
    For Each hoja In Application.Worksheets
        MsgBox hoja.Name
    Next hoja
    
End Sub

Sub recorrerFilas()
    Dim rango As range
    Dim rango_recorrer As range
    Set rango = ThisWorkbook.Sheets("Hoja1").range("A1:D11")
    
    For Each rango_recorrer In rango
        Debug.Print rango_recorrer.Value
    Next rango_recorrer
End Sub

Sub validarCeldasNumericas()
    Dim rango As range
    Dim rango_recorrer As range
    Set rango = ThisWorkbook.Sheets("Hoja1").range("A1").CurrentRegion
    
    For Each rango_recorrer In rango
        If Not IsNumeric(rango_recorrer) Then
            rango_recorrer.Font.Bold = True
        End If
    Next rango_recorrer
End Sub

Sub MayorADiez()
    Dim rango As range
    Dim rango_recorrer As range
    Set rango = ThisWorkbook.Sheets("Hoja1").range("A1").CurrentRegion
    
        For Each rango_recorrer In rango
            If IsNumeric(rango_recorrer) Then
                If rango_recorrer.Value > 10 Then
                    rango_recorrer.Interior.Color = RGB(72, 201, 176)
                Else
                    rango_recorrer.Interior.Color = RGB(231, 76, 60)
                End If
            End If
        Next rango_recorrer
End Sub

Sub NumHojasOcultas()
Dim contador As Integer
Dim hoja As Worksheet

contador = 0

For Each hoja In Application.Worksheets
    If hoja.Visible = False Then
        contador = contador + 1
    End If
Next hoja

Debug.Print "El numero de hojas ocultas es: " & contador

End Sub
