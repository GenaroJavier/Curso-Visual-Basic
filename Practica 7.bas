Sub practicaIf()
Dim Numero1 As Integer
Dim PosicionNum1 As Range
    Set PosicionNum1 = ThisWorkbook.Sheets("Practica 5").Range("F3")
    Numero1 = PosicionNum1.Value
    
    If Numero1 >= 10 Then
        PosicionNum1.Interior.Color = RGB(46, 204, 113)
    Else
        PosicionNum1.Interior.Color = RGB(231, 76, 60)
    End If
End Sub

Sub Descuentos()
    Dim ubi_cantidad As Range
    Dim cantidad As Integer
    Dim ubi_descuento As Range
    
    Set ubi_cantidad = ThisWorkbook.Sheets("Practica 5").Range("A23")
    Set ubi_descuento = ThisWorkbook.Sheets("Practica 5").Range("C23")
    
    cantidad = ubi_cantidad.Value
    
    If cantidad < 10 Then
        ubi_descuento.Value = 0
        ElseIf cantidad >= 10 And cantidad < 20 Then
            ubi_descuento.Value = 0.1
            ElseIf cantidad >= 20 Then
                ubi_descuento.Value = 0.2
    End If

End Sub
