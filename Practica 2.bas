Sub Propiedades()

'=====> Value
Range("A1").Value = "Hola"
MsgBox Range("A1").Value

'=====> Address
MsgBox ActiveCell.Address
MsgBox Selection.Address

'=====> Columns
MsgBox Range("A5:F12").Columns.Count

'=====> Rows
MsgBox Range("A5:F12").Rows.Count

'=====> CurrentRegion
Range("C2").CurrentRegion.Select
MsgBox Range("A1").CurrentRegion.Address

'=====> EntireRow
Range("A4").EntireRow.Select

'=====> EntireColumn
Range("D3").EntireColumn.Select

'=====>Font
Range("A1").CurrentRegion.Font.Name = "Calibri"
Range("A1").CurrentRegion.Font.Size = 12

'=====>Formula
Range("E2:E17").Formula = "=SUM(12*34)"

'=====>Formula Local
Range("F2:F17").Formula2Local = "=SUMA(56*3)"

'=====>HasFormula
MsgBox ActiveCell.HasFormula

'=====>Interior
Range("A1").Interior.Color = vbGreen
Range("A2").Interior.Color = RGB(200, 123, 73)
Range("A3").Interior.ColorIndex = 29

'=====>Offset
Range("A1").Offset(5, 1).Interior.Color = RGB(21, 164, 232)

'=====>Resize
Range("A1:F20").Resize(5, 4).Activate

End Sub

Sub ColorearTabla()
    Dim Rango As Range
    Set Rango = Range("A1").CurrentRegion
    
    Rango.Resize(1, Rango.Columns.Count).Interior.Color = RGB(21, 67, 96)
    Rango.Resize(1, Rango.Columns.Count).Font.Color = RGB(251, 252, 252)
    Rango.Offset(1, 0).Resize(Rango.Rows.Count - 1, 1).Interior.Color = RGB(31, 97, 141)
    Rango.Offset(1, 0).Resize(Rango.Rows.Count - 1, 1).Font.Color = RGB(251, 252, 252)
    Rango.Offset(1, 1).Resize(Rango.Rows.Count - 1, 4).Interior.Color = RGB(169, 204, 227)
    Rango.Offset(1, 5).Resize(Rango.Rows.Count - 1, 1).Interior.Color = RGB(52, 152, 219)
    Rango.Offset(1, 5).Resize(Rango.Rows.Count - 1, 1).Font.Color = RGB(251, 252, 252)
    
    Cells.EntireColumn.AutoFit
End Sub
