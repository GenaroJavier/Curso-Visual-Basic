Sub AgregarComentarios()
Range("A1").AddComment ("Comentario de prueba")
End Sub
Sub Autofiltro()
Range("A1").AutoFilter
End Sub

Sub AjustarAnchoColumna()

'Ajusta las celdas del rango
[A1].CurrentRegion.Columns.AutoFit

'Ajusta las celdas de toda la hoja
Cells.Columns.AutoFit
End Sub

Sub LimpiarCeldas()
'Range("A1").CurrentRegion.ClearContents
'Range("A1").CurrentRegion.ClearFormats
Range("A1").CurrentRegion.Clear
End Sub

Sub Copiar()

'=========Con esta forma puedes copiar un rango a otro, pero involucra mas pasos
'Range("A1:B10").Copy
'Range("D1").Select
'ActiveSheet.Paste
'Application.CutCopyMode = False
'Range("D1").CurrentRegion.EntireColumn.AutoFit


'=======De esta forma es mas facil y eficiente
Range("A1:B10").Copy Destination:=Range("D1")
Range("D1").CurrentRegion.EntireColumn.AutoFit
End Sub

Sub PegadoEspecial()
Range("A1:B10").Copy
Range("D1").PasteSpecial xlPasteValues
'Range("D1").PasteSpecial xlPasteFormats
Application.CutCopyMode = False
Range("D1").CurrentRegion.Columns.AutoFit
End Sub

Sub Ordenar()
'Ordenar Nombre (Tabla 1)
Range("A1:B10").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes

'Ordenar las dos columnas (Tabla 2)
'Range("A13:B24").Sort Key1:=Range("A13"), Order1:=xlAscending, Key2:=Range("B13"), Order2:=xlDescending, Header:=xlYes
End Sub
