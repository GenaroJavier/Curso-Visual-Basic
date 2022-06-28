Sub SeleccionarRangos()
    'Limpiar el contenido y formato de la celda
    Cells.Clear
    
    [A1] = "valor 1"
    Range("A2").Value = "valor 2"
    Range("A3:C3").Value = "valor 3"
    Range("A4, C4, E4").Value = "valor 4"
    Range("A5:C5,E5:G5").Value = "valor 5"
    Range("A6", "D6").Value = "valor 6"
    
    'Cambiar el color de fondo de una celda o rango de celdas
    Range("A5, C5, E5").Interior.Color = RGB(47, 239, 52)
    
    'Uso de la funcion Cell
    Cells(7, 2).Value = "valor 7"
    Range("C" & 8 & ":" & "E" & 8).Value = "valor 8"
    Range(Cells(9, 3), Cells(9, 6)).Value = "valor 9"
    Range("A10:E11").Cells(2, 3).Value = "valor 10"
    
    'Offset
    Range("A1").Offset(11, 3).Value = "valor 11"
    
    'Cambiar ancho de las celdas
    Range("14:15").EntireRow.RowHeight = 30
    Rows("14:15").RowHeight = 40
    
    'Cambiar largo de las celdas
    Range("D:D").EntireColumn.ColumnWidth = 5
    Columns(4).ColumnWidth = 20
    
    Range("CeldasPrueba").Value = "valor 12"
    
    'Ajusta el tamaño de las columnas
    Cells.EntireColumn.AutoFit
End Sub
