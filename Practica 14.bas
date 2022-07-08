Sub AgruparTablas()
    Dim hoja As Worksheet
    Dim nuevaHoja As Worksheet
    Dim datos As Range
    Dim rangoDatos As Range
    
    'Creación de nueva hoja
    Set nuevaHoja = Application.Sheets.Add(After:=Sheets("Hoja5"))
    nuevaHoja.Name = "Datos Agrupados"
    
    'Copiado de encabezados
    Sheets("Hoja1").Range("A1:G1").Copy Destination:=nuevaHoja.Range("A1")
    
    'Recorrido de hojas
    For Each hoja In ThisWorkbook.Sheets
        If hoja.Name <> "Menú" And hoja.Name <> "Datos Agrupados" Then
        
            'Copiado de datos
            Set datos = hoja.Range("A1").CurrentRegion
            datos.Offset(1, 0).Resize(datos.Rows.Count - 1, datos.Columns.Count).Copy
    
            'Obtener posicion disponible en la nueva hoja
            Set rangoDatos = nuevaHoja.Range("A1").CurrentRegion
            
            'Pegar valores
            rangoDatos.Offset(rangoDatos.Rows.Count, 0).PasteSpecial xlPasteValues
        End If
    Next hoja
    
    Application.CutCopyMode = False
    nuevaHoja.Range("A1").CurrentRegion.Columns.AutoFit
    nuevaHoja.Range("A1").Select
    
End Sub
