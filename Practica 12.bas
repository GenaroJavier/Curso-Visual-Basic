Option Explicit

Sub GenerarIndice()
    Dim contador As Integer
    Dim nuevoNombre, respuesta
    Dim hoja As Worksheet
    Dim cantidadHojas As Integer
    
    ThisWorkbook.Sheets.Add before:=Sheets(1)
    Set hoja = ThisWorkbook.ActiveSheet
    
    respuesta = MsgBox("¿Deseas cambiar el nombre de la nueva hoja?", vbYesNoCancel + vbQuestion, "Mensaje importante")
    
    If respuesta = vbYes Then
        Application.Dialogs(xlDialogWorkbookName).Show
    End If
    
    With hoja.Range("A1")
        .Value = "INDICE"
        .Font.Bold = True
    End With
    
    For contador = 1 To Application.Worksheets.Count
        hoja.Range("A1").Offset(contador, 0).Value = Application.Sheets(contador).Name
    Next contador
End Sub
