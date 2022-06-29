Sub Referencias_Hojas()

'Formas de seleccionar hojas
    
    'Con numero de la hoja
    Application.Worksheets(3).Select
    Worksheets(2).Select
    Sheets(1).Select

    'Con el nombre de la hoja
    Sheets("Practica 1").Select
    Application.Sheets("Practica 3").Select
    
    'Obtener el codigo de la hoja ej ("Hoja1)
    MsgBox Sheets(1).CodeName
    
    'Obtener el nombre de la hoja
    MsgBox ActiveSheet.Name
    
'Colocar un valor en una hoja de acuerdo a un rango
Hoja3.Range("A1").Value = "Mensaje de prueba"

'Colocar un valor en una celda activa
ActiveCell.Value = "Mensaje de prueba 2"


'ejercicio mover a una hoja oculta
Hoja3.Visible = False
Hoja2.Range("A1").CurrentRegion.Copy Hoja3.Range("A1")
Hoja3.Visible = True
Hoja3.Range("A1").CurrentRegion.Columns.AutoFit

'Cuenta cuantas hojas tiene el libro (sin importar si estan ocultas)
MsgBox Worksheets.Count

'Agrega una nueva hoja al libro
Worksheets.Add
            
End Sub
