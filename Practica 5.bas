Sub Referencia_Archivos()

'Libro activo
MsgBox ActiveWorkbook.Name

'Libro que contiene las macros
MsgBox ThisWorkbook.Name

'Te devuelve el nombre adjunto con la ruta de ubicación del archivo
MsgBox ThisWorkbook.FullName

'Te devuelve la ruta de ubicación del archivo
MsgBox ThisWorkbook.Path

'Cuenta el numero de archivos abiertos
MsgBox Workbooks.Count

'MsgBox Workbooks("Curso Macros en Excel.xlsm").Path
MsgBox Workbooks("ArchivoDePrueba").Path

'Colocar un valor en el libro que contiene la macro
ThisWorkbook.Sheets("Practica 5").Range("A1").Value = "Prueba 1"

'Colocar un valor en el libro activo
ActiveWorkbook.Sheets(1).Range("A1").Value = "Prueba 2"

'Abrir un archivo
Workbooks.Open ("C:\Users\GJPerez\Desktop\archivoDePrueba.xlsx")

'Cerrar archivo
Workbooks("archivoDePrueba.xlsx").Close False

'Ejercicio Copiar una tabla en un archivo nuevo y guardarlo

Hoja2.Range("A1").CurrentRegion.Copy
Workbooks.Add
ActiveSheet.Paste
Cells.Columns.AutoFit
Application.CutCopyMode = False
ActiveWorkbook.SaveAs "C:\Users\GJPerez\Desktop\archivoNuevoCreado.xlsx", xlOpenXMLStrictWorkbook
ActiveWorkbook.Close True

End Sub
