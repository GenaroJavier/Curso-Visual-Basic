'Modulo 1
Sub EjemploWith1()
    With ThisWorkbook.Sheets("Hoja1").Range("A1").CurrentRegion.Font
        .Name = "Calibri"
        .Size = 14
        .Bold = True
    End With
End Sub

Sub EjemploWith2()
Dim rango As Range
Set rango = ThisWorkbook.Sheets("Hoja1").Range("A1").CurrentRegion

    With rango.Font
        .Name = "Calibri"
        .Size = 14
        .Bold = True
    End With
End Sub

'========================================================================0

'Modulo 2
Option Explicit
Dim nombre As String
Dim tamaño As Byte
Dim negritas As Boolean

Sub aplicarFormato()
Dim rango As Range
Set rango = ThisWorkbook.Sheets("Hoja1").Range("A1").CurrentRegion

'Propiedades actuales
    With rango.Font
        nombre = .Name
        tamaño = .Size
        negritas = .Bold
    End With
    
'Nuevas propiedades del formato
    With rango.Font
        .Name = "Calibri"
        .Size = 14
        .Bold = True
    End With

End Sub

Sub resetearRango()
    Dim rango As Range
    Set rango = ThisWorkbook.Sheets("Hoja1").Range("A1").CurrentRegion
    
    With rango.Font
        .Name = nombre
        .Size = tamaño
        .Bold = negritas
    End With
End Sub
