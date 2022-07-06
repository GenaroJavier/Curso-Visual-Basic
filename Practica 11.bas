Option Explicit

Sub convertirTexto()
    Dim opcion As Variant
    Dim texto As String
    Dim rango As Range
    
    Set rango = ThisWorkbook.Sheets("Practica 2").Range("M4")
    
    texto = "Elige una opción:" & vbNewLine & _
            vbNewLine & "1. MAYUSCULAS" & _
            vbNewLine & "2. minusculas"
            
    opcion = InputBox(texto, "Convertir texto", 1)
    
    Select Case opcion
        Case 1
            rango.Value = VBA.UCase(rango.Value)
        Case 2
            rango.Value = VBA.LCase(rango.Value)
        Case Else
            MsgBox "Debes seleccionar una opción, valida", vbExclamation, "Información del sistema"
    End Select
End Sub
Sub ejemploSwitch()
    Select Case Weekday(Now)
        Case 1, 7
            MsgBox "Fin de semana"
        Case Else
            MsgBox "No es fin de semana"
    End Select
End Sub

Sub obtener_porcentaje_descuento()
Dim cantidad As Variant

cantidad = Sheets("Practica 2").Range("I12")

If cantidad = "" Or IsNumeric(cantidad) = False Then
MsgBox "Por favor verifica, los datos introducidos"
 Else
    Select Case cantidad
        Case Is < 10
            Range("K12").Value = 0
        Case Is <= 19
            Range("K12").Value = 0.1
        Case Is >= 20
            Range("K12").Value = 0.2
    End Select
End If
End Sub

Sub obtener_porcentaje_descuento2()
Dim cantidad As Variant

cantidad = Sheets("Practica 2").Range("I12")

If cantidad = "" Or IsNumeric(cantidad) = False Then
MsgBox "Por favor verifica, los datos introducidos"
 Else
    Select Case cantidad
        Case 1 To 9
            Range("K12").Value = 0
        Case 10 To 19
            Range("K12").Value = 0.1
        Case Else
            Range("K12").Value = 0.2
    End Select
End If
End Sub
