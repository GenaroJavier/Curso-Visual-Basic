'===============================================================================================
'Modulo 1

Option Explicit

'Cuando colocamos la palabra private al inicio de una macro, estamos especificando que esa macro solo queremos que sea llamada
'desde el modulo donde se encuentra la macro
Private Sub macro1()
    MsgBox "Mensaje 1"
End Sub


Sub macro4(ByVal nombre As String)
    MsgBox "Mensaje 4" & valor1
End Sub

'Por otro lado, si nosotros ponemos la palabra reservada, (public) esa macro puede ser llamada desde otro modulos,
'incluso no es necesario ponerle public
Public Sub macro2()
    'Formas de mandar a llamar una macro desde otra macro
    
    'Colocar el nombre de la macro
    macro1
    
    'Colocar la palabra reservada Call (Nombre macro)
    Call macro1
    
    'Este forma nos permite asignarle parametros, importante es solo para mandarle argumentos
    Application.Run macro4
    
    MsgBox "Mensaje 2"
End Sub


'===================================================================================================
'Modulo 2

Sub macro3()
    macro2
    MsgBox "Mensaje 3"
End Sub
