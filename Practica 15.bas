'================================== Macro a utilizar ========================================
Sub convertirMayusculas()
    Dim celda As Range
    
    For Each celda In Selection
        celda.Value = UCase(celda.Value)
    Next celda
End Sub


'=========================== Macro que ejecuta la otra macro de otro archivo =======================
Sub consumirMacroDesdeOtroArchivo()
Dim nombre As String

    'Abrimos el archivo que contiene la macro
    Workbooks.Open "C:\Users\GJPerez\Desktop\Respaldo\Lista de empleados.xlsm"
    
    'Le asignamos el nombre del archivo a la variable nombre
    nombre = ActiveWorkbook.Name
    
    'Nos volvemos a colocar en el archivo donde queremos ejecutar la macro
    Application.Workbooks("archivoNuevoCreado.xlsm").Activate
    
    'Ejecutamos la macro, colocando el nombre del archivo contenedor y el nombre de la macro
    Application.Run "'Lista de empleados.xlsm'!convertirMayusculas"
    
    'Cerramos el archivo de la macro, sin guardar cambios
    Application.Workbooks(nombre).Close False
    
End Sub
