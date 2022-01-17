Option Explicit

Sub CambiarNombre()
Dim NombreNuevo As String
Dim NombreAnterior As String
Dim Celda As Range

On Error Resume Next

For Each Celda In Selection
    NombreAnterior = Celda.Value
    NombreNuevo = Celda.Offset(0, 3).Value
    
    Name NombreAnterior As NombreNuevo
    
Next Celda

On Error GoTo 0

End Sub
