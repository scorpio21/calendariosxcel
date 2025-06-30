Attribute VB_Name = "CrearListaEmpleados"
Sub CrearListaEmpleadosEnB3()
    Dim ws As Worksheet
    Set ws = Worksheets("Principal")
    ws.Range("B3").Validation.Delete ' Limpia validación previa
    ws.Range("B3").Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, _
        Formula1:="Carmelo,María,José,Ángela,Luisito"
End Sub


