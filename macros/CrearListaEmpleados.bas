Attribute VB_Name = "CrearListaEmpleados"
Sub CrearListaEmpleadosEnB3()
    Dim ws As Worksheet
    Set ws = Worksheets("Principal")
    ws.Range("B3").Validation.Delete ' Limpia validaci�n previa
    ws.Range("B3").Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, _
        Formula1:="Carmelo,Mar�a,Jos�,�ngela,Luisito"
End Sub


