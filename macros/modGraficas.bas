Attribute VB_Name = "modGraficas"
' ================== modGraficas.bas ==================
' Gráfica de turnos por empleado

Public Sub GraficaTurnos()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets("Turnos")
    If ws Is Nothing Then MsgBox "No existe la hoja 'Turnos'.": Exit Sub
    On Error GoTo 0

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    Dim wsGraf As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("GraficaTurnos").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsGraf = Worksheets.Add(After:=ws)
    wsGraf.Name = "GraficaTurnos"

    wsGraf.Cells(1, 1).Value = "Empleado"
    wsGraf.Cells(1, 2).Value = "Turnos"

    Dim empleados As Variant
    empleados = Array("Carmelo", "María", "José", "Ángela", "Luisito")
    Dim i As Integer, rowGraf As Integer
    rowGraf = 2

    For i = 3 To 7
        wsGraf.Cells(rowGraf + i - 3, 1).Value = empleados(i - 3)
        wsGraf.Cells(rowGraf + i - 3, 2).Formula = "=COUNTIF('" & ws.Name & "'!" & ws.Cells(2, i).Address(False, False) & ":" & ws.Cells(lastRow, i).Address(False, False) & ",""<>-"")"
    Next i

    wsGraf.Columns("A:B").AutoFit

    Dim ch As ChartObject
    Set ch = wsGraf.ChartObjects.Add(Left:=100, Top:=50, Width:=400, Height:=250)
    ch.Chart.ChartType = xlColumnClustered
    ch.Chart.SetSourceData Source:=wsGraf.Range("A1:B6")
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "Turnos por Empleado"
    ch.Chart.Axes(xlCategory).HasTitle = True
    ch.Chart.Axes(xlCategory).AxisTitle.Text = "Empleado"
    ch.Chart.Axes(xlValue).HasTitle = True
    ch.Chart.Axes(xlValue).AxisTitle.Text = "Total Turnos"
End Sub
