Attribute VB_Name = "modResumen"
' ================== modResumen.bas ==================
' Resumen de ganancias y gráfica semanal en la misma hoja

Public Sub GenerarTurnosYResumenGanancias()
    Dim wsTurnos As Worksheet
    Set wsTurnos = ThisWorkbook.Worksheets("Turnos")

    Dim wsResumen As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("ResumenGanancias").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsResumen = Worksheets.Add(After:=wsTurnos)
    wsResumen.Name = "ResumenGanancias"

    wsResumen.Cells(1, 1).Value = "Semana (Año-Semana)"
    wsResumen.Cells(1, 2).Value = "Carmelo + María (€)"
    wsResumen.Cells(1, 3).Value = "José (€)"
    wsResumen.Cells(1, 4).Value = "Ángela + Luisito (€)"
    wsResumen.Cells(1, 5).Value = "Total semanal (€)"
    wsResumen.Rows(1).Font.Bold = True

    Dim dicSemanas As Object: Set dicSemanas = CreateObject("Scripting.Dictionary")
    Dim weekRow As Long, lastRow As Long
    Dim row As Long
    Dim fecha As Date, semanaStr As String
    Dim carmelo As Double, maria As Double, jose As Double, angela As Double, luisito As Double
    Dim turnoCarmelo As String, turnoMaria As String, turnoJose As String, turnoAngela As String, turnoLuisito As String

    lastRow = wsTurnos.Cells(wsTurnos.Rows.Count, 1).End(xlUp).row

    For row = 2 To lastRow
        fecha = wsTurnos.Cells(row, 1).Value
        If Not (fecha >= DateSerial(2025, 9, 1) And fecha <= DateSerial(2025, 9, 15)) Then
            semanaStr = Year(fecha) & "-S" & Format(Application.WorksheetFunction.WeekNum(fecha, 2), "00")
            If Not dicSemanas.Exists(semanaStr) Then
                dicSemanas.Add semanaStr, dicSemanas.Count + 2
                wsResumen.Cells(dicSemanas(semanaStr), 1).Value = semanaStr
            End If

            turnoCarmelo = wsTurnos.Cells(row, 3).Value
            turnoMaria = wsTurnos.Cells(row, 4).Value
            turnoJose = wsTurnos.Cells(row, 5).Value
            turnoAngela = wsTurnos.Cells(row, 6).Value
            turnoLuisito = wsTurnos.Cells(row, 7).Value

            If turnoCarmelo = "08:00–00:00" Or turnoCarmelo = "09:00–00:00" Then
                carmelo = carmelo + 100
            ElseIf turnoCarmelo = "17:00–00:00" Then
                carmelo = carmelo + 50
            End If

            If turnoMaria = "08:00–00:00" Or turnoMaria = "09:00–00:00" Then
                maria = maria + 100
            ElseIf turnoMaria = "17:00–00:00" Then
                maria = maria + 50
            End If

            If turnoJose = "08:00–00:00" Or turnoJose = "09:00–00:00" Then
                jose = jose + 100
            End If

            If turnoAngela = "08:00–00:00" Or turnoAngela = "09:00–00:00" Then
                angela = angela + 100
            ElseIf turnoAngela = "08:00–17:00" Then
                angela = angela + 50
            End If

            If turnoLuisito = "08:00–00:00" Or turnoLuisito = "09:00–00:00" Then
                luisito = luisito + 100
            ElseIf turnoLuisito = "08:00–17:00" Then
                luisito = luisito + 50
            End If

            If row = lastRow Or _
               (row < lastRow And (Year(wsTurnos.Cells(row + 1, 1).Value) & "-S" & Format(Application.WorksheetFunction.WeekNum(wsTurnos.Cells(row + 1, 1).Value, 2), "00")) <> semanaStr) Then
                weekRow = dicSemanas(semanaStr)
                wsResumen.Cells(weekRow, 2).Value = carmelo + maria
                wsResumen.Cells(weekRow, 3).Value = jose
                wsResumen.Cells(weekRow, 4).Value = angela + luisito
                wsResumen.Cells(weekRow, 5).Value = wsResumen.Cells(weekRow, 2).Value + wsResumen.Cells(weekRow, 3).Value + wsResumen.Cells(weekRow, 4).Value
                carmelo = 0: maria = 0: jose = 0: angela = 0: luisito = 0
            End If
        End If
    Next row

    wsResumen.Columns("A:E").AutoFit
    MsgBox "Resumen profesional de ganancias semanales generado en la hoja 'ResumenGanancias'."
End Sub

Public Sub GraficaGanancias()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets("ResumenGanancias")
    If ws Is Nothing Then MsgBox "No existe la hoja 'ResumenGanancias'.": Exit Sub
    On Error GoTo 0

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    ' Eliminar cualquier gráfico anterior en la hoja
    Dim obj As Object
    For Each obj In ws.ChartObjects
        obj.Delete
    Next obj

    ' Insertar el gráfico en la misma hoja
    Dim ch As ChartObject
    Set ch = ws.ChartObjects.Add(Left:=ws.Range("G2").Left, Top:=ws.Range("G2").Top, Width:=600, Height:=300)
    ch.Chart.ChartType = xlLineMarkers
    ch.Chart.SetSourceData Source:=ws.Range("A1:E" & lastRow)
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "Ganancias Semanales"
    ch.Chart.Axes(xlCategory).HasTitle = True
    ch.Chart.Axes(xlCategory).AxisTitle.Text = "Semana"
    ch.Chart.Axes(xlValue).HasTitle = True
    ch.Chart.Axes(xlValue).AxisTitle.Text = "€"
End Sub

