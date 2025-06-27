' ========== CALENDARIO 2025 ================
Sub Calendario2025_Cadiz_Imagen()
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("Calendario2025").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Calendario2025"

    ws.Rows("1").RowHeight = 36
    ws.Rows("2").RowHeight = 6
    Dim i As Integer
    For i = 3 To 10: ws.Rows(i).RowHeight = 18: Next i
    ws.Rows("11").RowHeight = 6
    For i = 12 To 19: ws.Rows(i).RowHeight = 18: Next i
    ws.Rows("20").RowHeight = 6
    For i = 21 To 28: ws.Rows(i).RowHeight = 18: Next i
    ws.Rows("29").RowHeight = 6
    For i = 30 To 37: ws.Rows(i).RowHeight = 18: Next i
    ws.Rows("38").RowHeight = 6

    Dim col As Integer
    For col = 1 To 7: ws.Columns(col).ColumnWidth = 3.55: Next col
    ws.Columns(8).ColumnWidth = 1.18
    For col = 9 To 15: ws.Columns(col).ColumnWidth = 3.55: Next col
    ws.Columns(16).ColumnWidth = 1.18
    For col = 17 To 23: ws.Columns(col).ColumnWidth = 3.55: Next col

    ws.Range("A1:G1").Merge
    With ws.Range("A1:G1")
        .Value = "  2   0   2   5"
        .Font.Size = 32
        .Font.Bold = True
        .Font.Name = "Arial Black"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(220, 220, 220)
        .Font.Color = RGB(0, 128, 255)
    End With

    ws.Range("A3:G3").Merge: ws.Range("A3").Value = "ENERO"
    ws.Range("I3:O3").Merge: ws.Range("I3").Value = "FEBRERO"
    ws.Range("Q3:W3").Merge: ws.Range("Q3").Value = "MARZO"
    ws.Range("A12:G12").Merge: ws.Range("A12").Value = "ABRIL"
    ws.Range("I12:O12").Merge: ws.Range("I12").Value = "MAYO"
    ws.Range("Q12:W12").Merge: ws.Range("Q12").Value = "JUNIO"
    ws.Range("A21:G21").Merge: ws.Range("A21").Value = "JULIO"
    ws.Range("I21:O21").Merge: ws.Range("I21").Value = "AGOSTO"
    ws.Range("Q21:W21").Merge: ws.Range("Q21").Value = "SEPTIEMBRE"
    ws.Range("A30:G30").Merge: ws.Range("A30").Value = "OCTUBRE"
    ws.Range("I30:O30").Merge: ws.Range("I30").Value = "NOVIEMBRE"
    ws.Range("Q30:W30").Merge: ws.Range("Q30").Value = "DICIEMBRE"

    Dim mesesRangos As Variant
    mesesRangos = Array("A3:G3", "I3:O3", "Q3:W3", "A12:G12", "I12:O12", "Q12:W12", _
                        "A21:G21", "I21:O21", "Q21:W21", "A30:G30", "I30:O30", "Q30:W30")
    Dim r As Variant
    For Each r In mesesRangos
        With ws.Range(r)
            .Font.Bold = True
            .Font.Size = 14
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.ColorIndex = xlNone
        End With
    Next r

    Dim dias As Variant: dias = Array("L", "M", "M", "J", "V", "S", "D")
    Dim filasDias As Variant: filasDias = Array(4, 13, 22, 31)
    Dim colMeses As Variant: colMeses = Array(1, 9, 17)
    Dim iFila As Integer, iCol As Integer, iDia As Integer
    For iFila = 0 To 3
        For iCol = 0 To 2
            For iDia = 0 To 6
                ws.Cells(filasDias(iFila), colMeses(iCol) + iDia).Value = dias(iDia)
                ws.Cells(filasDias(iFila), colMeses(iCol) + iDia).Font.Bold = True
                ws.Cells(filasDias(iFila), colMeses(iCol) + iDia).HorizontalAlignment = xlCenter
                ws.Cells(filasDias(iFila), colMeses(iCol) + iDia).Font.Size = 11
            Next iDia
        Next iCol
    Next iFila

    Dim festivos As Variant
    festivos = Array("2025-01-01", "2025-01-06", "2025-02-19", "2025-02-28", "2025-03-03", "2025-04-17", _
                     "2025-04-18", "2025-05-01", "2025-08-15", "2025-10-07", "2025-10-13", "2025-11-01", _
                     "2025-12-06", "2025-12-08", "2025-12-25")
    Dim noLaborables As Variant
    noLaborables = Array("2025-01-13", "2025-01-20", "2025-01-27", "2025-02-10", "2025-02-17", "2025-02-24", _
        "2025-03-10", "2025-03-17", "2025-03-24", "2025-03-31", "2025-04-04", "2025-04-07", "2025-04-14", "2025-04-21", _
        "2025-04-28", "2025-05-05", "2025-05-12", "2025-05-19", "2025-05-26", "2025-06-02", "2025-06-09", "2025-06-16", _
        "2025-06-23", "2025-06-30", "2025-07-07", "2025-07-14", "2025-07-21", "2025-07-28", "2025-08-04", "2025-08-11", _
        "2025-08-18", "2025-08-25", "2025-09-01", "2025-09-08", "2025-09-15", "2025-09-22", "2025-09-29", "2025-10-06", _
        "2025-10-13", "2025-10-20", "2025-10-27", "2025-11-03", "2025-11-10", "2025-11-17", "2025-11-24", "2025-12-01", _
        "2025-12-08", "2025-12-15", "2025-12-22", "2025-12-29")

    Dim mesFila As Variant, mesCol As Variant
    mesFila = Array(5, 5, 5, 14, 14, 14, 23, 23, 23, 32, 32, 32)
    mesCol = Array(1, 9, 17, 1, 9, 17, 1, 9, 17, 1, 9, 17)
    Dim m As Integer, f As Integer, c As Integer, d As Integer
    Dim fecha As Date, diasMes As Integer, primerDia As Integer
    Dim fechaStr As String
    For m = 0 To 11
        fecha = DateSerial(2025, m + 1, 1)
        diasMes = Day(DateSerial(2025, m + 2, 0))
        primerDia = Weekday(fecha, vbMonday)
        f = mesFila(m)
        c = mesCol(m) + primerDia - 2
        For d = 1 To diasMes
            ws.Cells(f, c).Value = d
            ws.Cells(f, c).HorizontalAlignment = xlCenter
            ws.Cells(f, c).Font.Size = 11
            ws.Cells(f, c).Font.Name = "Arial"
            fechaStr = "2025-" & Format(m + 1, "00") & "-" & Format(d, "00")
            If IsInArray(fechaStr, festivos) Then
                ws.Cells(f, c).Font.Color = vbRed
                ws.Cells(f, c).Font.Bold = True
            ElseIf IsInArray(fechaStr, noLaborables) Then
                ws.Cells(f, c).Font.Color = RGB(0, 0, 192)
                ws.Cells(f, c).Font.Bold = True
            Else
                ws.Cells(f, c).Font.Color = vbBlack
                ws.Cells(f, c).Font.Bold = False
            End If
            c = c + 1
            If c > mesCol(m) + 6 Then
                c = mesCol(m)
                f = f + 1
            End If
        Next d
    Next m

    Dim filaIni As Integer, filaFin As Integer, colIni As Integer, colFin As Integer
    For iFila = 0 To 3
        For iCol = 0 To 2
            filaIni = 3 + iFila * 9
            filaFin = filaIni + 8
            colIni = colMeses(iCol)
            colFin = colIni + 6
            With ws.Range(ws.Cells(filaIni, colIni), ws.Cells(filaFin, colFin))
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlThick
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlThick
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThick
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlThick
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
            End With
        Next iCol
    Next iFila

    Dim leyendaRow As Integer
    leyendaRow = 41
    ws.Cells(leyendaRow, 5).Value = "¦"
    ws.Cells(leyendaRow, 5).Font.Color = vbRed
    ws.Cells(leyendaRow, 6).Value = "FESTIVOS GENERALES"
    ws.Cells(leyendaRow, 9).Value = "¦"
    ws.Cells(leyendaRow, 9).Font.Color = RGB(0, 0, 192)
    ws.Cells(leyendaRow, 10).Value = "NO LABORABLES EN DEMAGRISA"
    ws.Range(ws.Cells(leyendaRow, 5), ws.Cells(leyendaRow, 12)).Font.Size = 11
    ws.Range(ws.Cells(leyendaRow, 5), ws.Cells(leyendaRow, 12)).Font.Bold = True
End Sub

Function IsInArray(val As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

' ========== TURNOS CON CABECERA CORRECTA Y CICLO QUE CAMBIA EL 28/07/2025 ================
Sub GenerarTurnosCicloAvanzado()
    Dim ws As Worksheet
    Dim fecha As Date, fechaInicio As Date, fechaFin As Date
    Dim row As Long, i As Integer
    Dim cicloCambio As Date
    Dim diaSemana As Integer
    Dim turnos(1 To 5) As String
    Dim empleados As Variant: empleados = Array("Carmelo", "María", "José", "Ángela", "Luisito")
    Dim observacion As String, horario As String, tipoTurno As String

    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("Turnos").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = Worksheets.Add
    ws.Name = "Turnos"
    ws.Cells(1, 1).Resize(1, 9).Value = Array("Fecha", "Día", "Carmelo", "María", "José", "Ángela", "Luisito", "Horario", "Turno/Observación")
    ws.Rows(1).Font.Bold = True

    fechaInicio = DateSerial(2025, 6, 5)
    fechaFin = DateSerial(2025, 12, 31)
    cicloCambio = DateSerial(2025, 7, 28) ' Cambio el 28/07/2025
    row = 2

    For fecha = fechaInicio To fechaFin
        diaSemana = Weekday(fecha, vbMonday) ' Lunes=1, Domingo=7
        For i = 1 To 5: turnos(i) = "-": Next i
        horario = ""
        observacion = ""
        tipoTurno = ""

        If fecha < cicloCambio Then
            tipoTurno = "fin de semana"
            Select Case diaSemana
                Case 1, 2 ' Lunes, Martes
                    turnos(1) = "-"    ' Carmelo
                    turnos(2) = "-"    ' María
                    turnos(3) = "-"    ' José
                    turnos(4) = "08:00–00:00" ' Ángela
                    turnos(5) = "08:00–00:00" ' Luisito
                    observacion = "Descansan Carmelo, María y José"
                Case 3 ' Miércoles
                    turnos(1) = "-"    ' Carmelo
                    turnos(2) = "-"    ' María
                    turnos(3) = "08:00–00:00" ' José
                    turnos(4) = "08:00–17:00" ' Ángela
                    turnos(5) = "08:00–17:00" ' Luisito
                Case 4, 5 ' Jueves y Viernes
                    turnos(1) = "17:00–00:00" ' Carmelo
                    turnos(2) = "17:00–00:00" ' María
                    turnos(3) = "08:00–00:00"  ' José
                    turnos(4) = "08:00–17:00"  ' Ángela
                    turnos(5) = "08:00–17:00"  ' Luisito
                Case 6, 7 ' Sábado y Domingo
                    turnos(1) = "09:00–00:00" ' Carmelo
                    turnos(2) = "09:00–00:00" ' María
                    turnos(3) = "09:00–00:00" ' José
                    turnos(4) = "-"           ' Ángela
                    turnos(5) = "-"           ' Luisito
            End Select
        Else
            tipoTurno = "semanal"
            Select Case diaSemana
                Case 1, 2 ' Lunes, Martes
                    turnos(1) = "08:00–00:00" ' Carmelo
                    turnos(2) = "08:00–00:00" ' María
                    turnos(3) = "08:00–00:00" ' José
                    turnos(4) = "-"           ' Ángela
                    turnos(5) = "-"           ' Luisito
                Case 3 ' Miércoles
                    turnos(1) = "-"           ' Carmelo
                    turnos(2) = "-"           ' María
                    turnos(3) = "08:00–00:00" ' José
                    turnos(4) = "08:00–17:00" ' Ángela
                    turnos(5) = "08:00–17:00" ' Luisito
                Case 4, 5 ' Jueves y Viernes
                    turnos(1) = "17:00–00:00" ' Carmelo
                    turnos(2) = "17:00–00:00" ' María
                    turnos(3) = "08:00–00:00" ' José
                    turnos(4) = "08:00–17:00" ' Ángela
                    turnos(5) = "08:00–17:00" ' Luisito
                Case 6, 7 ' Sábado y Domingo
                    turnos(1) = "-"           ' Carmelo
                    turnos(2) = "-"           ' María
                    turnos(3) = "-"           ' José
                    turnos(4) = "09:00–00:00" ' Ángela
                    turnos(5) = "09:00–00:00" ' Luisito
            End Select
        End If

        For i = 1 To 5
            If turnos(i) <> "-" Then
                If horario <> "" Then horario = horario & " | "
                horario = horario & empleados(i - 1) & ": " & turnos(i)
            End If
        Next i

        ws.Cells(row, 1).Value = fecha
        ws.Cells(row, 2).Value = Format(fecha, "dddd")
        ws.Cells(row, 3).Value = turnos(1)
        ws.Cells(row, 4).Value = turnos(2)
        ws.Cells(row, 5).Value = turnos(3)
        ws.Cells(row, 6).Value = turnos(4)
        ws.Cells(row, 7).Value = turnos(5)
        ws.Cells(row, 8).Value = horario
        ws.Cells(row, 9).Value = tipoTurno & IIf(observacion <> "", " - " & observacion, "")
        row = row + 1
    Next fecha

    ws.Columns("A:I").AutoFit
    MsgBox "¡Turnos generados correctamente según tus reglas!"
End Sub

' ========== RESUMEN DE GANANCIAS POR SEMANA (CORREGIDO) ================
Sub GenerarTurnosYResumenGanancias()
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
    Dim semanaKey As Variant

    lastRow = wsTurnos.Cells(wsTurnos.Rows.Count, 1).End(xlUp).row

    For row = 2 To lastRow
        fecha = wsTurnos.Cells(row, 1).Value
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

        ' Cuando cambia la semana o es el último registro, vuelca y resetea
        If row = lastRow Or _
           (row < lastRow And (Year(wsTurnos.Cells(row + 1, 1).Value) & "-S" & Format(Application.WorksheetFunction.WeekNum(wsTurnos.Cells(row + 1, 1).Value, 2), "00")) <> semanaStr) Then
            weekRow = dicSemanas(semanaStr)
            wsResumen.Cells(weekRow, 2).Value = carmelo + maria
            wsResumen.Cells(weekRow, 3).Value = jose
            wsResumen.Cells(weekRow, 4).Value = angela + luisito
            wsResumen.Cells(weekRow, 5).Value = wsResumen.Cells(weekRow, 2).Value + wsResumen.Cells(weekRow, 3).Value + wsResumen.Cells(weekRow, 4).Value
            carmelo = 0: maria = 0: jose = 0: angela = 0: luisito = 0
        End If
    Next row

    wsResumen.Columns("A:E").AutoFit
    MsgBox "Resumen profesional de ganancias semanales generado en la hoja 'ResumenGanancias'."
End Sub

' ========== MACRO TODO EN UNO ================
Sub GenerarTodo()
    Call Calendario2025_Cadiz_Imagen
    Call GenerarTurnosCicloAvanzado
    Call GenerarTurnosYResumenGanancias
    Call GraficaTurnos
    Call GraficaGanancias
End Sub
