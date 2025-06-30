Attribute VB_Name = "modCalendario"
' ================== modCalendario.bas ==================
' Genera el calendario anual con festivos, colores y leyenda

Public Sub Calendario2025_Cadiz_Imagen()
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

    ws.Range("A1:W1").Merge
    With ws.Range("A1:W1")
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

    ' === FESTIVOS 2025 (Cádiz/Andalucía) y nombres ===
    Dim festivos As Variant, festivosNombres As Variant
    festivos = Array( _
        "2025-01-01", "2025-01-06", "2025-02-28", "2025-03-03", _
        "2025-04-17", "2025-04-18", "2025-05-01", "2025-08-15", _
        "2025-10-07", "2025-10-13", "2025-11-01", "2025-12-06", _
        "2025-12-08", "2025-12-25")
    festivosNombres = Array( _
        "Año Nuevo", "Epifanía del Señor", "Día de Andalucía", "Lunes de Carnaval", _
        "Jueves Santo", "Viernes Santo", "Fiesta del Trabajo", "Asunción de la Virgen", _
        "Virgen del Rosario", "Día siguiente al día de la Hispanidad", "Todos los Santos", _
        "Constitución Española", "Inmaculada Concepción", "Navidad")

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
    Dim fechaStr As String, diaSemana As Integer

    For m = 0 To 11
        fecha = DateSerial(2025, m + 1, 1)
        diasMes = Day(DateSerial(2025, m + 2, 0))
        primerDia = Weekday(fecha, vbMonday)
        f = mesFila(m)
        c = mesCol(m) + ((primerDia - 1) Mod 7)
        For d = 1 To diasMes
            fechaStr = "2025-" & Format(m + 1, "00") & "-" & Format(d, "00")
            diaSemana = ((c - mesCol(m)) Mod 7) + 1

            ws.Cells(f, c).Value = d
            ws.Cells(f, c).HorizontalAlignment = xlCenter
            ws.Cells(f, c).Font.Size = 11
            ws.Cells(f, c).Font.Name = "Arial"
            ws.Cells(f, c).Font.Color = vbBlack
            ws.Cells(f, c).Font.Bold = False
            ws.Cells(f, c).Interior.ColorIndex = xlNone

            Dim festivoIdx As Integer
            festivoIdx = GetFestivoIndex(fechaStr, festivos)
            If festivoIdx > -1 Then
                ws.Cells(f, c).Interior.Color = RGB(173, 216, 230) ' Celeste
                ws.Cells(f, c).Font.Bold = True
                ws.Cells(f + 1, c).Value = festivosNombres(festivoIdx)
                ws.Cells(f + 1, c).Font.Color = RGB(0, 120, 180)
                ws.Cells(f + 1, c).Font.Size = 7
                ws.Cells(f + 1, c).HorizontalAlignment = xlCenter
                ws.Cells(f + 1, c).WrapText = True
            End If

            If diaSemana = 7 Then
                ws.Cells(f, c).Interior.Color = RGB(255, 200, 200) ' Rojo claro para domingos
                ws.Cells(f, c).Font.Color = vbRed
                ws.Cells(f, c).Font.Bold = True
            End If

            If festivoIdx > -1 And diaSemana = 7 Then
                ws.Cells(f, c).Interior.Color = RGB(255, 170, 200)
            End If

            If IsInArray(fechaStr, noLaborables) Then
                ws.Cells(f, c).Font.Color = RGB(0, 0, 192)
                ws.Cells(f, c).Font.Bold = True
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

    ' Leyenda completa con nombres de festivos
    Dim leyendaRow As Integer
    leyendaRow = 43
    With ws.Range("E" & leyendaRow & ":H" & leyendaRow)
        .Merge
        .Value = "¦ FESTIVOS GENERALES (Celeste):"
        .Font.Color = vbBlack
        .Interior.Color = RGB(173, 216, 230)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With

    Dim festRow As Integer
    festRow = leyendaRow + 1
    Dim fj As Integer, festTxt As String
    For fj = 0 To UBound(festivos)
        festTxt = festivos(fj) & " - " & festivosNombres(fj)
        ws.Range("E" & (festRow + fj) & ":H" & (festRow + fj)).Merge
        ws.Range("E" & (festRow + fj) & ":H" & (festRow + fj)).Value = festTxt
        ws.Range("E" & (festRow + fj) & ":H" & (festRow + fj)).Font.Size = 9
        ws.Range("E" & (festRow + fj) & ":H" & (festRow + fj)).Font.Color = RGB(0, 120, 180)
        ws.Range("E" & (festRow + fj) & ":H" & (festRow + fj)).HorizontalAlignment = xlLeft
    Next fj

    With ws.Range("J" & leyendaRow & ":M" & leyendaRow)
        .Merge
        .Value = "¦ NO LABORABLES (Azul)"
        .Font.Color = RGB(0, 0, 192)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With

    With ws.Range("O" & leyendaRow & ":Q" & leyendaRow)
        .Merge
        .Value = "DOMINGOS (Rojo claro)"
        .Font.Color = vbRed
        .Font.Bold = True
        .Interior.Color = RGB(255, 200, 200)
        .HorizontalAlignment = xlLeft
    End With
End Sub

Public Function IsInArray(val As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

Public Function GetFestivoIndex(val As String, arr As Variant) As Integer
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then
            GetFestivoIndex = i
            Exit Function
        End If
    Next i
    GetFestivoIndex = -1
End Function

