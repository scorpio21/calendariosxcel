Attribute VB_Name = "modTurnos"
' ================== modTurnos.bas ==================
' Genera la hoja de turnos con formato especial y encabezado fijo

Public Sub GenerarTurnosCicloAvanzado()
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
    ws.Cells(1, 1).Resize(1, 8).Value = Array("Fecha", "Día", "Carmelo", "María", "José", "Ángela", "Luisito", "Horario")
    ws.Rows(1).Font.Bold = True

    ' Inmovilizar encabezado
    ws.Activate
    ws.Rows(2).Select
    ActiveWindow.FreezePanes = True

    fechaInicio = DateSerial(2025, 6, 5)
    fechaFin = DateSerial(2025, 12, 31)
    cicloCambio = DateSerial(2025, 7, 28)
    row = 2

    For fecha = fechaInicio To fechaFin
        diaSemana = Weekday(fecha, vbMonday)
        If fecha >= DateSerial(2025, 9, 1) And fecha <= DateSerial(2025, 9, 15) Then
            ws.Cells(row, 1).Value = fecha
            ws.Cells(row, 2).Value = Format(fecha, "dddd")
            For i = 1 To 5
                ws.Cells(row, 2 + i).Value = "Vacaciones"
                With ws.Cells(row, 2 + i)
                    .Interior.Color = RGB(255, 255, 0)   ' Amarillo
                    .Font.Color = RGB(0, 0, 192)        ' Azul
                    .Font.Bold = True
                End With
            Next i
            ws.Cells(row, 8).Value = ""
            ws.Cells(row, 9).Value = "Vacaciones"
            With ws.Cells(row, 9)
                .Interior.Color = RGB(255, 255, 0)
                .Font.Color = RGB(0, 0, 192)
                .Font.Bold = True
            End With
            row = row + 1
        Else
            For i = 1 To 5: turnos(i) = "-": Next i
            horario = ""
            observacion = ""
            tipoTurno = ""

            If fecha < cicloCambio Then
                tipoTurno = "fin de semana"
                Select Case diaSemana
                    Case 1, 2
                        turnos(1) = "-"
                        turnos(2) = "-"
                        turnos(3) = "-"
                        turnos(4) = "08:00–00:00"
                        turnos(5) = "08:00–00:00"
                        observacion = "Descansan Carmelo, María y José"
                    Case 3
                        turnos(1) = "-"
                        turnos(2) = "-"
                        turnos(3) = "08:00–00:00"
                        turnos(4) = "08:00–17:00"
                        turnos(5) = "08:00–17:00"
                    Case 4, 5
                        turnos(1) = "17:00–00:00"
                        turnos(2) = "17:00–00:00"
                        turnos(3) = "08:00–00:00"
                        turnos(4) = "08:00–17:00"
                        turnos(5) = "08:00–17:00"
                    Case 6, 7
                        turnos(1) = "09:00–00:00"
                        turnos(2) = "09:00–00:00"
                        turnos(3) = "09:00–00:00"
                        turnos(4) = "-"
                        turnos(5) = "-"
                End Select
            Else
                tipoTurno = "semanal"
                Select Case diaSemana
                    Case 1, 2
                        turnos(1) = "08:00–00:00"
                        turnos(2) = "08:00–00:00"
                        turnos(3) = "08:00–00:00"
                        turnos(4) = "-"
                        turnos(5) = "-"
                    Case 3
                        turnos(1) = "-"
                        turnos(2) = "-"
                        turnos(3) = "08:00–00:00"
                        turnos(4) = "08:00–17:00"
                        turnos(5) = "08:00–17:00"
                    Case 4, 5
                        turnos(1) = "17:00–00:00"
                        turnos(2) = "17:00–00:00"
                        turnos(3) = "08:00–00:00"
                        turnos(4) = "08:00–17:00"
                        turnos(5) = "08:00–17:00"
                    Case 6, 7
                        turnos(1) = "-"
                        turnos(2) = "-"
                        turnos(3) = "-"
                        turnos(4) = "09:00–00:00"
                        turnos(5) = "09:00–00:00"
                End Select
            End If

            For i = 1 To 5
                ws.Cells(row, 2 + i).Value = turnos(i)
                If diaSemana = 6 Or diaSemana = 7 Then
                    With ws.Cells(row, 2 + i)
                        .Interior.Color = RGB(255, 0, 0)    ' Rojo
                        .Font.Color = RGB(0, 0, 192)        ' Azul
                        .Font.Bold = True
                    End With
                Else
                    With ws.Cells(row, 2 + i)
                        .Interior.ColorIndex = xlNone
                        .Font.Color = vbBlack
                        .Font.Bold = False
                    End With
                End If
                If turnos(i) <> "-" Then
                    If horario <> "" Then horario = horario & " | "
                    horario = horario & empleados(i - 1) & ": " & turnos(i)
                End If
            Next i

            ws.Cells(row, 1).Value = fecha
            ws.Cells(row, 2).Value = Format(fecha, "dddd")
            ws.Cells(row, 8).Value = horario

            If diaSemana = 6 Or diaSemana = 7 Then
                ws.Cells(row, 9).Value = tipoTurno & IIf(observacion <> "", " - " & observacion, "")
                With ws.Cells(row, 9)
                    .Interior.Color = RGB(255, 0, 0)
                    .Font.Color = RGB(0, 0, 192)
                    .Font.Bold = True
                End With
            Else
                ws.Cells(row, 9).Value = tipoTurno & IIf(observacion <> "", " - " & observacion, "")
                With ws.Cells(row, 9)
                    .Interior.ColorIndex = xlNone
                    .Font.Color = vbBlack
                    .Font.Bold = False
                End With
            End If
            row = row + 1
        End If
    Next fecha

    ws.Columns("A:I").AutoFit
    MsgBox "¡Turnos generados correctamente según tus reglas y formato solicitado!"
End Sub
