Attribute VB_Name = "modConsultaEmpleado"
' ================== modConsultaEmpleado.bas ==================
' Consulta de semanas, fechas y sueldo por empleado en hoja Principal

Public Sub MostrarResumenEmpleado()
    Dim wsTurnos As Worksheet, wsPrin As Worksheet
    Dim nombre As String
    Dim lastRow As Long, fila As Long, i As Long
    Dim fecha As String, semanaStr As String, sueldo As Double
    Dim semanaTrabajada As Object: Set semanaTrabajada = CreateObject("Scripting.Dictionary")
    Dim idxEmpleado As Integer

    Set wsTurnos = Worksheets("Turnos")
    Set wsPrin = Worksheets("Principal")
    nombre = wsPrin.Range("B3").Value
    If nombre = "" Then
        MsgBox "Selecciona un nombre en la celda B3.", vbExclamation
        Exit Sub
    End If

    wsPrin.Range("A6:D1000").ClearContents

    lastRow = wsTurnos.Cells(wsTurnos.Rows.Count, 1).End(xlUp).row
    fila = 6

    Dim nombresArr As Variant
    nombresArr = Array("Carmelo", "María", "José", "Ángela", "Luisito")
    idxEmpleado = -1
    For i = 0 To UBound(nombresArr)
        If nombresArr(i) = nombre Then idxEmpleado = i + 1
    Next i
    If idxEmpleado = -1 Then
        MsgBox "Empleado no reconocido.", vbCritical
        Exit Sub
    End If

    For i = 2 To lastRow
        If wsTurnos.Cells(i, 2 + idxEmpleado).Value <> "-" And wsTurnos.Cells(i, 2 + idxEmpleado).Value <> "Vacaciones" Then
            fecha = wsTurnos.Cells(i, 1).Value
            semanaStr = Year(wsTurnos.Cells(i, 1).Value) & "-S" & Format(Application.WorksheetFunction.WeekNum(wsTurnos.Cells(i, 1).Value, 2), "00")
            wsPrin.Cells(fila, 1).Value = fecha
            wsPrin.Cells(fila, 2).Value = nombre
            wsPrin.Cells(fila, 3).Value = semanaStr

            Select Case wsTurnos.Cells(i, 2 + idxEmpleado).Value
                Case "08:00–00:00", "09:00–00:00"
                    sueldo = 100
                Case "17:00–00:00", "08:00–17:00"
                    sueldo = 50
                Case Else
                    sueldo = 0
            End Select
            wsPrin.Cells(fila, 4).Value = sueldo

            If Not semanaTrabajada.Exists(semanaStr) Then
                semanaTrabajada.Add semanaStr, True
            End If

            fila = fila + 1
        End If
    Next i

    wsPrin.Range("F3").Value = "Semanas trabajadas:"
    wsPrin.Range("G3").Value = semanaTrabajada.Count
    wsPrin.Range("F4").Value = "Sueldo total (€):"
    wsPrin.Range("G4").Value = Application.Sum(wsPrin.Range("D6:D" & fila - 1))
End Sub

Attribute VB_Name = "modConsultaEmpleado"
' ================== modConsultaEmpleado.bas ==================
' Consulta de semanas, fechas y sueldo por empleado en hoja Principal

Public Sub MostrarResumenEmpleado()
    Dim wsTurnos As Worksheet, wsPrin As Worksheet
    Dim nombre As String
    Dim lastRow As Long, fila As Long, i As Long
    Dim fecha As String, semanaStr As String, sueldo As Double
    Dim semanaTrabajada As Object: Set semanaTrabajada = CreateObject("Scripting.Dictionary")
    Dim idxEmpleado As Integer

    Set wsTurnos = Worksheets("Turnos")
    Set wsPrin = Worksheets("Principal")
    nombre = wsPrin.Range("B3").Value
    If nombre = "" Then
        MsgBox "Selecciona un nombre en la celda B3.", vbExclamation
        Exit Sub
    End If

    wsPrin.Range("A6:D1000").ClearContents

    lastRow = wsTurnos.Cells(wsTurnos.Rows.Count, 1).End(xlUp).row
    fila = 6

    Dim nombresArr As Variant
    nombresArr = Array("Carmelo", "María", "José", "Ángela", "Luisito")
    idxEmpleado = -1
    For i = 0 To UBound(nombresArr)
        If nombresArr(i) = nombre Then idxEmpleado = i + 1
    Next i
    If idxEmpleado = -1 Then
        MsgBox "Empleado no reconocido.", vbCritical
        Exit Sub
    End If

    For i = 2 To lastRow
        If wsTurnos.Cells(i, 2 + idxEmpleado).Value <> "-" And wsTurnos.Cells(i, 2 + idxEmpleado).Value <> "Vacaciones" Then
            fecha = wsTurnos.Cells(i, 1).Value
            semanaStr = Year(wsTurnos.Cells(i, 1).Value) & "-S" & Format(Application.WorksheetFunction.WeekNum(wsTurnos.Cells(i, 1).Value, 2), "00")
            wsPrin.Cells(fila, 1).Value = fecha
            wsPrin.Cells(fila, 2).Value = nombre
            wsPrin.Cells(fila, 3).Value = semanaStr

            Select Case wsTurnos.Cells(i, 2 + idxEmpleado).Value
                Case "08:00–00:00", "09:00–00:00"
                    sueldo = 100
                Case "17:00–00:00", "08:00–17:00"
                    sueldo = 50
                Case Else
                    sueldo = 0
            End Select
            wsPrin.Cells(fila, 4).Value = sueldo

            If Not semanaTrabajada.Exists(semanaStr) Then
                semanaTrabajada.Add semanaStr, True
            End If

            fila = fila + 1
        End If
    Next i

    wsPrin.Range("F3").Value = "Semanas trabajadas:"
    wsPrin.Range("G3").Value = semanaTrabajada.Count
    wsPrin.Range("F4").Value = "Sueldo total (€):"
    wsPrin.Range("G4").Value = Application.Sum(wsPrin.Range("D6:D" & fila - 1))
End Sub
