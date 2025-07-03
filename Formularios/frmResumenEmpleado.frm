VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmResumenEmpleado 
   Caption         =   "UserForm1"
   ClientHeight    =   4910
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7540
   OleObjectBlob   =   "frmResumenEmpleado.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmResumenEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Constante para alineación derecha (por compatibilidad)
Const fmTextAlignRight = 3

Private Sub UserForm_Initialize()
    ' Tooltips para controles
    cmbEmpleado.ControlTipText = "Selecciona el empleado a consultar"
    btnConsultar.ControlTipText = "Haz clic para actualizar el resumen"
    btnCerrar.ControlTipText = "Cierra este formulario"
    
    ' Alineación derecha de los TextBox numéricos
    txtSemanas.TextAlign = fmTextAlignRight
    txtSueldo.TextAlign = fmTextAlignRight

    ' Cargar empleados en ComboBox
    With cmbEmpleado
        .Clear
        .AddItem "Carmelo"
        .AddItem "María"
        .AddItem "José"
        .AddItem "Ángela"
        .AddItem "Luisito"
    End With

    ' Cargar meses en ComboBox
    Dim i As Integer
    cmbMes.Clear
    cmbMes.AddItem "Todos"
    For i = 1 To 12
        cmbMes.AddItem Format(DateSerial(Year(Date), i, 1), "mmmm")
    Next i
    cmbMes.ListIndex = 0 ' "Todos" seleccionado por defecto
    cmbMes.ControlTipText = "Filtra los turnos por mes"

    ' Inicializar resumen
    txtSemanas.Value = ""
    txtSueldo.Value = ""
    
    ' Configurar columnas del ListView
    With lvwDetalle
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Fecha", 90
        .ColumnHeaders.Add , , "Semana", 90
        .ColumnHeaders.Add , , "Sueldo", 60
        .ColumnHeaders.Add , , "Observación", 100
    End With
End Sub

Private Sub btnConsultar_Click()
    Dim wsTurnos As Worksheet
    Dim nombre As String
    Dim lastRow As Long, i As Long
    Dim fecha As Date, semanaStr As String, sueldo As Double
    Dim semanaTrabajada As Object: Set semanaTrabajada = CreateObject("Scripting.Dictionary")
    Dim idxEmpleado As Integer
    Dim item As ListItem
    Dim obs As String
    Dim totalSueldo As Double
    Dim mesSeleccionado As Integer

    Set wsTurnos = Worksheets("Turnos")
    nombre = cmbEmpleado.Value
    If nombre = "" Then
        MsgBox "Selecciona un empleado.", vbExclamation
        Exit Sub
    End If

    lvwDetalle.ListItems.Clear
    txtSemanas.Value = ""
    txtSueldo.Value = ""

    lastRow = wsTurnos.Cells(wsTurnos.Rows.Count, 1).End(xlUp).row

    ' Determinar columna del empleado
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

    totalSueldo = 0

    ' Filtro de mes
    If cmbMes.Value = "Todos" Then
        mesSeleccionado = 0
    Else
        mesSeleccionado = cmbMes.ListIndex
    End If

    For i = 2 To lastRow
        fecha = wsTurnos.Cells(i, 1).Value
        If (mesSeleccionado = 0 Or Month(fecha) = mesSeleccionado) Then
            If wsTurnos.Cells(i, 2 + idxEmpleado).Value <> "-" And wsTurnos.Cells(i, 2 + idxEmpleado).Value <> "Vacaciones" Then
                semanaStr = Year(fecha) & "-S" & Format(Application.WorksheetFunction.WeekNum(fecha, 2), "00")
                Select Case wsTurnos.Cells(i, 2 + idxEmpleado).Value
                    Case "08:00–00:00", "09:00–00:00"
                        sueldo = 100
                    Case "17:00–00:00", "08:00–17:00"
                        sueldo = 50
                    Case Else
                        sueldo = 0
                End Select

                obs = ""
                If Weekday(fecha, vbMonday) = 6 Then
                    obs = "Sábado"
                ElseIf Weekday(fecha, vbMonday) = 7 Then
                    obs = "Domingo"
                ElseIf EsFestivo(fecha) Then
                    obs = "Festivo"
                End If

                Set item = lvwDetalle.ListItems.Add(, , Format(fecha, "yyyy-mm-dd"))
                item.SubItems(1) = semanaStr
                item.SubItems(2) = sueldo
                item.SubItems(3) = obs

                totalSueldo = totalSueldo + sueldo
                If Not semanaTrabajada.Exists(semanaStr) Then
                    semanaTrabajada.Add semanaStr, True
                End If
            End If
        End If
    Next i

    txtSemanas.Value = semanaTrabajada.Count
    txtSueldo.Value = totalSueldo
End Sub

Private Sub btnCerrar_Click()
    Unload Me
End Sub

' --- Función para detectar festivos ---
Private Function EsFestivo(fecha As Date) As Boolean
    Dim festivos As Variant
    festivos = Array(#1/1/2025#, #1/6/2025#, #2/28/2025#, #3/3/2025#, # _
        4/17/2025#, #4/18/2025#, #5/1/2025#, #8/15/2025#, #10/7/2025#, # _
        10/13/2025#, #11/1/2025#, #12/6/2025#, #12/8/2025#, #12/25/2025#)
    Dim i As Integer
    For i = LBound(festivos) To UBound(festivos)
        If fecha = festivos(i) Then
            EsFestivo = True
            Exit Function
        End If
    Next i
    EsFestivo = False
End Function

