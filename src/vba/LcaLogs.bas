Attribute VB_Name = "LcaLogs"
Option Explicit

' "Logs" sheet: every trace that grows with the simulation lives here, not on the LCA dashboard.
' One column per source, newest line first, plus the last VerificarTodo report.
Public Const LCA_LOGS_SHEET As String = "Logs"
Private Const MAX_LINES As Long = 128
Private Const FIRST_ROW As Long = 6
Private Const BG As Long = 2758415          ' RGB(15, 23, 42)
Private Const CARD As Long = 3877150        ' RGB(30, 41, 59)
Private Const HEAD As Long = 6897694        ' RGB(30, 64, 105)
Private Const INK As Long = 15788258        ' RGB(226, 232, 240)
Private Const MUTED As Long = 12100500      ' RGB(148, 163, 184)
Private Const ACCENT As Long = 16569213     ' RGB(125, 211, 252)
Private Const LAMP_ON As Long = 2477306     ' RGB(250, 204, 37)

Private mTests As String
Private mTestsStamp As String

Public Sub ConstruirLogs()
    Dim ws As Worksheet, titles As Variant, i As Long
    Set ws = LogsSheet(True)
    ws.Cells.Clear
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    ws.Cells.Interior.Color = BG
    ws.Cells.Font.Name = "Consolas"
    ws.Cells.Font.Size = 9
    ws.Cells.Font.Color = INK
    ws.Columns(1).ColumnWidth = 5
    ws.Range(ws.Columns(2), ws.Columns(9)).ColumnWidth = 62
    With ws.Range("B1")
        .Value = "LCA  |  LOGS DEL SIMULADOR"
        .Font.Name = "Segoe UI"
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = ACCENT
    End With
    ws.Range("B2").Value = "Trazas completas de cada modulo (la mas reciente arriba, max " & CStr(MAX_LINES) & " lineas) y el ultimo VerificarTodo. " & _
                           "Se actualizan solas al usar la maquina P12; tambien con ACTUALIZAR."
    ws.Range("B2").Font.Color = MUTED
    ws.Range("B2").Font.Name = "Segoe UI"
    Button ws, "B3", "ACTUALIZAR", "LogsActualizar", HEAD
    Button ws, "C3", "< VOLVER AL TABLERO", "LcaIrMaquina", HEAD
    Button ws, "D3", "VERIFICAR TODO (sin dialogos)", "LogsVerificarTodo", RGB(22, 101, 52)
    titles = Array("#", "P12  EJECUCION (instrucciones)", "P05/P06  BUS (ciclos)", "P09  CACHES (accesos)", _
                   "P08  PAGINACION (faults / CLOCK / swap)", "P10  CACHE + PAGINACION", "P13  TECLADO", "P14  LCD", _
                   "PRUEBAS  (ultimo VerificarTodo)")
    For i = 0 To 8
        With ws.Cells(FIRST_ROW - 1, i + 1)
            .Value = titles(i)
            .Font.Name = "Segoe UI"
            .Font.Bold = True
            .Font.Size = 10
            .Font.Color = BG
            .Interior.Color = ACCENT
        End With
    Next i
    ws.Range(ws.Cells(FIRST_ROW, 2), ws.Cells(FIRST_ROW + MAX_LINES - 1, 9)).Interior.Color = CARD
    ws.Range(ws.Cells(FIRST_ROW, 2), ws.Cells(FIRST_ROW + MAX_LINES - 1, 9)).Borders(xlInsideVertical).Color = BG
    ws.Range(ws.Cells(FIRST_ROW, 2), ws.Cells(FIRST_ROW + MAX_LINES - 1, 9)).Borders(xlInsideVertical).Weight = xlMedium
    For i = 1 To MAX_LINES
        ws.Cells(FIRST_ROW + i - 1, 1).Value = i
        ws.Cells(FIRST_ROW + i - 1, 1).Font.Color = MUTED
    Next i
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
    ws.Range("A" & CStr(FIRST_ROW)).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    LogsActualizar
End Sub

' Copies every engine trace into its column. Engines that were never reset are shown as "-".
Public Sub LogsActualizar()
    Dim ws As Worksheet, data() As Variant, r As Long, c As Long, n As Long, k As Long
    Set ws = LogsSheet(False)
    If ws Is Nothing Then Exit Sub
    ReDim data(1 To MAX_LINES, 1 To 8)
    For c = 1 To 8
        For r = 1 To MAX_LINES
            data(r, c) = vbNullString
        Next r
    Next c
    On Error Resume Next
    n = ExecTraceCount()
    For r = 1 To Min(n, MAX_LINES): data(r, 1) = ExecTraceLine(r - 1): Next r
    n = BusTraceCount()
    For r = 1 To Min(n, MAX_LINES): data(r, 2) = BusTraceLine(n - r): Next r
    n = CacheTraceCount()
    For r = 1 To Min(n, MAX_LINES): data(r, 3) = CacheTraceLine(n - r): Next r
    n = PagerTraceCount()
    For r = 1 To Min(n, MAX_LINES): data(r, 4) = PagerTraceLine(n - r): Next r
    n = CachePathTraceCount()
    For r = 1 To Min(n, MAX_LINES): data(r, 5) = CachePathTraceLine(r - 1): Next r
    n = KeyboardTraceCount()
    For r = 1 To Min(n, MAX_LINES): data(r, 6) = KeyboardTraceLine(r - 1): Next r
    n = LcdTraceCount()
    For r = 1 To Min(n, MAX_LINES): data(r, 7) = LcdTraceLine(r - 1): Next r
    On Error GoTo 0
    If Len(mTests) > 0 Then
        Dim lines As Variant
        lines = Split(mTestsStamp & vbLf & Replace(mTests, vbCr, ""), vbLf)
        k = 1
        For r = LBound(lines) To UBound(lines)
            If k > MAX_LINES Then Exit For
            If Len(Trim$(CStr(lines(r)))) > 0 Then
                data(k, 8) = CStr(lines(r))
                k = k + 1
            End If
        Next r
    End If
    ws.Range(ws.Cells(FIRST_ROW, 2), ws.Cells(FIRST_ROW + MAX_LINES - 1, 9)).Value = data
    ws.Range("E3").Value = "Actualizado: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Range("E3").Font.Color = MUTED
End Sub

' Keeps the last VerificarTodo report for the PRUEBAS column.
Public Sub LogsGuardarPruebas(ByVal report As String)
    mTests = report
    mTestsStamp = "VerificarTodo  " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    LogsActualizar
End Sub

Public Sub LogsVerificarTodo()
    Call LcaVerificarTodoTexto
    LogsActualizar
    LcaGoto LogsSheet(True), 0
    LogsSheet(True).Range("I" & CStr(FIRST_ROW)).Select
End Sub

Public Sub LcaIrLogs()
    If LogsSheet(False) Is Nothing Then
        ConstruirLogs
    Else
        LogsActualizar
        LogsSheet(False).Activate
    End If
End Sub

Public Function LogsSheet(ByVal create As Boolean) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = LCA_LOGS_SHEET Then
            Set LogsSheet = ws
            Exit Function
        End If
    Next ws
    If create Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = LCA_LOGS_SHEET
        ws.Tab.Color = LAMP_ON
        Set LogsSheet = ws
    End If
End Function

Private Function Min(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then Min = a Else Min = b
End Function

Private Sub Button(ByVal ws As Worksheet, ByVal address As String, ByVal caption As String, _
                   ByVal macroName As String, ByVal fill As Long)
    Dim area As Range, s As Shape
    Set area = ws.Range(address)
    ws.Rows(area.Row).RowHeight = 26
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left + 4, area.Top + 2, area.Width - 8, area.Height - 4)
    s.Name = "LOG_" & macroName
    s.Fill.ForeColor.RGB = fill
    s.Line.ForeColor.RGB = ACCENT
    With s.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = caption
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Sub
