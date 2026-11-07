Attribute VB_Name = "LcaCachePanel"
Option Explicit

Private Const SHEET_NAME As String = "P09 Caches"
Private Const DEMO_STEPS As Long = 11
Private mDemoStep As Long
Private mDemoStarted As Boolean
Private mDemoStatus As String

Public Sub ConstruirPanelCaches()
    Dim ws As Worksheet
    Dim i As Long
    Set ws = CacheSheet(True)
    ws.Cells.UnMerge
    ws.Cells.Clear
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B:R").ColumnWidth = 12
    ws.Rows("1:42").RowHeight = 23
    ws.Range("A1:S43").Interior.Color = RGB(15, 23, 42)
    ws.Range("A1:S43").Font.Name = "Consolas"
    ws.Range("A1:S43").Font.Color = RGB(226, 232, 240)
    ws.Range("B2:R3").Merge
    With ws.Range("B2")
        .Value = "P09  /  CACHES L1I + L1D + L2"
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 22
        .Font.Color = RGB(125, 211, 252)
    End With
    ws.Range("B4:R4").Merge
    ws.Range("B4").Value = "Directo / direct-mapped | linea 4 B | write-through + write-allocate | etiquetas fisicas / physical tags"
    AddCacheButton ws, "B6:H7", "REINICIAR / RESET", "CacheDemoReset"
    AddCacheButton ws, "I6:R7", "SIGUIENTE ACCESO / NEXT ACCESS", "CacheDemoNext"
    ws.Range("B9:R9").Merge
    ws.Range("B9").Value = "CONTADORES POR NIVEL / PER-LEVEL COUNTERS"
    ws.Range("B10").Value = "NIVEL / LEVEL"
    ws.Range("D10").Value = "HIT"
    ws.Range("E10").Value = "MISS"
    ws.Range("F10").Value = "TOTAL"
    ws.Range("G10").Value = "% HIT"
    ws.Range("B10:G10").Font.Bold = True
    ws.Range("B11:G13").Interior.Color = RGB(30, 41, 59)
    ws.Range("B15:R15").Merge
    ws.Range("B15").Value = "LINEAS / LINES  (V valida, TAG, bytes 0..3, PA cubiertas / covered)"
    ws.Range("B16").Value = "CACHE"
    ws.Range("C16").Value = "IDX"
    ws.Range("D16").Value = "V"
    ws.Range("E16").Value = "TAG"
    ws.Range("F16").Value = "B0"
    ws.Range("G16").Value = "B1"
    ws.Range("H16").Value = "B2"
    ws.Range("I16").Value = "B3"
    ws.Range("J16").Value = "PA"
    ws.Range("B16:J16").Font.Bold = True
    ws.Range("B17:J28").Interior.Color = RGB(30, 41, 59)
    ws.Range("B30:R30").Merge
    ws.Range("B30").Value = "TRAZA REAL DEL MOTOR / LIVE ENGINE TRACE (ultimos 8 / last 8)"
    For i = 31 To 38
        ws.Range(ws.Cells(i, 2), ws.Cells(i, 18)).Merge
    Next i
    ws.Range("B31:R38").Interior.Color = RGB(30, 41, 59)
    ws.Range("B40:R41").Merge
    ws.Range("B40:R41").WrapText = True
    ws.Range("B40").Font.Color = RGB(125, 211, 252)
    CacheDemoReset
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
End Sub

Public Sub MostrarPanelCaches()
    Dim ws As Worksheet
    Set ws = CacheSheet(False)
    If ws Is Nothing Then
        ConstruirPanelCaches
    Else
        ws.Activate
        CachePanelActualizar
    End If
End Sub

' Redraws counters, lines and trace from the engine; does nothing if the sheet is not built.
Public Sub CachePanelActualizar()
    Dim ws As Worksheet
    Dim level As Long, idx As Long, n As Long, row As Long, i As Long, total As Long
    Dim shown As Long
    Set ws = CacheSheet(False)
    If ws Is Nothing Then Exit Sub
    For level = 0 To 2
        row = 11 + level
        total = CacheHits(level) + CacheMisses(level)
        ws.Cells(row, 2).Value = CacheLevelName(level)
        ws.Cells(row, 4).Value = CacheHits(level)
        ws.Cells(row, 5).Value = CacheMisses(level)
        ws.Cells(row, 6).Value = total
        If total = 0 Then
            ws.Cells(row, 7).Value = "-"
        Else
            ws.Cells(row, 7).Value = Format$(100 * CacheHits(level) / total, "0.0") & "%"
        End If
    Next level
    row = 17
    For level = 0 To 2
        n = CacheLineCount(level)
        For idx = 0 To n - 1
            ws.Cells(row, 2).Value = CacheLevelName(level)
            ws.Cells(row, 3).Value = idx
            If CacheLineValid(level, idx) Then
                ws.Cells(row, 4).Value = 1
                ws.Cells(row, 5).Value = CacheLineTag(level, idx)
                For i = 0 To 3
                    ws.Cells(row, 6 + i).Value = Right$("00" & Hex$(CacheLineByte(level, idx, i)), 2) & "h"
                Next i
                ws.Cells(row, 10).Value = Right$("00" & Hex$((CacheLineTag(level, idx) * n + idx) * LCA_CACHE_LINE_BYTES), 2) & "h.." & _
                    Right$("00" & Hex$((CacheLineTag(level, idx) * n + idx) * LCA_CACHE_LINE_BYTES + 3), 2) & "h"
            Else
                ws.Cells(row, 4).Value = 0
                ws.Range(ws.Cells(row, 5), ws.Cells(row, 10)).Value = "-"
            End If
            row = row + 1
        Next idx
    Next level
    n = CacheTraceCount()
    For i = 0 To 7
        ws.Cells(31 + i, 2).Value = vbNullString
        If n > i Then ws.Cells(31 + i, 2).Value = CacheTraceLine(n - 1 - i)
    Next i
    ws.Range("B40").Value = mDemoStatus
End Sub

Public Sub CacheDemoReset()
    RamReset
    CacheReset
    RamWriteUser &H40, &H10
    RamWriteUser &H41, &H0
    RamWriteUser &H42, &H2A
    RamWriteUser &H43, &H0
    RamWriteUser &H48, &H33
    RamWriteUser &H6C, 1
    RamWriteUser &H6D, 2
    RamWriteUser &H6E, 3
    RamWriteUser &H6F, 4
    mDemoStep = 0
    mDemoStarted = True
    mDemoStatus = "RAM 40-43h = 10 00 2A 00. SIGUIENTE ejecuta un acceso real (fetch 40-43h, datos 6Ch/6Dh/70h, conflicto 48h)."
    CachePanelActualizar
End Sub

' Exactly one real cache access per click, using physical addresses.
Public Sub CacheDemoNext()
    On Error GoTo failed
    If Not mDemoStarted Then CacheDemoReset
    Select Case mDemoStep
        Case 0: CacheRead &H40, LCA_CACHE_FETCH
        Case 1: CacheRead &H41, LCA_CACHE_FETCH
        Case 2: CacheRead &H42, LCA_CACHE_FETCH
        Case 3: CacheRead &H43, LCA_CACHE_FETCH
        Case 4: CacheRead &H6D, LCA_CACHE_DATA
        Case 5: CacheRead &H6C, LCA_CACHE_DATA
        Case 6: CacheWrite &H6D, &H55
        Case 7: CacheRead &H6D, LCA_CACHE_DATA
        Case 8: CacheWrite &H70, &H77
        Case 9: CacheRead &H48, LCA_CACHE_FETCH
        Case 10: CacheRead &H40, LCA_CACHE_FETCH
        Case Else
            mDemoStatus = "Fin / Done. RESET para repetir / to repeat."
            CachePanelActualizar
            Exit Sub
    End Select
    mDemoStep = mDemoStep + 1
    mDemoStatus = "Acceso / access " & CStr(mDemoStep) & "/" & CStr(DEMO_STEPS) & ": " & CacheLastEvent()
    CachePanelActualizar
    Exit Sub
failed:
    mDemoStatus = "ERROR: " & Err.Description
    CachePanelActualizar
End Sub

Private Function CacheSheet(ByVal create As Boolean) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SHEET_NAME Then
            Set CacheSheet = ws
            Exit Function
        End If
    Next ws
    If create Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_NAME
        Set CacheSheet = ws
    End If
End Function

Private Sub AddCacheButton(ByVal ws As Worksheet, ByVal address As String, _
                           ByVal caption As String, ByVal macroName As String)
    Dim area As Range
    Dim button As Shape
    Set area = ws.Range(address)
    Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left, area.Top, area.Width - 4, area.Height - 2)
    button.Name = "P09_" & macroName
    button.Fill.ForeColor.RGB = RGB(37, 99, 235)
    button.Line.Visible = msoFalse
    button.TextFrame2.TextRange.Text = caption
    button.TextFrame2.TextRange.Font.Name = "Segoe UI"
    button.TextFrame2.TextRange.Font.Size = 12
    button.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    button.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    button.TextFrame2.VerticalAnchor = msoAnchorMiddle
    button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Sub
