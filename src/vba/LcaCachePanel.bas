Attribute VB_Name = "LcaCachePanel"
Option Explicit

Private Const ROW_BASE As Long = LCA_BASE_CACHE
Private Const DEMO_STEPS As Long = 11
Private mDemoStep As Long
Private mDemoStarted As Boolean
Private mDemoStatus As String

Public Sub ConstruirPanelCaches()
    Dim ws As Worksheet
    Dim i As Long
    Set ws = CacheSheet(True)
    LcaPrepareBand ws, ROW_BASE, 43, "P09_"
    ws.Rows(RA("1:42")).RowHeight = 23
    ws.Range(RA("A1:AZ43")).Interior.Color = RGB(15, 23, 42)
    ws.Range(RA("A1:AZ43")).Font.Name = "Consolas"
    ws.Range(RA("A1:AZ43")).Font.Color = RGB(226, 232, 240)
    ws.Range(RA("B2:R3")).Merge
    With ws.Range(RA("B2"))
        .Value = "P09  /  CACHES L1I + L1D + L2"
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 22
        .Font.Color = RGB(125, 211, 252)
    End With
    ws.Range(RA("B4:R4")).Merge
    ws.Range(RA("B4")).Value = "Directo / direct-mapped | linea 4 B | write-through + write-allocate | etiquetas fisicas / physical tags"
    AddCacheButton ws, "B6:H7", "REINICIAR / RESET", "CacheDemoReset"
    AddCacheButton ws, "I6:R7", "SIGUIENTE ACCESO / NEXT ACCESS", "CacheDemoNext"
    ws.Range(RA("B9:R9")).Merge
    ws.Range(RA("B9")).Value = "CONTADORES POR NIVEL / PER-LEVEL COUNTERS"
    ws.Range(RA("B10")).Value = "NIVEL / LEVEL"
    ws.Range(RA("D10")).Value = "HIT"
    ws.Range(RA("E10")).Value = "MISS"
    ws.Range(RA("F10")).Value = "TOTAL"
    ws.Range(RA("G10")).Value = "% HIT"
    ws.Range(RA("B10:G10")).Font.Bold = True
    ws.Range(RA("B11:G13")).Interior.Color = RGB(30, 41, 59)
    ws.Range(RA("B15:R15")).Merge
    ws.Range(RA("B15")).Value = "LINEAS / LINES  (V valida, TAG, bytes 0..3, PA cubiertas / covered)"
    ws.Range(RA("B16")).Value = "CACHE"
    ws.Range(RA("C16")).Value = "IDX"
    ws.Range(RA("D16")).Value = "V"
    ws.Range(RA("E16")).Value = "TAG"
    ws.Range(RA("F16")).Value = "B0"
    ws.Range(RA("G16")).Value = "B1"
    ws.Range(RA("H16")).Value = "B2"
    ws.Range(RA("I16")).Value = "B3"
    ws.Range(RA("J16")).Value = "PA"
    ws.Range(RA("B16:J16")).Font.Bold = True
    ws.Range(RA("B17:J28")).Interior.Color = RGB(30, 41, 59)
    ws.Range(RA("B10:J28")).HorizontalAlignment = xlLeft
    ws.Range(RA("B30:R30")).Merge
    ws.Range(RA("B30")).Value = "TRAZA REAL DEL MOTOR / LIVE ENGINE TRACE (ultimos 8 / last 8)"
    For i = 31 To 38
        ws.Range(ws.Cells(ROW_BASE + i, 2), ws.Cells(ROW_BASE + i, 18)).Merge
    Next i
    ws.Range(RA("B31:R38")).Interior.Color = RGB(30, 41, 59)
    ws.Range(RA("B40:R41")).Merge
    ws.Range(RA("B40:R41")).WrapText = True
    ws.Range(RA("B40")).Font.Color = RGB(125, 211, 252)
    CacheDemoReset
    LcaGoto ws, ROW_BASE
End Sub

Public Sub MostrarPanelCaches()
    Dim ws As Worksheet
    Set ws = CacheSheet(False)
    If ws Is Nothing Then
        ConstruirPanelCaches
    Else
        LcaGoto ws, ROW_BASE
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
        ws.Cells(ROW_BASE + row, 2).Value = CacheLevelName(level)
        ws.Cells(ROW_BASE + row, 4).Value = CacheHits(level)
        ws.Cells(ROW_BASE + row, 5).Value = CacheMisses(level)
        ws.Cells(ROW_BASE + row, 6).Value = total
        If total = 0 Then
            ws.Cells(ROW_BASE + row, 7).Value = "-"
        Else
            ws.Cells(ROW_BASE + row, 7).Value = Format$(100 * CacheHits(level) / total, "0.0") & "%"
        End If
    Next level
    row = 17
    For level = 0 To 2
        n = CacheLineCount(level)
        For idx = 0 To n - 1
            ws.Cells(ROW_BASE + row, 2).Value = CacheLevelName(level)
            ws.Cells(ROW_BASE + row, 3).Value = idx
            If CacheLineValid(level, idx) Then
                ws.Cells(ROW_BASE + row, 4).Value = 1
                ws.Cells(ROW_BASE + row, 5).Value = CacheLineTag(level, idx)
                For i = 0 To 3
                    ws.Cells(ROW_BASE + row, 6 + i).Value = Right$("00" & Hex$(CacheLineByte(level, idx, i)), 2) & "h"
                Next i
                ws.Cells(ROW_BASE + row, 10).Value = Right$("00" & Hex$((CacheLineTag(level, idx) * n + idx) * LCA_CACHE_LINE_BYTES), 2) & "h.." & _
                    Right$("00" & Hex$((CacheLineTag(level, idx) * n + idx) * LCA_CACHE_LINE_BYTES + 3), 2) & "h"
            Else
                ws.Cells(ROW_BASE + row, 4).Value = 0
                ws.Range(ws.Cells(ROW_BASE + row, 5), ws.Cells(ROW_BASE + row, 10)).Value = "-"
            End If
            row = row + 1
        Next idx
    Next level
    n = CacheTraceCount()
    For i = 0 To 7
        ws.Cells(ROW_BASE + 31 + i, 2).Value = vbNullString
        If n > i Then ws.Cells(ROW_BASE + 31 + i, 2).Value = CacheTraceLine(n - 1 - i)
    Next i
    ws.Range(RA("B40")).Value = mDemoStatus
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
    Set CacheSheet = LcaBand("P09_CacheDemoReset", create)
End Function

' Row shift of this panel inside the shared LCA sheet.
Private Function RA(ByVal address As String) As String
    RA = LcaRows(address, ROW_BASE)
End Function

Private Sub AddCacheButton(ByVal ws As Worksheet, ByVal address As String, _
                           ByVal caption As String, ByVal macroName As String)
    Dim area As Range
    Dim button As Shape
    Set area = ws.Range(RA(address))
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
