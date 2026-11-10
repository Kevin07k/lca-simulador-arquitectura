Attribute VB_Name = "LcaLayout"
Option Explicit

' One worksheet ("LCA") hosts every panel as a vertical band, in issue order. Each panel module keeps
' its own cell addresses relative to row 1 and shifts them by its band base (ROW_BASE) with LcaRows.
Public Const LCA_SHEET As String = "LCA"
Public Const LCA_BASE_CPU As Long = 0           ' rows   1..44  controls and registers (P04)
Public Const LCA_BASE_DIAGRAM As Long = 45      ' rows  46..87  CPU block diagram (P04)
Public Const LCA_BASE_BUS As Long = 88          ' rows  89..122 buses and RAM (P05/P06)
Public Const LCA_BASE_MMU As Long = 123         ' rows 124..156 page table and translation (P07)
Public Const LCA_BASE_PAGER As Long = 157       ' rows 158..192 page faults and swap (P08)
Public Const LCA_BASE_CACHE As Long = 193       ' rows 194..236 caches (P09)
Public Const LCA_BASE_CACHE_PATH As Long = 237  ' rows 238..284 caches + paging (P10)
Public Const LCA_BASE_EDITOR As Long = 285      ' rows 286..383 editor, assembler and loader (P11)

' Builds the whole workbook view: removes the old one-sheet-per-panel layout and anything left
' on the LCA sheet by an older band order, then builds every band top to bottom.
Public Sub ConstruirLca()
    Dim ws As Worksheet
    Dim item As Variant
    Dim previous As Boolean
    Dim i As Long
    previous = Application.DisplayAlerts
    Application.DisplayAlerts = False
    For Each item In Array("CPU", "CPU Diagrama", "P05 Buses", "P09 Caches", "P07 Paginas", "P08 Swap")
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name = CStr(item) Then
                ws.Delete
                Exit For
            End If
        Next ws
    Next item
    Application.DisplayAlerts = previous
    Set ws = LcaSheet(True)
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.Validation.Delete
    ConstruirDiagramaCpu
    ConstruirPanelCpu
    ConstruirPanelBuses
    ConstruirPanelPaginas
    ConstruirPanelSwap
    ConstruirPanelCaches
    ConstruirPanelCachePaginacion
    ConstruirPanelEditor
    LcaGoto ws, LCA_BASE_CPU
End Sub

Public Function LcaSheet(ByVal create As Boolean) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = LCA_SHEET Then
            Set LcaSheet = ws
            Exit Function
        End If
    Next ws
    If create Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = LCA_SHEET
        LcaSetColumns ws
        Set LcaSheet = ws
    End If
End Function

' Returns the LCA sheet only when a band is already built (its marker shape exists), or on create.
Public Function LcaBand(ByVal marker As String, ByVal create As Boolean) As Worksheet
    Dim ws As Worksheet
    Dim item As Shape
    Set ws = LcaSheet(create)
    If ws Is Nothing Then Exit Function
    If create Then
        Set LcaBand = ws
        Exit Function
    End If
    For Each item In ws.Shapes
        If item.Name = marker Then
            Set LcaBand = ws
            Exit Function
        End If
    Next item
End Function

' One shared grid for all bands.
Public Sub LcaSetColumns(ByVal ws As Worksheet)
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B:S").ColumnWidth = 14
End Sub

' Shifts every row number of an A1 address ("B2:R3", "1:40", "B7") by base.
' Addresses containing "$" are already absolute (Range.Address) and are returned unchanged.
Public Function LcaRows(ByVal address As String, ByVal base As Long) As String
    Dim i As Long
    Dim ch As String
    Dim digits As String
    Dim result As String
    If InStr(address, "$") > 0 Then
        LcaRows = address
        Exit Function
    End If
    For i = 1 To Len(address)
        ch = Mid$(address, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        Else
            If Len(digits) > 0 Then
                result = result & CStr(CLng(digits) + base)
                digits = vbNullString
            End If
            result = result & ch
        End If
    Next i
    If Len(digits) > 0 Then result = result & CStr(CLng(digits) + base)
    LcaRows = result
End Function

' Clears one band (cells, merges, validation) and its shapes; other bands are untouched.
Public Sub LcaPrepareBand(ByVal ws As Worksheet, ByVal base As Long, ByVal lastRow As Long, _
                          ByVal shapePrefix As String)
    Dim i As Long
    Dim band As Range
    Set band = ws.Range(LcaRows("A1:AZ" & CStr(lastRow), base))
    band.UnMerge
    band.Clear
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(shapePrefix)) = shapePrefix Then ws.Shapes(i).Delete
    Next i
    LcaSetColumns ws
End Sub

Public Sub LcaGoto(ByVal ws As Worksheet, ByVal base As Long)
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = base + 1
    ws.Cells(base + 1, 1).Select
End Sub

' Navigation entry points used by the panel headers.
Public Sub LcaIrCpu()
    LcaGoto LcaSheet(True), LCA_BASE_CPU
End Sub

Public Sub LcaIrDiagrama()
    LcaGoto LcaSheet(True), LCA_BASE_DIAGRAM
End Sub

Public Sub LcaIrBuses()
    LcaGoto LcaSheet(True), LCA_BASE_BUS
End Sub

Public Sub LcaIrCaches()
    LcaGoto LcaSheet(True), LCA_BASE_CACHE
End Sub

Public Sub LcaIrPaginas()
    LcaGoto LcaSheet(True), LCA_BASE_MMU
End Sub

Public Sub LcaIrSwap()
    LcaGoto LcaSheet(True), LCA_BASE_PAGER
End Sub

Public Sub LcaIrCachePaginacion()
    LcaGoto LcaSheet(True), LCA_BASE_CACHE_PATH
End Sub

Public Sub LcaIrEditor()
    LcaGoto LcaSheet(True), LCA_BASE_EDITOR
End Sub
