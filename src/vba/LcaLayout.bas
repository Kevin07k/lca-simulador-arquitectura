Attribute VB_Name = "LcaLayout"
Option Explicit

' One worksheet ("LCA") laid out as a 2-D grid of blocks, like a dashboard. Every panel keeps its own
' addresses relative to A1 of a 20-column block (A..T) and shifts them by its block origin
' (ROW_BASE, COL_BASE) with LcaShift / LcaCell. Column letters beyond T are clamped to the block.
'
'   COLUMN 1 (A..T)          COLUMN 2 (U..AN)         COLUMN 3 (AO..BH)      COLUMN 4 (BI..CB)
'   cover (portada) 1-26     P11 editor    1-98       P13 keyboard  1-36     P07 pages             1-33
'   P12 machine    28-95     P10 cache+pag 100-146    P14 LCD      38-73     P08 swap+CLOCK+map   35-104
'   P04 CPU+detail 97-168                             P05/P06 buses 75-108   P09 caches+chart    106-148
'   P04 diagram   170-211    P15 arithmetic demos 151-238 (column 2, below P10)
'                            P16 echo keyboard->LCD 148-202 (column 3, below the buses)
'                            P17 memory demo 150-217 (column 4, below the P09 caches)
' Long traces live on the separate "Logs" sheet (LcaLogs).
Public Const LCA_SHEET As String = "LCA"
Public Const LCA_BLOCK_COLS As Long = 20

Public Const LCA_BASE_COVER As Long = 0
Public Const LCA_COL_COVER As Long = 0
Public Const LCA_BASE_MACHINE As Long = 27
Public Const LCA_COL_MACHINE As Long = 0
Public Const LCA_BASE_CPU As Long = 96
Public Const LCA_COL_CPU As Long = 0
Public Const LCA_BASE_DIAGRAM As Long = 169
Public Const LCA_COL_DIAGRAM As Long = 0

Public Const LCA_BASE_EDITOR As Long = 0
Public Const LCA_COL_EDITOR As Long = 20
Public Const LCA_BASE_CACHE_PATH As Long = 99
Public Const LCA_COL_CACHE_PATH As Long = 20

Public Const LCA_BASE_KEYBOARD As Long = 0
Public Const LCA_COL_KEYBOARD As Long = 40
Public Const LCA_BASE_LCD As Long = 37
Public Const LCA_COL_LCD As Long = 40
Public Const LCA_BASE_BUS As Long = 74
Public Const LCA_COL_BUS As Long = 40

Public Const LCA_BASE_MMU As Long = 0
Public Const LCA_COL_MMU As Long = 60
Public Const LCA_BASE_PAGER As Long = 34
Public Const LCA_COL_PAGER As Long = 60
Public Const LCA_BASE_CACHE As Long = 105
Public Const LCA_COL_CACHE As Long = 60

Public Const LCA_BASE_ARITH As Long = 150
Public Const LCA_COL_ARITH As Long = 20

Public Const LCA_BASE_ECHO As Long = 147         ' below the P05/P06 bus detail (rows 109-146)
Public Const LCA_COL_ECHO As Long = 40

Public Const LCA_BASE_MEM As Long = 149           ' below the P09 caches (rows 106-148)
Public Const LCA_COL_MEM As Long = 60

Private Const GRID_LAST_ROW As Long = 245
Private Const GRID_LAST_COL As Long = 80
Private Const BG As Long = 2758415          ' RGB(15, 23, 42)

' Builds the whole dashboard. Two passes: the first settles the shared row heights of blocks that sit
' side by side, the second places every shape on its final cells. Removes the old sheets (Inicio, the
' P04 test report, the one-sheet-per-panel layout), keeps the logo, builds the Logs sheet and opens
' on the cover. The workbook ends with two sheets: LCA and Logs.
Public Sub ConstruirLca()
    Dim ws As Worksheet, item As Variant, previous As Boolean, i As Long, pass As Long
    Application.ScreenUpdating = False
    Set ws = LcaSheet(True)
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).Name <> LCA_LOGO_SHAPE Then ws.Shapes(i).Delete
    Next i
    PreserveLogo ws
    previous = Application.DisplayAlerts
    Application.DisplayAlerts = False
    For Each item In Array("CPU", "CPU Diagrama", "P05 Buses", "P09 Caches", "P07 Paginas", "P08 Swap", "P04 Tests", "Inicio")
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name = CStr(item) Then
                ws.Delete
                Exit For
            End If
        Next ws
    Next item
    Application.DisplayAlerts = previous
    Set ws = LcaSheet(True)
    If ws.Index <> 1 Then ws.Move Before:=ThisWorkbook.Worksheets(1)
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).Name <> LCA_LOGO_SHAPE Then ws.Shapes(i).Delete
    Next i
    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.Validation.Delete
    ws.Rows.RowHeight = 12
    ws.Range(ws.Cells(1, 1), ws.Cells(GRID_LAST_ROW, GRID_LAST_COL)).Interior.Color = BG
    For pass = 1 To 2
        ConstruirPortada
        ConstruirPanelEditor
        ConstruirDiagramaCpu
        ConstruirPanelCpu
        ConstruirCpuDetalle
        ConstruirPanelBuses
        ConstruirBusDetalle
        ConstruirPanelPaginas
        ConstruirMmuDetalle
        ConstruirPanelSwap
        ConstruirMemoriaDetalle
        ConstruirPanelCaches
        ConstruirPanelCachePaginacion
        ConstruirPanelTeclado
        ConstruirPanelLcd
        ConstruirPanelMaquina
        ConstruirPanelAritmetica
        ConstruirPanelEco
        ConstruirPanelMemoria
    Next pass
    LcaPulirBandas ws
    ConstruirLogs
    MemoriaDetalleActualizar
    PanelActualizar
    ws.Activate
    Application.ScreenUpdating = True
    LcaIrPortada
End Sub

' Moves the team logo from the old Inicio sheet into LCA (once) so the cover can show it.
Private Sub PreserveLogo(ByVal lca As Worksheet)
    Dim inicio As Worksheet, s As Shape, p As Shape
    For Each s In lca.Shapes
        If s.Name = LCA_LOGO_SHAPE Then Exit Sub
    Next s
    For Each inicio In ThisWorkbook.Worksheets
        If inicio.Name = "Inicio" Then Exit For
    Next inicio
    If inicio Is Nothing Then Exit Sub
    Dim before As Long
    On Error Resume Next
    For Each s In inicio.Shapes
        If s.Type = msoPicture Then
            before = lca.Shapes.Count
            s.Copy
            lca.Activate
            lca.Range("B2").Select
            lca.Paste
            If lca.Shapes.Count = before + 1 Then
                Set p = lca.Shapes(lca.Shapes.Count)
                If p.Type = msoPicture Then p.Name = LCA_LOGO_SHAPE Else p.Delete
            End If
            Exit For
        End If
    Next s
    Application.CutCopyMode = False
    On Error GoTo 0
End Sub

' Visual polish shared by the older text-only blocks (P05/P06, P07, P08, P09): centred tables,
' section titles in the common style and a "MAQUINA" button on every block header.
Public Sub LcaPulirBandas(ByVal ws As Worksheet)
    Dim blocks As Variant, i As Long, area As Range, s As Shape, item As Variant
    For Each item In Array(Array(LCA_BASE_MMU, LCA_COL_MMU, "B12:H28"), Array(LCA_BASE_PAGER, LCA_COL_PAGER, "B12:H28"), _
                           Array(LCA_BASE_CACHE, LCA_COL_CACHE, "B10:G13"), Array(LCA_BASE_CACHE, LCA_COL_CACHE, "B16:J28"))
        With ws.Range(LcaShift(CStr(item(2)), CLng(item(0)), CLng(item(1))))
            .HorizontalAlignment = xlCenter
            .Borders.Color = RGB(51, 65, 85)
            .Borders.Weight = xlThin
        End With
    Next item
    For Each item In Array(Array(LCA_BASE_BUS, LCA_COL_BUS, "B9"), Array(LCA_BASE_BUS, LCA_COL_BUS, "B13"), _
                           Array(LCA_BASE_BUS, LCA_COL_BUS, "B19"), Array(LCA_BASE_MMU, LCA_COL_MMU, "J12"), _
                           Array(LCA_BASE_MMU, LCA_COL_MMU, "J18"), Array(LCA_BASE_PAGER, LCA_COL_PAGER, "J12"), _
                           Array(LCA_BASE_CACHE, LCA_COL_CACHE, "B9"), Array(LCA_BASE_CACHE, LCA_COL_CACHE, "B15"), _
                           Array(LCA_BASE_CACHE, LCA_COL_CACHE, "B30"))
        With ws.Range(LcaShift(CStr(item(2)), CLng(item(0)), CLng(item(1))))
            .Font.Name = "Segoe UI"
            .Font.Bold = True
            .Font.Size = 11
            .Font.Color = RGB(125, 211, 252)
        End With
    Next item
    With ws.Range(LcaShift("B10:R29", LCA_BASE_BUS, LCA_COL_BUS))
        .Font.Name = "Consolas"
        .Font.Size = 11
        .IndentLevel = 1
    End With
    blocks = Array(Array(LCA_BASE_CPU, LCA_COL_CPU), Array(LCA_BASE_DIAGRAM, LCA_COL_DIAGRAM), _
                   Array(LCA_BASE_BUS, LCA_COL_BUS), Array(LCA_BASE_MMU, LCA_COL_MMU), _
                   Array(LCA_BASE_PAGER, LCA_COL_PAGER), Array(LCA_BASE_CACHE, LCA_COL_CACHE), _
                   Array(LCA_BASE_CACHE_PATH, LCA_COL_CACHE_PATH), Array(LCA_BASE_EDITOR, LCA_COL_EDITOR), _
                   Array(LCA_BASE_KEYBOARD, LCA_COL_KEYBOARD), Array(LCA_BASE_LCD, LCA_COL_LCD), _
                   Array(LCA_BASE_ARITH, LCA_COL_ARITH), Array(LCA_BASE_ECHO, LCA_COL_ECHO), Array(LCA_BASE_MEM, LCA_COL_MEM))
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 5) = "NAVM_" Then ws.Shapes(i).Delete
    Next i
    For i = 0 To UBound(blocks)
        Set area = ws.Cells(CLng(blocks(i)(0)) + 1, 19 + CLng(blocks(i)(1)))
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left + 2, area.Top + 1, area.Width - 4, 14)
        s.Name = "NAVM_" & CStr(i)
        s.Fill.ForeColor.RGB = RGB(250, 204, 37)
        s.Line.Visible = msoFalse
        With s.TextFrame2
            .MarginTop = 0
            .MarginBottom = 0
            .TextRange.Text = "< MAQUINA"
            .TextRange.Font.Name = "Segoe UI"
            .TextRange.Font.Size = 8
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = BG
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
        End With
        s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!LcaIrMaquina"
    Next i
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

' Returns the LCA sheet only when a block is already built (its marker shape exists), or on create.
Public Function LcaBand(ByVal marker As String, ByVal create As Boolean) As Worksheet
    Dim ws As Worksheet, item As Shape
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

' Column widths of one 20-column block: margin, 18 content columns, gutter.
Public Sub LcaSetColumns(ByVal ws As Worksheet, Optional ByVal colBase As Long = 0)
    ws.Columns(colBase + 1).ColumnWidth = 3
    ws.Range(ws.Columns(colBase + 2), ws.Columns(colBase + 19)).ColumnWidth = 14
    ws.Columns(colBase + 20).ColumnWidth = 3
End Sub

' Shifts an A1 address of a block ("B2:R3", "1:40", "B7", "A1:AZ43") by its origin: rows + rowBase,
' columns + colBase (columns past T are clamped to T so a block never paints its neighbour).
' Addresses containing "$" are already absolute (Range.Address) and are returned unchanged.
Public Function LcaShift(ByVal address As String, ByVal rowBase As Long, ByVal colBase As Long) As String
    Dim i As Long, ch As String, digits As String, letters As String, result As String
    If InStr(address, "$") > 0 Then
        LcaShift = address
        Exit Function
    End If
    For i = 1 To Len(address) + 1
        If i <= Len(address) Then ch = UCase$(Mid$(address, i, 1)) Else ch = vbNullString
        If ch >= "A" And ch <= "Z" And Len(ch) = 1 Then
            If Len(digits) > 0 Then
                result = result & CStr(CLng(digits) + rowBase)
                digits = vbNullString
            End If
            letters = letters & ch
        ElseIf ch >= "0" And ch <= "9" And Len(ch) = 1 Then
            If Len(letters) > 0 Then
                result = result & ShiftedColumn(letters, colBase)
                letters = vbNullString
            End If
            digits = digits & ch
        Else
            If Len(letters) > 0 Then result = result & ShiftedColumn(letters, colBase)
            If Len(digits) > 0 Then result = result & CStr(CLng(digits) + rowBase)
            letters = vbNullString
            digits = vbNullString
            result = result & ch
        End If
    Next i
    LcaShift = result
End Function

' Row-only shift kept for older callers (no column change, no clamping).
Public Function LcaRows(ByVal address As String, ByVal base As Long) As String
    Dim i As Long, ch As String, digits As String, result As String
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

Private Function ShiftedColumn(ByVal letters As String, ByVal colBase As Long) As String
    Dim n As Long, i As Long
    For i = 1 To Len(letters)
        n = n * 26 + (Asc(Mid$(letters, i, 1)) - 64)
    Next i
    If n > LCA_BLOCK_COLS Then n = LCA_BLOCK_COLS
    ShiftedColumn = ColumnName(n + colBase)
End Function

Private Function ColumnName(ByVal n As Long) As String
    Do While n > 0
        ColumnName = Chr$(65 + ((n - 1) Mod 26)) & ColumnName
        n = (n - 1) \ 26
    Loop
End Function

' Rows shared by blocks side by side keep the tallest height any block asks for.
Public Sub LcaRowHeight(ByVal rowsRange As Range, ByVal height As Double)
    Dim r As Range
    For Each r In rowsRange.Rows
        If r.RowHeight < height Then r.RowHeight = height
    Next r
End Sub

' Clears one block (cells, merges) and its shapes; the other blocks are untouched.
Public Sub LcaPrepareBand(ByVal ws As Worksheet, ByVal base As Long, ByVal lastRow As Long, _
                          ByVal shapePrefix As String, Optional ByVal colBase As Long = 0)
    Dim i As Long, band As Range
    Set band = ws.Range(LcaShift("A1:T" & CStr(lastRow), base, colBase))
    band.UnMerge
    band.Clear
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(shapePrefix)) = shapePrefix Then ws.Shapes(i).Delete
    Next i
    LcaSetColumns ws, colBase
End Sub

Public Sub LcaGoto(ByVal ws As Worksheet, ByVal base As Long, Optional ByVal colBase As Long = 0)
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.Zoom = 80
    ActiveWindow.ScrollRow = base + 1
    ActiveWindow.ScrollColumn = colBase + 1
    ws.Cells(base + 1, colBase + 1).Select
End Sub

' Navigation entry points used by the buttons of every block.
Public Sub LcaIrCpu()
    LcaGoto LcaSheet(True), LCA_BASE_CPU, LCA_COL_CPU
End Sub

Public Sub LcaIrDiagrama()
    LcaGoto LcaSheet(True), LCA_BASE_DIAGRAM, LCA_COL_DIAGRAM
End Sub

Public Sub LcaIrBuses()
    LcaGoto LcaSheet(True), LCA_BASE_BUS, LCA_COL_BUS
End Sub

Public Sub LcaIrCaches()
    LcaGoto LcaSheet(True), LCA_BASE_CACHE, LCA_COL_CACHE
End Sub

Public Sub LcaIrPaginas()
    LcaGoto LcaSheet(True), LCA_BASE_MMU, LCA_COL_MMU
End Sub

Public Sub LcaIrSwap()
    LcaGoto LcaSheet(True), LCA_BASE_PAGER, LCA_COL_PAGER
End Sub

Public Sub LcaIrCachePaginacion()
    LcaGoto LcaSheet(True), LCA_BASE_CACHE_PATH, LCA_COL_CACHE_PATH
End Sub

Public Sub LcaIrEditor()
    LcaGoto LcaSheet(True), LCA_BASE_EDITOR, LCA_COL_EDITOR
End Sub

Public Sub LcaIrTeclado()
    LcaGoto LcaSheet(True), LCA_BASE_KEYBOARD, LCA_COL_KEYBOARD
End Sub

Public Sub LcaIrAritmetica()
    LcaGoto LcaSheet(True), LCA_BASE_ARITH, LCA_COL_ARITH
    ArithPanelRefrescar
End Sub

Public Sub LcaIrEco()
    LcaGoto LcaSheet(True), LCA_BASE_ECHO, LCA_COL_ECHO
    EcoPanelRefrescar
End Sub

Public Sub LcaIrMemoria()
    LcaGoto LcaSheet(True), LCA_BASE_MEM, LCA_COL_MEM
    MemPanelRefrescar
End Sub

Public Sub LcaIrLcd()
    LcaGoto LcaSheet(True), LCA_BASE_LCD, LCA_COL_LCD
End Sub

Public Sub LcaIrMaquina()
    LcaGoto LcaSheet(True), LCA_BASE_MACHINE, LCA_COL_MACHINE
    MaquinaRefrescar
End Sub

' Zoomed-out view of the whole grid of blocks.
Public Sub LcaVistaGeneral()
    Dim ws As Worksheet
    Set ws = LcaSheet(True)
    ws.Activate
    ActiveWindow.Zoom = 25
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
End Sub
