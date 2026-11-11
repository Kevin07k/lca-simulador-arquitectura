Attribute VB_Name = "LcaLcdPanel"
Option Explicit

' P14 band of the LCA sheet: the 8x1 LCD drawn from RAM E8..EFh, the cursor, the raw video bytes,
' a RAM = screen check, the Ld_LCD bus cycle and the LCD trace. Every button calls the real engine.
Private Const ROW_BASE As Long = LCA_BASE_LCD
Private Const COL_BASE As Long = LCA_COL_LCD
Private Const BAND_ROWS As Long = 36
Private mStatus As String

Public Sub ConstruirPanelLcd()
    Dim ws As Worksheet
    Dim i As Long
    Set ws = LcdSheet(True)
    LcaPrepareBand ws, ROW_BASE, BAND_ROWS, "P14_", COL_BASE
    LcaRowHeight ws.Rows(RA("1:36")), 23
    ws.Range(RA("A1:AZ36")).Interior.Color = RGB(15, 23, 42)
    ws.Range(RA("A1:AZ36")).Font.Name = "Consolas"
    ws.Range(RA("A1:AZ36")).Font.Color = RGB(226, 232, 240)
    ws.Range(RA("B2:R3")).Merge
    With ws.Range(RA("B2"))
        .Value = "P14  /  LCD 8x1 + BUFFER DE VIDEO  /  VIDEO BUFFER"
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 22
        .Font.Color = RGB(125, 211, 252)
    End With
    ws.Range(RA("B4:R4")).Merge
    ws.Range(RA("B4")).Value = "8 bytes de video en RAM E8-EFh (PFN 7), despues de la cola E0-E7h | OUT R: IO=1 W=1 Ld_LCD=1 | cursor mod 8"
    AddLcdButton ws, "B6:E7", "REINICIAR / RESET", "LcdPanelReset"
    AddLcdButton ws, "F6:I7", "OUT AX (caracter)", "LcdPanelOut"
    AddLcdButton ws, "J6:M7", "ECO: IN AX + OUT AX", "LcdPanelEco"
    AddLcdButton ws, "N6:R7", "9 CARACTERES / CHARS", "LcdPanelNueve"
    ws.Range(RA("B9")).Value = "CARACTER"
    ws.Range(RA("C9:D9")).Merge
    With ws.Range(RA("C9:D9"))
        .Interior.Color = RGB(230, 240, 250)
        .Font.Color = RGB(15, 23, 42)
        .NumberFormat = "@"
        .HorizontalAlignment = xlCenter
    End With
    ws.Range(RA("C9")).Value = "A"
    ws.Range(RA("E9:R9")).Merge
    ws.Range(RA("E9")).Value = "<- AX = este caracter y OUT AX / AX = this char, then OUT AX.  ECO usa la cola del teclado (P13)."
    ws.Range(RA("B11:R11")).Merge
    ws.Range(RA("B11")).Value = "PANTALLA / SCREEN  (leida de la RAM / read from RAM;  ^ = cursor: siguiente OUT / next OUT)"
    LcaRowHeight ws.Rows(ROW_BASE + 12), 46
    ws.Range(RA("B12:K12")).Interior.Color = RGB(63, 98, 18)
    For i = 0 To LCA_LCD_SIZE - 1
        With CL(ws, ROW_BASE + 12, 3 + i)
            .Interior.Color = RGB(163, 196, 66)
            .Font.Color = RGB(20, 46, 16)
            .Font.Size = 30
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .NumberFormat = "@"
        End With
        CL(ws, ROW_BASE + 13, 3 + i).HorizontalAlignment = xlCenter
        CL(ws, ROW_BASE + 13, 3 + i).Font.Color = RGB(250, 204, 21)
        CL(ws, ROW_BASE + 15, 3 + i).Value = Right$("00" & Hex$(LCA_LCD_BASE_PA + i), 2) & "h"
    Next i
    ws.Range(RA("C12:J12")).Borders.Color = RGB(63, 98, 18)
    ws.Range(RA("B15")).Value = "RAM PA"
    ws.Range(RA("B16")).Value = "BYTE"
    ws.Range(RA("B15:J16")).HorizontalAlignment = xlCenter
    ws.Range(RA("B15:J15")).Font.Bold = True
    ws.Range(RA("C15:J16")).Borders.Color = RGB(51, 65, 85)
    ws.Range(RA("B18:R18")).Merge
    ws.Range(RA("B19:R19")).Merge
    ws.Range(RA("B21:R21")).Merge
    ws.Range(RA("B21")).Value = "ULTIMO CICLO DE BUS / LAST BUS CYCLE  (OUT: IO=1 W=1 Ld_LCD=1)"
    ws.Range(RA("B22:R23")).Merge
    ws.Range(RA("B24:R25")).Merge
    ws.Range(RA("B22:R25")).Interior.Color = RGB(30, 41, 59)
    ws.Range(RA("B22:R25")).WrapText = True
    ws.Range(RA("B27:R27")).Merge
    ws.Range(RA("B27")).Value = "TRAZA DEL LCD / LCD TRACE (ultimos 6 / last 6)"
    For i = 28 To 33
        ws.Range(CL(ws, ROW_BASE + i, 2), CL(ws, ROW_BASE + i, 18)).Merge
    Next i
    ws.Range(RA("B28:R33")).Interior.Color = RGB(30, 41, 59)
    ws.Range(RA("B35:R36")).Merge
    ws.Range(RA("B35:R36")).WrapText = True
    ws.Range(RA("B35")).Font.Color = RGB(125, 211, 252)
    LcdPanelReset
    LcaGoto ws, ROW_BASE, COL_BASE
End Sub

Public Sub MostrarPanelLcd()
    Dim ws As Worksheet
    Set ws = LcdSheet(False)
    If ws Is Nothing Then
        ConstruirPanelLcd
    Else
        LcaGoto ws, ROW_BASE, COL_BASE
        LcdPanelActualizar
    End If
End Sub

Public Sub LcdPanelReset()
    On Error GoTo failed
    MemoryPathReset
    LcdReset
    mStatus = "LCD = 8 espacios, cursor 0. OUT AX escribe el caracter; ECO lee el teclado y lo muestra."
    RefreshAll
    Exit Sub
failed:
    mStatus = "ERROR: " & Err.Description
    LcdPanelActualizar
End Sub

' AX = character typed in the CHAR cell, then OUT AX.
Public Sub LcdPanelOut()
    Dim ws As Worksheet, text As String, code As Long, cell As Long
    On Error GoTo failed
    EnsureLcd
    Set ws = LcdSheet(False)
    If ws Is Nothing Then Exit Sub
    text = CStr(ws.Range(RA("C9")).Value)
    If Len(text) = 0 Then
        mStatus = "Escriba un caracter / type one character."
    Else
        code = AscW(Left$(text, 1))
        If code < 0 Or code > 255 Then Err.Raise LCA_ERR_LCD, "LcaLcdPanel", "LCD_NOT_ASCII"
        CpuWriteRegister LcaAX, code
        cell = LcdOut(LcaAX)
        AnimarOut ws, Left$(text, 1), cell
        mStatus = "OUT AX: " & Right$("00" & Hex$(code), 2) & "h en la celda / in cell " & CStr(cell) & _
                  "; cursor = " & CStr(LcdCursor()) & "."
    End If
    RefreshAll
    Exit Sub
failed:
    mStatus = "ERROR: " & Err.Description
    LcdPanelActualizar
End Sub

' P16 preview: IN AX from the P13 queue, then OUT AX. An empty queue is a wait, nothing is written.
Public Sub LcdPanelEco()
    Dim cell As Long
    On Error GoTo failed
    EnsureLcd
    If Not KeyboardReady() Then KeyboardReset
    If KeyboardIn(LcaAX) Then
        cell = LcdOut(LcaAX)
        mStatus = "ECO: IN AX = " & Right$("00" & Hex$(CpuReadRegister(LcaAX)), 2) & "h -> OUT AX en la celda " & _
                  CStr(cell) & ". Teclas restantes en la cola: " & CStr(KeyboardCount()) & "."
    Else
        mStatus = "ECO: IN_EMPTY, la cola del teclado esta vacia; no se escribe en el LCD. Teclee en la banda P13."
    End If
    RefreshAll
    Exit Sub
failed:
    mStatus = "ERROR: " & Err.Description
    LcdPanelActualizar
End Sub

' Nine characters after a reset: the eighth fills the last cell and the ninth replaces the first.
Public Sub LcdPanelNueve()
    Dim text As String, i As Long
    On Error GoTo failed
    MemoryPathReset
    LcdReset
    text = "HOLA LCD!"
    For i = 1 To Len(text)
        CpuWriteRegister LcaAX, Asc(Mid$(text, i, 1))
        Call LcdOut(LcaAX)
    Next i
    mStatus = "9 OUT de 'HOLA LCD!': el 8.o llena la ultima celda, el 9.o ('!') reemplaza la primera -> '" & _
              LcdText() & "', cursor = " & CStr(LcdCursor()) & "."
    RefreshAll
    Exit Sub
failed:
    mStatus = "ERROR: " & Err.Description
    LcdPanelActualizar
End Sub

' Redraws the band from RAM; does nothing if the band is not built.
Public Sub LcdPanelActualizar()
    Dim ws As Worksheet
    Dim i As Long, n As Long
    Dim externalCycle As LcaBusCycle
    Set ws = LcdSheet(False)
    If ws Is Nothing Then Exit Sub
    If Not LcdReady() Then LcdReset
    For i = 0 To LCA_LCD_SIZE - 1
        ' A leading apostrophe is Excel's text prefix, so it must be doubled to be displayed.
        If LcdChar(i) = "'" Then
            CL(ws, ROW_BASE + 12, 3 + i).Value = "''"
        Else
            CL(ws, ROW_BASE + 12, 3 + i).Value = LcdChar(i)
        End If
        CL(ws, ROW_BASE + 13, 3 + i).Value = IIf(i = LcdCursor(), "^", vbNullString)
        CL(ws, ROW_BASE + 16, 3 + i).Value = Right$("00" & Hex$(LcdCell(i)), 2) & "h"
    Next i
    ws.Range(RA("B18")).Value = "cursor=" & CStr(LcdCursor()) & "   OUT=" & CStr(LcdWrites()) & _
        "   vueltas / wraps=" & CStr(LcdWraps()) & "   texto / text='" & LcdText() & "'"
    If LcdPanelCoherente() Then
        ws.Range(RA("B19")).Value = "RAM E8-EFh = pantalla / screen: OK"
        ws.Range(RA("B19")).Font.Color = RGB(134, 239, 172)
    Else
        ws.Range(RA("B19")).Value = "RAM E8-EFh <> pantalla / screen: ERROR"
        ws.Range(RA("B19")).Font.Color = RGB(248, 113, 113)
    End If
    ws.Range(RA("B22")).Value = ControlLastWord()
    externalCycle = BusSnapshot(LCA_BUS_EXTERNAL)
    ws.Range(RA("B24")).Value = BusFormat(externalCycle)
    n = LcdTraceCount()
    For i = 0 To 5
        CL(ws, ROW_BASE + 28 + i, 2).Value = vbNullString
        If n > i Then CL(ws, ROW_BASE + 28 + i, 2).Value = LcdTraceLine(i)
    Next i
    ws.Range(RA("B35")).Value = mStatus
End Sub

' True when every drawn cell (character and byte) matches RAM E8..EFh. True if the band is not built.
Public Function LcdPanelCoherente() As Boolean
    Dim ws As Worksheet, i As Long
    Set ws = LcdSheet(False)
    LcdPanelCoherente = True
    If ws Is Nothing Then Exit Function
    For i = 0 To LCA_LCD_SIZE - 1
        If CStr(CL(ws, ROW_BASE + 12, 3 + i).Value) <> LcdDisplayChar(RamReadIO(LCA_LCD_BASE_PA + i)) Then
            LcdPanelCoherente = False
        ElseIf CStr(CL(ws, ROW_BASE + 16, 3 + i).Value) <> Right$("00" & Hex$(RamReadIO(LCA_LCD_BASE_PA + i)), 2) & "h" Then
            LcdPanelCoherente = False
        End If
    Next i
End Function

' P18: this block, the CPU, the buses, the machine and every other view show the new state.
Private Sub RefreshAll()
    Call LcaSincronizar
End Sub

Private Sub EnsureLcd()
    If Not LcdReady() Then LcdReset
End Sub

' OUT: the character leaves the input box, crosses the bus (Ld_LCD) and lights its screen cell.
Private Sub AnimarOut(ByVal ws As Worksheet, ByVal ch As String, ByVal cell As Long)
    Dim target As Range
    If ws Is Nothing Then Exit Sub
    If Not FlowOn() Then Exit Sub
    Set target = CL(ws, ROW_BASE + 12, 3 + cell)
    Application.ScreenUpdating = True
    FlowSend ws, ch, FlowPacketColor(), ws.Range(RA("C9:D9")), ws.Range(RA("B22:R23")), target
    FlowPulseRange target, RGB(190, 242, 100), 2
End Sub

Private Function LcdSheet(ByVal create As Boolean) As Worksheet
    Set LcdSheet = LcaBand("P14_LcdPanelReset", create)
End Function

' Row shift of this panel inside the shared LCA sheet.
Private Function RA(ByVal address As String) As String
    RA = LcaShift(address, ROW_BASE, COL_BASE)
End Function

Private Sub AddLcdButton(ByVal ws As Worksheet, ByVal address As String, _
                         ByVal caption As String, ByVal macroName As String)
    Dim area As Range
    Dim button As Shape
    Set area = ws.Range(RA(address))
    Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left, area.Top, area.Width - 4, area.Height - 2)
    button.Name = "P14_" & macroName
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

' Cell of this block: row is absolute (ROW_BASE already added), column is relative to the block.
Private Function CL(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Range
    Set CL = ws.Cells(r, c + COL_BASE)
End Function
