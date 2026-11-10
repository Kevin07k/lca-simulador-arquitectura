Attribute VB_Name = "LcaKeyboardPanel"
Option Explicit

' P13 band of the LCA sheet: text input, the eight queue slots E0..E7h with head/tail, En_KBD bus
' cycle and the keyboard trace. Every button calls the real LcaKeyboard engine.
Private Const ROW_BASE As Long = LCA_BASE_KEYBOARD
Private Const COL_BASE As Long = LCA_COL_KEYBOARD
Private Const BAND_ROWS As Long = 36
Private mStatus As String

Public Sub ConstruirPanelTeclado()
    Dim ws As Worksheet
    Dim i As Long
    Set ws = KeyboardSheet(True)
    LcaPrepareBand ws, ROW_BASE, BAND_ROWS, "P13_", COL_BASE
    LcaRowHeight ws.Rows(RA("1:36")), 23
    ws.Range(RA("A1:AZ36")).Interior.Color = RGB(15, 23, 42)
    ws.Range(RA("A1:AZ36")).Font.Name = "Consolas"
    ws.Range(RA("A1:AZ36")).Font.Color = RGB(226, 232, 240)
    ws.Range(RA("B2:R3")).Merge
    With ws.Range(RA("B2"))
        .Value = "P13  /  TECLADO + COLA FIFO  /  KEYBOARD + FIFO QUEUE"
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 22
        .Font.Color = RGB(125, 211, 252)
    End With
    ws.Range(RA("B4:R4")).Merge
    ws.Range(RA("B4")).Value = "8 bytes circulares en RAM E0-E7h (PFN 7) | lleno = KBD_FULL | vacio = IN_EMPTY (espera)"
    AddKeyboardButton ws, "B6:E7", "REINICIAR / RESET", "TecladoReset"
    AddKeyboardButton ws, "F6:I7", "TECLEAR / TYPE", "TecladoEscribir"
    AddKeyboardButton ws, "J6:M7", "IN AX", "TecladoInAX"
    AddKeyboardButton ws, "N6:R7", "LLENAR / FILL (9)", "TecladoLlenar"
    ws.Range(RA("B9")).Value = "TEXTO / TEXT"
    ws.Range(RA("C9:H9")).Merge
    With ws.Range(RA("C9:H9"))
        .Interior.Color = RGB(230, 240, 250)
        .Font.Color = RGB(15, 23, 42)
        .NumberFormat = "@"
    End With
    ws.Range(RA("C9")).Value = "HOLA"
    ws.Range(RA("I9:R9")).Merge
    ws.Range(RA("I9")).Value = "<- escriba y pulse TECLEAR / type then press TYPE"
    ws.Range(RA("B11:R11")).Merge
    ws.Range(RA("B11")).Value = "COLA / QUEUE  (H = head: siguiente IN / next IN,  T = tail: siguiente tecla / next key)"
    ws.Range(RA("B12")).Value = "PA"
    ws.Range(RA("B13")).Value = "BYTE"
    ws.Range(RA("B14")).Value = "CHAR"
    ws.Range(RA("B15")).Value = "H / T"
    For i = 0 To LCA_KBD_SIZE - 1
        CL(ws, ROW_BASE + 12, 3 + i).Value = Right$("00" & Hex$(LCA_KBD_BASE_PA + i), 2) & "h"
    Next i
    ws.Range(RA("B12:J12")).Font.Bold = True
    ws.Range(RA("B12:J15")).HorizontalAlignment = xlCenter
    ws.Range(RA("C12:J15")).Borders.Color = RGB(51, 65, 85)
    ws.Range(RA("B17:R17")).Merge
    ws.Range(RA("B19:R19")).Merge
    ws.Range(RA("B19")).Value = "ULTIMO CICLO DE BUS / LAST BUS CYCLE  (IN: IO=1 R=1 En_KBD=1)"
    ws.Range(RA("B20:R21")).Merge
    ws.Range(RA("B22:R23")).Merge
    ws.Range(RA("B20:R23")).Interior.Color = RGB(30, 41, 59)
    ws.Range(RA("B20:R23")).WrapText = True
    ws.Range(RA("B25:R25")).Merge
    ws.Range(RA("B25")).Value = "TRAZA DEL TECLADO / KEYBOARD TRACE (ultimos 8 / last 8)"
    For i = 26 To 33
        ws.Range(CL(ws, ROW_BASE + i, 2), CL(ws, ROW_BASE + i, 18)).Merge
    Next i
    ws.Range(RA("B26:R33")).Interior.Color = RGB(30, 41, 59)
    ws.Range(RA("B35:R36")).Merge
    ws.Range(RA("B35:R36")).WrapText = True
    ws.Range(RA("B35")).Font.Color = RGB(125, 211, 252)
    TecladoReset
    LcaGoto ws, ROW_BASE, COL_BASE
End Sub

Public Sub MostrarPanelTeclado()
    Dim ws As Worksheet
    Set ws = KeyboardSheet(False)
    If ws Is Nothing Then
        ConstruirPanelTeclado
    Else
        LcaGoto ws, ROW_BASE, COL_BASE
        TecladoPanelActualizar
    End If
End Sub

Public Sub TecladoReset()
    On Error GoTo failed
    MemoryPathReset
    KeyboardReset
    mStatus = "Cola vacia. TECLEAR encola el texto; IN AX consume un byte; con la cola vacia IN espera."
    RefreshAll
    Exit Sub
failed:
    mStatus = "ERROR: " & Err.Description
    TecladoPanelActualizar
End Sub

Public Sub TecladoEscribir()
    Dim ws As Worksheet, text As String, accepted As Long
    On Error GoTo failed
    EnsureKeyboard
    Set ws = KeyboardSheet(False)
    If ws Is Nothing Then Exit Sub
    text = CStr(ws.Range(RA("C9")).Value)
    If Len(text) = 0 Then
        mStatus = "Escriba al menos un caracter / type at least one character."
    Else
        accepted = KeyboardPushText(text)
        mStatus = CStr(accepted) & " de / of " & CStr(Len(text)) & " teclas encoladas / keys queued"
        If accepted < Len(text) Then mStatus = mStatus & "; KBD_FULL rechazo / rejected " & CStr(Len(text) - accepted)
    End If
    RefreshAll
    Exit Sub
failed:
    mStatus = "ERROR: " & Err.Description
    TecladoPanelActualizar
End Sub

Public Sub TecladoInAX()
    Dim cpu As LcaCpuState
    On Error GoTo failed
    EnsureKeyboard
    If KeyboardIn(LcaAX) Then
        cpu = CpuSnapshot()
        mStatus = "IN AX: AX = " & Right$("00" & Hex$(cpu.AX), 2) & "h. El byte se consumio una sola vez."
    Else
        mStatus = "IN_EMPTY: espera. PC, AX y head no cambian; TECLEAR y vuelva a pulsar IN AX (reanudar)."
    End If
    RefreshAll
    Exit Sub
failed:
    mStatus = "ERROR: " & Err.Description
    TecladoPanelActualizar
End Sub

' Nine keys into an empty queue: eight are accepted and the ninth shows KBD_FULL.
Public Sub TecladoLlenar()
    Dim accepted As Long
    On Error GoTo failed
    MemoryPathReset
    KeyboardReset
    accepted = KeyboardPushText("ABCDEFGHI")
    mStatus = "ABCDEFGHI: " & CStr(accepted) & " aceptadas / accepted; 'I' -> KBD_FULL, la cola no cambia."
    RefreshAll
    Exit Sub
failed:
    mStatus = "ERROR: " & Err.Description
    TecladoPanelActualizar
End Sub

' Redraws the band from the engine; does nothing if the band is not built.
Public Sub TecladoPanelActualizar()
    Dim ws As Worksheet
    Dim i As Long, n As Long, value As Long, marker As String
    Dim externalCycle As LcaBusCycle
    Set ws = KeyboardSheet(False)
    If ws Is Nothing Then Exit Sub
    If Not KeyboardReady() Then KeyboardReset
    For i = 0 To LCA_KBD_SIZE - 1
        value = KeyboardSlot(i)
        CL(ws, ROW_BASE + 13, 3 + i).Value = Right$("00" & Hex$(value), 2) & "h"
        If value >= 32 And value <= 126 Then
            CL(ws, ROW_BASE + 14, 3 + i).Value = "'" & Chr$(value)
        Else
            CL(ws, ROW_BASE + 14, 3 + i).Value = "-"
        End If
        marker = vbNullString
        If i = KeyboardHead() Then marker = "H"
        If i = KeyboardTail() Then marker = marker & IIf(Len(marker) > 0, " ", vbNullString) & "T"
        CL(ws, ROW_BASE + 15, 3 + i).Value = marker
        If KeyboardSlotQueued(i) Then
            ws.Range(CL(ws, ROW_BASE + 13, 3 + i), CL(ws, ROW_BASE + 15, 3 + i)).Interior.Color = RGB(22, 101, 52)
        Else
            ws.Range(CL(ws, ROW_BASE + 13, 3 + i), CL(ws, ROW_BASE + 15, 3 + i)).Interior.Color = RGB(30, 41, 59)
        End If
    Next i
    ws.Range(RA("B17")).Value = "count=" & CStr(KeyboardCount()) & "/8   head=" & CStr(KeyboardHead()) & _
        "   tail=" & CStr(KeyboardTail()) & "   aceptadas=" & CStr(KeyboardAccepted()) & _
        "   rechazadas=" & CStr(KeyboardRejected()) & "   consumidas=" & CStr(KeyboardConsumed()) & _
        "   esperas=" & CStr(KeyboardWaits()) & "   estado=" & KeyboardLastStatus()
    ws.Range(RA("B20")).Value = ControlLastWord()
    externalCycle = BusSnapshot(LCA_BUS_EXTERNAL)
    ws.Range(RA("B22")).Value = BusFormat(externalCycle)
    n = KeyboardTraceCount()
    For i = 0 To 7
        CL(ws, ROW_BASE + 26 + i, 2).Value = vbNullString
        If n > i Then CL(ws, ROW_BASE + 26 + i, 2).Value = KeyboardTraceLine(i)
    Next i
    ws.Range(RA("B35")).Value = mStatus
End Sub

Private Sub RefreshAll()
    TecladoPanelActualizar
    PanelActualizar          ' CPU band shows the new AX
    BusPanelActualizar       ' bus band shows the IN cycle
End Sub

Private Sub EnsureKeyboard()
    If Not KeyboardReady() Then KeyboardReset
End Sub

Private Function KeyboardSheet(ByVal create As Boolean) As Worksheet
    Set KeyboardSheet = LcaBand("P13_TecladoReset", create)
End Function

' Row shift of this panel inside the shared LCA sheet.
Private Function RA(ByVal address As String) As String
    RA = LcaShift(address, ROW_BASE, COL_BASE)
End Function

Private Sub AddKeyboardButton(ByVal ws As Worksheet, ByVal address As String, _
                              ByVal caption As String, ByVal macroName As String)
    Dim area As Range
    Dim button As Shape
    Set area = ws.Range(RA(address))
    Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left, area.Top, area.Width - 4, area.Height - 2)
    button.Name = "P13_" & macroName
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
