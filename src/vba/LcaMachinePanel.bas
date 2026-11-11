Attribute VB_Name = "LcaMachinePanel"
Option Explicit

' P12 block of the LCA dashboard: the whole machine drawn with shapes, animated and clickable.
'   CPU (registers, PC/IR, MAR/MDR, flags, ALU, control unit with FETCH/DECODE/EXECUTE/STORE stages)
'   -> MMU (16-page grid) -> L1I/L1D/L2 lines -> RAM (8 frames) <-> SWAP disk, external bus with live
'   signal lamps, keyboard queue with on-screen keys and an 8-cell LCD with cursor.
' Every PASO animates the instruction: the stages light up in order and a data packet travels along the
' real path (fetch, fill, page fault, LD/ST, IN, OUT). Clicking a component opens its detail block.
Private Const ROW_BASE As Long = LCA_BASE_MACHINE
Private Const COL_BASE As Long = LCA_COL_MACHINE
Private Const BAND_ROWS As Long = 68
Private Const DIAGRAM_ROW As Long = 12
Private Const DESIGN_W As Double = 1380
Private Const LIST_GAP As Double = 700      ' listing starts this many points below the diagram origin
Private Const LIST_ROWS As Long = 16

Private Const BG As Long = 2758415          ' RGB(15, 23, 42)
Private Const CARD As Long = 3877150        ' RGB(30, 41, 59)
Private Const HEAD As Long = 6897694        ' RGB(30, 64, 105)
Private Const INK As Long = 15788258        ' RGB(226, 232, 240)
Private Const MUTED As Long = 12100500      ' RGB(148, 163, 184)
Private Const ACCENT As Long = 16569213     ' RGB(125, 211, 252)
Private Const BUTTON As Long = 15426341     ' RGB(37, 99, 235)
Private Const OK_GREEN As Long = 3433750    ' RGB(22, 101, 52)
Private Const WARN As Long = 1592985        ' RGB(153, 78, 24)
Private Const BAD As Long = 1908095         ' RGB(127, 29, 29)
Private Const EDGE As Long = 5587251        ' RGB(51, 65, 85)
Private Const CODE_FILL As Long = 9116461   ' RGB(45, 27, 139)  violet: code pages
Private Const DATA_FILL As Long = 7954958   ' RGB(14, 98, 121)  teal: data pages
Private Const IO_FILL As Long = 3747162     ' RGB(90, 45, 57)
Private Const LINE_ON As Long = 8508206     ' RGB(46, 211, 129) valid cache line
Private Const LAMP_ON As Long = 2477306     ' RGB(250, 204, 37) amber
Private Const PINK As Long = 10040319       ' RGB(255, 51, 153) data packet
Private Const LCD_BG As Long = 2384672      ' RGB(32, 99, 36)
Private Const LCD_LIT As Long = 4441661     ' RGB(61, 197, 67)
Private Const LCD_INK As Long = 9502671     ' RGB(207, 255, 144)

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private mX0 As Double
Private mY0 As Double
Private mSX As Double
Private mMessage As String
Private mSpeed As Long          ' 0 = no animation, 1 slow, 2 normal, 3 fast
Private mSpeedSet As Boolean
Private mFills0 As Long
Private mFaults0 As Long

' --- Build ---------------------------------------------------------------------------------------

Public Sub ConstruirPanelMaquina()
    Dim ws As Worksheet, i As Long, previous As Boolean, lr As Long
    previous = Application.ScreenUpdating
    Application.ScreenUpdating = False
    EnsureSpeed
    Set ws = MachineSheet(True)
    LcaPrepareBand ws, ROW_BASE, BAND_ROWS, "P12_", COL_BASE
    LcaRowHeight ws.Rows(RA("1:" & CStr(BAND_ROWS))), 18
    With ws.Range(RA("A1:AZ" & CStr(BAND_ROWS)))
        .Interior.Color = BG
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Color = INK
        .VerticalAlignment = xlCenter
    End With

    ' Header
    LcaRowHeight ws.Rows(ROW_BASE + 2), 34
    ws.Range(RA("B2:L2")).Merge
    With ws.Range(RA("B2"))
        .Value = "LCA  |  MAQUINA COMPLETA EN EJECUCION"
        .Font.Size = 24
        .Font.Bold = True
        .Font.Color = ACCENT
    End With
    ws.Range(RA("B3:S3")).Merge
    With ws.Range(RA("B3"))
        .Value = "P12  FETCH > DECODE > EXECUTE > STORE  |  PASO anima cada instruccion: el punto rosa es el dato viajando. " & _
                 "Haz clic en cualquier componente (CPU, MMU, caches, RAM, swap, bus, teclado, LCD) para ver su detalle."
        .Font.Color = MUTED
    End With
    NavBar ws

    ' Controls
    Btn ws, "B7:C8", "CARGAR EDITOR", "MaquinaCargarEditor", HEAD
    Btn ws, "D7:E8", "EJEMPLO SUMA", "MaquinaEjemplo", HEAD
    Btn ws, "F7:G8", "DEMO ECO", "MaquinaEco", HEAD
    Btn ws, "H7:I8", "PASO / STEP", "MaquinaPaso", BUTTON
    Btn ws, "J7:K8", "EJECUTAR / RUN", "MaquinaEjecutar", OK_GREEN
    Btn ws, "L7:M8", "RESET", "MaquinaReset", EDGE
    Btn ws, "N7:O8", "VELOCIDAD", "MaquinaVelocidad", CARD
    ws.Range(RA("P7:S8")).Merge
    With ws.Range(RA("P7"))
        .HorizontalAlignment = xlCenter
        .Font.Size = 15
        .Font.Bold = True
    End With
    ws.Range(RA("B10")).Value = "TECLAS:"
    ws.Range(RA("B10")).Font.Bold = True
    ws.Range(RA("B10")).Font.Color = ACCENT
    ws.Range(RA("C10:D10")).Merge
    With ws.Range(RA("C10:D10"))
        .NumberFormat = "@"
        .Interior.Color = RGB(230, 240, 250)
        .Font.Color = BG
        .Font.Name = "Consolas"
        .Value = "HOLA"
        .Locked = False
    End With
    Btn ws, "E10:E10", "TECLEAR", "MaquinaTeclear", HEAD
    ws.Range(RA("F10:S10")).Merge
    With ws.Range(RA("F10"))
        .Font.Name = "Consolas"
        .Font.Size = 12
        .Font.Bold = True
        .Interior.Color = CARD
        .IndentLevel = 1
    End With

    ' Diagram
    mX0 = CL(ws, ROW_BASE + DIAGRAM_ROW, 2).Left
    mY0 = CL(ws, ROW_BASE + DIAGRAM_ROW, 2).Top
    mSX = (CL(ws, ROW_BASE + DIAGRAM_ROW, 20).Left - mX0) / DESIGN_W
    DrawCpu ws
    DrawMmu ws
    DrawCaches ws
    DrawRam ws
    DrawSideDevices ws
    DrawBus ws
    DrawArrows ws
    ws.Shapes("P12_packet").Visible = msoFalse

    ' Listing and trace: first rows below the drawing, whatever the row heights shared with other blocks.
    lr = ListRow(ws)
    Title ws, "B" & CStr(lr - 1) & ":H" & CStr(lr - 1), "PROGRAMA CARGADO  (fila resaltada = PC)"
    Title ws, "J" & CStr(lr - 1) & ":S" & CStr(lr - 1), "INDICADORES EN VIVO  (la traza completa de cada instruccion esta en la hoja Logs)"
    For i = 0 To LIST_ROWS - 1
        ws.Range(CL(ws, ROW_BASE + lr + i, 4), CL(ws, ROW_BASE + lr + i, 8)).Merge
    Next i
    DrawIndicators ws, lr
    With ws.Range(CL(ws, ROW_BASE + lr, 2), CL(ws, ROW_BASE + lr + LIST_ROWS - 1, 19))
        .Font.Name = "Consolas"
        .Font.Size = 9
    End With

    mMessage = vbNullString
    If Not LoaderLoaded() Then
        MaquinaEjemplo
    Else
        MaquinaRefrescar
    End If
    Application.ScreenUpdating = previous
End Sub

Public Sub MostrarPanelMaquina()
    If MachineSheet(False) Is Nothing Then
        ConstruirPanelMaquina
    Else
        LcaGoto MachineSheet(False), ROW_BASE, COL_BASE
        MaquinaRefrescar
    End If
End Sub

Private Sub NavBar(ByVal ws As Worksheet)
    Dim labels As Variant, macros As Variant, i As Long, x As Double, w As Double, top As Double, n As Long
    labels = Array("PORTADA", "VISTA GENERAL", "CPU", "DIAGRAMA", "BUSES", "PAGINAS", "SWAP", "CACHES", "CACHE+PAG", _
                   "EDITOR", "TECLADO", "LCD", "DEMOS P15", "ECO P16", "LOGS")
    macros = Array("LcaIrPortada", "LcaVistaGeneral", "LcaIrCpu", "LcaIrDiagrama", "LcaIrBuses", "LcaIrPaginas", "LcaIrSwap", _
                   "LcaIrCaches", "LcaIrCachePaginacion", "LcaIrEditor", "LcaIrTeclado", "LcaIrLcd", "LcaIrAritmetica", "LcaIrEco", "LcaIrLogs")
    n = UBound(labels) + 1
    x = CL(ws, ROW_BASE + 5, 2).Left
    w = (CL(ws, ROW_BASE + 5, 20).Left - x) / n
    top = CL(ws, ROW_BASE + 5, 2).Top
    ws.Range(RA("B4")).Value = "IR A / GO TO:"
    ws.Range(RA("B4")).Font.Color = MUTED
    ws.Range(RA("B4")).Font.Size = 8
    For i = 0 To n - 1
        MakeButton ws, "nav_" & CStr(i), x + i * w + 2, top, w - 4, 20, CStr(labels(i)), CStr(macros(i)), _
                   IIf(i = 0, HEAD, CARD), 9
    Next i
End Sub

Private Sub DrawCpu(ByVal ws As Worksheet)
    Dim i As Long, names As Variant, stages As Variant
    Frame ws, "cpu", 0, 0, 370, 440, "CPU  (clic: panel P04)", "LcaIrCpu"
    names = Array("AX", "BX", "CX", "DX")
    For i = 0 To 3
        Box ws, "reg" & CStr(i), 12 + (i Mod 2) * 176, 34 + (i \ 2) * 70, 168, 62, CARD, CStr(names(i)), 15, True
    Next i
    Box ws, "pc", 12, 180, 110, 52, HEAD, "PC", 12, True
    Box ws, "ir", 128, 180, 110, 52, HEAD, "IR", 12, True
    Box ws, "flagZ", 248, 186, 36, 36, CARD, "Z", 11, True, msoShapeOval
    Box ws, "flagC", 288, 186, 36, 36, CARD, "C", 11, True, msoShapeOval
    Box ws, "flagS", 328, 186, 36, 36, CARD, "S", 11, True, msoShapeOval
    Box ws, "marv", 12, 242, 110, 44, CARD, "MAR_V", 9
    Box ws, "marp", 128, 242, 110, 44, CARD, "MAR_P", 9
    Box ws, "mdr", 244, 242, 114, 44, CARD, "MDR", 9
    Box ws, "alu", 60, 296, 250, 56, EDGE, "ALU", 11, True, msoShapeFlowchartManualOperation
    ws.Shapes("P12_alu").OnAction = MacroRef("LcaIrDiagrama")
    Box ws, "ctrl", 12, 360, 346, 30, CARD, "UNIDAD DE CONTROL", 9
    stages = Array("FETCH", "DECODE", "EXECUTE", "STORE")
    For i = 0 To 3
        Box ws, "stg" & CStr(i), 12 + i * 87, 398, 82, 30, CARD, CStr(stages(i)), 9, True
    Next i
End Sub

Private Sub DrawMmu(ByVal ws As Worksheet)
    Dim v As Long
    Frame ws, "mmu", 400, 0, 260, 330, "MMU  |  TABLA DE PAGINAS", "LcaIrPaginas"
    For v = 0 To 15
        Box ws, "pte" & CStr(v), 408 + (v Mod 4) * 62, 34 + (v \ 4) * 70, 58, 64, CARD, "VPN " & CStr(v), 9
    Next v
    Box ws, "xlat", 400, 340, 260, 60, CARD, "VA -> PA", 11, True
End Sub

Private Sub DrawCaches(ByVal ws As Worksheet)
    Dim i As Long
    Frame ws, "cache", 690, 0, 300, 430, "CACHES  (lineas de 4 B)", "LcaIrCaches"
    Frame ws, "l1i", 700, 30, 136, 120, "L1I", "LcaIrCaches"
    Frame ws, "l1d", 844, 30, 136, 120, "L1D", "LcaIrCaches"
    For i = 0 To 1
        Box ws, "l1i_" & CStr(i), 708, 66 + i * 40, 120, 34, CARD, "-", 8
        Box ws, "l1d_" & CStr(i), 852, 66 + i * 40, 120, 34, CARD, "-", 8
    Next i
    Frame ws, "l2", 700, 160, 280, 262, "L2 unificada", "LcaIrCachePaginacion"
    For i = 0 To 7
        Box ws, "l2_" & CStr(i), 708 + (i Mod 2) * 136, 194 + (i \ 2) * 56, 128, 50, CARD, "-", 8
    Next i
End Sub

Private Sub DrawRam(ByVal ws As Worksheet)
    Dim p As Long
    Frame ws, "ram", 1020, 0, 200, 430, "RAM  |  8 MARCOS x 32 B", "LcaIrSwap"
    For p = 0 To 7
        Box ws, "pfn" & CStr(p), 1028, 30 + p * 50, 184, 46, CARD, "PFN " & CStr(p), 9
    Next p
End Sub

Private Sub DrawSideDevices(ByVal ws As Worksheet)
    Dim i As Long, keys As Variant, k As String
    Box ws, "swap", 1240, 0, 140, 150, HEAD, "SWAP", 11, True, msoShapeCan
    ws.Shapes("P12_swap").OnAction = MacroRef("LcaIrSwap")
    Frame ws, "legend", 1240, 170, 140, 260, "LEYENDA", ""
    Box ws, "lg0", 1248, 200, 124, 26, CODE_FILL, "pagina de codigo", 8
    Box ws, "lg1", 1248, 230, 124, 26, DATA_FILL, "pagina de datos", 8
    Box ws, "lg2", 1248, 260, 124, 26, IO_FILL, "E/S reservada", 8
    Box ws, "lg3", 1248, 290, 124, 26, LINE_ON, "linea valida", 8
    Box ws, "lg4", 1248, 320, 124, 26, WARN, "fallo / miss", 8
    Box ws, "lg5", 1248, 350, 124, 26, PINK, "dato viajando", 8
    Box ws, "lg6", 1248, 380, 124, 40, CARD, "amarillo = camino" & vbCr & "o senal activa", 8

    ' Keyboard: queue slots, status and clickable keys.
    Frame ws, "kbd", 700, 520, 300, 160, "TECLADO  |  cola FIFO E0-E7h", "LcaIrTeclado"
    For i = 0 To 7
        Box ws, "kq" & CStr(i), 708 + i * 36, 550, 32, 32, CARD, "", 10, True, msoShapeRectangle
    Next i
    Box ws, "kinfo", 708, 586, 284, 22, CARD, "", 8, False, msoShapeRectangle
    keys = Array("H", "O", "L", "A", "E", "S", "I", " ", "0", "1", "2", "3", "4", "5", "6", "7")
    For i = 0 To 15
        k = CStr(keys(i))
        Box ws, "key_" & Right$("0" & Hex$(Asc(k)), 2), 708 + (i Mod 8) * 36, 614 + (i \ 8) * 32, 32, 28, HEAD, _
            IIf(k = " ", "_", k), 10, True
        ws.Shapes("P12_key_" & Right$("0" & Hex$(Asc(k)), 2)).OnAction = MacroRef("MaquinaTecla")
        ws.Shapes("P12_key_" & Right$("0" & Hex$(Asc(k)), 2)).Line.ForeColor.RGB = ACCENT
    Next i

    ' LCD: eight character cells, a cursor bar and status.
    Frame ws, "lcdf", 1020, 520, 360, 160, "LCD 8x1  |  video E8-EFh", "LcaIrLcd"
    Box ws, "lcdglass", 1028, 546, 344, 72, RGB(20, 60, 24), "", 8, False, msoShapeRectangle
    For i = 0 To 7
        Box ws, "lcd" & CStr(i), 1034 + i * 42, 552, 38, 54, LCD_BG, " ", 24, True, msoShapeRectangle
        ws.Shapes("P12_lcd" & CStr(i)).TextFrame2.TextRange.Font.Name = "Consolas"
        ws.Shapes("P12_lcd" & CStr(i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = LCD_INK
        ws.Shapes("P12_lcd" & CStr(i)).Line.ForeColor.RGB = RGB(20, 60, 24)
    Next i
    Box ws, "lcdcur", 1036, 609, 34, 5, LCD_INK, "", 6, False, msoShapeRectangle
    Box ws, "lcdinfo", 1028, 628, 344, 44, CARD, "", 8, False, msoShapeRectangle
End Sub

Private Sub DrawBus(ByVal ws As Worksheet)
    Dim i As Long, labels As Variant
    Box ws, "bus", 0, 462, 1380, 30, HEAD, "BUS EXTERNO  |  direcciones  -  datos  -  control   (clic: P05/P06)", 10, True, msoShapeRectangle
    ws.Shapes("P12_bus").OnAction = MacroRef("LcaIrBuses")
    labels = Array("R", "W", "C", "IO", "EN", "LD", "KBD", "LCD")
    For i = 0 To 7
        Box ws, "lamp" & CStr(i), 6 + i * 45, 512, 38, 38, CARD, CStr(labels(i)), 8, True, msoShapeOval
    Next i
    Box ws, "cycle", 0, 562, 680, 56, CARD, "", 8, False, msoShapeRectangle
    Box ws, "speed", 0, 626, 680, 50, CARD, "", 9, False, msoShapeRectangle
    ' The data packet that travels during the animation (kept hidden otherwise).
    Box ws, "packet", 0, 0, 22, 22, PINK, "", 6, False, msoShapeOval
    With ws.Shapes("P12_packet")
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Weight = 2
        .Glow.Color.RGB = PINK
        .Glow.Radius = 10
    End With
End Sub

Private Sub DrawArrows(ByVal ws As Worksheet)
    Arrow ws, "a_va", 370, 120, 400, 120, False
    Arrow ws, "a_pa", 660, 120, 690, 120, False
    Arrow ws, "a_fill", 1020, 120, 990, 120, False
    Arrow ws, "a_swap", 1240, 75, 1220, 75, True
    Arrow ws, "a_cpu_bus", 185, 440, 185, 462, True
    Arrow ws, "a_ram_bus", 1120, 430, 1120, 462, True
    Arrow ws, "a_kbd", 850, 520, 850, 492, False
    Arrow ws, "a_lcd", 1200, 492, 1200, 520, False
End Sub

' --- Actions ---------------------------------------------------------------------------------

Public Sub MaquinaEjemplo()
    LoadProgram LoaderSampleProgram(), "Ejemplo SUMA cargado: AX = 5+4+3+2+1 con un bucle JNZ, luego ST/LD en 100h. Pulsa PASO."
End Sub

Public Sub MaquinaEco()
    LoadProgram Array("; Demo ECO: cada tecla se muestra en el LCD", "eco:  IN AX       ; espera una tecla", _
                      "      OUT AX      ; la escribe en el LCD", "      JMP eco"), _
                "Demo ECO cargada: pulsa teclas en el teclado de abajo (o TECLEAR) y luego EJECUTAR."
End Sub

Public Sub MaquinaCargarEditor()
    If LoaderLoadSource(EditorRead()) Then
        mMessage = "Programa del editor (P11) cargado: " & CStr(LoaderInstructionCount()) & " instrucciones."
    Else
        mMessage = "El programa del editor tiene " & CStr(AsmErrorCount()) & " error(es): revisalos en EDITOR."
    End If
    MaquinaRefrescar
End Sub

Public Sub MaquinaPaso()
    EnsureSpeed
    Dim ws As Worksheet
    On Error GoTo failed
    If Not LoaderLoaded() Then MaquinaEjemplo
    Set ws = MachineSheet(False)
    BeginStepCapture
    Call ExecStep
    mMessage = vbNullString
    MaquinaRefrescar
    If mSpeed > 0 And Not ws Is Nothing Then AnimateLastStep ws
    AfterAction
    Exit Sub
failed:
    mMessage = "ERROR: " & Err.Description
    MaquinaRefrescar
End Sub

' With animation: every instruction is drawn (fast). Without: the engine runs to the end at once.
Public Sub MaquinaEjecutar()
    EnsureSpeed
    Dim ws As Worksheet, n As Long, keepSpeed As Long
    On Error GoTo failed
    If Not LoaderLoaded() Then MaquinaEjemplo
    Set ws = MachineSheet(False)
    If mSpeed = 0 Or ws Is Nothing Then
        n = ExecRun(500)
    Else
        keepSpeed = mSpeed
        If mSpeed < 3 Then mSpeed = mSpeed + 1
        Do While n < 300
            BeginStepCapture
            If Not ExecStep() Then Exit Do
            n = n + 1
            MaquinaRefrescar
            AnimateLastStep ws
            If ExecStatus() = EXEC_HALTED Then Exit Do
        Loop
        mSpeed = keepSpeed
    End If
    mMessage = CStr(n) & " instrucciones ejecutadas."
    AfterAction
    MaquinaRefrescar
    Exit Sub
failed:
    mMessage = "ERROR: " & Err.Description
    MaquinaRefrescar
End Sub

Public Sub MaquinaReset()
    On Error GoTo failed
    If Not LoaderLoaded() Then MaquinaEjemplo
    ExecReset
    mMessage = "Reset: CPU, buses, RAM, caches, tabla de paginas, teclado y LCD reiniciados; PC = 000h."
    AfterAction
    MaquinaRefrescar
    Exit Sub
failed:
    mMessage = "ERROR: " & Err.Description
    MaquinaRefrescar
End Sub

Public Sub MaquinaTeclear()
    Dim ws As Worksheet, text As String, n As Long
    On Error GoTo failed
    If Not LoaderLoaded() Then MaquinaEjemplo
    Set ws = MachineSheet(False)
    text = CStr(ws.Range(RA("C10")).Value)
    n = KeyboardPushText(text)
    mMessage = CStr(n) & " tecla(s) en la cola" & IIf(n < Len(text), " (" & CStr(Len(text) - n) & " rechazada(s): KBD_FULL)", "") & "."
    AfterAction
    MaquinaRefrescar
    Exit Sub
failed:
    mMessage = "ERROR: " & Err.Description
    MaquinaRefrescar
End Sub

' On-screen key: the shape name carries the ASCII code (P12_key_41 = "A").
Public Sub MaquinaTecla()
    EnsureSpeed
    Dim ws As Worksheet, nm As String, code As Long, accepted As Boolean
    On Error GoTo failed
    If Not LoaderLoaded() Then MaquinaEjemplo
    Set ws = MachineSheet(False)
    nm = CStr(Application.Caller)
    code = CLng("&H" & Right$(nm, 2))
    accepted = KeyboardPush(code)
    mMessage = IIf(accepted, "Tecla '" & Chr$(code) & "' (" & AsmHex2(code) & ") en la cola.", "KBD_FULL: la cola tiene 8 teclas, se rechaza.")
    MaquinaRefrescar
    If mSpeed > 0 Then
        Flash ws, Mid$(nm, 5), LAMP_ON, HEAD, 2
        If accepted Then Flash ws, "kq" & CStr((KeyboardTail() + 7) Mod 8), LAMP_ON, DATA_FILL, 2
    End If
    AfterAction
    Exit Sub
failed:
    mMessage = "ERROR: " & Err.Description
    MaquinaRefrescar
End Sub

Public Sub MaquinaVelocidad()
    FlowCycleSpeed
    EnsureSpeed
    mMessage = "Animacion: " & SpeedText() & " (rige tambien en los demas bloques)."
    MaquinaRefrescar
End Sub

Private Sub LoadProgram(ByVal source As Variant, ByVal message As String)
    EditorWrite source
    If LoaderLoadSource(source) Then
        mMessage = message
        EditorRefresh
    Else
        mMessage = "Error al ensamblar el programa de demostracion."
    End If
    AfterAction
    MaquinaRefrescar
End Sub

' --- Animation ----------------------------------------------------------------------------------

Private Sub BeginStepCapture()
    mFills0 = CacheRamFills()
    mFaults0 = PagerFaultCount()
End Sub

' Replays the last instruction on the drawing: FETCH (VA -> MMU -> L1I, fills and page faults),
' DECODE (IR), EXECUTE (ALU / L1D / bus to keyboard or LCD / PC), STORE (written register).
Private Sub AnimateLastStep(ByVal ws As Worksheet)
    Dim kind As String, s As LcaPathStep, dataFill As Boolean, dataFault As Boolean
    Dim fetchFill As Boolean, fetchFault As Boolean, fills As Long, faults As Long, r As Long
    kind = ExecLastKind()
    If Len(kind) = 0 Then Exit Sub
    Application.ScreenUpdating = True
    mX0 = CL(ws, ROW_BASE + DIAGRAM_ROW, 2).Left          ' geometry is module state: rebuild it (lost on reopen)
    mY0 = CL(ws, ROW_BASE + DIAGRAM_ROW, 2).Top
    mSX = (CL(ws, ROW_BASE + DIAGRAM_ROW, 20).Left - mX0) / DESIGN_W
    s = CachePathLast()
    fills = CacheRamFills() - mFills0
    faults = PagerFaultCount() - mFaults0
    If kind = "LD" Or kind = "ST" Then
        dataFill = (s.RamCycles >= 4)
        dataFault = s.PageFault
    End If
    fetchFill = (fills - IIf(dataFill, 1, 0)) > 0
    fetchFault = (faults - IIf(dataFault, 1, 0)) > 0

    ' FETCH
    Stage ws, 0
    MovePacket ws, Array(185, 120, 530, 200)
    If fetchFault Then SwapIn ws
    MovePacket ws, Array(530, 200, 768, 90)
    If fetchFill Then MovePacket ws, Array(768, 90, 840, 300, 1120, 215, 840, 300, 768, 90)
    MovePacket ws, Array(768, 90, 530, 120, 185, 205)
    HidePacket ws
    If kind = "ERROR" Then
        Flash ws, "ctrl", BAD, CARD, 3
        Stage ws, -1
        Exit Sub
    End If

    ' DECODE
    Stage ws, 1
    Flash ws, "ir", LAMP_ON, HEAD, 1

    ' EXECUTE
    Stage ws, 2
    Select Case kind
        Case "ALU"
            Flash ws, "alu", LAMP_ON, EDGE, 3
        Case "LD", "ST"
            MovePacket ws, Array(185, 100, 530, 200)
            If dataFault Then SwapIn ws
            MovePacket ws, Array(530, 200, 912, 90)
            If dataFill Or kind = "ST" Then MovePacket ws, Array(912, 90, 840, 300, 1120, 215)
            If kind = "LD" Then MovePacket ws, Array(1120, 215, 840, 300, 912, 90, 185, 100)
            HidePacket ws
        Case "IN"
            MovePacket ws, Array(850, 566, 850, 477, 185, 477, 185, 100)
            HidePacket ws
        Case "OUT"
            MovePacket ws, Array(185, 100, 185, 477, 1200, 477, 1200, 580)
            HidePacket ws
            Flash ws, "lcd" & CStr((LcdCursor() + 7) Mod 8), LCD_LIT, LCD_BG, 3
        Case "JMP"
            Flash ws, "pc", IIf(ExecLastTaken(), LAMP_ON, EDGE), HEAD, 2
        Case "WAIT"
            Flash ws, "kinfo", WARN, CARD, 3
        Case "HLT"
            Flash ws, "ctrl", OK_GREEN, CARD, 2
    End Select

    ' STORE
    Stage ws, 3
    For r = 0 To 3
        If ws.Shapes("P12_reg" & CStr(r)).Fill.ForeColor.RGB = WARN Then Flash ws, "reg" & CStr(r), LAMP_ON, WARN, 2
    Next r
    Flash ws, "pc", ACCENT, HEAD, 1
    BlinkCursor ws, 2
    Stage ws, -1
End Sub

' Page fault: the page travels from the swap disk into a RAM frame.
Private Sub SwapIn(ByVal ws As Worksheet)
    ws.Shapes("P12_a_swap").Line.ForeColor.RGB = LAMP_ON
    ws.Shapes("P12_packet").Fill.ForeColor.RGB = WARN
    MovePacket ws, Array(1310, 75, 1120, 75, 1120, 215)
    ws.Shapes("P12_packet").Fill.ForeColor.RGB = PINK
End Sub

Private Sub Stage(ByVal ws As Worksheet, ByVal active As Long)
    Dim i As Long
    For i = 0 To 3
        With ws.Shapes("P12_stg" & CStr(i))
            .Fill.ForeColor.RGB = IIf(i = active, LAMP_ON, CARD)
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = IIf(i = active, BG, MUTED)
        End With
    Next i
    Pause 3
End Sub

' Moves the packet through a list of design points (x1, y1, x2, y2, ...).
Private Sub MovePacket(ByVal ws As Worksheet, ByVal points As Variant)
    Dim i As Long, f As Long, frames As Long, x As Double, y As Double, p As Shape
    Set p = ws.Shapes("P12_packet")
    p.Visible = msoTrue
    frames = Choose(mSpeed + 1, 1, 9, 6, 3)
    For i = 0 To UBound(points) - 3 Step 2
        For f = 1 To frames
            x = points(i) + (points(i + 2) - points(i)) * f / frames
            y = points(i + 1) + (points(i + 3) - points(i + 1)) * f / frames
            p.Left = mX0 + x * mSX - p.Width / 2
            p.Top = mY0 + y - p.Height / 2
            Pause 1
        Next f
    Next i
End Sub

Private Sub HidePacket(ByVal ws As Worksheet)
    ws.Shapes("P12_packet").Visible = msoFalse
End Sub

Private Sub Flash(ByVal ws As Worksheet, ByVal name As String, ByVal onColor As Long, ByVal offColor As Long, ByVal times As Long)
    Dim i As Long, s As Shape
    Set s = ws.Shapes("P12_" & name)
    For i = 1 To times
        s.Fill.ForeColor.RGB = onColor
        Pause 3
        s.Fill.ForeColor.RGB = offColor
        Pause 2
    Next i
End Sub

Private Sub BlinkCursor(ByVal ws As Worksheet, ByVal times As Long)
    Dim i As Long
    For i = 1 To times
        ws.Shapes("P12_lcdcur").Visible = msoFalse
        Pause 3
        ws.Shapes("P12_lcdcur").Visible = msoTrue
        Pause 3
    Next i
End Sub

Private Sub Pause(ByVal ticks As Long)
    Dim ms As Long
    If mSpeed = 0 Then Exit Sub
    ms = Choose(mSpeed, 45, 25, 10) * ticks
    DoEvents
    Sleep ms
End Sub

Private Function SpeedText() As String
    SpeedText = Choose(mSpeed + 1, "SIN ANIMACION", "LENTA", "NORMAL", "RAPIDA")
End Function

' --- Refresh -----------------------------------------------------------------------------------

Public Sub MaquinaRefrescar()
    EnsureSpeed
    Dim ws As Worksheet, previous As Boolean
    Set ws = MachineSheet(False)
    If ws Is Nothing Then Exit Sub
    If Not LoaderLoaded() Then Exit Sub
    previous = Application.ScreenUpdating
    Application.ScreenUpdating = False
    RefreshStatus ws
    RefreshCpu ws
    RefreshMmu ws
    RefreshCaches ws
    RefreshRam ws
    RefreshDevices ws
    RefreshBus ws
    RefreshArrows ws
    RefreshListing ws
    Application.ScreenUpdating = previous
    PanelActualizar
End Sub

Private Sub RefreshStatus(ByVal ws As Worksheet)
    Dim st As String, fill As Long, label As String
    st = ExecStatus()
    Select Case st
        Case EXEC_HALTED: label = "DETENIDO (HLT)": fill = OK_GREEN
        Case EXEC_WAIT: label = "ESPERANDO TECLA (IN)": fill = WARN
        Case EXEC_ERROR: label = "ERROR: " & ExecErrorText(): fill = BAD
        Case EXEC_LIMIT: label = "PAUSA (limite de pasos)": fill = WARN
        Case Else: label = "LISTO  |  paso " & CStr(ExecSteps()): fill = HEAD
    End Select
    ws.Range(RA("P7")).Value = label
    ws.Range(RA("P7:S8")).Interior.Color = fill
    ws.Range(RA("F10")).Value = ExecLastText() & IIf(Len(mMessage) > 0, "     |  " & mMessage, "")
    SetBox ws, "btn_MaquinaVelocidad", "VELOCIDAD" & vbCr & SpeedText()
    SetBox ws, "speed", "Animacion " & SpeedText() & "   |   VELOCIDAD cambia: normal > lenta > sin animacion > rapida" & vbCr & _
           "EJECUTAR anima cada instruccion hasta HLT o una espera de IN (max 300).", CARD
End Sub

Private Sub RefreshCpu(ByVal ws As Worksheet)
    Dim cpu As LcaCpuState, i As Long, v As Long, names As Variant, kind As String
    cpu = CpuSnapshot()
    names = Array("AX", "BX", "CX", "DX")
    For i = 0 To 3
        v = CpuReadRegister(i)
        SetBox ws, "reg" & CStr(i), CStr(names(i)) & vbCr & AsmHex2(v) & "   " & CStr(v), _
               IIf(RegisterTouched(i), WARN, CARD)
    Next i
    SetBox ws, "pc", "PC" & vbCr & AsmHex3(cpu.PC), HEAD
    SetBox ws, "ir", "IR" & vbCr & AsmHex2(cpu.IR), HEAD
    Lamp ws, "flagZ", CpuFlagIsSet(LCA_FLAG_ZF)
    Lamp ws, "flagC", CpuFlagIsSet(LCA_FLAG_CF)
    Lamp ws, "flagS", CpuFlagIsSet(LCA_FLAG_SF)
    SetBox ws, "marv", "MAR_V" & vbCr & AsmHex3(cpu.MAR_V)
    SetBox ws, "marp", "MAR_P" & vbCr & AsmHex2(cpu.MAR_P)
    SetBox ws, "mdr", "MDR" & vbCr & AsmHex2(cpu.MDR)
    kind = ExecLastKind()
    If kind = "ALU" Then
        SetBox ws, "alu", "ALU" & vbCr & InstructionAt(ExecLastPC()), EDGE
    Else
        SetBox ws, "alu", "ALU" & vbCr & "en reposo", EDGE
    End If
    If ExecSteps() = 0 And ExecStatus() = EXEC_READY Then
        SetBox ws, "ctrl", "CONTROL: listo para FETCH en PC = " & AsmHex3(cpu.PC), CARD
    Else
        SetBox ws, "ctrl", "CONTROL: " & InstructionAt(ExecLastPC()) & "  (" & KindText(kind) & ")", CARD
    End If
End Sub

Private Sub RefreshMmu(ByVal ws As Worksheet)
    Dim v As Long, e As LcaPageEntry, fill As Long, text As String, codeVpn As Long, dataVpn As Long
    Dim s As LcaPathStep
    codeVpn = ExecLastPC() \ 32
    dataVpn = -1
    If ExecLastDataVA() >= 0 Then dataVpn = ExecLastDataVA() \ 32
    For v = 0 To 15
        e = MmuPage(v)
        If v = 15 Then
            text = "VPN 15" & vbCr & "E/S"
            fill = IO_FILL
        ElseIf e.Present Then
            text = "VPN " & CStr(v) & vbCr & "PFN " & CStr(e.PFN) & vbCr & IIf(e.Dirty, "D ", "") & IIf(e.Referenced, "R", "")
            fill = IIf(v <= 7, CODE_FILL, DATA_FILL)
        Else
            text = "VPN " & CStr(v) & vbCr & "en swap"
            fill = CARD
        End If
        With ws.Shapes("P12_pte" & CStr(v)).Line
            If ExecSteps() > 0 And (v = dataVpn Or (v = codeVpn And dataVpn < 0)) Then
                If ExecLastFault() Then fill = WARN
                .ForeColor.RGB = LAMP_ON
                .Weight = 2.5
            Else
                .ForeColor.RGB = EDGE
                .Weight = 1
            End If
        End With
        SetBox ws, "pte" & CStr(v), text, fill
    Next v
    s = CachePathLast()
    If s.PA >= 0 And Not s.Failed Then
        SetBox ws, "xlat", "VA " & AsmHex3(s.VA) & "  ->  PA " & AsmHex2(s.PA) & vbCr & _
               "VPN " & CStr(s.VPN) & " -> PFN " & CStr(s.PFN) & IIf(s.PageFault, "   (PAGE_FAULT)", ""), _
               IIf(s.PageFault, WARN, CARD)
    Else
        SetBox ws, "xlat", "VA -> PA" & vbCr & "sin accesos todavia", CARD
    End If
End Sub

Private Sub RefreshCaches(ByVal ws As Worksheet)
    Dim i As Long
    SetBox ws, "l1i", "L1I   " & CStr(CacheHits(LCA_CACHE_L1I)) & " hit / " & CStr(CacheMisses(LCA_CACHE_L1I)) & " miss"
    SetBox ws, "l1d", "L1D   " & CStr(CacheHits(LCA_CACHE_L1D)) & " hit / " & CStr(CacheMisses(LCA_CACHE_L1D)) & " miss"
    SetBox ws, "l2", "L2 unificada   " & CStr(CacheHits(LCA_CACHE_L2)) & " hit / " & CStr(CacheMisses(LCA_CACHE_L2)) & " miss"
    For i = 0 To 1
        CacheLineBox ws, "l1i_" & CStr(i), LCA_CACHE_L1I, i
        CacheLineBox ws, "l1d_" & CStr(i), LCA_CACHE_L1D, i
    Next i
    For i = 0 To 7
        CacheLineBox ws, "l2_" & CStr(i), LCA_CACHE_L2, i
    Next i
End Sub

Private Sub CacheLineBox(ByVal ws As Worksheet, ByVal name As String, ByVal level As Long, ByVal idx As Long)
    Dim n As Long, base As Long, k As Long, bytes As String
    If CacheLineValid(level, idx) Then
        n = CacheLineCount(level)
        base = (CacheLineTag(level, idx) * n + idx) * LCA_CACHE_LINE_BYTES
        For k = 0 To 3
            bytes = bytes & Right$("00" & Hex$(CacheLineByte(level, idx, k)), 2) & " "
        Next k
        SetBox ws, name, CStr(idx) & ": PA " & AsmHex2(base) & "  (PFN " & CStr(base \ 32) & ")" & vbCr & RTrim$(bytes), LINE_ON
        ws.Shapes("P12_" & name).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = BG
    Else
        SetBox ws, name, CStr(idx) & ": vacia", CARD
        ws.Shapes("P12_" & name).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = MUTED
    End If
End Sub

Private Sub RefreshRam(ByVal ws As Worksheet)
    Dim p As Long, owner As Long, e As LcaPageEntry, text As String, fill As Long, lines As Long
    For p = 0 To 7
        lines = CacheFrameLines(LCA_CACHE_L1I, p) + CacheFrameLines(LCA_CACHE_L1D, p) + CacheFrameLines(LCA_CACHE_L2, p)
        If p = 7 Then
            text = "PFN 7  |  E/S reservada" & vbCr & "KBD E0-E7h   LCD E8-EFh"
            fill = IO_FILL
        Else
            owner = PagerFrameOwner(p)
            If owner < 0 Then
                text = "PFN " & CStr(p) & "  |  libre"
                fill = CARD
            Else
                e = MmuPage(owner)
                text = "PFN " & CStr(p) & "  |  VPN " & CStr(owner) & IIf(owner <= 7, " codigo", " datos") & vbCr & _
                       "dirty " & Abs(CLng(e.Dirty)) & "  ref " & Abs(CLng(e.Referenced)) & "  lineas " & CStr(lines) & _
                       IIf(PagerHand() = p, "   < CLOCK", "")
                fill = IIf(owner <= 7, CODE_FILL, DATA_FILL)
            End If
        End If
        SetBox ws, "pfn" & CStr(p), text, fill
    Next p
    SetBox ws, "swap", "SWAP" & vbCr & "15 pag x 32 B" & vbCr & "faults " & CStr(PagerFaultCount()) & vbCr & _
           "writeback " & CStr(PagerWritebackCount()), IIf(ExecLastFault(), WARN, HEAD)
End Sub

Private Sub RefreshDevices(ByVal ws As Worksheet)
    Dim i As Long, fill As Long, txt As String, b As Long, cell As Shape
    For i = 0 To 7
        b = KeyboardSlot(i)
        If KeyboardSlotQueued(i) Then
            fill = DATA_FILL
            txt = LcdDisplayChar(b)
        Else
            fill = CARD
            txt = vbNullString
        End If
        SetBox ws, "kq" & CStr(i), txt, fill
        ws.Shapes("P12_kq" & CStr(i)).Line.ForeColor.RGB = IIf(i = KeyboardHead(), LAMP_ON, EDGE)
    Next i
    SetBox ws, "kinfo", "en cola " & CStr(KeyboardCount()) & "/8   head " & CStr(KeyboardHead()) & "   tail " & _
           CStr(KeyboardTail()) & "   leidas " & CStr(KeyboardConsumed()), CARD
    For i = 0 To 7
        SetBox ws, "lcd" & CStr(i), LcdChar(i), LCD_BG
    Next i
    Set cell = ws.Shapes("P12_lcd" & CStr(LcdCursor()))
    With ws.Shapes("P12_lcdcur")
        .Left = cell.Left + 2
        .Visible = msoTrue
    End With
    SetBox ws, "lcdinfo", "Texto: [" & LcdText() & "]   cursor " & CStr(LcdCursor()) & "   escrituras " & CStr(LcdWrites()) & _
           vbCr & "La barra marca donde escribe el proximo OUT", CARD
End Sub

' The external bus floats (High-Z) after a fetch latch, so the lamps show the last real external cycle.
Private Sub RefreshBus(ByVal ws As Worksheet)
    Dim line As String, tokens As Variant, i As Long, bits As String, p As Long
    line = LastExternalCycle()
    SetBox ws, "cycle", "Ultimo ciclo del bus externo:" & vbCr & line
    p = InStr(line, "| R")
    If p > 0 Then bits = " " & Mid$(line, p + 2) & " "
    tokens = Array("R", "W", "C", "I", "E", "L", "KBD", "LCD")
    For i = 0 To 7
        Lamp ws, "lamp" & CStr(i), (Len(bits) > 0 And InStr(bits, " " & tokens(i) & "1 ") > 0)
    Next i
End Sub

Private Sub RefreshArrows(ByVal ws As Worksheet)
    Dim kind As String, ran As Boolean, fills As Boolean
    kind = ExecLastKind()
    ran = ExecSteps() > 0 Or kind = "WAIT" Or kind = "ERROR"
    fills = CacheRamFills() > mFills0
    SetArrow ws, "a_va", ran And kind <> "ERROR"
    SetArrow ws, "a_pa", ran And kind <> "ERROR"
    SetArrow ws, "a_fill", fills
    SetArrow ws, "a_ram_bus", fills Or kind = "ST"
    SetArrow ws, "a_cpu_bus", fills Or kind = "ST" Or kind = "IN" Or kind = "OUT"
    SetArrow ws, "a_swap", ExecLastFault()
    SetArrow ws, "a_kbd", kind = "IN"
    SetArrow ws, "a_lcd", kind = "OUT"
End Sub

Private Sub RefreshListing(ByVal ws As Worksheet)
    Dim first As Long, i As Long, addr As Long, row As Long, pc As Long, n As Long, lr As Long
    lr = ListRow(ws)
    pc = CpuSnapshot().PC
    n = LoaderInstructionCount() + 1
    If n > 64 Then n = 64
    first = (pc \ 4) - 4
    If first > n - LIST_ROWS Then first = n - LIST_ROWS
    If first < 0 Then first = 0
    For i = 0 To LIST_ROWS - 1
        row = ROW_BASE + lr + i
        With ws.Range(CL(ws, row, 2), CL(ws, row, 8))
            .Value = vbNullString
            .Interior.Color = CARD
            .Font.Color = INK
        End With
        If first + i < n Then
            addr = (first + i) * 4
            CL(ws, row, 2).Value = AsmHex3(addr)
            CL(ws, row, 3).Value = Right$("00" & Hex$(LoaderImageByte(addr)), 2) & " " & _
                                   Right$("00" & Hex$(LoaderImageByte(addr + 1)), 2) & " " & _
                                   Right$("00" & Hex$(LoaderImageByte(addr + 2)), 2) & " " & _
                                   Right$("00" & Hex$(LoaderImageByte(addr + 3)), 2)
            CL(ws, row, 4).Value = InstructionAt(addr)
            If addr = pc Then
                ws.Range(CL(ws, row, 2), CL(ws, row, 8)).Interior.Color = LAMP_ON
                ws.Range(CL(ws, row, 2), CL(ws, row, 8)).Font.Color = BG
                CL(ws, row, 4).Value = InstructionAt(addr) & "      <- PC"
            End If
        End If
    Next i
    RefreshIndicators ws, lr
End Sub

' Six KPI tiles and a hit/miss chart per cache level (data cells hidden under the chart).
Private Sub DrawIndicators(ByVal ws As Worksheet, ByVal lr As Long)
    Dim area As Range, x0 As Double, y0 As Double, w As Double, i As Long, co As ChartObject, names As Variant
    Set area = ws.Range(CL(ws, ROW_BASE + lr, 10), CL(ws, ROW_BASE + lr + LIST_ROWS - 1, 19))
    area.Interior.Color = BG
    x0 = area.Left
    y0 = area.Top
    w = area.Width / 3
    names = Array("INSTRUCCIONES", "FALLOS DE PAGINA", "LLENADOS DE LINEA", "ACIERTOS L1I", "ACIERTOS L1D", "E/S  teclado | LCD")
    For i = 0 To 5
        RawBox ws, "kpi" & CStr(i), x0 + (i Mod 3) * w + 4, y0 + (i \ 3) * 62, w - 8, 56, CARD, CStr(names(i)) & vbCr & "0", 10
        ws.Shapes("P12_kpi" & CStr(i)).TextFrame2.TextRange.Paragraphs(2).Font.Size = 20
        ws.Shapes("P12_kpi" & CStr(i)).TextFrame2.TextRange.Paragraphs(2).Font.Bold = msoTrue
        ws.Shapes("P12_kpi" & CStr(i)).Line.ForeColor.RGB = ACCENT
    Next i
    For i = 0 To 2
        CL(ws, ROW_BASE + lr + 12 + i, 11).Value = CacheLevelName(i)
        CL(ws, ROW_BASE + lr + 12 + i, 12).Value = 0
        CL(ws, ROW_BASE + lr + 12 + i, 13).Value = 0
    Next i
    Set co = ws.ChartObjects.Add(x0 + 4, y0 + 130, area.Width - 8, area.Top + area.Height - (y0 + 130) - 4)
    co.Name = "P12_chart"
    co.Placement = xlMove
    With co.Chart
        .ChartType = xlBarClustered
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        With .SeriesCollection.NewSeries
            .Name = "hit"
            .Values = ws.Range(CL(ws, ROW_BASE + lr + 12, 12), CL(ws, ROW_BASE + lr + 14, 12))
            .XValues = ws.Range(CL(ws, ROW_BASE + lr + 12, 11), CL(ws, ROW_BASE + lr + 14, 11))
            .Format.Fill.ForeColor.RGB = LINE_ON
            .HasDataLabels = True
            .DataLabels.Font.Color = INK
        End With
        With .SeriesCollection.NewSeries
            .Name = "miss"
            .Values = ws.Range(CL(ws, ROW_BASE + lr + 12, 13), CL(ws, ROW_BASE + lr + 14, 13))
            .XValues = ws.Range(CL(ws, ROW_BASE + lr + 12, 11), CL(ws, ROW_BASE + lr + 14, 11))
            .Format.Fill.ForeColor.RGB = RGB(251, 146, 60)
            .HasDataLabels = True
            .DataLabels.Font.Color = INK
        End With
        .HasTitle = True
        .ChartTitle.Text = "CACHES: aciertos vs fallos"
        .ChartTitle.Font.Color = ACCENT
        .ChartTitle.Font.Size = 11
        .HasLegend = True
        .Legend.Position = xlLegendPositionTop
        .Legend.Font.Color = INK
        .ChartArea.Format.Fill.ForeColor.RGB = CARD
        .ChartArea.Format.Line.ForeColor.RGB = ACCENT
        .PlotArea.Format.Fill.ForeColor.RGB = CARD
        .Axes(xlValue).TickLabels.Font.Color = MUTED
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = EDGE
        .Axes(xlCategory).TickLabels.Font.Color = INK
        .Axes(xlCategory).TickLabels.Font.Bold = True
    End With
End Sub

Private Sub RefreshIndicators(ByVal ws As Worksheet, ByVal lr As Long)
    Dim i As Long, values As Variant, l1i As Long, l1d As Long
    l1i = CacheHits(LCA_CACHE_L1I) + CacheMisses(LCA_CACHE_L1I)
    l1d = CacheHits(LCA_CACHE_L1D) + CacheMisses(LCA_CACHE_L1D)
    ' IIf evaluates both branches, so the percentages are computed only when there were accesses.
    values = Array(CStr(ExecSteps()), CStr(PagerFaultCount()), CStr(CacheRamFills()), _
                   HitRate(CacheHits(LCA_CACHE_L1I), l1i), HitRate(CacheHits(LCA_CACHE_L1D), l1d), _
                   CStr(KeyboardConsumed()) & " | " & CStr(LcdWrites()))
    For i = 0 To 5
        With ws.Shapes("P12_kpi" & CStr(i)).TextFrame2.TextRange.Paragraphs(2)
            If .Text <> CStr(values(i)) Then .Text = CStr(values(i))
        End With
    Next i
    For i = 0 To 2
        CL(ws, ROW_BASE + lr + 12 + i, 12).Value = CacheHits(i)
        CL(ws, ROW_BASE + lr + 12 + i, 13).Value = CacheMisses(i)
    Next i
End Sub

Private Function HitRate(ByVal hits As Long, ByVal total As Long) As String
    If total = 0 Then
        HitRate = "-"
    Else
        HitRate = Format$(100 * hits / total, "0") & " %"
    End If
End Function

' Absolute-position box (points), used for the indicator tiles.
Private Sub RawBox(ByVal ws As Worksheet, ByVal name As String, ByVal left As Double, ByVal top As Double, _
                   ByVal w As Double, ByVal h As Double, ByVal fill As Long, ByVal text As String, ByVal size As Single)
    Dim s As Shape
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, left, top, w, h)
    s.Name = "P12_" & name
    s.Fill.ForeColor.RGB = fill
    s.Line.ForeColor.RGB = EDGE
    With s.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = text
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = size
        .TextRange.Font.Fill.ForeColor.RGB = INK
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    s.Placement = xlMove
End Sub

' --- Helpers -------------------------------------------------------------------------------------

Private Function InstructionAt(ByVal addr As Long) As String
    If addr < 0 Or addr > 252 Then
        InstructionAt = "-"
    Else
        InstructionAt = AsmDisassemble(LoaderImageByte(addr), LoaderImageByte(addr + 1), _
                                       LoaderImageByte(addr + 2), LoaderImageByte(addr + 3))
    End If
End Function

Private Function RegisterTouched(ByVal code As Long) As Boolean
    Dim addr As Long, op As Long
    If ExecSteps() = 0 Then Exit Function
    addr = ExecLastPC()
    If addr < 0 Or addr > 252 Then Exit Function
    op = LoaderImageByte(addr)
    Select Case op
        Case &H10, &H11, &H12, &H20, &H21, &H23, &H24, &H40
            RegisterTouched = (LoaderImageByte(addr + 1) = code) And ExecLastKind() <> "WAIT" And ExecLastKind() <> "ERROR"
    End Select
End Function

Private Function KindText(ByVal kind As String) As String
    Select Case kind
        Case "ALU": KindText = "operacion en la ALU"
        Case "LD": KindText = "lectura de datos por L1D"
        Case "ST": KindText = "escritura write-through por L1D"
        Case "JMP": KindText = IIf(ExecLastTaken(), "salto tomado", "salto no tomado")
        Case "IN": KindText = "entrada desde el teclado"
        Case "OUT": KindText = "salida al LCD"
        Case "WAIT": KindText = "IN espera una tecla"
        Case "HLT": KindText = "fin del programa"
        Case "ERROR": KindText = "error: " & ExecErrorText()
        Case "MOV": KindText = "copia a registro"
        Case Else: KindText = LCase$(kind)
    End Select
End Function

Private Function LastExternalCycle() As String
    Dim i As Long, line As String
    For i = BusTraceCount() - 1 To 0 Step -1
        line = BusTraceLine(i)
        If InStr(line, "| EXTERNAL |") > 0 And InStr(line, "drivers=0") = 0 Then
            LastExternalCycle = line
            Exit Function
        End If
    Next i
    LastExternalCycle = "High-Z (sin transferencias externas)"
End Function

Private Function MacroRef(ByVal macroName As String) As String
    MacroRef = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Function

' Shape factories: coordinates in design units (x scaled to the block width, y in points).
Private Sub Frame(ByVal ws As Worksheet, ByVal name As String, ByVal x As Double, ByVal y As Double, _
                  ByVal w As Double, ByVal h As Double, ByVal title As String, ByVal macroName As String)
    Dim s As Shape
    Set s = Box(ws, name, x, y, w, h, BG, title, 10, True)
    s.Line.ForeColor.RGB = ACCENT
    s.Line.Weight = 1.25
    s.TextFrame2.VerticalAnchor = msoAnchorTop
    s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ACCENT
    s.TextFrame2.MarginTop = 4
    If Len(macroName) > 0 Then s.OnAction = MacroRef(macroName)
End Sub

Private Function Box(ByVal ws As Worksheet, ByVal name As String, ByVal x As Double, ByVal y As Double, _
                     ByVal w As Double, ByVal h As Double, ByVal fill As Long, ByVal text As String, _
                     ByVal size As Single, Optional ByVal bold As Boolean = False, _
                     Optional ByVal kind As Long = msoShapeRoundedRectangle) As Shape
    Dim s As Shape
    Set s = ws.Shapes.AddShape(kind, mX0 + x * mSX, mY0 + y, w * mSX, h)
    s.Name = "P12_" & name
    s.Fill.ForeColor.RGB = fill
    s.Line.ForeColor.RGB = EDGE
    s.Line.Weight = 1
    If kind = msoShapeRoundedRectangle Then s.Adjustments.Item(1) = 0.12
    With s.TextFrame2
        .MarginLeft = 3
        .MarginRight = 3
        .MarginTop = 1
        .MarginBottom = 1
        .WordWrap = msoTrue
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = text
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = size
        .TextRange.Font.Bold = IIf(bold, msoTrue, msoFalse)
        .TextRange.Font.Fill.ForeColor.RGB = INK
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    s.Placement = xlMove
    Set Box = s
End Function

Private Sub SetBox(ByVal ws As Worksheet, ByVal name As String, ByVal text As String, Optional ByVal fill As Long = -1)
    With ws.Shapes("P12_" & name)
        If .TextFrame2.TextRange.Text <> text Then .TextFrame2.TextRange.Text = text
        If fill >= 0 Then .Fill.ForeColor.RGB = fill
    End With
End Sub

Private Sub Lamp(ByVal ws As Worksheet, ByVal name As String, ByVal isOn As Boolean)
    With ws.Shapes("P12_" & name)
        .Fill.ForeColor.RGB = IIf(isOn, LAMP_ON, CARD)
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = IIf(isOn, BG, MUTED)
    End With
End Sub

Private Sub Arrow(ByVal ws As Worksheet, ByVal name As String, ByVal x1 As Double, ByVal y1 As Double, _
                  ByVal x2 As Double, ByVal y2 As Double, ByVal both As Boolean)
    Dim s As Shape
    Set s = ws.Shapes.AddConnector(msoConnectorStraight, mX0 + x1 * mSX, mY0 + y1, mX0 + x2 * mSX, mY0 + y2)
    s.Name = "P12_" & name
    With s.Line
        .ForeColor.RGB = EDGE
        .Weight = 4
        .EndArrowheadStyle = msoArrowheadTriangle
        If both Then .BeginArrowheadStyle = msoArrowheadTriangle
    End With
End Sub

Private Sub SetArrow(ByVal ws As Worksheet, ByVal name As String, ByVal active As Boolean)
    ws.Shapes("P12_" & name).Line.ForeColor.RGB = IIf(active, LAMP_ON, EDGE)
End Sub

Private Sub MakeButton(ByVal ws As Worksheet, ByVal name As String, ByVal left As Double, ByVal top As Double, _
                       ByVal w As Double, ByVal h As Double, ByVal caption As String, ByVal macroName As String, _
                       ByVal fill As Long, ByVal size As Single)
    Dim s As Shape
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, left, top, w, h)
    s.Name = "P12_" & name
    s.Fill.ForeColor.RGB = fill
    s.Line.ForeColor.RGB = ACCENT
    s.Line.Weight = 0.75
    With s.TextFrame2
        .MarginLeft = 2
        .MarginRight = 2
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = caption
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = size
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    s.Shadow.Visible = msoTrue
    s.Shadow.OffsetX = 1.5
    s.Shadow.OffsetY = 1.5
    s.Shadow.Transparency = 0.5
    s.OnAction = MacroRef(macroName)
End Sub

Private Sub Btn(ByVal ws As Worksheet, ByVal address As String, ByVal caption As String, _
                ByVal macroName As String, ByVal fill As Long)
    Dim area As Range
    Set area = ws.Range(RA(address))
    MakeButton ws, "btn_" & macroName, area.Left + 2, area.Top + 2, area.Width - 4, area.Height - 4, caption, macroName, fill, 11
End Sub

Private Sub Title(ByVal ws As Worksheet, ByVal address As String, ByVal text As String)
    With ws.Range(RA(address))
        .Merge
        .Value = text
        .Font.Bold = True
        .Font.Color = ACCENT
        .Borders(xlEdgeBottom).Color = HEAD
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
End Sub

' Module state is lost when the workbook is reopened: the default animation speed is NORMAL.
Private Sub EnsureSpeed()
    mSpeed = FlowSpeed()          ' one shared speed for the machine and every animated block
    mSpeedSet = True
End Sub

' Block row where the listing starts: first row at least LIST_GAP points below the diagram origin.
Private Function ListRow(ByVal ws As Worksheet) As Long
    Dim r As Long, y0 As Double
    y0 = CL(ws, ROW_BASE + DIAGRAM_ROW, 2).Top
    r = DIAGRAM_ROW
    Do While CL(ws, ROW_BASE + r, 2).Top - y0 < LIST_GAP And r < BAND_ROWS - LIST_ROWS - 1
        r = r + 1
    Loop
    ListRow = r + 1
End Function

Private Function MachineSheet(ByVal create As Boolean) As Worksheet
    Set MachineSheet = LcaBand("P12_btn_MaquinaPaso", create)
End Function

Private Function RA(ByVal address As String) As String
    RA = LcaShift(address, ROW_BASE, COL_BASE)
End Function

' Cell of this block: row is absolute (ROW_BASE already added), column is relative to the block.
Private Function CL(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Range
    Set CL = ws.Cells(r, c + COL_BASE)
End Function

' Heavier views (Logs sheet, RAM map, CLOCK) are refreshed once per user action, not per animation frame.
Private Sub AfterAction()
    On Error Resume Next
    LogsActualizar
    MemoriaDetalleActualizar
    BusDetalleActualizar
End Sub
