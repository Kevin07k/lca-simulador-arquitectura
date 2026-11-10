Attribute VB_Name = "LcaMachinePanel"
Option Explicit

' P12 band of the LCA sheet: the whole machine drawn with shapes and updated after every instruction.
'   CPU (registers, PC/IR, MAR/MDR, flags, ALU, control unit) -> MMU (16-page grid) -> L1I/L1D/L2 lines
'   -> RAM (8 frames) <-> SWAP disk, external bus with live signal lamps, keyboard queue and 8x1 LCD.
' Active paths of the last instruction light up. Listing (current PC) and execution trace below.
Private Const ROW_BASE As Long = LCA_BASE_MACHINE
Private Const BAND_ROWS As Long = 66
Private Const DIAGRAM_ROW As Long = 12
Private Const DESIGN_W As Double = 1380
Private Const LIST_ROW As Long = 50
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
Private Const LAMP_ON As Long = 2477306     ' RGB(250, 204, 37) amber lamp
Private Const ARROW_OFF As Long = 5587251
Private Const ARROW_ON As Long = 2477306
Private Const LCD_BG As Long = 2384672      ' RGB(32, 99, 36)
Private Const LCD_INK As Long = 9502671     ' RGB(207, 255, 144)

Private mX0 As Double
Private mY0 As Double
Private mSX As Double
Private mPrevFills As Long
Private mMessage As String

' --- Build ---------------------------------------------------------------------------------------

Public Sub ConstruirPanelMaquina()
    Dim ws As Worksheet, i As Long, previous As Boolean
    previous = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set ws = MachineSheet(True)
    LcaPrepareBand ws, ROW_BASE, BAND_ROWS, "P12_"
    ws.Rows(RA("1:" & CStr(BAND_ROWS))).RowHeight = 18
    With ws.Range(RA("A1:AZ" & CStr(BAND_ROWS)))
        .Interior.Color = BG
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Color = INK
        .VerticalAlignment = xlCenter
    End With

    ' Header
    ws.Rows(ROW_BASE + 2).RowHeight = 34
    ws.Range(RA("B2:L2")).Merge
    With ws.Range(RA("B2"))
        .Value = "LCA  |  MAQUINA COMPLETA EN EJECUCION"
        .Font.Size = 24
        .Font.Bold = True
        .Font.Color = ACCENT
    End With
    ws.Range(RA("B3:S3")).Merge
    With ws.Range(RA("B3"))
        .Value = "P12  ciclo FETCH > DECODE > EXECUTE > STORE  |  CPU 8 bits, VA 9 bits, paginas de 32 B, caches L1I/L1D/L2, " & _
                 "teclado FIFO y LCD 8x1. Las flechas amarillas muestran el camino de la ultima instruccion."
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
    ws.Range(RA("N7:S8")).Merge
    With ws.Range(RA("N7"))
        .HorizontalAlignment = xlCenter
        .Font.Size = 16
        .Font.Bold = True
    End With
    ws.Range(RA("B10")).Value = "TECLAS:"
    ws.Range(RA("B10")).Font.Bold = True
    ws.Range(RA("B10")).Font.Color = ACCENT
    ws.Range(RA("C10:D10")).Merge
    With ws.Range(RA("C10:D10"))
        .NumberFormat = "@"
        .Interior.Color = RGB(230, 240, 250)
        .Font.Color = RGB(15, 23, 42)
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
    mX0 = ws.Cells(ROW_BASE + DIAGRAM_ROW, 2).Left
    mY0 = ws.Cells(ROW_BASE + DIAGRAM_ROW, 2).Top
    mSX = (ws.Cells(ROW_BASE + DIAGRAM_ROW, 20).Left - mX0) / DESIGN_W
    DrawCpu ws
    DrawMmu ws
    DrawCaches ws
    DrawRam ws
    DrawSideDevices ws
    DrawBus ws
    DrawArrows ws

    ' Listing and trace
    Title ws, "B" & CStr(LIST_ROW - 1) & ":H" & CStr(LIST_ROW - 1), "PROGRAMA CARGADO  (fila resaltada = PC)"
    Title ws, "J" & CStr(LIST_ROW - 1) & ":S" & CStr(LIST_ROW - 1), "TRAZA DE EJECUCION  (registros, memoria y senales; mas reciente arriba)"
    For i = 0 To LIST_ROWS - 1
        ws.Range(ws.Cells(ROW_BASE + LIST_ROW + i, 4), ws.Cells(ROW_BASE + LIST_ROW + i, 8)).Merge
        ws.Range(ws.Cells(ROW_BASE + LIST_ROW + i, 10), ws.Cells(ROW_BASE + LIST_ROW + i, 19)).Merge
    Next i
    With ws.Range(ws.Cells(ROW_BASE + LIST_ROW, 2), ws.Cells(ROW_BASE + LIST_ROW + LIST_ROWS - 1, 19))
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
    LcaGoto ws, ROW_BASE
End Sub

Public Sub MostrarPanelMaquina()
    If MachineSheet(False) Is Nothing Then
        ConstruirPanelMaquina
    Else
        LcaGoto MachineSheet(False), ROW_BASE
        MaquinaRefrescar
    End If
End Sub

Private Sub NavBar(ByVal ws As Worksheet)
    Dim labels As Variant, macros As Variant, i As Long, x As Double, w As Double, top As Double
    labels = Array("CPU", "DIAGRAMA", "BUSES", "PAGINAS", "SWAP", "CACHES", "CACHE+PAG", "EDITOR", "TECLADO", "LCD")
    macros = Array("LcaIrCpu", "LcaIrDiagrama", "LcaIrBuses", "LcaIrPaginas", "LcaIrSwap", "LcaIrCaches", _
                   "LcaIrCachePaginacion", "LcaIrEditor", "LcaIrTeclado", "LcaIrLcd")
    x = ws.Cells(ROW_BASE + 5, 2).Left
    w = (ws.Cells(ROW_BASE + 5, 20).Left - x) / 10
    top = ws.Cells(ROW_BASE + 5, 2).Top
    ws.Range(RA("B4")).Value = "IR A / GO TO:"
    ws.Range(RA("B4")).Font.Color = MUTED
    ws.Range(RA("B4")).Font.Size = 8
    For i = 0 To 9
        MakeButton ws, "nav_" & CStr(i), x + i * w + 2, top, w - 4, 20, CStr(labels(i)), CStr(macros(i)), CARD, 9
    Next i
End Sub

Private Sub DrawCpu(ByVal ws As Worksheet)
    Dim i As Long, names As Variant
    Frame ws, "cpu", 0, 0, 370, 430, "CPU"
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
    Box ws, "alu", 60, 298, 250, 58, EDGE, "ALU", 11, True, msoShapeFlowchartManualOperation
    Box ws, "ctrl", 12, 366, 346, 52, CARD, "UNIDAD DE CONTROL", 9
End Sub

Private Sub DrawMmu(ByVal ws As Worksheet)
    Dim v As Long
    Frame ws, "mmu", 400, 0, 260, 330, "MMU  |  TABLA DE PAGINAS"
    For v = 0 To 15
        Box ws, "pte" & CStr(v), 408 + (v Mod 4) * 62, 34 + (v \ 4) * 70, 58, 64, CARD, "VPN " & CStr(v), 9
    Next v
    Box ws, "xlat", 400, 340, 260, 60, CARD, "VA -> PA", 11, True
End Sub

Private Sub DrawCaches(ByVal ws As Worksheet)
    Dim i As Long
    Frame ws, "cache", 690, 0, 300, 430, "CACHES  (lineas de 4 B)"
    Frame ws, "l1i", 700, 30, 136, 120, "L1I"
    Frame ws, "l1d", 844, 30, 136, 120, "L1D"
    For i = 0 To 1
        Box ws, "l1i_" & CStr(i), 708, 66 + i * 40, 120, 34, CARD, "-", 8
        Box ws, "l1d_" & CStr(i), 852, 66 + i * 40, 120, 34, CARD, "-", 8
    Next i
    Frame ws, "l2", 700, 160, 280, 262, "L2 unificada"
    For i = 0 To 7
        Box ws, "l2_" & CStr(i), 708 + (i Mod 2) * 136, 194 + (i \ 2) * 56, 128, 50, CARD, "-", 8
    Next i
End Sub

Private Sub DrawRam(ByVal ws As Worksheet)
    Dim p As Long
    Frame ws, "ram", 1020, 0, 200, 430, "RAM  |  8 MARCOS x 32 B"
    For p = 0 To 7
        Box ws, "pfn" & CStr(p), 1028, 30 + p * 50, 184, 46, CARD, "PFN " & CStr(p), 9
    Next p
End Sub

Private Sub DrawSideDevices(ByVal ws As Worksheet)
    Dim i As Long
    Box ws, "swap", 1240, 0, 140, 150, HEAD, "SWAP", 11, True, msoShapeCan
    Frame ws, "legend", 1240, 170, 140, 260, "LEYENDA"
    Box ws, "lg0", 1248, 200, 124, 26, CODE_FILL, "pagina de codigo", 8
    Box ws, "lg1", 1248, 230, 124, 26, DATA_FILL, "pagina de datos", 8
    Box ws, "lg2", 1248, 260, 124, 26, IO_FILL, "E/S reservada", 8
    Box ws, "lg3", 1248, 290, 124, 26, LINE_ON, "linea valida", 8
    Box ws, "lg4", 1248, 320, 124, 26, WARN, "fallo / miss", 8
    Box ws, "lg5", 1248, 350, 124, 26, LAMP_ON, "camino / senal activa", 8
    Box ws, "lg6", 1248, 380, 124, 40, CARD, "PASO = 1 instruccion" & vbCr & "EJECUTAR = hasta HLT", 8
    ' Keyboard and LCD hang from the external bus.
    Frame ws, "kbd", 700, 520, 300, 96, "TECLADO  |  cola FIFO E0-E7h"
    For i = 0 To 7
        Box ws, "kq" & CStr(i), 708 + i * 36, 552, 32, 32, CARD, "", 9, False, msoShapeRectangle
    Next i
    Box ws, "kinfo", 708, 588, 284, 22, CARD, "", 8, False, msoShapeRectangle
    Frame ws, "lcdf", 1020, 520, 360, 96, "LCD 8x1  |  video E8-EFh"
    Box ws, "lcd", 1030, 550, 340, 56, LCD_BG, "        ", 24, True, msoShapeRectangle
    ws.Shapes("P12_lcd").TextFrame2.TextRange.Font.Name = "Consolas"
    ws.Shapes("P12_lcd").TextFrame2.TextRange.Font.Fill.ForeColor.RGB = LCD_INK
End Sub

Private Sub DrawBus(ByVal ws As Worksheet)
    Dim i As Long, labels As Variant
    Box ws, "bus", 0, 462, 1380, 30, HEAD, "BUS EXTERNO  |  direcciones  -  datos  -  control", 10, True, msoShapeRectangle
    labels = Array("R", "W", "C", "IO", "EN", "LD", "KBD", "LCD")
    For i = 0 To 7
        Box ws, "lamp" & CStr(i), 6 + i * 45, 512, 38, 38, CARD, CStr(labels(i)), 8, True, msoShapeOval
    Next i
    Box ws, "cycle", 0, 560, 680, 56, CARD, "", 8, False, msoShapeRectangle
End Sub

Private Sub DrawArrows(ByVal ws As Worksheet)
    Arrow ws, "a_va", 370, 120, 400, 120, False
    Arrow ws, "a_pa", 660, 120, 690, 120, False
    Arrow ws, "a_fill", 1020, 120, 990, 120, False
    Arrow ws, "a_swap", 1240, 75, 1220, 75, True
    Arrow ws, "a_cpu_bus", 185, 430, 185, 462, True
    Arrow ws, "a_ram_bus", 1120, 430, 1120, 462, True
    Arrow ws, "a_kbd", 850, 520, 850, 492, False
    Arrow ws, "a_lcd", 1200, 492, 1200, 520, False
End Sub

' --- Actions ---------------------------------------------------------------------------------

Public Sub MaquinaEjemplo()
    LoadProgram LoaderSampleProgram(), "Ejemplo SUMA cargado: AX = 5+4+3+2+1 con un bucle JNZ, luego ST/LD en 100h."
End Sub

Public Sub MaquinaEco()
    LoadProgram Array("; Demo ECO: cada tecla se muestra en el LCD", "eco:  IN AX       ; espera una tecla", _
                      "      OUT AX      ; la escribe en el LCD", "      JMP eco"), _
                "Demo ECO cargada. Pulsa TECLEAR y luego EJECUTAR: cada tecla viaja teclado -> AX -> LCD."
End Sub

Public Sub MaquinaCargarEditor()
    Dim lines As Variant
    lines = EditorRead()
    If LoaderLoadSource(lines) Then
        mMessage = "Programa del editor (P11) cargado: " & CStr(LoaderInstructionCount()) & " instrucciones."
    Else
        mMessage = "El programa del editor tiene " & CStr(AsmErrorCount()) & " error(es): revisalos en la banda EDITOR."
    End If
    MaquinaRefrescar
End Sub

Public Sub MaquinaPaso()
    On Error GoTo failed
    If Not LoaderLoaded() Then MaquinaEjemplo
    mPrevFills = CacheRamFills()
    Call ExecStep
    mMessage = vbNullString
    MaquinaRefrescar
    Exit Sub
failed:
    mMessage = "ERROR: " & Err.Description
    MaquinaRefrescar
End Sub

Public Sub MaquinaEjecutar()
    Dim n As Long
    On Error GoTo failed
    If Not LoaderLoaded() Then MaquinaEjemplo
    mPrevFills = CacheRamFills()
    n = ExecRun(500)
    mMessage = CStr(n) & " instrucciones ejecutadas."
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
    mPrevFills = 0
    mMessage = "Reset: CPU, buses, RAM, caches, tabla de paginas, teclado y LCD reiniciados; PC = 000h."
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
    MaquinaRefrescar
    Exit Sub
failed:
    mMessage = "ERROR: " & Err.Description
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
    mPrevFills = 0
    MaquinaRefrescar
End Sub

' --- Refresh -----------------------------------------------------------------------------------

Public Sub MaquinaRefrescar()
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
    ws.Range(RA("N7")).Value = label
    ws.Range(RA("N7:S8")).Interior.Color = fill
    ws.Range(RA("F10")).Value = ExecLastText() & IIf(Len(mMessage) > 0, "     |  " & mMessage, "")
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
    SetBox ws, "pc", "PC" & vbCr & AsmHex3(cpu.PC)
    SetBox ws, "ir", "IR" & vbCr & AsmHex2(cpu.IR)
    Lamp ws, "flagZ", CpuFlagIsSet(LCA_FLAG_ZF)
    Lamp ws, "flagC", CpuFlagIsSet(LCA_FLAG_CF)
    Lamp ws, "flagS", CpuFlagIsSet(LCA_FLAG_SF)
    SetBox ws, "marv", "MAR_V" & vbCr & AsmHex3(cpu.MAR_V)
    SetBox ws, "marp", "MAR_P" & vbCr & AsmHex2(cpu.MAR_P)
    SetBox ws, "mdr", "MDR" & vbCr & AsmHex2(cpu.MDR)
    kind = ExecLastKind()
    If kind = "ALU" Then
        SetBox ws, "alu", "ALU" & vbCr & InstructionAt(ExecLastPC()), LAMP_ON
        ws.Shapes("P12_alu").TextFrame2.TextRange.Font.Fill.ForeColor.RGB = BG
    Else
        SetBox ws, "alu", "ALU" & vbCr & "en reposo", EDGE
        ws.Shapes("P12_alu").TextFrame2.TextRange.Font.Fill.ForeColor.RGB = INK
    End If
    If ExecSteps() = 0 And ExecStatus() = EXEC_READY Then
        SetBox ws, "ctrl", "UNIDAD DE CONTROL" & vbCr & "listo para FETCH en PC = " & AsmHex3(cpu.PC)
    Else
        SetBox ws, "ctrl", "UNIDAD DE CONTROL  |  FETCH > DECODE > EXECUTE > STORE" & vbCr & _
               InstructionAt(ExecLastPC()) & "  (" & KindText(kind) & ")"
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
        If ExecSteps() > 0 And (v = dataVpn Or (v = codeVpn And dataVpn < 0)) Then
            If ExecLastFault() Then fill = WARN
            ws.Shapes("P12_pte" & CStr(v)).Line.ForeColor.RGB = LAMP_ON
            ws.Shapes("P12_pte" & CStr(v)).Line.Weight = 2.5
        Else
            ws.Shapes("P12_pte" & CStr(v)).Line.ForeColor.RGB = EDGE
            ws.Shapes("P12_pte" & CStr(v)).Line.Weight = 1
        End If
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
    Dim i As Long, fill As Long, txt As String, b As Long
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
           CStr(KeyboardTail()) & "   leidas " & CStr(KeyboardConsumed())
    SetBox ws, "lcd", LcdText()
    SetBox ws, "lcdf", "LCD 8x1  |  video E8-EFh   cursor " & CStr(LcdCursor())
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
    fills = CacheRamFills() > mPrevFills
    SetArrow ws, "a_va", ran And kind <> "ERROR"
    SetArrow ws, "a_pa", ran And kind <> "ERROR"
    SetArrow ws, "a_fill", fills
    SetArrow ws, "a_ram_bus", fills Or kind = "ST"
    SetArrow ws, "a_cpu_bus", fills Or kind = "ST" Or kind = "IN" Or kind = "OUT"
    SetArrow ws, "a_swap", ExecLastFault()
    SetArrow ws, "a_kbd", kind = "IN"
    SetArrow ws, "a_lcd", kind = "OUT"
    mPrevFills = CacheRamFills()
End Sub

Private Sub RefreshListing(ByVal ws As Worksheet)
    Dim first As Long, i As Long, addr As Long, row As Long, pc As Long, n As Long
    pc = CpuSnapshot().PC
    n = LoaderInstructionCount() + 1                      ' include the first HLT of the fill
    If n > 64 Then n = 64
    first = (pc \ 4) - 4
    If first > n - LIST_ROWS Then first = n - LIST_ROWS
    If first < 0 Then first = 0
    For i = 0 To LIST_ROWS - 1
        row = ROW_BASE + LIST_ROW + i
        ws.Range(ws.Cells(row, 2), ws.Cells(row, 8)).Value = vbNullString
        ws.Range(ws.Cells(row, 2), ws.Cells(row, 8)).Interior.Color = CARD
        ws.Range(ws.Cells(row, 2), ws.Cells(row, 8)).Font.Color = INK
        If first + i < n Then
            addr = (first + i) * 4
            ws.Cells(row, 2).Value = AsmHex3(addr)
            ws.Cells(row, 3).Value = Right$("00" & Hex$(LoaderImageByte(addr)), 2) & " " & _
                                     Right$("00" & Hex$(LoaderImageByte(addr + 1)), 2) & " " & _
                                     Right$("00" & Hex$(LoaderImageByte(addr + 2)), 2) & " " & _
                                     Right$("00" & Hex$(LoaderImageByte(addr + 3)), 2)
            ws.Cells(row, 4).Value = InstructionAt(addr)
            If addr = pc Then
                ws.Range(ws.Cells(row, 2), ws.Cells(row, 8)).Interior.Color = LAMP_ON
                ws.Range(ws.Cells(row, 2), ws.Cells(row, 8)).Font.Color = BG
                ws.Cells(row, 4).Value = InstructionAt(addr) & "      <- PC"
            End If
        End If
        ws.Cells(row, 10).Value = vbNullString
        ws.Range(ws.Cells(row, 10), ws.Cells(row, 19)).Interior.Color = CARD
        If i < ExecTraceCount() Then
            ws.Cells(row, 10).Value = ExecTraceLine(i)
            If InStr(ExecTraceLine(i), "ERROR") > 0 Then
                ws.Range(ws.Cells(row, 10), ws.Cells(row, 19)).Interior.Color = BAD
            ElseIf InStr(ExecTraceLine(i), "PAGE_FAULT") > 0 Or InStr(ExecTraceLine(i), "IN_EMPTY") > 0 Then
                ws.Range(ws.Cells(row, 10), ws.Cells(row, 19)).Interior.Color = WARN
            ElseIf i = 0 Then
                ws.Range(ws.Cells(row, 10), ws.Cells(row, 19)).Interior.Color = HEAD
            End If
        End If
    Next i
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

' Shape factories: coordinates in design units (x scaled to the band width, y in points).
Private Sub Frame(ByVal ws As Worksheet, ByVal name As String, ByVal x As Double, ByVal y As Double, _
                  ByVal w As Double, ByVal h As Double, ByVal title As String)
    Dim s As Shape
    Set s = Box(ws, name, x, y, w, h, BG, title, 10, True)
    s.Line.ForeColor.RGB = ACCENT
    s.Line.Weight = 1.25
    s.TextFrame2.VerticalAnchor = msoAnchorTop
    s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ACCENT
    s.TextFrame2.MarginTop = 4
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
    s.Placement = xlFreeFloating
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
        .ForeColor.RGB = ARROW_OFF
        .Weight = 4
        .EndArrowheadStyle = msoArrowheadTriangle
        If both Then .BeginArrowheadStyle = msoArrowheadTriangle
    End With
End Sub

Private Sub SetArrow(ByVal ws As Worksheet, ByVal name As String, ByVal active As Boolean)
    ws.Shapes("P12_" & name).Line.ForeColor.RGB = IIf(active, ARROW_ON, ARROW_OFF)
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
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = caption
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = size
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
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

Private Function MachineSheet(ByVal create As Boolean) As Worksheet
    Set MachineSheet = LcaBand("P12_btn_MaquinaPaso", create)
End Function

Private Function RA(ByVal address As String) As String
    RA = LcaRows(address, ROW_BASE)
End Function
