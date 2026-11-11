Attribute VB_Name = "LcaArithPanel"
Option Explicit

' P15 block of the LCA dashboard: five arithmetic programs run on the real CPU and drawn as a data factory.
'   - four register TANKS (AX..DX) that fill and drain with their value, bit string underneath;
'   - the ALU in the middle, three flag lamps, data memory ([100h], [101h]) on the right;
'   - every instruction animates the real path of its data with LcaFlow packets: operands travel from
'     their tanks into the ALU, the result travels back; MOV/LD/ST carry the byte; a taken jump loops back;
'   - loop gauge, program listing with the PC row, a line chart of the registers over time and a table of the
'     last 14 steps (the documented register trace). The full trace is also on the Logs sheet.
Private Const ROW_BASE As Long = LCA_BASE_ARITH
Private Const COL_BASE As Long = LCA_COL_ARITH
Private Const BAND_ROWS As Long = 88
Private Const DIAGRAM_ROW As Long = 19
Private Const DESIGN_W As Double = 1380
Private Const DIAGRAM_H As Double = 500
Private Const TRACE_ROWS As Long = 14
Private Const DATA_COL As Long = 85          ' chart data (hidden columns CG..CK)
Private Const MAX_LOG As Long = 99

Private Const BG As Long = 2758415           ' RGB(15, 23, 42)
Private Const CARD As Long = 3877150         ' RGB(30, 41, 59)
Private Const HEAD As Long = 6897694         ' RGB(30, 64, 105)
Private Const INK As Long = 15788258         ' RGB(226, 232, 240)
Private Const MUTED As Long = 12100500       ' RGB(148, 163, 184)
Private Const ACCENT As Long = 16569213      ' RGB(125, 211, 252)
Private Const LAMP_ON As Long = 2477306      ' RGB(250, 204, 37)
Private Const EDGE As Long = 5587251         ' RGB(51, 65, 85)
Private Const OK_GREEN As Long = 3433750     ' RGB(22, 101, 52)
Private Const WARN As Long = 1592985         ' RGB(153, 78, 24)
Private Const BAD As Long = 1908095          ' RGB(127, 29, 29)
Private Const BUTTON As Long = 15426341      ' RGB(37, 99, 235)

Private mX0 As Double
Private mY0 As Double
Private mSX As Double
Private mCur As Long
Private mStarted As Boolean
Private mMessage As String
Private mLoops As Long
Private mLogN As Long
Private mLogInstr(0 To MAX_LOG) As String
Private mLogEff(0 To MAX_LOG) As String
Private mLogReg(0 To MAX_LOG, 0 To 3) As Long
Private mLogFl(0 To MAX_LOG) As Long

' --- Build -------------------------------------------------------------------------------------

Public Sub ConstruirPanelAritmetica()
    Dim ws As Worksheet, i As Long, tr As Long, prev As Boolean, k As Long
    prev = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set ws = LcaSheet(True)
    LcaPrepareBand ws, ROW_BASE, BAND_ROWS, "P15_", COL_BASE
    LcaRowHeight ws.Rows(RA("1:" & CStr(BAND_ROWS))), 18
    With ws.Range(RA("A1:T" & CStr(BAND_ROWS)))
        .Interior.Color = BG
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Color = INK
        .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Columns(DATA_COL), ws.Columns(DATA_COL + 4)).Hidden = True

    LcaRowHeight ws.Rows(ROW_BASE + 2), 34
    ws.Range(RA("B2:N2")).Merge
    With ws.Range(RA("B2"))
        .Value = "P15  |  DEMOS ARITMETICOS  -  como viajan los datos"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = ACCENT
    End With
    ws.Range(RA("B3:S3")).Merge
    With ws.Range(RA("B3"))
        .Value = "Cinco programas reales en AX, BX, CX y DX con bucles. Cada instruccion anima como se mueve el dato: " & _
                 "la bolita rosa viaja desde donde esta el dato hasta donde se usa."
        .Font.Color = MUTED
    End With
    For i = 0 To ARITH_COUNT - 1
        Btn ws, Choose(i + 1, "B6:D7", "E6:G7", "H6:J7", "K6:M7", "N6:P7"), ArithName(i), "ArithCargar" & CStr(i), HEAD
    Next i
    Btn ws, "Q6:S7", "VER EN MAQUINA", "LcaIrMaquina", EDGE
    Btn ws, "B9:D10", "PASO / STEP", "ArithPaso", BUTTON
    Btn ws, "E9:G10", "EJECUTAR / RUN", "ArithEjecutar", OK_GREEN
    Btn ws, "H9:J10", "RESET", "ArithReset", EDGE
    Btn ws, "K9:M10", "VELOCIDAD", "ArithVelocidad", CARD
    ws.Range(RA("N9:S10")).Merge
    With ws.Range(RA("N9"))
        .HorizontalAlignment = xlCenter
        .Font.Size = 15
        .Font.Bold = True
    End With

    ' Cards: input / expected / actual, algorithm, message.
    InfoCard ws, "B12:G12", "B13:G14", "ENTRADA / INPUT"
    InfoCard ws, "H12:M12", "H13:M14", "RESULTADO ESPERADO"
    InfoCard ws, "N12:S12", "N13:S14", "RESULTADO OBTENIDO"
    ws.Range(RA("B15:S17")).Merge
    With ws.Range(RA("B15"))
        .WrapText = True
        .Interior.Color = CARD
        .Font.Size = 10
        .IndentLevel = 1
    End With
    ws.Range(RA("B18:S18")).Merge
    ws.Range(RA("B18")).Font.Color = ACCENT

    Geometry ws
    DrawDiagram ws

    ' Lower part: listing + chart, trace table.
    tr = TableRow(ws)
    Title ws, "B" & CStr(tr - 1) & ":H" & CStr(tr - 1), "PROGRAMA  (fila amarilla = PC, el inicio del bucle esta marcado)"
    Title ws, "J" & CStr(tr - 1) & ":S" & CStr(tr - 1), "TRAZA DE REGISTROS EN EL TIEMPO  (una marca por instruccion)"
    For i = 0 To 15
        ws.Range(CL(ws, ROW_BASE + tr + i, 4), CL(ws, ROW_BASE + tr + i, 8)).Merge
    Next i
    With ws.Range(CL(ws, ROW_BASE + tr, 2), CL(ws, ROW_BASE + tr + 15, 8))
        .Font.Name = "Consolas"
        .Font.Size = 10
        .Interior.Color = CARD
    End With
    AddChart ws, tr
    Title ws, "B" & CStr(tr + 17) & ":S" & CStr(tr + 17), "TABLA DE LA TRAZA  (ultimos " & CStr(TRACE_ROWS) & " pasos: instruccion ejecutada y registros resultantes)"
    Dim hdr As Variant
    hdr = Array("PASO", "", "", "", "", "AX", "BX", "CX", "DX", "ZF", "CF", "SF", "EFECTO")
    ws.Range(CL(ws, ROW_BASE + tr + 18, 3), CL(ws, ROW_BASE + tr + 18, 6)).Merge
    ws.Range(CL(ws, ROW_BASE + tr + 18, 14), CL(ws, ROW_BASE + tr + 18, 19)).Merge
    For k = 0 To 12
        If Len(hdr(k)) > 0 Then CL(ws, ROW_BASE + tr + 18, 2 + k).Value = hdr(k)
    Next k
    CL(ws, ROW_BASE + tr + 18, 3).Value = "INSTRUCCION"
    With ws.Range(CL(ws, ROW_BASE + tr + 18, 2), CL(ws, ROW_BASE + tr + 18, 19))
        .Interior.Color = HEAD
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    For i = 1 To TRACE_ROWS
        ws.Range(CL(ws, ROW_BASE + tr + 18 + i, 3), CL(ws, ROW_BASE + tr + 18 + i, 6)).Merge
        ws.Range(CL(ws, ROW_BASE + tr + 18 + i, 14), CL(ws, ROW_BASE + tr + 18 + i, 19)).Merge
    Next i
    With ws.Range(CL(ws, ROW_BASE + tr + 19, 2), CL(ws, ROW_BASE + tr + 18 + TRACE_ROWS, 19))
        .Interior.Color = CARD
        .Font.Name = "Consolas"
        .Font.Size = 9
        .Borders(xlInsideHorizontal).Color = BG
    End With
    ws.Range(CL(ws, ROW_BASE + tr + 19, 2), CL(ws, ROW_BASE + tr + 18 + TRACE_ROWS, 2)).HorizontalAlignment = xlCenter
    ws.Range(CL(ws, ROW_BASE + tr + 19, 7), CL(ws, ROW_BASE + tr + 18 + TRACE_ROWS, 13)).HorizontalAlignment = xlCenter

    mStarted = False
    mCur = 0
    ArithSeleccionar 0
    Application.ScreenUpdating = prev
End Sub

Public Sub MostrarPanelAritmetica()
    If ArithSheet(False) Is Nothing Then
        ConstruirPanelAritmetica
    Else
        LcaGoto ArithSheet(False), ROW_BASE, COL_BASE
        ArithPanelRefrescar
    End If
End Sub

Private Sub DrawDiagram(ByVal ws As Worksheet)
    Dim i As Long, x As Double, s As Shape
    Lbl ws, "lbl_instr", 20, 0, 420, 14, "PROXIMA INSTRUCCION (PC)"
    Box ws, "instr", 20, 16, 420, 72, CARD, "-", 20, True
    Lbl ws, "lbl_loop", 470, 0, 470, 14, "VUELTAS DEL BUCLE"
    Box ws, "loop", 470, 16, 470, 44, CARD, "sin empezar", 14, True
    Set s = Box(ws, "loopbg", 470, 66, 470, 22, BG, "", 8, False, msoShapeRectangle)
    s.Line.ForeColor.RGB = LAMP_ON
    Set s = Box(ws, "loopbar", 472, 68, 4, 18, LAMP_ON, "", 8, False, msoShapeRectangle)
    s.Line.Visible = msoFalse

    Box ws, "alu", 430, 130, 330, 110, EDGE, "ALU", 18, True, msoShapeFlowchartManualOperation
    Lbl ws, "lbl_flags", 800, 128, 190, 16, "FLAGS  ZF  CF  SF"
    Box ws, "fz", 800, 150, 46, 46, CARD, "Z", 14, True, msoShapeOval
    Box ws, "fc", 856, 150, 46, 46, CARD, "C", 14, True, msoShapeOval
    Box ws, "fs", 912, 150, 46, 46, CARD, "S", 14, True, msoShapeOval

    Set s = Box(ws, "mem", 1000, 0, 360, 244, BG, "MEMORIA DE DATOS  (VPN 8)", 10, True)
    s.Line.ForeColor.RGB = ACCENT
    s.TextFrame2.VerticalAnchor = msoAnchorTop
    s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ACCENT
    Box ws, "mem0", 1020, 36, 320, 92, CARD, "[100h]", 14, True
    Box ws, "mem1", 1020, 138, 320, 92, CARD, "[101h]", 14, True
    Box ws, "res", 1000, 262, 360, 118, CARD, "", 12, True
    Set s = Box(ws, "leg", 1000, 392, 360, 100, CARD, "COMO LEERLO" & vbCr & "bolita rosa = el dato viajando" & vbCr & _
                "tanque lleno = valor del registro" & vbCr & "amarillo = lo que se esta usando ahora" & vbCr & _
                "ALU = el calculo con sus numeros", 9)
    s.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    s.TextFrame2.TextRange.Paragraphs(1).Font.Bold = msoTrue

    For i = 0 To 3
        x = 20 + i * 240
        Set s = Box(ws, "tank" & CStr(i), x, 270, 220, 190, BG, "", 10)
        s.Line.ForeColor.RGB = RegColor(i)
        s.Line.Weight = 2.5
        Set s = Box(ws, "liq" & CStr(i), x + 10, 450, 200, 4, RegColor(i), "", 6, False, msoShapeRectangle)
        s.Line.Visible = msoFalse
        s.Fill.Transparency = 0.15
        Set s = Box(ws, "nm" & CStr(i), x + 10, 274, 200, 26, BG, Choose(i + 1, "AX", "BX", "CX", "DX"), 17, True)
        s.Fill.Visible = msoFalse
        s.Line.Visible = msoFalse
        s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RegColor(i)
        Set s = Box(ws, "val" & CStr(i), x + 10, 302, 200, 32, BG, "00h   0", 16, True)
        s.Fill.Visible = msoFalse
        s.Line.Visible = msoFalse
        Set s = Box(ws, "bits" & CStr(i), x, 464, 220, 24, BG, "0000 0000", 12, True)
        s.Fill.Visible = msoFalse
        s.Line.Visible = msoFalse
        s.TextFrame2.TextRange.Font.Name = "Consolas"
        s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = MUTED
    Next i
End Sub

' --- Actions -----------------------------------------------------------------------------------

Public Sub ArithCargar0()
    ArithSeleccionar 0
End Sub

Public Sub ArithCargar1()
    ArithSeleccionar 1
End Sub

Public Sub ArithCargar2()
    ArithSeleccionar 2
End Sub

Public Sub ArithCargar3()
    ArithSeleccionar 3
End Sub

Public Sub ArithCargar4()
    ArithSeleccionar 4
End Sub

' Loads program i, resets the machine and the trace, and shows its input / expected result.
Public Sub ArithSeleccionar(ByVal i As Long)
    Dim ws As Worksheet
    Set ws = ArithSheet(False)
    If ws Is Nothing Then Exit Sub
    Geometry ws
    If Not ArithLoad(i) Then
        mMessage = "No se pudo ensamblar el programa."
        RefreshAll ws
        Exit Sub
    End If
    mCur = i
    mStarted = True
    mLoops = 0
    ClearLog ws
    Record ws, "(inicio)", "valores iniciales"
    SetChartRange ws, ArithExpectedSteps(i) + 1, ArithChartMax(i)
    mMessage = ArithTitle(i) & ".  Pulsa PASO para ver una instruccion o EJECUTAR para todo el programa."
    RefreshAll ws
    AfterAction
End Sub

Public Sub ArithPaso()
    Dim ws As Worksheet
    Set ws = ArithSheet(False)
    If ws Is Nothing Then Exit Sub
    Geometry ws
    If Not mStarted Then ArithSeleccionar 0
    If ExecStatus() = EXEC_HALTED Then
        mMessage = "El programa ya termino (HLT). Pulsa RESET para repetirlo."
    ElseIf ExecStatus() = EXEC_ERROR Then
        mMessage = "ERROR: " & ExecErrorText() & ". Pulsa RESET."
    Else
        DoStep ws, True
    End If
    RefreshAll ws
    AfterAction
End Sub

Public Sub ArithEjecutar()
    Dim ws As Worksheet, n As Long, keep As Long, animate As Boolean
    Set ws = ArithSheet(False)
    If ws Is Nothing Then Exit Sub
    Geometry ws
    If Not mStarted Then ArithSeleccionar 0
    If ExecStatus() = EXEC_HALTED Then ArithSeleccionar mCur
    keep = FlowSpeed()
    animate = (keep > 0)
    If animate And keep < 3 Then FlowSetSpeed keep + 1
    Do While n < 300
        If Not DoStep(ws, animate) Then Exit Do
        n = n + 1
        If animate Then RefreshAll ws
        If ExecStatus() = EXEC_HALTED Then Exit Do
    Loop
    FlowSetSpeed keep
    mMessage = CStr(n) & " instrucciones ejecutadas hasta HLT."
    RefreshAll ws
    AfterAction
End Sub

Public Sub ArithReset()
    ArithSeleccionar mCur
End Sub

Public Sub ArithVelocidad()
    Dim ws As Worksheet
    FlowCycleSpeed
    mMessage = "Animacion: " & FlowSpeedText() & "  (la misma velocidad rige la maquina y los demas bloques)."
    Set ws = ArithSheet(False)
    If Not ws Is Nothing Then RefreshAll ws
End Sub

Public Sub ArithPanelRefrescar()
    Dim ws As Worksheet
    Set ws = ArithSheet(False)
    If ws Is Nothing Then Exit Sub
    If Not mStarted Then Exit Sub
    Geometry ws
    RefreshAll ws
End Sub

' One instruction. The pre-step register values feed the packets; tanks refill after the trip.
Private Function DoStep(ByVal ws As Worksheet, ByVal animate As Boolean) As Boolean
    Dim pre(0 To 3) As Long, i As Long, pc As Long, ok As Boolean, t As String
    pc = CpuSnapshot().PC
    For i = 0 To 3
        pre(i) = CpuReadRegister(i)
    Next i
    ok = ExecStep()
    If ExecStatus() = EXEC_ERROR Then
        mMessage = "ERROR: " & ExecErrorText()
        Exit Function
    End If
    If Not ok Then Exit Function
    If pc = ArithLoopHead(mCur) Then mLoops = mLoops + 1
    t = ExecLastText()
    Record ws, InstrOf(t), EffectOf(t)
    If animate Then AnimateStep ws, pre, pc
    DoStep = True
End Function

' --- Animation ---------------------------------------------------------------------------------

Private Sub AnimateStep(ByVal ws As Worksheet, ByRef pre() As Long, ByVal pc As Long)
    Dim op As Long, r As Long, x As Long, va As Long, post As Long, i As Long
    If Not FlowOn() Then Exit Sub
    Application.ScreenUpdating = True
    op = LoaderImageByte(pc)
    r = LoaderImageByte(pc + 1)
    x = LoaderImageByte(pc + 2)
    va = 256 * LoaderImageByte(pc + 3) + x
    FlowPulse ws, "P15_instr", LAMP_ON, 1
    Select Case op
        Case &H10
            FlowSpawn ws, "a", HexNum(x), RegColor(r), "P15_instr"
            FlowMoveTo ws, Array("a"), Array(TankPort(r))
            FlowKill ws, "a"
            AnimateLevel ws, r, pre(r), CpuReadRegister(r)
        Case &H11
            FlowSpawn ws, "a", CStr(pre(x)), RegColor(x), TankPort(x)
            FlowMoveTo ws, Array("a"), Array(TankPort(r))
            FlowKill ws, "a"
            AnimateLevel ws, r, pre(r), CpuReadRegister(r)
        Case &H12
            FlowSpawn ws, "a", CStr(CpuReadRegister(r)), FlowPacketColor(), "P15_" & MemName(va)
            FlowPulse ws, "P15_" & MemName(va), LAMP_ON, 1
            FlowMoveTo ws, Array("a"), Array(TankPort(r))
            FlowKill ws, "a"
            AnimateLevel ws, r, pre(r), CpuReadRegister(r)
        Case &H13
            FlowSpawn ws, "a", CStr(pre(r)), RegColor(r), TankPort(r)
            FlowMoveTo ws, Array("a"), Array("P15_" & MemName(va))
            FlowKill ws, "a"
            RefreshMemory ws
            FlowPulse ws, "P15_" & MemName(va), LAMP_ON, 2
        Case &H20 To &H24
            AnimateAlu ws, op, r, x, pre
        Case &H30 To &H34
            If ExecLastTaken() Then
                FlowSend ws, AsmHex3(va), FlowPacketColor(), "P15_instr", "P15_loop", "P15_instr"
                FlowPulse ws, "P15_loop", LAMP_ON, 2
            Else
                FlowPulse ws, "P15_loop", EDGE, 1
            End If
        Case &H1
            For i = 0 To 3
                FlowPulse ws, "P15_tank" & CStr(i), OK_GREEN, 1
            Next i
            FlowPulse ws, "P15_res", OK_GREEN, 2
    End Select
    SetText ws, "alu", "ALU"
End Sub

Private Sub AnimateAlu(ByVal ws As Worksheet, ByVal op As Long, ByVal r As Long, ByVal x As Long, ByRef pre() As Long)
    Dim twoOps As Boolean, a As Long, b As Long, post As Long, sym As String, eq As String, nm As String
    twoOps = (op <= &H22)
    a = pre(r)
    If twoOps Then b = pre(x) Else b = 1
    post = CpuReadRegister(r)
    sym = Choose(op - &H1F, "+", "-", "cmp", "+", "-")
    nm = Choose(op - &H1F, "ADD", "SUB", "CMP", "INC", "DEC")
    FlowSpawn ws, "a", CStr(a), RegColor(r), TankPort(r)
    If twoOps Then
        FlowSpawn ws, "b", CStr(b), RegColor(x), TankPort(x)
    Else
        FlowSpawn ws, "b", "1", FlowPacketColor(), PT(850, 90)
    End If
    FlowMoveTo ws, Array("a", "b"), Array(PT(485, 165), PT(705, 165))
    FlowKill ws, "a"
    FlowKill ws, "b"
    eq = CStr(a) & " " & sym & " " & CStr(b)
    If op = &H22 Then
        eq = eq & "   (solo flags)"
    Else
        eq = eq & " = " & CStr(post)
    End If
    SetText ws, "alu", nm & vbCr & eq
    FlowPulse ws, "P15_alu", LAMP_ON, 3
    SetFlags ws
    FlowPulse ws, "P15_fz", LAMP_ON, 1
    If op <> &H22 Then
        FlowSpawn ws, "c", CStr(post), RegColor(r), PT(595, 232)
        FlowMoveTo ws, Array("c"), Array(TankPort(r))
        FlowKill ws, "c"
        AnimateLevel ws, r, pre(r), post
    End If
End Sub

' The tank fills (or drains) to its new level.
Private Sub AnimateLevel(ByVal ws As Worksheet, ByVal r As Long, ByVal fromV As Long, ByVal toV As Long)
    Dim k As Long, v As Double
    For k = 1 To 8
        v = fromV + (toV - fromV) * k / 8
        SetLiquid ws, r, CLng(v)
        FlowPause 1
    Next k
    SetLiquid ws, r, toV
    SetText ws, "val" & CStr(r), AsmHex2(toV) & "   " & CStr(toV)
    SetText ws, "bits" & CStr(r), Bin8(toV)
    FlowPulse ws, "P15_tank" & CStr(r), RegColor(r), 1
End Sub

' --- Refresh -----------------------------------------------------------------------------------

Private Sub RefreshAll(ByVal ws As Worksheet)
    Dim prev As Boolean, i As Long, v As Long, st As String, label As String, fill As Long
    Dim pc As Long, total As Long, w As Double
    prev = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Geometry ws
    For i = 0 To 3
        v = CpuReadRegister(i)
        SetLiquid ws, i, v
        SetText ws, "val" & CStr(i), AsmHex2(v) & "   " & CStr(v)
        SetText ws, "bits" & CStr(i), Bin8(v)
    Next i
    SetFlags ws
    RefreshMemory ws
    pc = CpuSnapshot().PC
    SetText ws, "instr", InstructionAt(pc) & vbCr & "PC = " & AsmHex3(pc)
    ws.Shapes("P15_instr").TextFrame2.TextRange.Paragraphs(2).Font.Size = 10
    ws.Shapes("P15_instr").TextFrame2.TextRange.Paragraphs(2).Font.Bold = msoFalse
    total = ArithLoopTotal(mCur)
    If mLoops = 0 Then
        SetText ws, "loop", "vuelta 0 de " & CStr(total) & "  (aun no entra al bucle)"
    Else
        SetText ws, "loop", "vuelta " & CStr(IIf(mLoops > total, total, mLoops)) & " de " & CStr(total)
    End If
    w = 466 * IIf(mLoops > total, total, mLoops) / total
    If w < 4 Then w = 4
    ws.Shapes("P15_loopbar").Width = w * mSX

    st = ExecStatus()
    Select Case st
        Case EXEC_HALTED
            If ArithMatches(mCur) Then
                label = "CORRECTO": fill = OK_GREEN
            Else
                label = "TERMINO, DISTINTO": fill = BAD
            End If
        Case EXEC_ERROR: label = "ERROR: " & ExecErrorText(): fill = BAD
        Case Else
            If ExecSteps() = 0 Then
                label = "LISTO": fill = HEAD
            Else
                label = "EN CURSO  paso " & CStr(ExecSteps()): fill = BUTTON
            End If
    End Select
    ws.Range(RA("N9")).Value = label
    ws.Range(RA("N9:S10")).Interior.Color = fill
    SetBox ws, "btn_ArithVelocidad", "VELOCIDAD" & vbCr & FlowSpeedText()
    CL(ws, ROW_BASE + 13, 2).Value = ArithInputText(mCur)
    CL(ws, ROW_BASE + 13, 8).Value = ArithExpectedText(mCur)
    CL(ws, ROW_BASE + 13, 14).Value = ArithResultText(mCur)
    ws.Range(RA("N13:S14")).Interior.Color = IIf(st = EXEC_HALTED, fill, CARD)
    CL(ws, ROW_BASE + 15, 2).Value = "ALGORITMO:  " & ArithAlgorithm(mCur)
    CL(ws, ROW_BASE + 18, 2).Value = mMessage
    If st = EXEC_HALTED Then
        SetText ws, "res", label & vbCr & ArithResultText(mCur), fill
    ElseIf st = EXEC_ERROR Then
        SetText ws, "res", label, BAD
    Else
        SetText ws, "res", "EN CURSO" & vbCr & ArithResultText(mCur), CARD
    End If
    ws.Shapes("P15_res").TextFrame2.TextRange.Paragraphs(1).Font.Size = 18
    RefreshListing ws, pc
    RefreshTrace ws
    Application.ScreenUpdating = prev
End Sub

Private Sub RefreshMemory(ByVal ws As Worksheet)
    Dim i As Long, v As Long
    For i = 0 To 1
        v = ArithMemByte(i)
        If ArithExpectedMem(mCur, i) >= 0 Or v <> 0 Then
            SetText ws, "mem" & CStr(i), "[" & Choose(i + 1, "100h", "101h") & "]" & vbCr & AsmHex2(v) & "   " & CStr(v), IIf(v <> 0, HEAD, CARD)
        Else
            SetText ws, "mem" & CStr(i), "[" & Choose(i + 1, "100h", "101h") & "]" & vbCr & "sin usar", CARD
        End If
    Next i
End Sub

Private Sub SetFlags(ByVal ws As Worksheet)
    Lamp ws, "fz", CpuFlagIsSet(LCA_FLAG_ZF)
    Lamp ws, "fc", CpuFlagIsSet(LCA_FLAG_CF)
    Lamp ws, "fs", CpuFlagIsSet(LCA_FLAG_SF)
End Sub

Private Sub RefreshListing(ByVal ws As Worksheet, ByVal pc As Long)
    Dim tr As Long, i As Long, n As Long, addr As Long, row As Long
    tr = TableRow(ws)
    n = LoaderInstructionCount()
    For i = 0 To 15
        row = ROW_BASE + tr + i
        With ws.Range(CL(ws, row, 2), CL(ws, row, 8))
            .Value = vbNullString
            .Interior.Color = CARD
            .Font.Color = INK
        End With
        If i < n Then
            addr = i * 4
            CL(ws, row, 2).Value = AsmHex3(addr)
            CL(ws, row, 3).Value = Right$("00" & Hex$(LoaderImageByte(addr)), 2) & " " & Right$("00" & Hex$(LoaderImageByte(addr + 1)), 2) & _
                                   " " & Right$("00" & Hex$(LoaderImageByte(addr + 2)), 2) & " " & Right$("00" & Hex$(LoaderImageByte(addr + 3)), 2)
            CL(ws, row, 4).Value = InstructionAt(addr) & IIf(addr = ArithLoopHead(mCur), "      <- inicio del bucle", "")
            If addr = pc Then
                With ws.Range(CL(ws, row, 2), CL(ws, row, 8))
                    .Interior.Color = LAMP_ON
                    .Font.Color = BG
                End With
            End If
        End If
    Next i
End Sub

Private Sub RefreshTrace(ByVal ws As Worksheet)
    Dim tr As Long, i As Long, first As Long, k As Long, row As Long, r As Long
    tr = TableRow(ws)
    first = mLogN - TRACE_ROWS
    If first < 0 Then first = 0
    For i = 0 To TRACE_ROWS - 1
        row = ROW_BASE + tr + 19 + i
        k = first + i
        ws.Range(CL(ws, row, 2), CL(ws, row, 19)).Interior.Color = CARD
        ws.Range(CL(ws, row, 2), CL(ws, row, 19)).Font.Color = INK
        If k < mLogN Then
            CL(ws, row, 2).Value = k
            CL(ws, row, 3).Value = mLogInstr(k)
            For r = 0 To 3
                CL(ws, row, 7 + r).Value = mLogReg(k, r)
            Next r
            CL(ws, row, 11).Value = Abs(CLng((mLogFl(k) And 1) <> 0))
            CL(ws, row, 12).Value = Abs(CLng((mLogFl(k) And 2) <> 0))
            CL(ws, row, 13).Value = Abs(CLng((mLogFl(k) And 4) <> 0))
            CL(ws, row, 14).Value = mLogEff(k)
            If k = mLogN - 1 Then ws.Range(CL(ws, row, 2), CL(ws, row, 19)).Interior.Color = HEAD
            If InStr(mLogEff(k), "PAGE_FAULT") > 0 Then ws.Range(CL(ws, row, 2), CL(ws, row, 19)).Interior.Color = WARN
        Else
            ws.Range(CL(ws, row, 2), CL(ws, row, 19)).ClearContents
        End If
    Next i
End Sub

' --- Trace log (also the chart source) ------------------------------------------------------------

Private Sub ClearLog(ByVal ws As Worksheet)
    mLogN = 0
    ws.Range(ws.Cells(1, DATA_COL), ws.Cells(MAX_LOG + 2, DATA_COL + 4)).ClearContents
    ws.Cells(1, DATA_COL).Value = "paso"
    ws.Cells(1, DATA_COL + 1).Value = "AX"
    ws.Cells(1, DATA_COL + 2).Value = "BX"
    ws.Cells(1, DATA_COL + 3).Value = "CX"
    ws.Cells(1, DATA_COL + 4).Value = "DX"
End Sub

Private Sub Record(ByVal ws As Worksheet, ByVal instrText As String, ByVal effect As String)
    Dim r As Long, n As Long
    n = mLogN
    If n > MAX_LOG Then Exit Sub
    mLogInstr(n) = instrText
    mLogEff(n) = effect
    mLogFl(n) = CLng(CpuSnapshot().flags)
    ws.Cells(2 + n, DATA_COL).Value = n
    For r = 0 To 3
        mLogReg(n, r) = CpuReadRegister(r)
        ws.Cells(2 + n, DATA_COL + 1 + r).Value = mLogReg(n, r)
    Next r
    mLogN = n + 1
End Sub

Private Function InstrOf(ByVal t As String) As String
    Dim p As Long
    p = InStr(t, "  ->  ")
    If p > 7 Then InstrOf = Trim$(Mid$(t, 7, p - 7)) Else InstrOf = Trim$(t)
End Function

Private Function EffectOf(ByVal t As String) As String
    Dim p As Long
    p = InStr(t, "  ->  ")
    If p > 0 Then EffectOf = Mid$(t, p + 6)
End Function

Private Sub AddChart(ByVal ws As Worksheet, ByVal tr As Long)
    Dim area As Range, co As ChartObject, k As Long
    Set area = ws.Range(CL(ws, ROW_BASE + tr, 10), CL(ws, ROW_BASE + tr + 15, 19))
    Set co = ws.ChartObjects.Add(area.Left + 2, area.Top, area.Width - 4, area.Height)
    co.Name = "P15_chart"
    co.Placement = xlMove
    ClearLog ws
    With co.Chart
        .ChartType = xlLineMarkers
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        For k = 0 To 3
            With .SeriesCollection.NewSeries
                .Name = Choose(k + 1, "AX", "BX", "CX", "DX")
                .Values = ws.Range(ws.Cells(2, DATA_COL + 1 + k), ws.Cells(22, DATA_COL + 1 + k))
                .XValues = ws.Range(ws.Cells(2, DATA_COL), ws.Cells(22, DATA_COL))
                .Format.Line.ForeColor.RGB = RegColor(k)
                .Format.Line.Weight = 2.5
                .MarkerStyle = xlMarkerStyleCircle
                .MarkerSize = 6
                .MarkerBackgroundColor = RegColor(k)
                .MarkerForegroundColor = RegColor(k)
            End With
        Next k
        .DisplayBlanksAs = xlNotPlotted
        .PlotVisibleOnly = False
        .HasTitle = False
        .HasLegend = True
        .Legend.Position = xlLegendPositionTop
        .Legend.Font.Color = INK
        .Legend.Font.Bold = True
        .ChartArea.Format.Fill.ForeColor.RGB = CARD
        .ChartArea.Format.Line.ForeColor.RGB = ACCENT
        .PlotArea.Format.Fill.ForeColor.RGB = CARD
        .Axes(xlValue).MinimumScale = 0
        .Axes(xlValue).TickLabels.Font.Color = MUTED
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = EDGE
        .Axes(xlCategory).TickLabels.Font.Color = MUTED
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "instruccion ejecutada (paso)"
        .Axes(xlCategory).AxisTitle.Font.Color = MUTED
    End With
End Sub

Private Sub SetChartRange(ByVal ws As Worksheet, ByVal points As Long, ByVal maxScale As Long)
    Dim k As Long
    On Error Resume Next
    With ws.ChartObjects("P15_chart").Chart
        For k = 1 To 4
            .SeriesCollection(k).Values = ws.Range(ws.Cells(2, DATA_COL + k), ws.Cells(1 + points, DATA_COL + k))
            .SeriesCollection(k).XValues = ws.Range(ws.Cells(2, DATA_COL), ws.Cells(1 + points, DATA_COL))
        Next k
        .Axes(xlValue).MaximumScale = maxScale
        .Axes(xlCategory).TickLabelSpacing = IIf(points > 30, 5, 1)
    End With
    On Error GoTo 0
End Sub

' --- Helpers -----------------------------------------------------------------------------------

Private Sub AfterAction()
    On Error Resume Next
    LogsActualizar
End Sub

Private Sub Geometry(ByVal ws As Worksheet)
    mX0 = CL(ws, ROW_BASE + DIAGRAM_ROW, 2).Left
    mY0 = CL(ws, ROW_BASE + DIAGRAM_ROW, 2).Top
    mSX = (CL(ws, ROW_BASE + DIAGRAM_ROW, 20).Left - mX0) / DESIGN_W
End Sub

' First block row whose top is below the drawing; its title sits on the row above.
Private Function TableRow(ByVal ws As Worksheet) As Long
    Dim r As Long, y0 As Double
    y0 = CL(ws, ROW_BASE + DIAGRAM_ROW, 2).Top
    r = DIAGRAM_ROW
    Do While CL(ws, ROW_BASE + r, 2).Top - y0 < DIAGRAM_H + 6 And r < BAND_ROWS - 36
        r = r + 1
    Loop
    TableRow = r + 1
End Function

Private Function PT(ByVal x As Double, ByVal y As Double) As Variant
    PT = Array(mX0 + x * mSX, mY0 + y)
End Function

Private Function TankPort(ByVal i As Long) As Variant
    TankPort = PT(20 + i * 240 + 110, 330)
End Function

Private Function MemName(ByVal va As Long) As String
    If va = &H101 Then MemName = "mem1" Else MemName = "mem0"
End Function

Private Function RegColor(ByVal i As Long) As Long
    RegColor = Choose(i + 1, RGB(56, 189, 248), RGB(52, 211, 153), RGB(251, 146, 60), RGB(232, 121, 249))
End Function

Private Function HexNum(ByVal v As Long) As String
    HexNum = Right$("00" & Hex$(v), 2) & "h"
End Function

Private Function Bin8(ByVal v As Long) As String
    Dim b As Long, s As String
    For b = 7 To 0 Step -1
        s = s & IIf((v And (2 ^ b)) <> 0, "1", "0")
        If b = 4 Then s = s & " "
    Next b
    Bin8 = s
End Function

Private Function InstructionAt(ByVal addr As Long) As String
    If addr < 0 Or addr > 252 Then
        InstructionAt = "-"
    Else
        InstructionAt = AsmDisassemble(LoaderImageByte(addr), LoaderImageByte(addr + 1), _
                                       LoaderImageByte(addr + 2), LoaderImageByte(addr + 3))
    End If
End Function

Private Sub SetLiquid(ByVal ws As Worksheet, ByVal i As Long, ByVal v As Long)
    Dim h As Double
    If v = 0 Then h = 0 Else h = 10 + 110 * v / 255
    With ws.Shapes("P15_liq" & CStr(i))
        If h < 1 Then
            .Visible = msoFalse
        Else
            .Visible = msoTrue
            .Height = h
            .Top = mY0 + 450 - h
        End If
    End With
End Sub

Private Sub Lamp(ByVal ws As Worksheet, ByVal name As String, ByVal isOn As Boolean)
    With ws.Shapes("P15_" & name)
        .Fill.ForeColor.RGB = IIf(isOn, LAMP_ON, CARD)
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = IIf(isOn, BG, MUTED)
    End With
End Sub

Private Sub SetText(ByVal ws As Worksheet, ByVal name As String, ByVal text As String, Optional ByVal fill As Long = -1)
    With ws.Shapes("P15_" & name)
        If .TextFrame2.TextRange.Text <> text Then .TextFrame2.TextRange.Text = text
        If fill >= 0 Then .Fill.ForeColor.RGB = fill
    End With
End Sub

Private Sub SetBox(ByVal ws As Worksheet, ByVal name As String, ByVal text As String)
    On Error Resume Next
    ws.Shapes("P15_" & name).TextFrame2.TextRange.Text = text
    On Error GoTo 0
End Sub

Private Sub Lbl(ByVal ws As Worksheet, ByVal name As String, ByVal x As Double, ByVal y As Double, _
                ByVal w As Double, ByVal h As Double, ByVal text As String)
    Dim s As Shape
    Set s = Box(ws, name, x, y, w, h, BG, text, 9, True)
    s.Fill.Visible = msoFalse
    s.Line.Visible = msoFalse
    s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = MUTED
End Sub

Private Function Box(ByVal ws As Worksheet, ByVal name As String, ByVal x As Double, ByVal y As Double, _
                     ByVal w As Double, ByVal h As Double, ByVal fill As Long, ByVal text As String, _
                     ByVal size As Single, Optional ByVal bold As Boolean = False, _
                     Optional ByVal kind As Long = msoShapeRoundedRectangle) As Shape
    Dim s As Shape
    Set s = ws.Shapes.AddShape(kind, mX0 + x * mSX, mY0 + y, w * mSX, h)
    s.Name = "P15_" & name
    s.Fill.ForeColor.RGB = fill
    s.Line.ForeColor.RGB = EDGE
    If kind = msoShapeRoundedRectangle Then s.Adjustments.Item(1) = 0.1
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

Private Sub Btn(ByVal ws As Worksheet, ByVal address As String, ByVal caption As String, _
                ByVal macroName As String, ByVal fill As Long)
    Dim area As Range, s As Shape
    Set area = ws.Range(RA(address))
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left + 2, area.Top + 2, area.Width - 4, area.Height - 4)
    s.Name = "P15_btn_" & macroName
    s.Fill.ForeColor.RGB = fill
    s.Line.ForeColor.RGB = ACCENT
    s.Line.Weight = 0.75
    With s.TextFrame2
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = caption
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 11
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    s.Shadow.Visible = msoTrue
    s.Shadow.OffsetX = 1.5
    s.Shadow.OffsetY = 1.5
    s.Shadow.Transparency = 0.5
    s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Sub

Private Sub InfoCard(ByVal ws As Worksheet, ByVal headAddr As String, ByVal bodyAddr As String, ByVal title As String)
    ws.Range(RA(headAddr)).Merge
    With ws.Range(RA(headAddr))
        .Value = title
        .Interior.Color = HEAD
        .Font.Bold = True
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
    End With
    ws.Range(RA(bodyAddr)).Merge
    With ws.Range(RA(bodyAddr))
        .Interior.Color = CARD
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .Font.Size = 11
        .Font.Bold = True
    End With
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

Private Function ArithSheet(ByVal create As Boolean) As Worksheet
    Set ArithSheet = LcaBand("P15_btn_ArithPaso", create)
End Function

Private Function RA(ByVal address As String) As String
    RA = LcaShift(address, ROW_BASE, COL_BASE)
End Function

' Cell of this block: row is absolute (ROW_BASE already added), column is relative to the block.
Private Function CL(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Range
    Set CL = ws.Cells(r, c + COL_BASE)
End Function
