Attribute VB_Name = "LcaMemPanel"
Option Explicit

' P17 block of the LCA dashboard (column 4, below the P09 caches): the memory demo running on the real CPU.
' It draws, live and animated with the shared LcaFlow packet:
'   - the 7 pageable RAM frames (owner VPN, dirty / referenced, CLOCK hand) and PFN 7 (I/O);
'   - the 15 virtual pages on the swap (resident, on disk, last fault, last victim);
'   - L1I, L1D and L2 lines with the physical line each one holds;
'   - the access sequence (one row per instruction: page fault, victim, write-back, L1/L2 result);
'   - the expected counters of the demo against the observed ones.
Private Const ROW_BASE As Long = LCA_BASE_MEM
Private Const COL_BASE As Long = LCA_COL_MEM
Private Const BAND_ROWS As Long = 68
Private Const R_CARD As Long = 8
Private Const R_MSG As Long = 12
Private Const R_FRAMES As Long = 14
Private Const R_SWAP As Long = 21
Private Const R_CACHE As Long = 28
Private Const R_TABLE As Long = 36
Private Const R_FIRST As Long = 37

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
Private Const DATA_FILL As Long = 7954958    ' RGB(14, 98, 121)

Private mMessage As String

' --- Build -------------------------------------------------------------------------------------

Public Sub ConstruirPanelMemoria()
    Dim ws As Worksheet, i As Long, prev As Boolean, e As Variant
    prev = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set ws = LcaSheet(True)
    LcaPrepareBand ws, ROW_BASE, BAND_ROWS, "P17_", COL_BASE
    LcaRowHeight ws.Rows(RA("1:" & CStr(BAND_ROWS))), 20
    With ws.Range(RA("A1:T" & CStr(BAND_ROWS)))
        .Interior.Color = BG
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Color = INK
        .VerticalAlignment = xlCenter
    End With
    LcaRowHeight ws.Rows(ROW_BASE + 2), 34
    ws.Range(RA("B2:R2")).Merge
    With ws.Range(RA("B2"))
        .Value = "P17  |  MEMORIA: PAGE FAULTS Y CACHE EN ACCION"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = ACCENT
    End With
    ws.Range(RA("B3:S3")).Merge
    With ws.Range(RA("B3"))
        .Value = "Siete paginas dispersas (faults y misses), una lectura repetida (hits), mas paginas que marcos (la CLOCK desaloja) y un dato modificado que se recupera del swap."
        .Font.Color = MUTED
    End With
    Btn ws, "B5:D6", "CARGAR PROGRAMA", "MemCargar", HEAD
    Btn ws, "E5:G6", "PASO / STEP", "MemPaso", BUTTON
    Btn ws, "H5:J6", "EJECUTAR / RUN", "MemEjecutar", OK_GREEN
    Btn ws, "K5:M6", "RESET", "MemReset", EDGE
    Btn ws, "N5:P6", "VELOCIDAD", "MemVelocidad", CARD
    Btn ws, "Q5:S6", "VER EN MAQUINA", "LcaIrMaquina", EDGE

    AddCard ws, "B", "E", "ESTADO / FASE"
    AddCard ws, "F", "G", "FALLOS DE PAGINA"
    AddCard ws, "H", "I", "EVICCIONES"
    AddCard ws, "J", "K", "WRITE-BACKS"
    AddCard ws, "L", "M", "L1D ACIERTOS"
    AddCard ws, "N", "O", "L1D FALLOS"
    AddCard ws, "P", "Q", "L1I A / F"
    AddCard ws, "R", "S", "L2 A / F"
    ws.Range(RA("B" & CStr(R_MSG) & ":S" & CStr(R_MSG))).Merge
    ws.Range(RA("B" & CStr(R_MSG))).Font.Color = ACCENT

    Title ws, "B" & CStr(R_FRAMES - 1) & ":S" & CStr(R_FRAMES - 1), "RAM: 7 MARCOS PAGINABLES (PFN 0-6) + PFN 7 E/S   (borde amarillo = cursor de la CLOCK;  D = dirty, R = referenciada)"
    For i = 0 To 7
        Slot ws, "fr" & CStr(i), R_FRAMES, R_FRAMES + 4, i, 8
    Next i
    Title ws, "B" & CStr(R_SWAP - 1) & ":S" & CStr(R_SWAP - 1), "SWAP: 15 PAGINAS VIRTUALES (VPN 0-14)   (azul = en RAM, gris = solo en disco, rojo = ultima victima, amarillo = ultima falta)"
    For i = 0 To 14
        Slot ws, "pg" & CStr(i), R_SWAP, R_SWAP + 3, i, 15
    Next i
    Title ws, "B" & CStr(R_CACHE - 1) & ":F" & CStr(R_CACHE - 1), "L1I (2 lineas)"
    Title ws, "G" & CStr(R_CACHE - 1) & ":K" & CStr(R_CACHE - 1), "L1D (2 lineas)"
    Title ws, "L" & CStr(R_CACHE - 1) & ":S" & CStr(R_CACHE - 1), "L2 (8 lineas)"
    For i = 0 To 1
        CacheBox ws, "ci" & CStr(i), "B", "F", i, 2
        CacheBox ws, "cd" & CStr(i), "G", "K", i, 2
    Next i
    For i = 0 To 7
        CacheBox ws, "c2" & CStr(i), "L", "S", i, 8
    Next i

    Title ws, "B" & CStr(R_TABLE - 1) & ":M" & CStr(R_TABLE - 1), "SECUENCIA DE ACCESOS  (una fila por instruccion)"
    Title ws, "N" & CStr(R_TABLE - 1) & ":S" & CStr(R_TABLE - 1), "CONTADORES ESPERADOS"
    e = Array("#", "PC", "INSTRUCCION", "", "", "DATO VA", "PAG", "FALLO", "DESALOJA", "", "L1", "L2")
    For i = 0 To 11
        CL(ws, ROW_BASE + R_TABLE, 2 + i).Value = e(i)
    Next i
    With ws.Range(RA("B" & CStr(R_TABLE) & ":M" & CStr(R_TABLE)))
        .Interior.Color = HEAD
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(CL(ws, ROW_BASE + R_FIRST, 2), CL(ws, ROW_BASE + R_FIRST + MEMDEMO_STEPS - 1, 13))
        .Interior.Color = CARD
        .Font.Name = "Consolas"
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
    End With
    ws.Range(CL(ws, ROW_BASE + R_TABLE, 14), CL(ws, ROW_BASE + R_TABLE, 19)).Interior.Color = HEAD
    ws.Range(RA("N" & CStr(R_TABLE) & ":S" & CStr(R_TABLE))).Font.Bold = True
    CL(ws, ROW_BASE + R_TABLE, 14).Value = "MEDIDA"
    CL(ws, ROW_BASE + R_TABLE, 17).Value = "ESPERADO"
    CL(ws, ROW_BASE + R_TABLE, 18).Value = "OBSERVADO"
    CL(ws, ROW_BASE + R_TABLE, 19).Value = "OK"
    With ws.Range(CL(ws, ROW_BASE + R_FIRST, 14), CL(ws, ROW_BASE + R_FIRST + 10, 19))
        .Interior.Color = CARD
        .Font.Name = "Consolas"
        .Font.Size = 9
    End With
    ws.Range(CL(ws, ROW_BASE + R_FIRST, 17), CL(ws, ROW_BASE + R_FIRST + 10, 19)).HorizontalAlignment = xlCenter
    ws.Range(RA("N" & CStr(R_FIRST + 12) & ":S" & CStr(R_FIRST + 29))).Merge
    With ws.Range(RA("N" & CStr(R_FIRST + 12)))
        .WrapText = True
        .VerticalAlignment = xlTop
        .Interior.Color = CARD
        .IndentLevel = 1
        .Value = "COMO LEERLO" & vbLf & _
                 "1) Fase 1: cada ST toca una pagina nueva: FALLO + L1D fallo. Con 3 paginas de codigo y 7 de datos hay 10 paginas y solo 7 marcos: desde el 8vo fallo la CLOCK desaloja y las paginas dirty se guardan en el swap." & vbLf & _
                 "2) Fase 2: LD [100h] repetido. La 1a vez la pagina 8 vuelve del swap con el dato A1h que se modifico; las otras 3 lecturas aciertan en L1D." & vbLf & _
                 "3) Fase 3: LD [1A0h] no causa falta (la pagina sigue en RAM) pero si un fallo de cache: pagina residente no significa linea en cache." & vbLf & _
                 "4) Fase 4: el dato recuperado se vuelve a guardar en [101h]."
    End With
    Refresh ws
    Application.ScreenUpdating = prev
End Sub

Public Sub MostrarPanelMemoria()
    If MemSheet(False) Is Nothing Then
        ConstruirPanelMemoria
    Else
        LcaGoto MemSheet(False), ROW_BASE, COL_BASE
        MemPanelRefrescar
    End If
End Sub

' --- Actions -----------------------------------------------------------------------------------

Public Sub MemCargar()
    Dim ws As Worksheet
    Set ws = MemSheet(False)
    If ws Is Nothing Then Exit Sub
    If MemDemoLoad() Then
        mMessage = "Programa de memoria cargado: maquina reiniciada (RAM vacia, caches vacias, ninguna pagina residente). Pulsa PASO o EJECUTAR."
    Else
        mMessage = "No se pudo ensamblar el programa de memoria."
    End If
    Refresh ws
    AfterAction
End Sub

Public Sub MemReset()
    MemCargar
End Sub

Public Sub MemVelocidad()
    Dim ws As Worksheet
    FlowCycleSpeed
    Set ws = MemSheet(False)
    If ws Is Nothing Then Exit Sub
    mMessage = "Animacion: " & FlowSpeedText() & "  (la misma velocidad rige todos los bloques)."
    Refresh ws
End Sub

Public Sub MemPaso()
    Dim ws As Worksheet
    Set ws = MemSheet(False)
    If ws Is Nothing Then Exit Sub
    EnsureLoaded
    If ExecStatus() = EXEC_HALTED Then
        mMessage = "El programa termino (HLT). Pulsa RESET para repetirlo."
    ElseIf ExecStatus() = EXEC_ERROR Then
        mMessage = "ERROR: " & ExecErrorText() & ". Pulsa RESET."
    ElseIf DoStep(ws, True) Then
        mMessage = "Paso: " & ExecLastText()
    End If
    Refresh ws
    AfterAction
End Sub

Public Sub MemEjecutar()
    Dim ws As Worksheet, n As Long, keep As Long, animate As Boolean
    Set ws = MemSheet(False)
    If ws Is Nothing Then Exit Sub
    EnsureLoaded
    If ExecStatus() = EXEC_HALTED Then
        mMessage = "El programa ya termino (HLT). Pulsa RESET."
        Refresh ws
        Exit Sub
    End If
    keep = FlowSpeed()
    animate = (keep > 0)
    If animate And keep < 3 Then FlowSetSpeed keep + 1
    Do While n < 100
        If Not DoStep(ws, animate) Then Exit Do
        n = n + 1
        If animate Then Refresh ws
        If ExecStatus() = EXEC_HALTED Then Exit Do
    Loop
    FlowSetSpeed keep
    If ExecStatus() = EXEC_HALTED Then
        mMessage = "Programa terminado: " & CStr(ExecSteps()) & " instrucciones, " & CStr(PagerFaultCount()) & " page faults, " & _
                   CStr(PagerEvictionCount()) & " evicciones, " & CStr(PagerWritebackCount()) & " write-backs. BX recupero " & _
                   AsmHex2(CpuReadRegister(LcaBX)) & "."
    ElseIf ExecStatus() = EXEC_ERROR Then
        mMessage = "ERROR: " & ExecErrorText()
    End If
    Refresh ws
    AfterAction
End Sub

Public Sub MemPanelRefrescar()
    Dim ws As Worksheet
    Set ws = MemSheet(False)
    If ws Is Nothing Then Exit Sub
    Refresh ws
End Sub

Private Sub EnsureLoaded()
    If Not MemDemoLoaded() Then Call MemDemoLoad
End Sub

' One instruction; animates the page that comes in, the dirty victim that goes out and the cache path.
Private Function DoStep(ByVal ws As Worksheet, ByVal animate As Boolean) As Boolean
    Dim ok As Boolean, i As Long
    ok = MemDemoStep()
    If Not ok Then Exit Function
    DoStep = True
    i = MemDemoLogCount() - 1
    If animate And FlowOn() And i >= 0 Then AnimateAccess ws, i
End Function

Private Sub AnimateAccess(ByVal ws As Worksheet, ByVal i As Long)
    Dim vpn As Long, pfn As Long, victim As Long, e As LcaPageEntry, p As LcaPathStep, idx As Long
    Application.ScreenUpdating = True
    If MemDemoLogFaults(i) > 0 Then
        vpn = PagerLastFaultVPN()
        victim = MemDemoLogVictim(i)
        e = MmuPage(vpn)
        pfn = e.PFN
        If MemDemoLogWritebacks(i) > 0 And victim >= 0 Then
            FlowSend ws, "32 B", RGB(251, 146, 60), "P17_fr" & CStr(pfn), "P17_pg" & CStr(victim)
        End If
        FlowSend ws, "VPN" & CStr(vpn), FlowPacketColor(), "P17_pg" & CStr(vpn), "P17_fr" & CStr(pfn)
    End If
    If MemDemoLogVA(i) >= 0 Then
        p = CachePathLast()
        If p.PA >= 0 Then
            idx = (p.PA \ 4) Mod 2
            If MemDemoLogDataHit(i) = 1 Then
                FlowPulse ws, "P17_cd" & CStr(idx), RGB(52, 211, 153), 2
            Else
                FlowSend ws, "4 B", RGB(56, 189, 248), "P17_fr" & CStr(p.PA \ 32), "P17_c2" & CStr((p.PA \ 4) Mod 8), "P17_cd" & CStr(idx)
            End If
        End If
    End If
End Sub

' --- Refresh -----------------------------------------------------------------------------------

Private Sub Refresh(ByVal ws As Worksheet)
    Dim prev As Boolean
    prev = Application.ScreenUpdating
    Application.ScreenUpdating = False
    RefreshCards ws
    RefreshFrames ws
    RefreshPages ws
    RefreshCaches ws
    RefreshTable ws
    RefreshExpected ws
    ws.Range(RA("B" & CStr(R_MSG))).Value = mMessage
    Application.ScreenUpdating = prev
End Sub

Private Sub RefreshCards(ByVal ws As Worksheet)
    Dim state As String, fill As Long, pc As Long, ph As String
    If Not MemDemoLoaded() Then
        state = "SIN CARGAR": fill = CARD
    Else
        pc = CpuSnapshot().PC
        ph = MemDemoPhase(pc)
        Select Case ExecStatus()
            Case EXEC_HALTED: state = "TERMINADO (HLT)": fill = OK_GREEN
            Case EXEC_ERROR: state = "ERROR": fill = BAD
            Case Else
                If ExecSteps() = 0 Then
                    state = "LISTO": fill = HEAD
                Else
                    state = "FASE " & ph: fill = BUTTON
                End If
        End Select
    End If
    SetCard ws, "B", state, fill
    SetCard ws, "F", CStr(PagerFaultCount()), IIf(PagerFaultCount() > 0, WARN, CARD)
    SetCard ws, "H", CStr(PagerEvictionCount()), IIf(PagerEvictionCount() > 0, WARN, CARD)
    SetCard ws, "J", CStr(PagerWritebackCount()), IIf(PagerWritebackCount() > 0, WARN, CARD)
    SetCard ws, "L", CStr(CacheHits(LCA_CACHE_L1D)), IIf(CacheHits(LCA_CACHE_L1D) > 0, OK_GREEN, CARD)
    SetCard ws, "N", CStr(CacheMisses(LCA_CACHE_L1D)), CARD
    SetCard ws, "P", CStr(CacheHits(LCA_CACHE_L1I)) & " / " & CStr(CacheMisses(LCA_CACHE_L1I)), CARD
    SetCard ws, "R", CStr(CacheHits(LCA_CACHE_L2)) & " / " & CStr(CacheMisses(LCA_CACHE_L2)), CARD
End Sub

Private Sub RefreshFrames(ByVal ws As Worksheet)
    Dim k As Long, owner As Long, e As LcaPageEntry, s As Shape, hand As Long
    hand = PagerHand()
    For k = 0 To 6
        Set s = ws.Shapes("P17_fr" & CStr(k))
        owner = PagerFrameOwner(k)
        If owner < 0 Then
            Fill s, "PFN " & CStr(k) & vbCr & "libre", CARD, EDGE
        Else
            e = MmuPage(owner)
            Fill s, "PFN " & CStr(k) & vbCr & "VPN " & CStr(owner) & vbCr & IIf(e.Dirty, "D", "-") & IIf(e.Referenced, " R", " -"), _
                 IIf(e.Dirty, WARN, DATA_FILL), EDGE
        End If
        If k = hand Then
            s.Line.ForeColor.RGB = LAMP_ON
            s.Line.Weight = 3
        Else
            s.Line.Weight = 1
        End If
    Next k
    Fill ws.Shapes("P17_fr7"), "PFN 7" & vbCr & "E/S", RGB(76, 29, 149), EDGE
End Sub

Private Sub RefreshPages(ByVal ws As Worksheet)
    Dim v As Long, e As LcaPageEntry, s As Shape, fault As Long, victim As Long, edge As Long
    fault = PagerLastFaultVPN()
    victim = PagerLastVictimVPN()
    For v = 0 To 14
        Set s = ws.Shapes("P17_pg" & CStr(v))
        e = MmuPage(v)
        edge = EDGE
        If v = fault Then edge = LAMP_ON
        If v = victim Then edge = RGB(248, 113, 113)
        If e.Present Then
            Fill s, "VPN " & CStr(v) & vbCr & "PFN " & CStr(e.PFN), IIf(e.Dirty, WARN, DATA_FILL), edge
        Else
            Fill s, "VPN " & CStr(v) & vbCr & "disco", CARD, edge
        End If
        s.Line.Weight = IIf(edge = EDGE, 1, 3)
    Next v
End Sub

Private Sub RefreshCaches(ByVal ws As Worksheet)
    Dim i As Long
    For i = 0 To 1
        ShowLine ws.Shapes("P17_ci" & CStr(i)), LCA_CACHE_L1I, i, 2
        ShowLine ws.Shapes("P17_cd" & CStr(i)), LCA_CACHE_L1D, i, 2
    Next i
    For i = 0 To 7
        ShowLine ws.Shapes("P17_c2" & CStr(i)), LCA_CACHE_L2, i, 8
    Next i
End Sub

Private Sub ShowLine(ByVal s As Shape, ByVal level As Long, ByVal idx As Long, ByVal lines As Long)
    Dim lineNo As Long
    If CacheLineValid(level, idx) Then
        lineNo = CacheLineTag(level, idx) * lines + idx
        Fill s, "L" & CStr(idx) & vbCr & "PA " & AsmHex2(lineNo * 4), DATA_FILL, EDGE
    Else
        Fill s, "L" & CStr(idx) & vbCr & "vacia", CARD, EDGE
    End If
End Sub

Private Sub RefreshTable(ByVal ws As Worksheet)
    Dim i As Long, r As Long, n As Long, c As Long, fill As Long
    n = MemDemoLogCount()
    For i = 0 To MEMDEMO_STEPS - 1
        r = ROW_BASE + R_FIRST + i
        ws.Range(CL(ws, r, 2), CL(ws, r, 13)).ClearContents
        fill = CARD
        If i < n Then
            CL(ws, r, 2).Value = i + 1
            CL(ws, r, 3).Value = AsmHex3(MemDemoLogPC(i))
            CL(ws, r, 4).Value = AsmDisassemble(LoaderImageByte(MemDemoLogPC(i)), LoaderImageByte(MemDemoLogPC(i) + 1), _
                                                LoaderImageByte(MemDemoLogPC(i) + 2), LoaderImageByte(MemDemoLogPC(i) + 3))
            If MemDemoLogVA(i) >= 0 Then
                CL(ws, r, 7).Value = AsmHex3(MemDemoLogVA(i))
                CL(ws, r, 8).Value = MemDemoLogVA(i) \ 32
            End If
            If MemDemoLogFaults(i) > 0 Then
                CL(ws, r, 9).Value = "FALTA x" & CStr(MemDemoLogFaults(i))
                fill = RGB(88, 63, 25)
            End If
            If MemDemoLogVictim(i) >= 0 Then
                CL(ws, r, 10).Value = "VPN " & CStr(MemDemoLogVictim(i)) & IIf(MemDemoLogWritebacks(i) > 0, "  + WRITEBACK", "")
            End If
            CL(ws, r, 12).Value = CStr(MemDemoLogL1Hits(i)) & "A/" & CStr(MemDemoLogL1Misses(i)) & "F"
            CL(ws, r, 13).Value = IIf(MemDemoLogL2Misses(i) > 0, CStr(MemDemoLogL2Misses(i)) & " fallo", "-")
            If MemDemoLogDataHit(i) = 1 Then fill = OK_GREEN
        End If
        ws.Range(CL(ws, r, 2), CL(ws, r, 13)).Interior.Color = fill
        ws.Range(CL(ws, r, 4), CL(ws, r, 4)).HorizontalAlignment = xlLeft
        ws.Range(CL(ws, r, 10), CL(ws, r, 10)).HorizontalAlignment = xlLeft
    Next i
End Sub

' Expected counters of the demo (derived by hand, see docs/P17_MEMORY_DEMO.md) against the live ones.
Private Sub RefreshExpected(ByVal ws As Worksheet)
    Dim names As Variant, i As Long, done As Boolean, r As Long, obs As Long, ok As Boolean
    names = Array("Instrucciones", "Page faults", "Evicciones CLOCK", "Write-backs", "L1D aciertos", "L1D fallos", _
                  "L1I aciertos", "L1I fallos", "L2 aciertos", "L2 fallos", "Dato [100h] recuperado")
    done = MemDemoLoaded() And (ExecStatus() = EXEC_HALTED)
    For i = 0 To 10
        r = ROW_BASE + R_FIRST + i
        CL(ws, r, 14).Value = names(i)
        CL(ws, r, 17).Value = IIf(i = 10, AsmHex2(MemDemoExpectedCounter(i)), MemDemoExpectedCounter(i))
        obs = MemDemoObservedCounter(i)
        CL(ws, r, 18).Value = IIf(i = 10, AsmHex2(obs), obs)
        ok = (obs = MemDemoExpectedCounter(i))
        If done Then
            CL(ws, r, 19).Value = IIf(ok, "OK", "X")
            CL(ws, r, 19).Interior.Color = IIf(ok, OK_GREEN, BAD)
        Else
            CL(ws, r, 19).Value = "..."
            CL(ws, r, 19).Interior.Color = CARD
        End If
    Next i
End Sub

' --- Helpers -----------------------------------------------------------------------------------

' P18: every view of the shared machine follows the action of this block.
Private Sub AfterAction()
    Call LcaSincronizar
End Sub

Private Sub Fill(ByVal s As Shape, ByVal text As String, ByVal color As Long, ByVal edge As Long)
    If s.TextFrame2.TextRange.Text <> text Then s.TextFrame2.TextRange.Text = text
    s.Fill.ForeColor.RGB = color
    s.Line.ForeColor.RGB = edge
End Sub

' k-th of n equal boxes between rows r1 and r2, over columns B..S.
Private Sub Slot(ByVal ws As Worksheet, ByVal name As String, ByVal r1 As Long, ByVal r2 As Long, _
                 ByVal k As Long, ByVal n As Long)
    Dim x0 As Double, w As Double, y0 As Double, h As Double, s As Shape
    x0 = CL(ws, ROW_BASE + 1, 2).Left
    w = (CL(ws, ROW_BASE + 1, 20).Left - x0 - 4) / n
    y0 = CL(ws, ROW_BASE + r1, 2).Top + 2
    h = CL(ws, ROW_BASE + r2, 2).Top - y0 - 2
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x0 + k * w + 1, y0, w - 3, h)
    StyleBox s, "P17_" & name, IIf(n = 15, 9, 10)
End Sub

Private Sub CacheBox(ByVal ws As Worksheet, ByVal name As String, ByVal c1 As String, ByVal c2 As String, _
                     ByVal k As Long, ByVal n As Long)
    Dim a As Range, w As Double, s As Shape
    Set a = ws.Range(RA(c1 & CStr(R_CACHE) & ":" & c2 & CStr(R_CACHE + 3)))
    w = a.Width / n
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, a.Left + k * w + 1, a.Top + 2, w - 3, a.Height - 4)
    StyleBox s, "P17_" & name, 9
End Sub

Private Sub StyleBox(ByVal s As Shape, ByVal name As String, ByVal size As Single)
    s.Name = name
    s.Fill.ForeColor.RGB = CARD
    s.Line.ForeColor.RGB = EDGE
    s.Adjustments.Item(1) = 0.1
    s.Placement = xlMove
    With s.TextFrame2
        .MarginLeft = 2
        .MarginRight = 2
        .MarginTop = 1
        .MarginBottom = 1
        .WordWrap = msoTrue
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = " "
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = size
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = INK
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
End Sub

Private Sub AddCard(ByVal ws As Worksheet, ByVal c1 As String, ByVal c2 As String, ByVal title As String)
    ws.Range(RA(c1 & CStr(R_CARD) & ":" & c2 & CStr(R_CARD))).Merge
    With ws.Range(RA(c1 & CStr(R_CARD)))
        .Value = title
        .Interior.Color = HEAD
        .Font.Bold = True
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
    End With
    ws.Range(RA(c1 & CStr(R_CARD + 1) & ":" & c2 & CStr(R_CARD + 2))).Merge
    With ws.Range(RA(c1 & CStr(R_CARD + 1) & ":" & c2 & CStr(R_CARD + 2)))
        .Interior.Color = CARD
        .HorizontalAlignment = xlCenter
        .Font.Size = 13
        .Font.Bold = True
        .WrapText = True
        .NumberFormat = "@"
    End With
End Sub

Private Sub SetCard(ByVal ws As Worksheet, ByVal c1 As String, ByVal text As String, ByVal fill As Long)
    With ws.Range(RA(c1 & CStr(R_CARD + 1)))
        .Value = text
        .MergeArea.Interior.Color = fill
    End With
End Sub

Private Sub Btn(ByVal ws As Worksheet, ByVal address As String, ByVal caption As String, _
                ByVal macroName As String, ByVal fill As Long)
    Dim area As Range, s As Shape
    Set area = ws.Range(RA(address))
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left + 2, area.Top + 2, area.Width - 4, area.Height - 4)
    s.Name = "P17_btn_" & macroName
    s.Fill.ForeColor.RGB = fill
    s.Line.ForeColor.RGB = ACCENT
    s.Line.Weight = 0.75
    With s.TextFrame2
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = caption
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 10
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

Private Function MemSheet(ByVal create As Boolean) As Worksheet
    Set MemSheet = LcaBand("P17_btn_MemPaso", create)
End Function

Private Function RA(ByVal address As String) As String
    RA = LcaShift(address, ROW_BASE, COL_BASE)
End Function

' Cell of this block: row is absolute (ROW_BASE already added), column is relative to the block.
Private Function CL(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Range
    Set CL = ws.Cells(r, c + COL_BASE)
End Function

' --- Inspection used by the tests ----------------------------------------------------------------

Public Function MemPanelConstruido() As Boolean
    MemPanelConstruido = Not MemSheet(False) Is Nothing
End Function

Public Function MemPanelEstado() As String
    Dim ws As Worksheet
    Set ws = MemSheet(False)
    If ws Is Nothing Then Exit Function
    MemPanelEstado = CStr(ws.Range(RA("B" & CStr(R_CARD + 1))).Value)
End Function

' True when every drawn frame, page and counter card equals the engine state.
Public Function MemPanelCoherente() As Boolean
    Dim ws As Worksheet, k As Long, owner As Long, txt As String
    Set ws = MemSheet(False)
    If ws Is Nothing Then Exit Function
    For k = 0 To 6
        owner = PagerFrameOwner(k)
        txt = ws.Shapes("P17_fr" & CStr(k)).TextFrame2.TextRange.Text
        If owner < 0 Then
            If InStr(txt, "libre") = 0 Then Exit Function
        ElseIf InStr(txt, "VPN " & CStr(owner)) = 0 Then
            Exit Function
        End If
    Next k
    For k = 0 To 14
        txt = ws.Shapes("P17_pg" & CStr(k)).TextFrame2.TextRange.Text
        If MmuPage(k).Present Then
            If InStr(txt, "PFN " & CStr(MmuPage(k).PFN)) = 0 Then Exit Function
        ElseIf InStr(txt, "disco") = 0 Then
            Exit Function
        End If
    Next k
    If CStr(ws.Range(RA("F" & CStr(R_CARD + 1))).Value) <> CStr(PagerFaultCount()) Then Exit Function
    If CStr(ws.Range(RA("L" & CStr(R_CARD + 1))).Value) <> CStr(CacheHits(LCA_CACHE_L1D)) Then Exit Function
    MemPanelCoherente = True
End Function
