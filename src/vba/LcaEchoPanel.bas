Attribute VB_Name = "LcaEchoPanel"
Option Explicit

' P16 block of the LCA dashboard (column 3, below the P05/P06 buses): the keyboard-to-LCD echo program
' running on the real CPU. It shows, live:
'   - the state of the machine (READY / WAIT_IN / HLT) and the counters (AX, CX shown, queued, accepted,
'     rejected with KBD_FULL, IN_EMPTY waits, [100h]);
'   - the keyboard queue E0..E7h with head/tail and the LCD E8..EFh with its cursor, both read from RAM;
'   - the ORDER of the characters: which queue slot each key left and which LCD cell it reached;
'   - the program listing with the PC row, the defined behaviour and the execution trace.
' IN and OUT are animated with the shared LcaFlow packet (slot -> AX -> LCD cell).
Private Const ROW_BASE As Long = LCA_BASE_ECHO
Private Const COL_BASE As Long = LCA_COL_ECHO
Private Const BAND_ROWS As Long = 55
Private Const ORDER_ROWS As Long = 12
Private Const TRACE_ROWS As Long = 8
Private Const R_CARD As Long = 11
Private Const R_MSG As Long = 15
Private Const R_IO As Long = 17
Private Const R_ORDER As Long = 23
Private Const R_RULES As Long = 38
Private Const R_TRACE As Long = 46

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
Private Const LCD_BG As Long = 3381555       ' RGB(51, 153, 51)
Private Const LCD_INK As Long = 1586439      ' RGB(7, 52, 24)

Private mMessage As String

' --- Build -------------------------------------------------------------------------------------

Public Sub ConstruirPanelEco()
    Dim ws As Worksheet, i As Long, prev As Boolean, rules As Variant
    prev = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set ws = LcaSheet(True)
    LcaPrepareBand ws, ROW_BASE, BAND_ROWS, "P16_", COL_BASE
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
        .Value = "P16  |  ECO TECLADO -> LCD  -  IN / OUT en la CPU real"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = ACCENT
    End With
    ws.Range(RA("B3:S3")).Merge
    With ws.Range(RA("B3"))
        .Value = "Cada tecla entra a la cola E0-E7h; IN AX la saca (la mas antigua primero) y OUT AX la escribe en la celda del cursor " & _
                 "del LCD E8-EFh. La tecla '.' termina el programa."
        .Font.Color = MUTED
    End With

    Btn ws, "B5:D6", "CARGAR ECO", "EcoCargar", HEAD
    Btn ws, "E5:G6", "PASO / STEP", "EcoPaso", BUTTON
    Btn ws, "H5:J6", "EJECUTAR / RUN", "EcoEjecutar", OK_GREEN
    Btn ws, "K5:M6", "RESET", "EcoReset", EDGE
    Btn ws, "N5:P6", "VELOCIDAD", "EcoVelocidad", CARD
    Btn ws, "Q5:S6", "VER EN MAQUINA", "LcaIrMaquina", EDGE

    ws.Range(RA("B8:B9")).Merge
    With ws.Range(RA("B8"))
        .Value = "TECLAS"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Range(RA("C8:G9")).Merge
    With ws.Range(RA("C8:G9"))
        .Interior.Color = RGB(230, 240, 250)
        .Font.Color = BG
        .Font.Name = "Consolas"
        .Font.Size = 14
        .Font.Bold = True
        .NumberFormat = "@"
        .HorizontalAlignment = xlCenter
    End With
    ws.Range(RA("C8")).Value = "HOLA LCA"
    Btn ws, "H8:J9", "TECLEAR / TYPE", "EcoTeclear", BUTTON
    Btn ws, "K8:M9", "DEMO 8 CARACTERES", "EcoDemoOcho", RGB(236, 72, 153)
    Btn ws, "N8:P9", "+ 'S.'  9no Y FIN", "EcoDemoFin", RGB(236, 72, 153)
    Btn ws, "Q8:S9", "COLA LLENA (9)", "EcoDemoLlena", RGB(236, 72, 153)

    ' Cards.
    AddCard ws, "B", "E", "ESTADO DE LA CPU"
    AddCard ws, "F", "G", "AX (tecla)"
    AddCard ws, "H", "I", "CX mostrados"
    AddCard ws, "J", "K", "EN COLA"
    AddCard ws, "L", "M", "ACEPTADAS"
    AddCard ws, "N", "O", "RECHAZADAS KBD_FULL"
    AddCard ws, "P", "Q", "ESPERAS IN_EMPTY"
    AddCard ws, "R", "S", "[100h] al terminar"
    ws.Range(RA("B" & CStr(R_MSG) & ":S" & CStr(R_MSG))).Merge
    ws.Range(RA("B" & CStr(R_MSG))).Font.Color = ACCENT

    ' Queue and LCD.
    Title ws, "B" & CStr(R_IO) & ":J" & CStr(R_IO), "COLA DEL TECLADO  E0-E7h   (H = proximo IN,  T = proxima tecla)"
    Title ws, "K" & CStr(R_IO) & ":S" & CStr(R_IO), "LCD 8x1  E8-EFh   (^ = cursor: celda del proximo OUT)"
    LcaRowHeight ws.Rows(ROW_BASE + R_IO + 2), 34
    For i = 0 To 3
        CL(ws, ROW_BASE + R_IO + 1 + i, 2).Value = Choose(i + 1, "PA", "TECLA", "BYTE", "H / T")
        CL(ws, ROW_BASE + R_IO + 1 + i, 11).Value = Choose(i + 1, "PA", "LCD", "BYTE", "CURSOR")
    Next i
    For i = 0 To 7
        CL(ws, ROW_BASE + R_IO + 1, 3 + i).Value = Right$("00" & Hex$(LCA_KBD_BASE_PA + i), 2) & "h"
        CL(ws, ROW_BASE + R_IO + 1, 12 + i).Value = Right$("00" & Hex$(LCA_LCD_BASE_PA + i), 2) & "h"
    Next i
    With ws.Range(RA("B" & CStr(R_IO + 1) & ":S" & CStr(R_IO + 4)))
        .HorizontalAlignment = xlCenter
        .Font.Name = "Consolas"
    End With
    ws.Range(RA("B" & CStr(R_IO + 1) & ":S" & CStr(R_IO + 1))).Font.Color = MUTED
    ws.Range(RA("B" & CStr(R_IO + 1) & ":B" & CStr(R_IO + 4))).Font.Color = MUTED
    ws.Range(RA("K" & CStr(R_IO + 1) & ":K" & CStr(R_IO + 4))).Font.Color = MUTED
    With ws.Range(RA("C" & CStr(R_IO + 2) & ":J" & CStr(R_IO + 3)))
        .Borders.Color = EDGE
        .NumberFormat = "@"
    End With
    ws.Range(RA("C" & CStr(R_IO + 2) & ":J" & CStr(R_IO + 2))).Font.Size = 16
    ws.Range(RA("C" & CStr(R_IO + 2) & ":J" & CStr(R_IO + 2))).Font.Bold = True
    With ws.Range(RA("L" & CStr(R_IO + 2) & ":S" & CStr(R_IO + 2)))
        .Interior.Color = LCD_BG
        .Font.Color = LCD_INK
        .Font.Size = 20
        .Font.Bold = True
        .NumberFormat = "@"
        .Borders.Color = LCD_INK
        .Borders.Weight = xlMedium
    End With
    ws.Range(RA("L" & CStr(R_IO + 3) & ":S" & CStr(R_IO + 3))).NumberFormat = "@"

    ' Order of the characters and listing.
    Title ws, "B" & CStr(R_ORDER) & ":J" & CStr(R_ORDER), "ORDEN DE LOS CARACTERES  (de la cola al LCD, el mas antiguo primero)"
    Title ws, "L" & CStr(R_ORDER) & ":S" & CStr(R_ORDER), "PROGRAMA  eco.asm   (fila amarilla = PC)"
    ws.Range(RA("E" & CStr(R_ORDER + 1) & ":F" & CStr(R_ORDER + 1))).Merge
    ws.Range(RA("G" & CStr(R_ORDER + 1) & ":H" & CStr(R_ORDER + 1))).Merge
    ws.Range(RA("I" & CStr(R_ORDER + 1) & ":J" & CStr(R_ORDER + 1))).Merge
    CL(ws, ROW_BASE + R_ORDER + 1, 2).Value = "#"
    CL(ws, ROW_BASE + R_ORDER + 1, 3).Value = "TECLA"
    CL(ws, ROW_BASE + R_ORDER + 1, 4).Value = "BYTE"
    CL(ws, ROW_BASE + R_ORDER + 1, 5).Value = "SALE DE (IN)"
    CL(ws, ROW_BASE + R_ORDER + 1, 7).Value = "LLEGA A (OUT)"
    CL(ws, ROW_BASE + R_ORDER + 1, 9).Value = "RESULTADO"
    With ws.Range(RA("B" & CStr(R_ORDER + 1) & ":J" & CStr(R_ORDER + 1)))
        .Interior.Color = HEAD
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    For i = 1 To ORDER_ROWS
        ws.Range(RA("E" & CStr(R_ORDER + 1 + i) & ":F" & CStr(R_ORDER + 1 + i))).Merge
        ws.Range(RA("G" & CStr(R_ORDER + 1 + i) & ":H" & CStr(R_ORDER + 1 + i))).Merge
        ws.Range(RA("I" & CStr(R_ORDER + 1 + i) & ":J" & CStr(R_ORDER + 1 + i))).Merge
    Next i
    With ws.Range(RA("B" & CStr(R_ORDER + 2) & ":J" & CStr(R_ORDER + 1 + ORDER_ROWS)))
        .Interior.Color = CARD
        .Font.Name = "Consolas"
        .HorizontalAlignment = xlCenter
        .NumberFormat = "@"
        .Borders(xlInsideHorizontal).Color = BG
    End With
    For i = 1 To 10
        ws.Range(RA("M" & CStr(R_ORDER + i) & ":S" & CStr(R_ORDER + i))).Merge
    Next i
    With ws.Range(RA("L" & CStr(R_ORDER + 1) & ":S" & CStr(R_ORDER + 10)))
        .Font.Name = "Consolas"
        .Interior.Color = CARD
    End With

    ' Defined behaviour.
    Title ws, "B" & CStr(R_RULES) & ":S" & CStr(R_RULES), "COMPORTAMIENTO DEFINIDO"
    rules = Array( _
        "1. Orden: las teclas salen en el mismo orden en que se teclearon (cola FIFO); cada tecla se usa una sola vez.", _
        "2. Sin entrada: IN con la cola vacia NO es error: la CPU queda en ESPERANDO TECLA (WAIT_IN), el PC se queda en el IN y no se escribe nada.", _
        "3. Continuar: se teclea mas y se pulsa EJECUTAR (o PASO): se repite el mismo IN y el eco sigue donde se quedo.", _
        "4. Ocho caracteres llenan el LCD: el octavo va a la celda 7 (EFh) y el cursor vuelve a 0. El noveno reemplaza al primero.", _
        "5. Cola llena: con 8 teclas sin leer, la novena se rechaza (KBD_FULL) y nunca llega al LCD.", _
        "6. Fin: la tecla '.' no se muestra; el programa guarda en [100h] cuantos caracteres mostro (CX) y se detiene en HLT.")
    For i = 0 To UBound(rules)
        ws.Range(RA("B" & CStr(R_RULES + 1 + i) & ":S" & CStr(R_RULES + 1 + i))).Merge
        With ws.Range(RA("B" & CStr(R_RULES + 1 + i)))
            .Value = rules(i)
            .Interior.Color = CARD
            .IndentLevel = 1
        End With
    Next i

    ' Trace.
    Title ws, "B" & CStr(R_TRACE) & ":S" & CStr(R_TRACE), "TRAZA DE EJECUCION  (ultimas " & CStr(TRACE_ROWS) & " lineas; la traza completa esta en la hoja Logs)"
    For i = 1 To TRACE_ROWS
        ws.Range(RA("B" & CStr(R_TRACE + i) & ":S" & CStr(R_TRACE + i))).Merge
    Next i
    With ws.Range(RA("B" & CStr(R_TRACE + 1) & ":S" & CStr(R_TRACE + TRACE_ROWS)))
        .Font.Name = "Consolas"
        .Font.Size = 9
        .Interior.Color = CARD
        .IndentLevel = 1
    End With

    mMessage = "Pulsa CARGAR ECO, escribe teclas y pulsa TECLEAR; luego EJECUTAR. O usa DEMO 8 CARACTERES."
    Refresh ws
    Application.ScreenUpdating = prev
End Sub

Public Sub MostrarPanelEco()
    If EcoSheet(False) Is Nothing Then ConstruirPanelEco
    LcaGoto EcoSheet(False), ROW_BASE, COL_BASE
    EcoPanelRefrescar
End Sub

' --- Actions -----------------------------------------------------------------------------------

Public Sub EcoCargar()
    Dim ws As Worksheet
    Set ws = EcoSheet(False)
    If ws Is Nothing Then Exit Sub
    If EchoLoad() Then
        mMessage = "Programa de eco cargado: maquina reiniciada, cola vacia y LCD en blanco. Teclea y pulsa EJECUTAR."
    Else
        mMessage = "No se pudo ensamblar el programa de eco."
    End If
    Refresh ws
    AfterAction
End Sub

Public Sub EcoTeclear()
    Dim ws As Worksheet, text As String, accepted As Long, tail0 As Long
    Set ws = EcoSheet(False)
    If ws Is Nothing Then Exit Sub
    EnsureLoaded
    text = CStr(ws.Range(RA("C8")).Value)
    If Len(text) = 0 Then
        mMessage = "Escribe una o mas teclas en TECLAS."
    Else
        tail0 = KeyboardTail()
        On Error GoTo badText
        accepted = EchoType(text)
        On Error GoTo 0
        AnimateKeys ws, text, tail0, accepted
        If accepted < Len(text) Then
            mMessage = CStr(accepted) & " de " & CStr(Len(text)) & " teclas entraron a la cola; " & CStr(Len(text) - accepted) & _
                       " rechazadas: la cola ya tenia 8 teclas sin leer (KBD_FULL)."
        Else
            mMessage = CStr(accepted) & " teclas en la cola. Pulsa EJECUTAR (o PASO) para que el programa las muestre."
        End If
    End If
    Refresh ws
    AfterAction
    Exit Sub
badText:
    mMessage = "Solo se aceptan caracteres ASCII (00-FFh)."
    Refresh ws
End Sub

Public Sub EcoPaso()
    Dim ws As Worksheet
    Set ws = EcoSheet(False)
    If ws Is Nothing Then Exit Sub
    EnsureLoaded
    If ExecStatus() = EXEC_HALTED Then
        mMessage = "El programa termino con '.' (HLT). Pulsa RESET o CARGAR ECO para empezar otra vez."
    ElseIf ExecStatus() = EXEC_ERROR Then
        mMessage = "ERROR: " & ExecErrorText() & ". Pulsa RESET."
    ElseIf DoStep(ws, True) Then
        mMessage = "Paso: " & ExecLastText()
    ElseIf ExecStatus() = EXEC_WAIT Then
        mMessage = "IN con la cola vacia: la CPU espera una tecla (WAIT_IN). Teclea y vuelve a pulsar PASO o EJECUTAR."
    End If
    Refresh ws
    AfterAction
End Sub

Public Sub EcoEjecutar()
    Dim ws As Worksheet, n As Long, keep As Long, animate As Boolean
    Set ws = EcoSheet(False)
    If ws Is Nothing Then Exit Sub
    EnsureLoaded
    If ExecStatus() = EXEC_HALTED Then
        mMessage = "El programa ya termino (HLT). Pulsa RESET o CARGAR ECO."
        Refresh ws
        Exit Sub
    End If
    keep = FlowSpeed()
    animate = (keep > 0)
    If animate And keep < 3 Then FlowSetSpeed keep + 1
    Do While n < 300
        If Not DoStep(ws, animate) Then Exit Do
        n = n + 1
        If ExecStatus() = EXEC_HALTED Then Exit Do
    Loop
    FlowSetSpeed keep
    Select Case ExecStatus()
        Case EXEC_HALTED
            mMessage = CStr(n) & " instrucciones. Se leyo '.': el programa guardo CX = " & CStr(EchoShown()) & " en [100h] y se detuvo (HLT)."
        Case EXEC_WAIT
            mMessage = CStr(n) & " instrucciones. Cola vacia: ESPERANDO TECLA en IN (PC = 008h). Teclea mas y pulsa EJECUTAR para continuar."
        Case EXEC_ERROR
            mMessage = "ERROR: " & ExecErrorText()
        Case Else
            mMessage = CStr(n) & " instrucciones ejecutadas."
    End Select
    Refresh ws
    AfterAction
End Sub

Public Sub EcoReset()
    Dim ws As Worksheet
    Set ws = EcoSheet(False)
    If ws Is Nothing Then Exit Sub
    If EchoLoaded() Then
        EchoReset
        mMessage = "RESET: PC = 000h, cola vacia, LCD en blanco."
    Else
        Call EchoLoad
        mMessage = "Programa de eco cargado."
    End If
    Refresh ws
    AfterAction
End Sub

Public Sub EcoVelocidad()
    Dim ws As Worksheet
    FlowCycleSpeed
    mMessage = "Animacion: " & FlowSpeedText() & "  (la misma velocidad rige la maquina y los demas bloques)."
    Set ws = EcoSheet(False)
    If Not ws Is Nothing Then Refresh ws
End Sub

' Exercise of eight characters: HOLA LCA fills the screen and the CPU waits for more input.
Public Sub EcoDemoOcho()
    Dim ws As Worksheet
    Set ws = EcoSheet(False)
    If ws Is Nothing Then Exit Sub
    Call EchoLoad
    ws.Range(RA("C8")).Value = "HOLA LCA"
    Refresh ws
    EcoTeclear
    EcoEjecutar
    mMessage = "Ocho caracteres: HOLA LCA llena el LCD y el cursor vuelve a la celda 0. Sin mas teclas, la CPU espera en IN. " & _
               "Pulsa  + 'S.'  para continuar."
    Refresh ws
End Sub

' Continuing: S is the ninth character (replaces the first cell), '.' ends the program.
Public Sub EcoDemoFin()
    Dim ws As Worksheet
    Set ws = EcoSheet(False)
    If ws Is Nothing Then Exit Sub
    If Not EchoLoaded() Or ExecStatus() = EXEC_HALTED Or EchoShown() <> 8 Then EcoDemoOcho
    ws.Range(RA("C8")).Value = "S."
    Refresh ws
    EcoTeclear
    EcoEjecutar
    mMessage = "Continuacion: 'S' es el noveno caracter y reemplaza la celda 0 (SOLA LCA); '.' no se muestra: [100h] = " & _
               CStr(ArithMemByte(0)) & " caracteres y HLT."
    Refresh ws
End Sub

' Queue full: nine keys typed before running; the ninth is rejected and never reaches the LCD.
Public Sub EcoDemoLlena()
    Dim ws As Worksheet
    Set ws = EcoSheet(False)
    If ws Is Nothing Then Exit Sub
    Call EchoLoad
    ws.Range(RA("C8")).Value = "ABCDEFGHI"
    Refresh ws
    EcoTeclear
    mMessage = "Nueve teclas sin que la CPU lea: la cola tiene 8 lugares y la 'I' se rechazo (KBD_FULL). Pulsa EJECUTAR: " & _
               "salen A..H en orden."
    Refresh ws
End Sub

Public Sub EcoPanelRefrescar()
    Dim ws As Worksheet
    Set ws = EcoSheet(False)
    If ws Is Nothing Then Exit Sub
    Refresh ws
End Sub

' --- Inspection used by the tests ----------------------------------------------------------------

Public Function EcoPanelConstruido() As Boolean
    EcoPanelConstruido = Not EcoSheet(False) Is Nothing
End Function

Public Function EcoPanelEstado() As String
    Dim ws As Worksheet
    Set ws = EcoSheet(False)
    If ws Is Nothing Then Exit Function
    EcoPanelEstado = CStr(CL(ws, ROW_BASE + R_CARD + 1, 2).Value)
End Function

' True when every drawn queue and LCD byte equals RAM E0..EFh and the LCD row shows LcdChar.
Public Function EcoPanelCoherente() As Boolean
    Dim ws As Worksheet, i As Long
    Set ws = EcoSheet(False)
    If ws Is Nothing Then Exit Function
    For i = 0 To 7
        If CStr(CL(ws, ROW_BASE + R_IO + 3, 3 + i).Value) <> HexByte(RamReadIO(LCA_KBD_BASE_PA + i)) Then Exit Function
        If CStr(CL(ws, ROW_BASE + R_IO + 3, 12 + i).Value) <> HexByte(RamReadIO(LCA_LCD_BASE_PA + i)) Then Exit Function
        If CStr(CL(ws, ROW_BASE + R_IO + 2, 12 + i).Value) <> LcdChar(i) Then Exit Function
    Next i
    EcoPanelCoherente = True
End Function

' --- Step and animation --------------------------------------------------------------------------

Private Function DoStep(ByVal ws As Worksheet, ByVal animate As Boolean) As Boolean
    Dim slot As Long, cell As Long
    slot = KeyboardHead()
    cell = LcdCursor()
    DoStep = EchoStep()
    If Not DoStep Or Not animate Or Not FlowOn() Then Exit Function
    Select Case ExecLastKind()
        Case "IN"
            Refresh ws
            Application.ScreenUpdating = True
            FlowSend ws, LcdDisplayChar(CpuReadRegister(0)), FlowPacketColor(), _
                     CL(ws, ROW_BASE + R_IO + 2, 3 + slot), ws.Range(RA("F" & CStr(R_CARD + 1) & ":G" & CStr(R_CARD + 2)))
        Case "OUT"
            Application.ScreenUpdating = True
            FlowSend ws, LcdDisplayChar(CpuReadRegister(0)), FlowPacketColor(), _
                     ws.Range(RA("F" & CStr(R_CARD + 1) & ":G" & CStr(R_CARD + 2))), CL(ws, ROW_BASE + R_IO + 2, 12 + cell)
            Refresh ws
            FlowPulseRange CL(ws, ROW_BASE + R_IO + 2, 12 + cell), LAMP_ON, 1
    End Select
End Function

Private Sub AnimateKeys(ByVal ws As Worksheet, ByVal text As String, ByVal tail0 As Long, ByVal accepted As Long)
    Dim i As Long
    If Not FlowOn() Then Exit Sub
    Application.ScreenUpdating = True
    For i = 0 To accepted - 1
        FlowSend ws, Mid$(text, i + 1, 1), FlowPacketColor(), ws.Range(RA("C8:G9")), CL(ws, ROW_BASE + R_IO + 2, 3 + ((tail0 + i) Mod 8))
    Next i
End Sub

' --- Refresh -------------------------------------------------------------------------------------

Private Sub Refresh(ByVal ws As Worksheet)
    Dim prev As Boolean, i As Long, st As String, fill As Long, row As Long, n As Long, first As Long
    Dim v As Long, queued As Boolean, pc As Long, addr As Long, mark As String, loaded As Boolean, item As Variant
    prev = Application.ScreenUpdating
    Application.ScreenUpdating = False
    loaded = EchoLoaded()

    ' Cards.
    st = EchoStateText()
    Select Case ExecStatus()
        Case EXEC_WAIT: fill = WARN
        Case EXEC_HALTED: fill = OK_GREEN
        Case EXEC_ERROR: fill = BAD
        Case Else: fill = HEAD
    End Select
    If Not loaded Then fill = CARD
    SetCard ws, "B", st, fill
    If loaded Then v = CpuReadRegister(0)
    SetCard ws, "F", IIf(loaded, HexByte(v) & "   " & LcdDisplayChar(v), "-"), CARD
    SetCard ws, "H", IIf(loaded, CStr(EchoShown()), "-"), CARD
    If loaded Then
        SetCard ws, "J", CStr(KeyboardCount()) & " de 8", IIf(KeyboardIsFull(), WARN, CARD)
        SetCard ws, "L", CStr(KeyboardAccepted()), CARD
        SetCard ws, "N", CStr(KeyboardRejected()), IIf(KeyboardRejected() > 0, BAD, CARD)
        SetCard ws, "P", CStr(KeyboardWaits()), IIf(ExecStatus() = EXEC_WAIT, WARN, CARD)
    Else
        For Each item In Array("J", "L", "N", "P")
            SetCard ws, CStr(item), "-", CARD
        Next item
    End If
    If loaded And ExecStatus() = EXEC_HALTED Then
        SetCard ws, "R", CStr(ArithMemByte(0)) & " caracteres", OK_GREEN
    Else
        SetCard ws, "R", "-", CARD
    End If
    CL(ws, ROW_BASE + R_MSG, 2).Value = mMessage
    SetBox ws, "btn_EcoVelocidad", "VELOCIDAD" & vbCr & FlowSpeedText()

    If Not loaded Then
        ws.Range(CL(ws, ROW_BASE + R_IO + 2, 3), CL(ws, ROW_BASE + R_IO + 4, 10)).ClearContents
        ws.Range(CL(ws, ROW_BASE + R_IO + 2, 12), CL(ws, ROW_BASE + R_IO + 4, 19)).ClearContents
        ws.Range(CL(ws, ROW_BASE + R_IO + 2, 3), CL(ws, ROW_BASE + R_IO + 3, 10)).Interior.Color = CARD
        GoTo order
    End If
    ' Queue E0..E7h.
    For i = 0 To 7
        v = RamReadIO(LCA_KBD_BASE_PA + i)
        queued = KeyboardSlotQueued(i)
        CL(ws, ROW_BASE + R_IO + 2, 3 + i).Value = IIf(queued, CellText(LcdDisplayChar(v)), "")
        CL(ws, ROW_BASE + R_IO + 3, 3 + i).Value = HexByte(v)
        mark = vbNullString
        If i = KeyboardHead() Then mark = "H"
        If i = KeyboardTail() Then mark = mark & IIf(Len(mark) > 0, " / ", "") & "T"
        CL(ws, ROW_BASE + R_IO + 4, 3 + i).Value = mark
        With ws.Range(CL(ws, ROW_BASE + R_IO + 2, 3 + i), CL(ws, ROW_BASE + R_IO + 3, 3 + i))
            .Interior.Color = IIf(queued, HEAD, CARD)
            .Font.Color = IIf(queued, INK, MUTED)
        End With
        CL(ws, ROW_BASE + R_IO + 4, 3 + i).Font.Color = IIf(Len(mark) > 0, LAMP_ON, MUTED)
    Next i

    ' LCD E8..EFh (always read back from RAM).
    For i = 0 To 7
        CL(ws, ROW_BASE + R_IO + 2, 12 + i).Value = CellText(LcdChar(i))
        CL(ws, ROW_BASE + R_IO + 3, 12 + i).Value = HexByte(RamReadIO(LCA_LCD_BASE_PA + i))
        CL(ws, ROW_BASE + R_IO + 4, 12 + i).Value = IIf(i = LcdCursor(), "^", "")
        CL(ws, ROW_BASE + R_IO + 4, 12 + i).Font.Color = LAMP_ON
    Next i

order:
    ' Order of the characters (last ORDER_ROWS).
    n = EchoOrderCount()
    If Not loaded Then n = 0
    first = n - ORDER_ROWS
    If first < 0 Then first = 0
    For i = 0 To ORDER_ROWS - 1
        row = ROW_BASE + R_ORDER + 2 + i
        ws.Range(CL(ws, row, 2), CL(ws, row, 10)).ClearContents
        ws.Range(CL(ws, row, 2), CL(ws, row, 10)).Interior.Color = CARD
        If first + i < n Then OrderRow ws, row, first + i, (first + i = n - 1)
    Next i

    ' Listing.
    pc = CpuSnapshot().PC
    For i = 0 To 9
        row = ROW_BASE + R_ORDER + 1 + i
        addr = 4 * i
        ws.Range(CL(ws, row, 12), CL(ws, row, 19)).Interior.Color = CARD
        ws.Range(CL(ws, row, 12), CL(ws, row, 19)).Font.Color = INK
        If loaded And i < LoaderInstructionCount() Then
            CL(ws, row, 12).Value = AsmHex3(addr)
            CL(ws, row, 13).Value = AsmDisassemble(LoaderImageByte(addr), LoaderImageByte(addr + 1), _
                                                   LoaderImageByte(addr + 2), LoaderImageByte(addr + 3)) & ListingNote(addr)
            If addr = pc Then
                ws.Range(CL(ws, row, 12), CL(ws, row, 19)).Interior.Color = IIf(ExecStatus() = EXEC_WAIT, WARN, LAMP_ON)
                ws.Range(CL(ws, row, 12), CL(ws, row, 19)).Font.Color = IIf(ExecStatus() = EXEC_WAIT, INK, BG)
            End If
        Else
            CL(ws, row, 12).Value = vbNullString
            CL(ws, row, 13).Value = IIf(i = 0 And Not loaded, "(pulsa CARGAR ECO)", vbNullString)
        End If
    Next i

    ' Trace.
    n = 0
    If loaded Then n = ExecTraceCount()
    For i = 0 To TRACE_ROWS - 1
        row = ROW_BASE + R_TRACE + TRACE_ROWS - i
        If i < n Then CL(ws, row, 2).Value = ExecTraceLine(i) Else CL(ws, row, 2).Value = vbNullString
    Next i
    Application.ScreenUpdating = prev
End Sub

Private Sub OrderRow(ByVal ws As Worksheet, ByVal row As Long, ByVal k As Long, ByVal newest As Boolean)
    Dim key As Long, cell As Long, slot As Long
    key = EchoOrderKey(k)
    slot = EchoOrderSlot(k)
    cell = EchoOrderCell(k)
    CL(ws, row, 2).Value = CStr(k + 1)
    CL(ws, row, 3).Value = CellText(LcdDisplayChar(key))
    CL(ws, row, 4).Value = HexByte(key)
    CL(ws, row, 5).Value = "cola " & CStr(slot) & "  (" & HexByte(LCA_KBD_BASE_PA + slot) & ")"
    Select Case cell
        Case -2
            CL(ws, row, 7).Value = "-"
            CL(ws, row, 9).Value = "FIN: el punto no se muestra"
        Case -1
            CL(ws, row, 7).Value = "..."
            CL(ws, row, 9).Value = "en AX, falta OUT"
        Case Else
            CL(ws, row, 7).Value = "LCD " & CStr(cell) & "  (" & HexByte(LCA_LCD_BASE_PA + cell) & ")"
            CL(ws, row, 9).Value = IIf(k >= 8, "reemplaza celda " & CStr(cell), "mostrada")
    End Select
    If newest Then ws.Range(CL(ws, row, 2), CL(ws, row, 10)).Interior.Color = HEAD
End Sub

Private Function ListingNote(ByVal addr As Long) As String
    Select Case addr
        Case ECHO_LOOP_HEAD: ListingNote = "      <- espera aqui si la cola esta vacia"
        Case &H14: ListingNote = "      <- escribe en el LCD"
        Case ECHO_HALT_PC: ListingNote = "      <- fin al leer '.'"
    End Select
End Function

' --- Helpers -----------------------------------------------------------------------------------

Private Sub EnsureLoaded()
    If Not EchoLoaded() Then
        Call EchoLoad
    End If
End Sub

Private Sub AfterAction()
    On Error Resume Next
    LogsActualizar
End Sub

Private Function HexByte(ByVal v As Long) As String
    HexByte = Right$("00" & Hex$(v), 2) & "h"
End Function

' A leading apostrophe would be taken by Excel as a text prefix: double it so it is shown.
Private Function CellText(ByVal s As String) As String
    If s = "'" Then CellText = "''" Else CellText = s
End Function

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

Private Sub SetBox(ByVal ws As Worksheet, ByVal name As String, ByVal text As String)
    On Error Resume Next
    ws.Shapes("P16_" & name).TextFrame2.TextRange.Text = text
    On Error GoTo 0
End Sub

Private Sub Btn(ByVal ws As Worksheet, ByVal address As String, ByVal caption As String, _
                ByVal macroName As String, ByVal fill As Long)
    Dim area As Range, s As Shape
    Set area = ws.Range(RA(address))
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left + 2, area.Top + 2, area.Width - 4, area.Height - 4)
    s.Name = "P16_btn_" & macroName
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

Private Function EcoSheet(ByVal create As Boolean) As Worksheet
    Set EcoSheet = LcaBand("P16_btn_EcoPaso", create)
End Function

Private Function RA(ByVal address As String) As String
    RA = LcaShift(address, ROW_BASE, COL_BASE)
End Function

' Cell of this block: row is absolute (ROW_BASE already added), column is relative to the block.
Private Function CL(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Range
    Set CL = ws.Cells(r, c + COL_BASE)
End Function
