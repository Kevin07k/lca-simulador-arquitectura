Attribute VB_Name = "LcaEchoTests"
Option Explicit

Private mChecks As Long

' P16 regression: the echo program runs on the real CPU with IN/OUT; characters come out in typing order,
' an empty queue is a wait that can be continued, eight characters fill the LCD, the ninth replaces the
' first, a full queue rejects the extra key, '.' ends the program; STEP and RUN agree; the .asm assembles
' to the loaded image; the P16 block draws the same queue and LCD as RAM.
Public Sub VerificarEcoP16()
    Dim n As Long, i As Long, stepLcd As String, stepRegs(0 To 3) As Long, stepCount As Long, r As Long
    On Error GoTo failed
    mChecks = 0

    ' --- A: load ------------------------------------------------------------------------------------
    Check EchoLoad(), "A: the echo program assembles and loads"
    Check EchoLoaded() And LoaderInstructionCount() = 10, "A: ten instructions loaded"
    Check ExecStatus() = EXEC_READY And EchoStateText() = "LISTO", "A: machine READY"
    Check LcdText() = Space$(8) And LcdCursor() = 0, "A: LCD blank, cursor 0"
    Check KeyboardIsEmpty() And EchoOrderCount() = 0, "A: queue empty, no characters yet"
    Check InStr(AsmDisassemble(LoaderImageByte(8), LoaderImageByte(9), LoaderImageByte(10), LoaderImageByte(11)), "IN") = 1, _
          "A: the loop head (08h) is IN AX"
    Check InStr(AsmDisassemble(LoaderImageByte(20), LoaderImageByte(21), LoaderImageByte(22), LoaderImageByte(23)), "OUT") = 1, _
          "A: 14h is OUT AX"

    ' --- B: empty queue = wait, not error ---------------------------------------------------------------
    n = EchoRun()
    Check n = 2 And ExecStatus() = EXEC_WAIT, "B: two MOVs, then IN waits (WAIT_IN)"
    Check CpuSnapshot().PC = ECHO_LOOP_HEAD, "B: PC stays on the IN (08h)"
    Check CpuReadRegister(1) = ECHO_END_KEY And CpuReadRegister(2) = 0, "B: BX = 2Eh ('.'), CX = 0"
    Check LcdText() = Space$(8) And LcdWrites() = 0, "B: nothing written to the LCD"
    Check KeyboardWaits() = 1 And EchoStateText() = "ESPERANDO TECLA (WAIT_IN)", "B: one IN_EMPTY wait recorded"
    n = EchoRun()
    Check n = 0 And ExecStatus() = EXEC_WAIT And CpuSnapshot().PC = ECHO_LOOP_HEAD And ExecSteps() = 2, _
          "B: running again without keys waits again on the same IN"
    Check KeyboardWaits() = 2, "B: second wait counted"

    ' --- C: continue after typing: the same IN completes, order kept ----------------------------------
    Check EchoType("HOLA") = 4 And KeyboardCount() = 4, "C: HOLA typed, four keys queued"
    n = EchoRun()
    Check n = 24 And ExecSteps() = 26, "C: four echo passes of six instructions (got " & CStr(n) & ")"
    Check ExecStatus() = EXEC_WAIT And CpuSnapshot().PC = ECHO_LOOP_HEAD, "C: waits again when the queue is empty"
    Check LcdText() = "HOLA    " And LcdCursor() = 4, "C: LCD shows HOLA, cursor 4"
    Check EchoShown() = 4 And KeyboardIsEmpty() And KeyboardConsumed() = 4, "C: CX = 4, every key used once"
    Check EchoOrderText() = "HOLA" And EchoOrderRoute() = "0>0;1>1;2>2;3>3;", "C: order HOLA, slot k -> cell k"

    ' --- D: eight characters fill the screen ---------------------------------------------------------
    Check EchoType(" LCA") = 4, "D: ' LCA' typed"
    n = EchoRun()
    Check n = 24 And ExecStatus() = EXEC_WAIT, "D: four more passes, then waits"
    Check LcdText() = "HOLA LCA", "D: eight characters: HOLA LCA"
    Check LcdCursor() = 0 And LcdWraps() = 1, "D: the eighth fills cell 7 and the cursor wraps to 0"
    Check EchoShown() = 8 And LcdWrites() = 8, "D: CX = 8"
    Check EchoOrderRoute() = "0>0;1>1;2>2;3>3;4>4;5>5;6>6;7>7;", "D: queue slot and LCD cell of every key"
    For i = 0 To 7
        Check RamReadIO(LCA_LCD_BASE_PA + i) = AscW(Mid$("HOLA LCA", i + 1, 1)), "D: RAM " & Hex$(LCA_LCD_BASE_PA + i) & "h = screen"
    Next i

    ' --- E: continue: the ninth replaces the first; '.' ends ------------------------------------------
    Check EchoType("S.") = 2, "E: 'S.' typed (slots 0 and 1 again: the queue is circular)"
    n = EchoRun()
    Check ExecStatus() = EXEC_HALTED And CpuSnapshot().PC = ECHO_HALT_PC, "E: '.' ends the program at HLT (24h)"
    Check LcdText() = "SOLA LCA" And LcdCursor() = 1, "E: the ninth character replaced the first: SOLA LCA"
    Check EchoShown() = 9 And LcdWrites() = 9, "E: '.' is not shown (nine characters written)"
    Check ArithMemByte(0) = 9, "E: [100h] = 9 characters shown"
    Check ExecSteps() = 6 * 9 + 7, "E: 61 instructions for nine characters plus '.' (got " & CStr(ExecSteps()) & ")"
    Check EchoOrderText() = "HOLA LCAS." And Right$(EchoOrderRoute(), 10) = "0>0;1>fin;", "E: order and route of the last two keys"
    Check EchoType("Z") = 1 And EchoRun() = 0 And KeyboardCount() = 1 And LcdText() = "SOLA LCA", _
          "E: after HLT a new key stays in the queue and nothing runs"

    ' --- F: typing before running: a full queue rejects the ninth key ------------------------------
    Check EchoLoad(), "F: reload"
    Check EchoType("ABCDEFGHI") = 8, "F: nine keys typed, eight accepted"
    Check KeyboardRejected() = 1 And KeyboardIsFull() And KeyboardLastStatus() = "KBD_FULL", "F: the ninth is rejected (KBD_FULL)"
    n = EchoRun()
    Check n = 2 + 6 * 8 And ExecStatus() = EXEC_WAIT, "F: eight passes, then waits"
    Check LcdText() = "ABCDEFGH" And EchoOrderText() = "ABCDEFGH", "F: shown in typing order; 'I' never arrived"
    Check EchoType(".") = 1 And EchoRun() = 5 And ExecStatus() = EXEC_HALTED, "F: '.' ends the program"
    Check ArithMemByte(0) = 8 And ExecSteps() = 55, "F: [100h] = 8, 55 instructions"
    Check PagerFaultCount() = 3, "F: page faults: code pages 0 and 1 and the data page (got " & CStr(PagerFaultCount()) & ")"

    ' --- G: STEP one by one and RUN give the same machine ------------------------------------------
    Check EchoLoad(), "G: reload"
    Call EchoType("HI.")
    stepCount = 0
    Do While stepCount < 200
        If Not EchoStep() Then Exit Do
        stepCount = stepCount + 1
        If ExecStatus() = EXEC_HALTED Then Exit Do
    Loop
    stepLcd = LcdText()
    For r = 0 To 3
        stepRegs(r) = CpuReadRegister(r)
    Next r
    Check EchoLoad(), "G: reload"
    Call EchoType("HI.")
    n = EchoRun()
    Check n = stepCount And n = 6 * 2 + 7, "G: STEP and RUN execute the same 19 instructions"
    Check LcdText() = stepLcd And stepLcd = "HI      ", "G: same LCD"
    For r = 0 To 3
        Check CpuReadRegister(r) = stepRegs(r), "G: same " & Choose(r + 1, "AX", "BX", "CX", "DX")
    Next r
    Check EchoOrderRoute() = "0>0;1>1;2>fin;", "G: same order"

    ' --- H: machine RESET restarts the program and the order log; the .asm is the loaded image -------
    EchoReset
    Check ExecSteps() = 0 And EchoOrderCount() = 0 And LcdText() = Space$(8) And KeyboardIsEmpty(), _
          "H: RESET clears steps, order, LCD and queue"
    n = LoaderInstructionCount()
    Check AsmAssembleText(EchoAsmText()), "H: EchoAsmText assembles"
    Check AsmInstructionCount() = n, "H: same instruction count"
    For i = 0 To 4 * n - 1
        If AsmByte(i) <> LoaderImageByte(i) Then Check False, "H: byte " & AsmHex2(i) & " differs"
    Next i
    mChecks = mChecks + 1
    Check EchoFile() = "eco.asm", "H: file name eco.asm"

    ' --- I: the P16 block draws the same queue and LCD as RAM (when the dashboard is built) ----------
    If EcoPanelConstruido() Then
        Check EchoLoad(), "I: reload"
        Call EchoType("ECO 16")
        Call EchoRun
        Call EchoType("XY")
        EcoPanelRefrescar
        Check EcoPanelCoherente(), "I: drawn queue and LCD equal RAM E0-EFh"
        Check InStr(EcoPanelEstado(), "WAIT_IN") > 0, "I: the block shows the wait"
    End If

    LcaReport "P16 PASS: " & CStr(mChecks) & " comprobaciones / assertions." & vbCrLf & _
              "Eco teclado -> LCD con IN/OUT: orden FIFO, espera y continuacion, ocho caracteres, noveno, cola llena, fin con '.'.", True
    Exit Sub
failed:
    LcaReport "P16 FAIL after " & CStr(mChecks) & " checks: " & Err.Description, False
End Sub

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    mChecks = mChecks + 1
    If Not condition Then Err.Raise vbObjectError + 2172, "LcaEchoTests", message
End Sub
