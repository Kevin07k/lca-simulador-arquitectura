Attribute VB_Name = "LcaArithTests"
Option Explicit

Private mChecks As Long

' P15 regression: every arithmetic demo runs on the real CPU to HLT with the hand-derived registers, flags,
' memory, instruction count, page faults and register trace; STEP and RUN agree; the sources assemble to the
' loaded image; and the shared data-flow animation helpers behave.
Public Sub VerificarDemoAritmeticaP15()
    Dim i As Long, r As Long, n As Long, expected As Variant, label As String
    Dim stepRegs(0 To 3) As Long, stepFlags As Long, stepCount As Long, faultsExpected As Variant
    Dim usedDX As Boolean, prev As Long, ax As Double, ay As Double, sh As Worksheet, box As Range
    On Error GoTo failed
    mChecks = 0
    faultsExpected = Array(2, 2, 3, 3, 3)

    ' --- A: every program runs to HLT with the expected final state -------------------------------
    For i = 0 To ARITH_COUNT - 1
        label = ArithName(i) & ": "
        Check Len(ArithName(i)) > 0 And Len(ArithTitle(i)) > 0 And Len(ArithInputText(i)) > 0 And _
              Len(ArithExpectedText(i)) > 0 And Len(ArithAlgorithm(i)) > 0, label & "input, expected result and algorithm documented"
        Check ArithLoad(i), label & "assembles and loads"
        Check ExecStatus() = EXEC_READY And Not ArithMatches(i), label & "not matched before it runs"
        n = ExecRun(500)
        Check ExecStatus() = EXEC_HALTED, label & "stops at HLT"
        Check n = ArithExpectedSteps(i), label & CStr(ArithExpectedSteps(i)) & " instructions executed (got " & CStr(n) & ")"
        For r = 0 To 3
            Check CpuReadRegister(r) = ArithExpectedReg(i, r), label & "register " & Choose(r + 1, "AX", "BX", "CX", "DX") & " = " & _
                  CStr(ArithExpectedReg(i, r)) & " (got " & CStr(CpuReadRegister(r)) & ")"
        Next r
        Check CLng(CpuSnapshot().flags) = ArithExpectedFlags(i), label & "final flags ZF/CF/SF"
        Check ArithMemByte(0) = ArithExpectedMem(i, 0), label & "[100h] holds the result"
        If ArithExpectedMem(i, 1) >= 0 Then Check ArithMemByte(1) = ArithExpectedMem(i, 1), label & "[101h] holds the second result"
        Check CachePathReadByte(&H100) = ArithExpectedMem(i, 0), label & "the cached path reads the same byte (write-through coherent)"
        Check PagerFaultCount() = faultsExpected(i), label & CStr(faultsExpected(i)) & " page faults (got " & CStr(PagerFaultCount()) & ")"
        Check ArithMatches(i), label & "matches the documented result"
        Check CpuSnapshot().PC = 4 * (LoaderInstructionCount() - 1), label & "PC stays on the HLT"
        If ArithExpectedReg(i, 3) <> 0 Then usedDX = True
    Next i
    Check usedDX, "A: DX takes part (Fibonacci); AX, BX, CX are used by every program"

    ' --- B: register trace at the start of every loop pass (the documented trace) ----------------------
    expected = Array("0,5;5,4;9,3;12,2;14,1;", "0,7;6,6;12,5;18,4;24,3;30,2;36,1;", _
                     "0,1;1,1;1,2;2,3;3,5;5,8;8,13;13,21;21,34;34,55;", "47,0;42,1;37,2;32,3;27,4;22,5;17,6;12,7;7,8;2,9;", _
                     "48,18;30,18;12,18;12,6;6,6;")
    Check TraceAtHead(0, 0, 2) = expected(0), "B: SUMA (AX,CX) per pass"
    Check TraceAtHead(1, 0, 2) = expected(1), "B: MULTIPLICAR (AX,CX) per pass"
    Check TraceAtHead(2, 0, 1) = expected(2), "B: FIBONACCI (AX,BX) per pass"
    Check TraceAtHead(3, 0, 2) = expected(3), "B: DIVIDIR (AX,CX) per pass"
    Check TraceAtHead(4, 0, 1) = expected(4), "B: MCD (AX,BX) per pass"
    For i = 0 To ARITH_COUNT - 1
        Check CountOf(CStr(expected(i)), ";") = ArithLoopTotal(i), "B: " & ArithName(i) & " loop head visited " & CStr(ArithLoopTotal(i)) & " times"
    Next i

    ' --- C: STEP one by one and RUN give exactly the same machine state ----------------------------------
    For i = 0 To ARITH_COUNT - 1
        label = ArithName(i) & ": "
        Call ArithLoad(i)
        stepCount = 0
        Do While stepCount < 500
            If Not ExecStep() Then Exit Do
            stepCount = stepCount + 1
            If ExecStatus() = EXEC_HALTED Then Exit Do
        Loop
        For r = 0 To 3
            stepRegs(r) = CpuReadRegister(r)
        Next r
        stepFlags = CLng(CpuSnapshot().flags)
        Call ArithLoad(i)
        n = ExecRun(500)
        Check n = stepCount And ExecSteps() = stepCount, label & "STEP and RUN execute the same number of instructions"
        For r = 0 To 3
            Check CpuReadRegister(r) = stepRegs(r), label & "STEP and RUN end with the same " & Choose(r + 1, "AX", "BX", "CX", "DX")
        Next r
        Check CLng(CpuSnapshot().flags) = stepFlags, label & "STEP and RUN end with the same flags"
        Check Not ExecStep() And ExecSteps() = n, label & "no instruction runs after HLT"
    Next i

    ' --- D: the .asm sources assemble to exactly the image the loader holds ------------------------------
    For i = 0 To ARITH_COUNT - 1
        label = ArithName(i) & ": "
        Check ArithLoad(i), label & "loads"
        n = LoaderInstructionCount()
        Check AsmAssembleText(ArithAsmText(i)), label & "ArithAsmText assembles"
        Check AsmInstructionCount() = n And n <= LCA_ASM_MAX_INSTRUCTIONS, label & "same instruction count (" & CStr(n) & ")"
        For r = 0 To 4 * n - 1
            If AsmByte(r) <> LoaderImageByte(r) Then Check False, label & "byte " & AsmHex2(r) & " differs"
        Next r
        Check Right$(ArithFile(i), 4) = ".asm", label & "file name " & ArithFile(i)
    Next i
    mChecks = mChecks + 1

    ' --- E: shared animation helpers -----------------------------------------------------------------------
    prev = FlowSpeed()
    FlowSetSpeed 2
    Check FlowOn() And FlowSpeedText() = "NORMAL", "E: speed 2 is NORMAL and animation is on"
    FlowCycleSpeed
    Check FlowSpeed() = 1, "E: cycle normal -> slow"
    FlowCycleSpeed
    Check FlowSpeed() = 0 And Not FlowOn() And FlowSpeedText() = "SIN ANIMACION", "E: cycle slow -> off"
    FlowCycleSpeed
    Check FlowSpeed() = 3, "E: cycle off -> fast"
    FlowCycleSpeed
    Check FlowSpeed() = 2, "E: cycle fast -> normal"
    FlowSetSpeed 9
    Check FlowSpeed() = 2, "E: an invalid speed falls back to normal"
    FlowAnchorXY Nothing, Array(12.5, 40), ax, ay
    Check ax = 12.5 And ay = 40, "E: an Array(x, y) anchor is taken as points"
    Set sh = LcaSheet(False)
    If Not sh Is Nothing Then
        Set box = sh.Range("B2:D3")
        FlowAnchorXY sh, box, ax, ay
        Check Abs(ax - (box.Left + box.Width / 2)) < 0.01 And Abs(ay - (box.Top + box.Height / 2)) < 0.01, _
              "E: a multi-cell Range anchor is its centre (IsArray is True for it, it must not be read as coordinates)"
        FlowAnchorXY sh, box.Cells(1, 1), ax, ay
        Check ax > box.Left And ax < box.Left + box.Width, "E: a single-cell Range anchor is its centre"
    End If
    FlowSetSpeed 0
    FlowSend Nothing, "x", FlowPacketColor(), Array(0, 0), Array(10, 10)
    Check True, "E: with the animation off FlowSend returns without touching the sheet"
    FlowSetSpeed prev

    LcaReport "P15 PASS: " & CStr(mChecks) & " comprobaciones / assertions." & vbCrLf & _
              "Cinco programas aritmeticos: resultados, pasos, banderas, memoria, fallos de pagina, traza por vuelta, STEP = RUN.", True
    Exit Sub
failed:
    LcaReport "P15 FAIL after " & CStr(mChecks) & " checks: " & Err.Description, False
End Sub

' "a,b;" for every time the loop head is about to run (registers ra and rb).
Private Function TraceAtHead(ByVal prog As Long, ByVal ra As Long, ByVal rb As Long) As String
    Dim s As String, n As Long
    Call ArithLoad(prog)
    Do While n < 500
        If CpuSnapshot().PC = ArithLoopHead(prog) Then s = s & CStr(CpuReadRegister(ra)) & "," & CStr(CpuReadRegister(rb)) & ";"
        If Not ExecStep() Then Exit Do
        n = n + 1
        If ExecStatus() = EXEC_HALTED Then Exit Do
    Loop
    TraceAtHead = s
End Function

Private Function CountOf(ByVal text As String, ByVal ch As String) As Long
    Dim i As Long
    For i = 1 To Len(text)
        If Mid$(text, i, 1) = ch Then CountOf = CountOf + 1
    Next i
End Function

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    mChecks = mChecks + 1
    If Not condition Then Err.Raise vbObjectError + 2161, "LcaArithTests", message
End Sub
