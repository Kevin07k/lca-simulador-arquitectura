Attribute VB_Name = "LcaExecutorTests"
Option Explicit

Private mChecks As Long

' P12 regression: full fetch/decode/execute/store cycle, flags and branches, loops, HLT,
' IN/OUT through the P13/P14 controllers, fatal errors without commit and the global reset.
Public Sub VerificarCicloP12()
    Dim cpu As LcaCpuState, n As Long, src As String, i As Long
    On Error GoTo failed
    mChecks = 0

    ' --- A: sample program (loop, ST, LD, HLT) -----------------------------------------------
    Check LoaderLoadSource(LoaderSampleProgram()), "A: sample loads"
    Check ExecStatus() = EXEC_READY And ExecSteps() = 0, "A: READY after load"
    n = ExecRun()
    cpu = CpuSnapshot()
    Check ExecStatus() = EXEC_HALTED And n = 20 And ExecSteps() = 20, "A: 20 instructions until HLT (got " & CStr(n) & ")"
    Check cpu.AX = &HF And cpu.BX = &HF And cpu.CX = 0, "A: AX = BX = 0Fh, CX = 0"
    Check cpu.PC = &H1C And ExecLastKind() = "HLT", "A: PC stays on HLT (01Ch)"
    Check CpuFlagIsSet(LCA_FLAG_ZF) And Not CpuFlagIsSet(LCA_FLAG_CF) And Not CpuFlagIsSet(LCA_FLAG_SF), _
          "A: flags from the last DEC (ZF=1) preserved by JNZ/ST/LD/HLT"
    Check PagerFaultCount() = 2, "A: two page faults (VPN0 code, VPN8 data)"
    Check CacheHits(LCA_CACHE_L1I) > 0 And CacheMisses(LCA_CACHE_L1D) >= 1, "A: fetches hit L1I, data used L1D"
    Check Not ExecStep() And ExecSteps() = 20, "A: no step after HLT"
    Check CachePathReadByte(&H100) = &HF, "A: memory [100h] = 0Fh"

    ' --- B: single steps --------------------------------------------------------------------
    ExecReset
    cpu = CpuSnapshot()
    Check cpu.PC = 0 And cpu.AX = 0 And ExecStatus() = EXEC_READY, "B: reset restores PC and registers"
    Check ExecStep() And CpuSnapshot().PC = 4 And CpuSnapshot().IR = &H10, "B: MOV AX advances PC 4"
    Check ExecStep() And CpuReadRegister(LcaCX) = 5, "B: MOV CX, #05h"
    Check ExecStep() And CpuReadRegister(LcaAX) = 5, "B: ADD AX, CX"
    Check ExecStep() And CpuReadRegister(LcaCX) = 4 And Not CpuFlagIsSet(LCA_FLAG_ZF), "B: DEC CX, ZF=0"
    Check ExecStep() And ExecLastTaken() And CpuSnapshot().PC = 8, "B: JNZ taken back to 008h"
    Check InStr(ExecTraceLine(0), "JNZ 008h") > 0 And InStr(ExecTraceLine(0), "PC=008h") > 0, "B: trace shows the jump"

    ' --- C: flags and every branch -----------------------------------------------------------
    src = Join(Array("MOV AX, #0FFh", "INC AX", "JC carry", "HLT", "carry: JZ zero", "HLT", _
                     "zero: MOV BX, #80h", "MOV CX, #01h", "CMP CX, BX", "JS neg", "HLT", _
                     "neg: SUB BX, CX", "JS bad", "JZ bad", "JC bad", "MOV DX, #33h", "MOV AX, BX", "HLT", _
                     "bad: MOV DX, #0EEh", "HLT"), vbLf)
    Check LoaderLoadText(src), "C: branch program loads"
    ExecRun
    Check ExecStatus() = EXEC_HALTED, "C: halts"
    Check CpuReadRegister(LcaDX) = &H33, "C: JC/JZ/JS taken and not-taken paths correct (DX=33h)"
    Check CpuReadRegister(LcaBX) = &H7F And CpuReadRegister(LcaCX) = 1 And CpuReadRegister(LcaAX) = &H7F, _
          "C: SUB 80h-01h = 7Fh, CMP did not write CX, MOV AX, BX"

    ' --- D: IN / OUT echo with wait and resume -----------------------------------------------
    Check LoaderLoadText("eco: IN AX" & vbLf & "OUT AX" & vbLf & "JMP eco"), "D: echo loads"
    Check ExecRun() = 0 And ExecStatus() = EXEC_WAIT And CpuSnapshot().PC = 0, "D: IN waits on empty queue, PC kept"
    Check KeyboardPushText("HI") = 2, "D: two keys typed"
    n = ExecRun()
    Check n = 6 And ExecStatus() = EXEC_WAIT And CpuSnapshot().PC = 0, "D: two echo rounds, then waits again"
    Check Left$(LcdText(), 2) = "HI" And KeyboardConsumed() = 2 And KeyboardCount() = 0, "D: LCD shows HI, each key used once"
    Check CpuReadRegister(LcaAX) = Asc("I"), "D: AX holds the last key"

    ' --- E: fatal errors stop without committing ---------------------------------------------
    Check LoaderLoadSource(LoaderSampleProgram()), "E: sample reloads"
    SwapWriteByte 0, 0, &H77
    Check Not ExecStep() And ExecStatus() = EXEC_ERROR And ExecErrorText() = "INVALID_OPCODE", "E: INVALID_OPCODE"
    Check CpuSnapshot().PC = 0 And ExecSteps() = 0, "E: PC stays on the offending instruction"
    Check Not ExecStep() And ExecRun() = 0, "E: machine stopped until reset"
    Check ErrorOf(&H10, 5, 1, 0) = "INVALID_REGISTER", "E: INVALID_REGISTER"
    Check ErrorOf(&H1, 0, 5, 0) = "INVALID_ENCODING", "E: INVALID_ENCODING (HLT with operand)"
    Check ErrorOf(&H12, 0, &HE8, 1) = "PROTECTION", "E: LD on the LCD address is PROTECTION"
    Check ErrorOf(&H30, 0, 2, 0) = "INVALID_BRANCH_TARGET", "E: unaligned branch target"
    Check CpuReadRegister(LcaAX) = 0 And CpuReadRegister(LcaBX) = 0, "E: no register committed"
    src = ""
    For i = 1 To 64
        src = src & "NOP" & vbLf
    Next i
    Check LoaderLoadText(src), "E: 64 NOPs load"
    Check ExecRun() = 64 And ExecStatus() = EXEC_ERROR And ExecErrorText() = "PC_OUT_OF_CODE", "E: PC_OUT_OF_CODE at 100h"
    Check CpuSnapshot().PC = &H100, "E: PC = 100h"

    ' --- F: global reset includes the keyboard and the LCD -----------------------------------
    Check LoaderLoadText("MOV AX, #41h" & vbLf & "OUT AX" & vbLf & "HLT"), "F: OUT program loads"
    ExecRun
    Check LcdCell(0) = &H41 And LcdCursor() = 1, "F: LCD shows A"
    Call KeyboardPush(&H5A)
    ExecReset
    Check LcdText() = "        " And LcdCursor() = 0, "F: reset clears the LCD"
    Check KeyboardCount() = 0 And KeyboardHead() = 0 And RamReadIO(&HE0) = 0, "F: reset empties the keyboard queue"
    Check ExecStatus() = EXEC_READY And ExecSteps() = 0 And CpuSnapshot().PC = 0, "F: executor READY"
    Check ExecRun() = 3 And LcdCell(0) = &H41, "F: program runs again after reset"

    LcaReport "P12 PASS: " & CStr(mChecks) & " comprobaciones / assertions." & vbCrLf & _
              "Ciclo fetch/decode/execute/store, saltos, bucles, HLT, IN/OUT y errores verificados.", True
    Exit Sub
failed:
    LcaReport "P12 FAIL after " & CStr(mChecks) & " checks: " & Err.Description, False
End Sub

' Loads a one-instruction image (patched into the backing store) and returns the decoder error.
Private Function ErrorOf(ByVal op As Long, ByVal r As Long, ByVal x As Long, ByVal h As Long) As String
    Call LoaderLoadText("HLT")
    SwapWriteByte 0, 0, op
    SwapWriteByte 0, 1, r
    SwapWriteByte 0, 2, x
    SwapWriteByte 0, 3, h
    Call ExecStep
    ErrorOf = ExecErrorText()
End Function

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    mChecks = mChecks + 1
    If Not condition Then Err.Raise vbObjectError + 2122, "LcaExecutorTests", message
End Sub
