Attribute VB_Name = "LcaExecutor"
Option Explicit

' P12: complete instruction cycle for the fixed P02 ISA.
'   FETCH  four bytes at PC through the P10 path (page fault -> L1I -> L2 -> RAM), IR = opcode;
'   DECODE validate opcode, register codes, unused fields, data addresses and branch targets;
'   EXECUTE ALU / MOV / LD / ST (L1D) / branches / IN (keyboard P13) / OUT (LCD P14);
'   STORE  commit registers, flags and PC. PC advances 4 except a taken jump, HLT or an IN wait.
' A fatal error stops the machine, leaves PC on the offending instruction and commits nothing.
' IN on an empty queue is a WAIT (not an error): PC and registers stay, RUN pauses until a key.
' No worksheet dependency.
Public Const LCA_ERR_EXEC As Long = vbObjectError + 2121
Public Const EXEC_READY As String = "READY"
Public Const EXEC_HALTED As String = "HALTED"
Public Const EXEC_WAIT As String = "WAIT_IN"
Public Const EXEC_ERROR As String = "ERROR"
Public Const EXEC_LIMIT As String = "LIMIT"
Private Const TRACE_MAX As Long = 128

Private mStatus As String
Private mErrorText As String
Private mSteps As Long
Private mLastKind As String        ' FETCH-only kinds: NOP HLT MOV LD ST ALU JMP IN OUT WAIT ERROR
Private mLastText As String
Private mLastPC As Long
Private mLastTaken As Boolean
Private mLastFault As Boolean
Private mLastDataVA As Long
Private mTrace(0 To TRACE_MAX - 1) As String
Private mTraceCount As Long

' Restores the loaded program and the whole machine (loader contract), then READY.
Public Sub ExecReset()
    LoaderResetExecution     ' calls ExecClearState
End Sub

' Clears only the executor state. Called by LoaderResetExecution.
Public Sub ExecClearState()
    Dim i As Long
    mStatus = EXEC_READY
    mErrorText = vbNullString
    mSteps = 0
    mLastKind = vbNullString
    mLastText = "RESET: PC = 000h"
    mLastPC = 0
    mLastTaken = False
    mLastFault = False
    mLastDataVA = -1
    mTraceCount = 0
    MicroClear
    For i = 0 To TRACE_MAX - 1
        mTrace(i) = vbNullString
    Next i
    AddTrace "RESET  PC=000h  AX=BX=CX=DX=00h  flags=000  cola/LCD vacios"
End Sub

' Executes one instruction. Returns True when it completed (HLT included); False on wait/error/stopped.
Public Function ExecStep() As Boolean
    Dim pc As Long, op As Long, r As Long, x As Long, h As Long, va As Long, value As Long
    Dim faultsBefore As Long, cpu As LcaCpuState, text As String, effect As String
    Dim errNo As Long, errText As String, a0 As Long, b0 As Long
    If mStatus = EXEC_HALTED Or mStatus = EXEC_ERROR Then Exit Function
    If Not LoaderLoaded() Then Err.Raise LCA_ERR_EXEC, "LcaExecutor", "NO_PROGRAM_LOADED"
    On Error GoTo failed
    cpu = CpuSnapshot()
    pc = cpu.PC
    mLastPC = pc
    mLastTaken = False
    mLastDataVA = -1
    faultsBefore = PagerFaultCount()

    ' FETCH
    If pc > &HFC Then Err.Raise LCA_ERR_EXEC, "LcaExecutor", "PC_OUT_OF_CODE"
    If (pc Mod 4) <> 0 Then Err.Raise LCA_ERR_EXEC, "LcaExecutor", "UNALIGNED_PC"
    CachePathFetchFour pc
    op = ControlFetched(0): r = ControlFetched(1): x = ControlFetched(2): h = ControlFetched(3)

    ' DECODE (nothing is committed before this passes)
    Decode op, r, x, h
    text = AsmDisassemble(op, r, x, h)
    va = 256 * h + x
    If r <= 3 Then a0 = CpuReadRegister(r)                  ' operands before the ALU/MOV changes them (P05b micro-word)
    If (op = &H11 Or (op >= &H20 And op <= &H22)) And x <= 3 Then b0 = CpuReadRegister(x)

    ' EXECUTE + STORE
    Select Case op
        Case &H0
            mLastKind = "NOP"
            AdvancePC pc
        Case &H1
            mLastKind = "HLT"
            mStatus = EXEC_HALTED
            effect = "HALT (PC se queda en " & AsmHex3(pc) & ")"
        Case &H10
            mLastKind = "MOV"
            CpuWriteRegister r, x
            effect = RegName(r) & "=" & AsmHex2(x)
            AdvancePC pc
        Case &H11
            mLastKind = "MOV"
            CpuWriteRegister r, CpuReadRegister(x)
            effect = RegName(r) & "=" & AsmHex2(CpuReadRegister(r))
            AdvancePC pc
        Case &H12
            mLastKind = "LD"
            mLastDataVA = va
            value = CachePathReadByte(va)
            CpuWriteRegister r, value
            effect = RegName(r) & "=[" & AsmHex3(va) & "]=" & AsmHex2(value) & PathNote()
            AdvancePC pc
        Case &H13
            mLastKind = "ST"
            mLastDataVA = va
            CachePathStoreByte va, r
            effect = "[" & AsmHex3(va) & "]=" & RegName(r) & "=" & AsmHex2(CpuReadRegister(r)) & PathNote()
            AdvancePC pc
        Case &H20 To &H24
            mLastKind = "ALU"
            value = AluExecute(op, r, x)
            If op = &H22 Then
                effect = "flags de " & RegName(r) & "-" & RegName(x) & "=" & AsmHex2(value)
            Else
                effect = RegName(r) & "=" & AsmHex2(value)
            End If
            AdvancePC pc
        Case &H30 To &H34
            mLastKind = "JMP"
            mLastTaken = BranchTaken(op)
            If mLastTaken Then
                CpuWriteControl LcaPC, va
                effect = "tomado / taken -> PC=" & AsmHex3(va)
            Else
                effect = "no tomado / not taken"
                AdvancePC pc
            End If
        Case &H40
            If KeyboardIn(r) Then
                mLastKind = "IN"
                effect = RegName(r) & "=" & AsmHex2(CpuReadRegister(r)) & " desde teclado (En_KBD=1)"
                AdvancePC pc
            Else
                mLastKind = "WAIT"
                mStatus = EXEC_WAIT
                mLastText = AsmHex3(pc) & "  " & text & "  ->  IN_EMPTY: espera una tecla (PC no avanza)"
                MicroRecord pc, op, r, x, h, False, True, a0, b0
                AddTrace "#" & CStr(mSteps + 1) & "  " & mLastText
                Exit Function
            End If
        Case &H41
            mLastKind = "OUT"
            value = LcdOut(r)
            effect = "LCD[" & CStr(value) & "]=" & AsmHex2(CpuReadRegister(r)) & " (Ld_LCD=1)"
            AdvancePC pc
    End Select

    MicroRecord pc, op, r, x, h, mLastTaken, False, a0, b0      ' P05b: full control word of every micro-step
    mLastFault = PagerFaultCount() > faultsBefore
    mSteps = mSteps + 1
    If mStatus <> EXEC_HALTED Then mStatus = EXEC_READY
    mLastText = AsmHex3(pc) & "  " & text & "  ->  " & effect
    If mLastFault Then mLastText = mLastText & "  [PAGE_FAULT]"
    cpu = CpuSnapshot()
    AddTrace "#" & CStr(mSteps) & "  " & mLastText & "  |  AX=" & AsmHex2(cpu.AX) & " BX=" & AsmHex2(cpu.BX) & _
             " CX=" & AsmHex2(cpu.CX) & " DX=" & AsmHex2(cpu.DX) & "  ZF" & FlagBit(LCA_FLAG_ZF) & _
             " CF" & FlagBit(LCA_FLAG_CF) & " SF" & FlagBit(LCA_FLAG_SF) & "  PC=" & AsmHex3(cpu.PC)
    ExecStep = True
    Exit Function
failed:
    errNo = Err.Number
    errText = Err.Description
    mStatus = EXEC_ERROR
    mErrorText = errText
    mLastKind = "ERROR"
    mLastText = AsmHex3(pc) & "  ERROR " & errText & "  (PC se queda / PC kept)"
    AddTrace "#" & CStr(mSteps + 1) & "  " & mLastText
    ' PC stays on the offending instruction; restore it in case a partial step moved it.
    On Error Resume Next
    CpuWriteControl LcaPC, pc
End Function

' Runs until HLT, error, IN wait or maxSteps. Returns the number of completed instructions.
Public Function ExecRun(Optional ByVal maxSteps As Long = 1000) As Long
    Dim n As Long
    If mStatus = EXEC_LIMIT Or mStatus = EXEC_WAIT Then mStatus = EXEC_READY
    Do While n < maxSteps
        If Not ExecStep() Then Exit Do
        n = n + 1
        If mStatus = EXEC_HALTED Then Exit Do
    Loop
    If n >= maxSteps And mStatus = EXEC_READY Then
        mStatus = EXEC_LIMIT
        AddTrace "RUN detenido tras " & CStr(maxSteps) & " instrucciones (EJECUTAR otra vez para seguir)"
    End If
    ExecRun = n
End Function

' --- Inspection ------------------------------------------------------------------------------

Public Function ExecStatus() As String
    If Len(mStatus) = 0 Then mStatus = EXEC_READY
    ExecStatus = mStatus
End Function

Public Function ExecErrorText() As String
    ExecErrorText = mErrorText
End Function

Public Function ExecSteps() As Long
    ExecSteps = mSteps
End Function

Public Function ExecLastKind() As String
    ExecLastKind = mLastKind
End Function

Public Function ExecLastText() As String
    ExecLastText = mLastText
End Function

Public Function ExecLastPC() As Long
    ExecLastPC = mLastPC
End Function

Public Function ExecLastTaken() As Boolean
    ExecLastTaken = mLastTaken
End Function

Public Function ExecLastFault() As Boolean
    ExecLastFault = mLastFault
End Function

Public Function ExecLastDataVA() As Long
    ExecLastDataVA = mLastDataVA
End Function

Public Function ExecTraceCount() As Long
    ExecTraceCount = mTraceCount
    If ExecTraceCount > TRACE_MAX Then ExecTraceCount = TRACE_MAX
End Function

' index 0 = newest line.
Public Function ExecTraceLine(ByVal index As Long) As String
    If index < 0 Or index >= ExecTraceCount() Then Err.Raise LCA_ERR_EXEC, "LcaExecutor", "TRACE_INDEX"
    ExecTraceLine = mTrace((mTraceCount - 1 - index) Mod TRACE_MAX)
End Function

' --- Internals --------------------------------------------------------------------------------

' P02 decoder: raises the contract error codes before anything is committed.
Private Sub Decode(ByVal op As Long, ByVal r As Long, ByVal x As Long, ByVal h As Long)
    Dim va As Long
    va = 256 * h + x
    Select Case op
        Case &H0, &H1
            If r <> 0 Or x <> 0 Or h <> 0 Then Fail "INVALID_ENCODING"
        Case &H10
            NeedReg r
            If h <> 0 Then Fail "INVALID_ENCODING"
        Case &H11, &H20, &H21, &H22
            NeedReg r
            NeedReg x
            If h <> 0 Then Fail "INVALID_ENCODING"
        Case &H12, &H13
            NeedReg r
            If h > 1 Then Fail "INVALID_ENCODING"
            If va < &H100 Or va > &H1DF Then Fail "PROTECTION"
        Case &H23, &H24, &H40, &H41
            NeedReg r
            If x <> 0 Or h <> 0 Then Fail "INVALID_ENCODING"
        Case &H30 To &H34
            If r <> 0 Or h > 1 Then Fail "INVALID_ENCODING"
            If va > &HFC Or (va Mod 4) <> 0 Then Fail "INVALID_BRANCH_TARGET"
        Case Else
            Fail "INVALID_OPCODE"
    End Select
End Sub

Private Sub NeedReg(ByVal code As Long)
    If code > 3 Then Fail "INVALID_REGISTER"
End Sub

Private Sub Fail(ByVal code As String)
    Err.Raise LCA_ERR_EXEC, "LcaExecutor", code
End Sub

Private Function BranchTaken(ByVal op As Long) As Boolean
    Select Case op
        Case &H30: BranchTaken = True
        Case &H31: BranchTaken = CpuFlagIsSet(LCA_FLAG_ZF)
        Case &H32: BranchTaken = Not CpuFlagIsSet(LCA_FLAG_ZF)
        Case &H33: BranchTaken = CpuFlagIsSet(LCA_FLAG_CF)
        Case &H34: BranchTaken = CpuFlagIsSet(LCA_FLAG_SF)
    End Select
End Function

Private Sub AdvancePC(ByVal pc As Long)
    CpuWriteControl LcaPC, pc + 4
End Sub

Private Function PathNote() As String
    Dim s As LcaPathStep
    s = CachePathLast()
    PathNote = "  (PA " & AsmHex2(s.PA) & ", L1D " & IIf(s.L1Misses > 0, "miss", "hit")
    If s.PageFault Then PathNote = PathNote & ", fallo de pagina"
    PathNote = PathNote & ")"
End Function

Private Function FlagBit(ByVal mask As Long) As String
    FlagBit = IIf(CpuFlagIsSet(mask), "1", "0")
End Function

Private Function RegName(ByVal code As Long) As String
    RegName = Choose(code + 1, "AX", "BX", "CX", "DX")
End Function

Private Sub AddTrace(ByVal text As String)
    mTrace(mTraceCount Mod TRACE_MAX) = text
    mTraceCount = mTraceCount + 1
End Sub
