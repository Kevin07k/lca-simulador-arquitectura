Attribute VB_Name = "LcaMicroTests"
Option Explicit

Private mChecks As Long

' Desktop Excel verification for P05b: the complete control word of every micro-step.
'   A: the 25-signal vector, names and device buffers;
'   B: every instruction class produces its exact micro-sequence (fetch T1..T4, execute, PC update);
'   C: the physical rule (one En_* at most, otherwise High-Z) holds on every micro-step of every demo program;
'   D: IN on an empty queue, HLT and a taken branch;
'   E: the drawn band shows the selected micro-step (when the LCA sheet is built).
Public Sub VerificarMicropalabraP05B()
    Dim ok As Boolean, i As Long, p As Long, n As Long, steps As Long, ws As Worksheet
    On Error GoTo failed
    mChecks = 0

    ' --- A: vector, names, buffers --------------------------------------------------------------
    Check MIC_SIGNALS = 25 And MIC_ENABLES = 10, "A: 25 signals, 10 tri-state enables"
    Check MicroSignalName(MIC_EN_PC) = "En_PC" And MicroSignalName(MIC_LD_MAR) = "Ld_MAR" And _
          MicroSignalName(MIC_MEM_R) = "Mem_R" And MicroSignalName(MIC_EN_KBD) = "En_KBD" And _
          MicroSignalName(MIC_LD_LCD) = "Ld_LCD" And MicroSignalName(MIC_LD_AX + 3) = "Ld_DX" And _
          MicroSignalName(MIC_EN_AX + 2) = "En_CX" And MicroSignalName(MIC_INC_PC) = "Inc_PC", "A: signal names match the brief"
    Check MicroDeviceName(MIC_EN_ALU) = "ALU" And MicroDeviceName(MIC_EN_MEM) = "MEM", "A: device names"

    ' --- B: sequences of every instruction class -------------------------------------------------
    Check LoaderLoadSource(Array( _
        "        MOV AX, #05h", _
        "        MOV BX, #03h", _
        "        ADD AX, BX", _
        "        SUB AX, BX", _
        "        CMP AX, BX", _
        "        INC CX", _
        "        DEC CX", _
        "        MOV DX, AX", _
        "        ST [100h], AX", _
        "        LD CX, [100h]", _
        "        JZ ok", _
        "        NOP", _
        "ok:     JNZ nope", _
        "        OUT AX", _
        "        IN BX", _
        "        HLT", _
        "nope:   HLT")), "B: test program assembles"
    Call KeyboardPush(&H41)

    StepOk "MOV AX,#05h"
    Check MicroCount() = 6, "B: MOV imm = 4 fetch + 1 execute + PC = 6 micro-steps"
    Fetch4
    Check Only(4, MIC_EN_IR, MIC_LD_AX), "B: T5 IR(imm) -> AX: En_IR, Ld_AX"
    Check MicroBusData(4) = "imm=05h", "B: immediate crosses the bus"
    Check Only(5, MIC_LD_PC, MIC_INC_PC) And MicroDriver(5) = "High-Z", "B: T6 PC+4: Ld_PC, Inc_PC, bus High-Z"

    StepOk "MOV BX,#03h"
    Check Only(4, MIC_EN_IR, MIC_LD_AX + 1), "B: MOV BX loads Ld_BX"

    StepOk "ADD AX,BX"
    Check MicroCount() = 7, "B: ADD = 7 micro-steps"
    Check Only(4, MIC_EN_AX, MIC_LD_A) And MicroBusData(4) = "AX=05h", "B: T5 AX -> A (En_AX, Ld_A)"
    Check Only(5, MIC_EN_ALU, MIC_LD_AX, MIC_LD_FLG) And MicroAluOp(5) = "ADD" And MicroBusData(5) = "res=08h", _
          "B: T6 ALU result -> AX and flags (En_ALU, Ld_AX, Ld_FLG)"

    StepOk "SUB AX,BX"
    Check MicroAluOp(5) = "SUB" And MicroBusData(4) = "AX=08h" And MicroBusData(5) = "res=05h", "B: SUB shows A=08h and result 05h"

    StepOk "CMP AX,BX"
    Check MicroCount() = 7 And Only(5, MIC_LD_FLG) And MicroDriver(5) = "High-Z" And MicroAluOp(5) = "CMP", _
          "B: CMP only loads the flags; the bus stays High-Z"
    Check Not MicroSignal(5, MIC_LD_AX) And Not MicroSignal(5, MIC_EN_ALU), "B: CMP does not write a register"

    StepOk "INC CX"
    Check Only(4, MIC_EN_AX + 2, MIC_LD_A) And Only(5, MIC_EN_ALU, MIC_LD_AX + 2, MIC_LD_FLG) And MicroAluOp(5) = "INC", "B: INC CX uses CX"
    StepOk "DEC CX"
    Check MicroAluOp(5) = "DEC" And MicroBusData(5) = "res=00h", "B: DEC CX leaves 00h"

    StepOk "MOV DX,AX"
    Check Only(4, MIC_EN_AX, MIC_LD_AX + 3), "B: MOV DX,AX: En_AX -> Ld_DX"

    StepOk "ST [100h],AX"
    Check MicroCount() = 8, "B: ST = 8 micro-steps"
    Check Only(4, MIC_EN_IR, MIC_LD_MAR) And MicroBusData(4) = "VA=100h", "B: ST T5 address -> MAR"
    Check Only(5, MIC_EN_AX, MIC_LD_MDR), "B: ST T6 AX -> MDR"
    Check Only(6, MIC_MEM_W, MIC_EN_MDR) And MicroDriver(6) = "MDR", "B: ST T7 MDR drives memory with Mem_W"

    StepOk "LD CX,[100h]"
    Check MicroCount() = 8 And Only(4, MIC_EN_IR, MIC_LD_MAR), "B: LD = 8 micro-steps, address -> MAR"
    Check Only(5, MIC_MEM_R, MIC_EN_MEM, MIC_LD_MDR) And MicroDriver(5) = "MEM", "B: LD T6 memory drives, Mem_R, Ld_MDR"
    Check Only(6, MIC_EN_MDR, MIC_LD_AX + 2), "B: LD T7 MDR -> CX"

    ' --- D: branches, OUT, IN, HLT ---------------------------------------------------------------------
    StepOk "JZ ok"
    Check MicroCount() = 5 And Only(4, MIC_EN_IR, MIC_LD_PC), "B: taken JZ = 5 micro-steps, IR(dest) -> PC, no PC+4"
    Check Not MicroSignal(4, MIC_INC_PC), "B: taken branch does not increment"
    StepOk "JNZ nope"
    Check MicroCount() = 6 And Only(4), "B: untaken JNZ: no transfer"
    Check Only(5, MIC_LD_PC, MIC_INC_PC), "B: untaken JNZ then PC+4"
    StepOk "OUT AX"
    Check Only(4, MIC_IO, MIC_MEM_W, MIC_EN_AX, MIC_LD_LCD) And MicroDriver(4) = "AX", "D: OUT: IO, Mem_W, En_AX, Ld_LCD"
    StepOk "IN BX"
    Check Only(4, MIC_IO, MIC_MEM_R, MIC_EN_KBD, MIC_LD_AX + 1) And MicroDriver(4) = "KBD", "D: IN: IO, Mem_R, En_KBD, Ld_BX"
    StepOk "HLT"
    Check MicroCount() = 5 And Not MicroSignal(4, MIC_LD_PC) And MicroDriver(4) = "High-Z", "D: HLT: no PC update"

    ' IN on an empty queue waits; the keyboard buffer stays in High-Z
    Check LoaderLoadSource(Array("        IN AX", "        HLT")), "D: IN program assembles"
    Check Not ExecStep() And ExecStatus() = EXEC_WAIT, "D: IN on an empty queue waits"
    Check MicroCount() = 5 And Only(4, MIC_IO, MIC_MEM_R) And MicroDriver(4) = "High-Z" And Not MicroSignal(4, MIC_EN_KBD), _
          "D: waiting IN: En_KBD stays 0 (High-Z), PC does not advance"
    Check MicroCheckRules() = vbNullString, "D: the waiting sequence respects the rules"

    ' --- C: the physical rule on every micro-step of the demo programs ---------------------------------
    For p = 0 To 4
        Check LoaderLoadSource(ArithSource(p)), "C: arithmetic demo " & CStr(p) & " loads"
        n = 0
        Do While n < 400
            If Not ExecStep() Then Exit Do
            n = n + 1
            Check MicroCheckRules() = vbNullString, "C: " & ArithName(p) & " step " & CStr(n) & ": " & MicroCheckRules()
            steps = steps + MicroCount()
            If ExecStatus() = EXEC_HALTED Then Exit Do
        Loop
        Check ExecStatus() = EXEC_HALTED, "C: " & ArithName(p) & " reaches HLT"
    Next p
    Check LoaderLoadSource(MemDemoSource()), "C: memory demo loads"
    n = 0
    Do While n < 400
        If Not ExecStep() Then Exit Do
        n = n + 1
        Check MicroCheckRules() = vbNullString, "C: memory demo step " & CStr(n) & ": " & MicroCheckRules()
        If ExecStatus() = EXEC_HALTED Then Exit Do
    Loop
    Check n = 30, "C: memory demo ran its 30 instructions with a legal control word each time"
    Check steps > 300, "C: more than 300 micro-steps checked"
    Check LoaderLoadSource(EchoSource()), "C: echo demo loads"
    Check KeyboardPushText("OK.") = 3, "C: keys queued"
    Do While n < 800 And ExecStatus() <> EXEC_HALTED
        If Not ExecStep() Then Exit Do
        n = n + 1
        Check MicroCheckRules() = vbNullString, "C: echo step: " & MicroCheckRules()
    Loop

    ' --- E: the band shows the selected micro-step -------------------------------------------------------
    Check LoaderLoadSource(Array("        MOV AX, #05h", "        MOV BX, #03h", "        ADD AX, BX", "        HLT")), "E: panel program loads"
    StepOk "MOV AX,#05h"
    StepOk "MOV BX,#03h"
    StepOk "ADD AX,BX"
    If MicroPanelConstruido() Then
        MicroPanelRefrescar
        MicroPanelSeleccionar 5
        Set ws = LcaSheet(False)
        Check CLng(ws.Cells(LCA_BASE_MICRO + 14, LCA_COL_MICRO + 3 + MIC_EN_ALU).Value) = 1 And _
              CLng(ws.Cells(LCA_BASE_MICRO + 14, LCA_COL_MICRO + 3 + MIC_EN_PC).Value) = 0, "E: band shows En_ALU = 1 and En_PC = 0 at T6"
        Check CStr(ws.Cells(LCA_BASE_MICRO + 26, LCA_COL_MICRO + 3 + MIC_EN_ALU).Value) = "1 CONDUCE" And _
              CStr(ws.Cells(LCA_BASE_MICRO + 26, LCA_COL_MICRO + 3 + MIC_EN_PC).Value) = "High-Z", "E: ALU buffer conducts, PC buffer is High-Z"
        Check CLng(ws.Cells(LCA_BASE_MICRO + 18, LCA_COL_MICRO + 3 + MIC_LD_AX - MIC_LD_PC).Value) = 1, "E: band shows Ld_AX = 1"
        MicroPanelSeleccionar 0
        Check CLng(ws.Cells(LCA_BASE_MICRO + 14, LCA_COL_MICRO + 3 + MIC_EN_PC).Value) = 1, "E: T1 shows En_PC = 1"
        Check CLng(ws.Cells(LCA_BASE_MICRO + 18, LCA_COL_MICRO + 3 + MIC_LD_MAR - MIC_LD_PC).Value) = 1, "E: T1 shows Ld_MAR = 1"
    End If
    ExecReset
    LcaReport "P05b PASS: " & CStr(mChecks) & " comprobaciones / assertions." & vbCrLf & _
              "Micro-palabra de 25 senales por micro-paso, un solo Enable (resto en High-Z) en cada paso de todas las instrucciones y demos.", True
    Exit Sub
failed:
    LcaReport "P05b FAIL after " & CStr(mChecks) & " checks: " & Err.Description, False
End Sub

' Executes one instruction and checks that it is the expected one (same mnemonic as its disassembly).
Private Sub StepOk(ByVal expected As String)
    Check ExecStep(), "step " & expected & " completes"
    Check Mnemonic(MicroInstruction()) = Mnemonic(expected), _
          "recorded instruction is " & expected & " (got " & MicroInstruction() & ")"
    Check MicroCheckRules() = vbNullString, expected & ": " & MicroCheckRules()
End Sub

Private Function Mnemonic(ByVal text As String) As String
    text = Trim$(text)
    If InStr(text, " ") > 0 Then text = Left$(text, InStr(text, " ") - 1)
    Mnemonic = UCase$(text)
End Function

' T1..T4 of the fetch, identical for every instruction.
Private Sub Fetch4()
    Check Only(0, MIC_EN_PC, MIC_LD_MAR) And MicroDriver(0) = "PC", "fetch T1: PC -> MAR"
    Check Only(1, MIC_MEM_R, MIC_EN_MEM, MIC_LD_MDR) And MicroDriver(1) = "MEM", "fetch T2: memory -> MDR with Mem_R"
    Check Only(2, MIC_EN_MDR, MIC_LD_IR) And MicroDriver(2) = "MDR", "fetch T3: MDR -> IR"
    Check Only(3) And MicroDriver(3) = "High-Z", "fetch T4: decode, no signal, bus High-Z"
End Sub

' True when exactly the listed signals are 1 in micro-step i (and the other of the 25 are 0).
Private Function Only(ByVal i As Long, ParamArray wanted() As Variant) As Boolean
    Dim s As Long, k As Long, expected As Boolean
    Only = True
    For s = 0 To MIC_SIGNALS - 1
        expected = False
        For k = LBound(wanted) To UBound(wanted)
            If CLng(wanted(k)) = s Then expected = True
        Next k
        If MicroSignal(i, s) <> expected Then Only = False
    Next s
End Function

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    mChecks = mChecks + 1
    If Not condition Then Err.Raise vbObjectError + 2212, "LcaMicroTests", message
End Sub
