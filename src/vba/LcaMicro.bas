Attribute VB_Name = "LcaMicro"
Option Explicit

' P05b: complete control word (micro-palabra de control) of the control unit, one per clock micro-step.
' After every instruction the executor records its micro-sequence here (fetch T1..T4 + execute + PC update).
' The 25 signals are the vector asked by the design brief:
'   ENABLE (output buffer to the bus, one source at a time):
'       En_PC En_MDR En_IR En_AX En_BX En_CX En_DX En_ALU En_MEM En_KBD
'   LOAD (receiver latches from the bus):
'       Ld_PC Ld_MAR Ld_MDR Ld_IR Ld_AX Ld_BX Ld_CX Ld_DX Ld_A Ld_FLG Ld_LCD
'   CONTROL:
'       Mem_R Mem_W IO Inc_PC
' Physical rule enforced by MicroCheckRules: at most ONE En_* is 1 in each micro-step; every other buffer is
' in High-Z. When no En_* is 1 the bus floats (High-Z). Ld_PC with Inc_PC takes the PC+4 incrementer, not the bus.
' The ALU reads its second operand through the register-file port (not the bus); En_ALU drives only the result.
' No worksheet dependency: the panel (LcaMicroPanel) only reads this state.
Public Const MIC_EN_PC As Long = 0
Public Const MIC_EN_MDR As Long = 1
Public Const MIC_EN_IR As Long = 2
Public Const MIC_EN_AX As Long = 3          ' En_AX..En_DX = 3..6
Public Const MIC_EN_ALU As Long = 7
Public Const MIC_EN_MEM As Long = 8
Public Const MIC_EN_KBD As Long = 9
Public Const MIC_LD_PC As Long = 10
Public Const MIC_LD_MAR As Long = 11
Public Const MIC_LD_MDR As Long = 12
Public Const MIC_LD_IR As Long = 13
Public Const MIC_LD_AX As Long = 14         ' Ld_AX..Ld_DX = 14..17
Public Const MIC_LD_A As Long = 18
Public Const MIC_LD_FLG As Long = 19
Public Const MIC_LD_LCD As Long = 20
Public Const MIC_MEM_R As Long = 21
Public Const MIC_MEM_W As Long = 22
Public Const MIC_IO As Long = 23
Public Const MIC_INC_PC As Long = 24
Public Const MIC_SIGNALS As Long = 25
Public Const MIC_ENABLES As Long = 10       ' signals 0..9 are the tri-state output enables
Public Const MIC_MAX_STEPS As Long = 14
Public Const LCA_ERR_MICRO As Long = vbObjectError + 2211

Private mSig(0 To MIC_MAX_STEPS - 1, 0 To MIC_SIGNALS - 1) As Byte
Private mLabel(0 To MIC_MAX_STEPS - 1) As String
Private mData(0 To MIC_MAX_STEPS - 1) As String
Private mAlu(0 To MIC_MAX_STEPS - 1) As String
Private mCount As Long
Private mInstr As String
Private mPC As Long
Private mSerial As Long

' --- Names ----------------------------------------------------------------------------------------

Public Function MicroSignalName(ByVal sig As Long) As String
    If sig < 0 Or sig >= MIC_SIGNALS Then Err.Raise LCA_ERR_MICRO, "LcaMicro", "SIGNAL_INDEX"
    MicroSignalName = Choose(sig + 1, "En_PC", "En_MDR", "En_IR", "En_AX", "En_BX", "En_CX", "En_DX", "En_ALU", "En_MEM", "En_KBD", _
                             "Ld_PC", "Ld_MAR", "Ld_MDR", "Ld_IR", "Ld_AX", "Ld_BX", "Ld_CX", "Ld_DX", "Ld_A", "Ld_FLG", "Ld_LCD", _
                             "Mem_R", "Mem_W", "IO", "Inc_PC")
End Function

' Device owning the tri-state buffer of enable signal 0..9.
Public Function MicroDeviceName(ByVal device As Long) As String
    If device < 0 Or device >= MIC_ENABLES Then Err.Raise LCA_ERR_MICRO, "LcaMicro", "DEVICE_INDEX"
    MicroDeviceName = Choose(device + 1, "PC", "MDR", "IR", "AX", "BX", "CX", "DX", "ALU", "MEM", "KBD")
End Function

' --- Recording --------------------------------------------------------------------------------------

Public Sub MicroClear()
    Dim i As Long, j As Long
    For i = 0 To MIC_MAX_STEPS - 1
        For j = 0 To MIC_SIGNALS - 1
            mSig(i, j) = 0
        Next j
        mLabel(i) = vbNullString
        mData(i) = vbNullString
        mAlu(i) = vbNullString
    Next i
    mCount = 0
    mInstr = vbNullString
    mPC = 0
    mSerial = mSerial + 1
End Sub

' Builds the micro-sequence of the instruction that the executor has just completed (or is waiting on).
' Registers already hold the result, so the bus data shown is the value that crossed the bus.
' taken: branch outcome. waiting: IN on an empty queue (the keyboard buffer stays in High-Z).
Public Sub MicroRecord(ByVal pc As Long, ByVal op As Long, ByVal r As Long, ByVal x As Long, ByVal h As Long, _
                       ByVal taken As Boolean, ByVal waiting As Boolean, _
                       ByVal a0 As Long, ByVal b0 As Long)
    Dim va As Long, rn As String, sn As String, advance As Boolean
    MicroClear
    mPC = pc
    mInstr = AsmDisassemble(op, r, x, h)
    va = 256 * h + x
    advance = True
    If r >= 0 And r <= 3 Then rn = RegN(r)
    If op = &H11 Or (op >= &H20 And op <= &H22) Then sn = RegN(x And 3)

    ' FETCH (the same four micro-steps for every instruction)
    AddStep "T1  PC -> MAR", "PC=" & H3(pc), "", MIC_EN_PC, MIC_LD_MAR
    AddStep "T2  MEM -> MDR (lectura: L1I / L2 / RAM)", "instr=" & H2(op) & " " & H2(r) & " " & H2(x) & " " & H2(h), "", _
            MIC_MEM_R, MIC_EN_MEM, MIC_LD_MDR
    AddStep "T3  MDR -> IR", "IR=" & H2(op), "", MIC_EN_MDR, MIC_LD_IR
    AddStep "T4  DECODIFICAR (la UC lee IR; el bus queda en High-Z)", "opcode " & H2(op), ""

    ' EXECUTE
    Select Case op
        Case &H0
            AddStep "T5  NOP: sin transferencia", "", ""
        Case &H1
            AddStep "T5  HLT: se detiene el reloj, PC se queda", "", ""
            advance = False
        Case &H10
            AddStep "T5  IR(inmediato) -> " & rn, "imm=" & H2(x), "", MIC_EN_IR, MIC_LD_AX + r
        Case &H11
            AddStep "T5  " & sn & " -> " & rn, sn & "=" & H2(b0), "", MIC_EN_AX + (x And 3), MIC_LD_AX + r
        Case &H12
            AddStep "T5  IR(direccion) -> MAR", "VA=" & H3(va), "", MIC_EN_IR, MIC_LD_MAR
            AddStep "T6  MEM -> MDR (lectura de datos: L1D / L2 / RAM)", "dato=" & H2(CpuReadRegister(r)), "", _
                    MIC_MEM_R, MIC_EN_MEM, MIC_LD_MDR
            AddStep "T7  MDR -> " & rn, rn & "=" & H2(CpuReadRegister(r)), "", MIC_EN_MDR, MIC_LD_AX + r
        Case &H13
            AddStep "T5  IR(direccion) -> MAR", "VA=" & H3(va), "", MIC_EN_IR, MIC_LD_MAR
            AddStep "T6  " & rn & " -> MDR", rn & "=" & H2(CpuReadRegister(r)), "", MIC_EN_AX + r, MIC_LD_MDR
            AddStep "T7  MDR -> MEM (escritura write-through)", "dato=" & H2(CpuReadRegister(r)), "", _
                    MIC_MEM_W, MIC_EN_MDR
        Case &H20, &H21
            AddStep "T5  " & rn & " -> A (entrada A de la ALU)", rn & "=" & H2(a0), "", MIC_EN_AX + r, MIC_LD_A
            AddStep "T6  ALU(A " & IIf(op = &H20, "+", "-") & " " & sn & ") -> " & rn & " + banderas", _
                    "res=" & H2(CpuReadRegister(r)), IIf(op = &H20, "ADD", "SUB"), MIC_EN_ALU, MIC_LD_AX + r, MIC_LD_FLG
        Case &H22
            AddStep "T5  " & rn & " -> A (entrada A de la ALU)", rn & "=" & H2(a0), "", MIC_EN_AX + r, MIC_LD_A
            AddStep "T6  ALU(A - " & sn & ") -> solo banderas (CMP no guarda resultado; bus High-Z)", "", "CMP", MIC_LD_FLG
        Case &H23, &H24
            AddStep "T5  " & rn & " -> A (entrada A de la ALU)", rn & "=" & H2(a0), "", MIC_EN_AX + r, MIC_LD_A
            AddStep "T6  ALU(A " & IIf(op = &H23, "+", "-") & " 1) -> " & rn & " + banderas", "res=" & H2(CpuReadRegister(r)), _
                    IIf(op = &H23, "INC", "DEC"), MIC_EN_ALU, MIC_LD_AX + r, MIC_LD_FLG
        Case &H30 To &H34
            If taken Then
                AddStep "T5  " & CondText(op) & " cumplida: IR(destino) -> PC", "PC=" & H3(va), "", MIC_EN_IR, MIC_LD_PC
                advance = False
            Else
                AddStep "T5  " & CondText(op) & " no cumplida: no hay transferencia", "", ""
            End If
        Case &H40
            If waiting Then
                AddStep "T5  IN: cola vacia, KBD sigue en High-Z (espera; PC no avanza)", "", "", MIC_IO, MIC_MEM_R
                advance = False
            Else
                AddStep "T5  KBD -> " & rn & "  (IO=1, En_KBD)", rn & "=" & H2(CpuReadRegister(r)), "", _
                        MIC_IO, MIC_MEM_R, MIC_EN_KBD, MIC_LD_AX + r
            End If
        Case &H41
            AddStep "T5  " & rn & " -> LCD  (IO=1, Ld_LCD)", rn & "=" & H2(CpuReadRegister(r)), "", _
                    MIC_IO, MIC_MEM_W, MIC_EN_AX + r, MIC_LD_LCD
    End Select

    ' PC UPDATE
    If advance Then
        AddStep "T" & CStr(mCount + 1) & "  PC + 4 -> PC (incrementador; el bus queda en High-Z)", "PC=" & H3(pc + 4), "", _
                MIC_LD_PC, MIC_INC_PC
    End If
End Sub

Private Sub AddStep(ByVal label As String, ByVal data As String, ByVal aluOp As String, ParamArray signals() As Variant)
    Dim i As Long
    If mCount >= MIC_MAX_STEPS Then Err.Raise LCA_ERR_MICRO, "LcaMicro", "TOO_MANY_STEPS"
    mLabel(mCount) = label
    mData(mCount) = data
    mAlu(mCount) = aluOp
    For i = LBound(signals) To UBound(signals)
        mSig(mCount, CLng(signals(i))) = 1
    Next i
    mCount = mCount + 1
End Sub

Private Function CondText(ByVal op As Long) As String
    CondText = Choose(op - &H30 + 1, "JMP (incondicional)", "JZ: ZF=1", "JNZ: ZF=0", "JC: CF=1", "JS: SF=1")
End Function

' --- Inspection ---------------------------------------------------------------------------------------

Public Function MicroCount() As Long
    MicroCount = mCount
End Function

' Increments each time a new sequence is recorded or cleared; the panel uses it to notice a new instruction.
Public Function MicroSerial() As Long
    MicroSerial = mSerial
End Function

Public Function MicroInstruction() As String
    MicroInstruction = mInstr
End Function

Public Function MicroInstructionPC() As Long
    MicroInstructionPC = mPC
End Function

Public Function MicroLabel(ByVal i As Long) As String
    RequireStep i
    MicroLabel = mLabel(i)
End Function

Public Function MicroBusData(ByVal i As Long) As String
    RequireStep i
    MicroBusData = mData(i)
End Function

Public Function MicroAluOp(ByVal i As Long) As String
    RequireStep i
    MicroAluOp = mAlu(i)
End Function

Public Function MicroSignal(ByVal i As Long, ByVal sig As Long) As Boolean
    RequireStep i
    If sig < 0 Or sig >= MIC_SIGNALS Then Err.Raise LCA_ERR_MICRO, "LcaMicro", "SIGNAL_INDEX"
    MicroSignal = (mSig(i, sig) <> 0)
End Function

' Number of output buffers conducting in the step (the rule says 0 or 1).
Public Function MicroEnableCount(ByVal i As Long) As Long
    Dim s As Long
    RequireStep i
    For s = 0 To MIC_ENABLES - 1
        MicroEnableCount = MicroEnableCount + mSig(i, s)
    Next s
End Function

' Name of the only device driving the bus in the step, or "High-Z" when none does.
Public Function MicroDriver(ByVal i As Long) As String
    Dim s As Long
    RequireStep i
    MicroDriver = "High-Z"
    For s = 0 To MIC_ENABLES - 1
        If mSig(i, s) <> 0 Then MicroDriver = MicroDeviceName(s)
    Next s
End Function

' State of device d's tri-state buffer in step i: "1" (conducting) or "Z" (High-Z).
Public Function MicroBufferState(ByVal i As Long, ByVal device As Long) As String
    RequireStep i
    MicroBufferState = IIf(MicroSignal(i, device), "1", "Z")
End Function

' Comma-free list of the receivers (Ld_*) active in the step, "-" when none.
Public Function MicroReceivers(ByVal i As Long) As String
    Dim s As Long
    RequireStep i
    For s = MIC_LD_PC To MIC_LD_LCD
        If mSig(i, s) <> 0 Then MicroReceivers = MicroReceivers & IIf(Len(MicroReceivers) > 0, " ", vbNullString) & Mid$(MicroSignalName(s), 4)
    Next s
    If Len(MicroReceivers) = 0 Then MicroReceivers = "-"
End Function

' 25-bit vector in the order En_PC..Inc_PC, grouped: ENABLE | LOAD | CONTROL.
Public Function MicroWord(ByVal i As Long) As String
    Dim s As Long
    RequireStep i
    For s = 0 To MIC_SIGNALS - 1
        If s = MIC_LD_PC Or s = MIC_MEM_R Then MicroWord = MicroWord & " "
        MicroWord = MicroWord & CStr(mSig(i, s))
    Next s
End Function

' One line per step: "T2  MEM -> MDR  [En_MEM Ld_MDR Mem_R]".
Public Function MicroSignalsOn(ByVal i As Long) As String
    Dim s As Long
    RequireStep i
    For s = 0 To MIC_SIGNALS - 1
        If mSig(i, s) <> 0 Then MicroSignalsOn = MicroSignalsOn & IIf(Len(MicroSignalsOn) > 0, " ", vbNullString) & MicroSignalName(s)
    Next s
    If Len(MicroSignalsOn) = 0 Then MicroSignalsOn = "(ninguna / none)"
End Function

' Verifies the physical rules on the whole recorded sequence. Returns "" when they hold, else the first violation.
'   - at most one En_* per micro-step;
'   - Mem_R and Mem_W never together;
'   - Ld_X needs a source: the bus driven by one En_*, or the PC incrementer (Inc_PC) for Ld_PC, or the ALU latch for Ld_FLG;
'   - Mem_R implies En_MEM, or En_KBD when IO=1 (the reading device is the only driver).
Public Function MicroCheckRules() As String
    Dim i As Long, s As Long, loads As Long, busLoad As Boolean
    For i = 0 To mCount - 1
        If MicroEnableCount(i) > 1 Then
            MicroCheckRules = "T" & CStr(i + 1) & ": mas de un buffer conduce (contencion)"
            Exit Function
        End If
        If mSig(i, MIC_MEM_R) <> 0 And mSig(i, MIC_MEM_W) <> 0 Then
            MicroCheckRules = "T" & CStr(i + 1) & ": Mem_R y Mem_W a la vez"
            Exit Function
        End If
        busLoad = False
        For s = MIC_LD_PC To MIC_LD_LCD
            If mSig(i, s) <> 0 And s <> MIC_LD_FLG Then busLoad = True
        Next s
        If busLoad And MicroEnableCount(i) = 0 And mSig(i, MIC_INC_PC) = 0 Then
            MicroCheckRules = "T" & CStr(i + 1) & ": un receptor carga pero nadie maneja el bus"
            Exit Function
        End If
        If mSig(i, MIC_INC_PC) <> 0 And MicroEnableCount(i) <> 0 Then
            MicroCheckRules = "T" & CStr(i + 1) & ": Inc_PC con un emisor en el bus"
            Exit Function
        End If
    Next i
End Function

' --- Helpers ------------------------------------------------------------------------------------------

Private Sub RequireStep(ByVal i As Long)
    If i < 0 Or i >= mCount Then Err.Raise LCA_ERR_MICRO, "LcaMicro", "STEP_INDEX"
End Sub

Private Function RegN(ByVal code As Long) As String
    RegN = Choose((code And 3) + 1, "AX", "BX", "CX", "DX")
End Function

Private Function H2(ByVal v As Long) As String
    H2 = Right$("00" & Hex$(v And 255), 2) & "h"
End Function

Private Function H3(ByVal v As Long) As String
    H3 = Right$("000" & Hex$(v And 511), 3) & "h"
End Function
