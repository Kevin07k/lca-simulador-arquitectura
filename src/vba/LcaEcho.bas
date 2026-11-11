Attribute VB_Name = "LcaEcho"
Option Explicit

' P16: keyboard-to-LCD echo program written in the P11 assembly and run on the real CPU (P12).
'   IN AX takes the oldest key of the P13 queue (E0..E7h); OUT AX writes it to the LCD cell under the
'   cursor (P14, E8..EFh). The '.' key (2Eh) ends the program: it is not shown, the number of characters
'   shown is stored in [100h] and the CPU halts.
' Defined behaviour:
'   - empty queue: IN is a wait (WAIT_IN), PC stays on the IN, nothing is written; after more keys are
'     typed, running again repeats that same IN and the echo continues where it stopped;
'   - the eighth character fills the last cell and the cursor wraps to 0; the ninth replaces the first;
'   - a ninth key typed while the eight slots are full is rejected by the keyboard (KBD_FULL);
'   - keys come out in the order they were typed (FIFO), each one exactly once.
' This module also records the order of the characters (which queue slot each key left, which LCD cell
' it reached). No worksheet dependency.
Public Const ECHO_END_KEY As Long = &H2E        ' "."
Public Const ECHO_LOOP_HEAD As Long = &H8       ' address of IN AX
Public Const ECHO_HALT_PC As Long = &H24        ' address of HLT
Public Const ECHO_LOG_MAX As Long = 64
Public Const LCA_ERR_ECHO As Long = vbObjectError + 2171

Private mLoadId As Long
Private mLogN As Long
Private mLogKey(0 To ECHO_LOG_MAX - 1) As Long
Private mLogSlot(0 To ECHO_LOG_MAX - 1) As Long
Private mLogCell(0 To ECHO_LOG_MAX - 1) As Long    ' -1 = not shown yet, -2 = end key (not shown)
Private mLastSteps As Long
Private mTyped As Long

Public Function EchoSource() As Variant
    EchoSource = Array( _
        "; ECO: cada tecla del teclado aparece en el LCD. La tecla '.' termina.", _
        "        MOV BX, #2Eh      ; BX = '.' (tecla de fin)", _
        "        MOV CX, #00h      ; CX cuenta los caracteres mostrados", _
        "eco:    IN AX             ; toma la tecla mas antigua; cola vacia = espera (WAIT_IN)", _
        "        CMP AX, BX        ; es el punto?", _
        "        JZ fin            ; si: termina sin mostrarlo", _
        "        OUT AX            ; no: la escribe en la celda del cursor del LCD", _
        "        INC CX            ; un caracter mas", _
        "        JMP eco", _
        "fin:    ST [100h], CX     ; guarda cuantos caracteres se mostraron", _
        "        HLT")
End Function

' Source text exactly as stored in programs/echo/eco.asm (LF line ends).
Public Function EchoAsmText() As String
    EchoAsmText = Join(EchoSource(), vbLf)
End Function

Public Function EchoFile() As String
    EchoFile = "eco.asm"
End Function

' Assembles and loads the echo program; the whole machine is reset (queue empty, LCD blank).
Public Function EchoLoad() As Boolean
    EchoLoad = LoaderLoadSource(EchoSource())
    If EchoLoad Then
        mLoadId = LoaderLoadCount()
        ClearLog
    End If
End Function

' True while the program loaded in the machine is this one (another load replaces it).
Public Function EchoLoaded() As Boolean
    EchoLoaded = (mLoadId > 0 And mLoadId = LoaderLoadCount() And LoaderLoaded())
End Function

' Restarts the loaded echo program (same as RESET of the machine) and clears the order log.
Public Sub EchoReset()
    If Not EchoLoaded() Then Err.Raise LCA_ERR_ECHO, "LcaEcho", "ECHO_NOT_LOADED"
    ExecReset
    ClearLog
End Sub

' Types text on the keyboard. Returns the keys accepted; the rest were rejected with KBD_FULL.
Public Function EchoType(ByVal text As String) As Long
    EchoType = KeyboardPushText(text)
    mTyped = mTyped + Len(text)
End Function

' One instruction of the machine, recording the order of the characters. True when it completed.
Public Function EchoStep() As Boolean
    Dim slot As Long, kind As String, cell As Long
    If Not EchoLoaded() Then Err.Raise LCA_ERR_ECHO, "LcaEcho", "ECHO_NOT_LOADED"
    SyncLog
    slot = KeyboardHead()
    cell = LcdCursor()
    EchoStep = ExecStep()
    If Not EchoStep Then Exit Function
    kind = ExecLastKind()
    If kind = "IN" Then
        If mLogN < ECHO_LOG_MAX Then
            mLogKey(mLogN) = CpuReadRegister(0)
            mLogSlot(mLogN) = slot
            mLogCell(mLogN) = IIf(mLogKey(mLogN) = ECHO_END_KEY, -2, -1)
            mLogN = mLogN + 1
        End If
    ElseIf kind = "OUT" Then
        If mLogN > 0 Then
            If mLogCell(mLogN - 1) = -1 Then mLogCell(mLogN - 1) = cell
        End If
    End If
    mLastSteps = ExecSteps()
End Function

' Runs until HLT, an IN wait, an error or maxSteps. Returns the number of completed instructions.
Public Function EchoRun(Optional ByVal maxSteps As Long = 500) As Long
    Dim n As Long
    Do While n < maxSteps
        If Not EchoStep() Then Exit Do
        n = n + 1
        If ExecStatus() = EXEC_HALTED Then Exit Do
    Loop
    EchoRun = n
End Function

' --- Inspection ------------------------------------------------------------------------------

' Characters written to the LCD by the program (CX).
Public Function EchoShown() As Long
    EchoShown = CpuReadRegister(2)
End Function

Public Function EchoTyped() As Long
    EchoTyped = mTyped
End Function

Public Function EchoOrderCount() As Long
    SyncLog
    EchoOrderCount = mLogN
End Function

Public Function EchoOrderKey(ByVal i As Long) As Long
    CheckIndex i
    EchoOrderKey = mLogKey(i)
End Function

' Queue slot (0..7, PA E0h + slot) the key left.
Public Function EchoOrderSlot(ByVal i As Long) As Long
    CheckIndex i
    EchoOrderSlot = mLogSlot(i)
End Function

' LCD cell (0..7, PA E8h + cell) the key reached; -1 = read but not shown yet; -2 = end key.
Public Function EchoOrderCell(ByVal i As Long) As Long
    CheckIndex i
    EchoOrderCell = mLogCell(i)
End Function

' Keys in the order the CPU read them, e.g. "HOLA." (non-printable bytes shown as "?").
Public Function EchoOrderText() As String
    Dim i As Long
    SyncLog
    For i = 0 To mLogN - 1
        EchoOrderText = EchoOrderText & LcdDisplayChar(mLogKey(i))
    Next i
End Function

' "slot>cell;" for every key, e.g. "0>0;1>1;". End key: "slot>fin;".
Public Function EchoOrderRoute() As String
    Dim i As Long
    SyncLog
    For i = 0 To mLogN - 1
        EchoOrderRoute = EchoOrderRoute & CStr(mLogSlot(i)) & ">" & _
                         IIf(mLogCell(i) = -2, "fin", IIf(mLogCell(i) = -1, "?", CStr(mLogCell(i)))) & ";"
    Next i
End Function

' Short state for the panel and the documentation.
Public Function EchoStateText() As String
    If Not EchoLoaded() Then
        EchoStateText = "SIN CARGAR"
        Exit Function
    End If
    Select Case ExecStatus()
        Case EXEC_WAIT: EchoStateText = "ESPERANDO TECLA (WAIT_IN)"
        Case EXEC_HALTED: EchoStateText = "TERMINADO (HLT)"
        Case EXEC_ERROR: EchoStateText = "ERROR: " & ExecErrorText()
        Case Else
            If ExecSteps() = 0 Then EchoStateText = "LISTO" Else EchoStateText = "EN CURSO"
    End Select
End Function

' --- Internals --------------------------------------------------------------------------------

Private Sub ClearLog()
    mLogN = 0
    mLastSteps = 0
    mTyped = 0
End Sub

' A machine RESET (P12 button or ExecReset) restarts the program: the old order no longer applies.
Private Sub SyncLog()
    If ExecSteps() < mLastSteps Then ClearLog
End Sub

Private Sub CheckIndex(ByVal i As Long)
    SyncLog
    If i < 0 Or i >= mLogN Then Err.Raise LCA_ERR_ECHO, "LcaEcho", "ECHO_INDEX"
End Sub
