Attribute VB_Name = "LcaKeyboardTests"
Option Explicit

Private mChecks As Long

' Desktop Excel verification for P13: FIFO in E0..E7h, KBD_FULL, IN_EMPTY wait/resume, En_KBD cycle.
Public Sub VerificarTecladoP13()
    Dim failure As Long
    Dim i As Long, traceBefore As Long
    Dim cycle As LcaBusCycle
    On Error GoTo failed
    mChecks = 0

    ' --- Section A: reset ---------------------------------------------------------------------
    CpuReset
    RamReset
    MemoryPathReset
    KeyboardReset
    Check KeyboardCount() = 0 And KeyboardHead() = 0 And KeyboardTail() = 0, "A: empty queue after reset"
    For i = 0 To 7
        Check RamReadIO(&HE0 + i) = 0, "A: slot E" & CStr(i) & "h cleared"
        Check RamReadIO(&HE8 + i) = &H20, "A: LCD byte E" & Hex$(8 + i) & "h untouched (space)"
    Next i

    ' --- Section B: IN on an empty queue is a wait, not an error -----------------------------
    CpuWriteRegister LcaAX, &H77
    traceBefore = BusTraceCount()
    Check Not KeyboardIn(LcaAX), "B: IN on empty queue returns False"
    Check KeyboardLastStatus() = "IN_EMPTY" And KeyboardWaits() = 1, "B: IN_EMPTY recorded as a wait"
    Check CpuReadRegister(LcaAX) = &H77, "B: AX unchanged while waiting"
    Check KeyboardHead() = 0 And KeyboardCount() = 0, "B: head/count unchanged while waiting"
    Check BusTraceCount() = traceBefore, "B: no bus cycle while waiting"

    ' --- Section C: P02 example, key 'A' then IN AX ---------------------------------------------
    Check KeyboardPush(&H41), "C: key A accepted"
    Check RamReadIO(&HE0) = &H41 And KeyboardCount() = 1 And KeyboardTail() = 1, "C: A stored at E0h"
    Check KeyboardPeek() = &H41 And KeyboardCount() = 1, "C: peek does not consume"
    Check KeyboardIn(LcaAX), "C: IN AX succeeds"
    Check CpuReadRegister(LcaAX) = &H41, "C: AX = 41h"
    Check KeyboardHead() = 1 And KeyboardCount() = 0, "C: head=1 count=0 after commit"
    Check BusTraceCount() = traceBefore + 1, "C: exactly one bus cycle for IN"
    cycle = BusSnapshot(LCA_BUS_EXTERNAL)
    Check cycle.IOSignal And cycle.ReadSignal And Not cycle.WriteSignal, "C: IO=1 R=1 W=0"
    Check cycle.EnKBD And Not cycle.LdLCD, "C: En_KBD=1 Ld_LCD=0"
    Check cycle.EnableSignal And cycle.LoadSignal And cycle.ClockSignal, "C: Enable, Load and Clock set"
    Check cycle.PA = &HE0 And cycle.VA = &H1E0, "C: PA E0h / VA 1E0h"
    Check cycle.Driver = "KBD" And cycle.Receiver = "AX" And cycle.Data = &H41, "C: KBD drives 41h into AX"
    Check InStr(1, ControlLastWord(), "KBD=1", vbBinaryCompare) > 0, "C: control word shows KBD=1"

    ' --- Section D: pause/resume consumes each byte exactly once -------------------------------
    Check Not KeyboardIn(LcaAX), "D: second IN waits"
    Check CpuReadRegister(LcaAX) = &H41 And KeyboardConsumed() = 1, "D: A not consumed twice"
    Check KeyboardPush(&H42), "D: key B arrives during the wait"
    Check RamReadIO(&HE1) = &H42, "D: B stored at E1h"
    Check KeyboardIn(LcaBX), "D: resumed IN BX"
    Check CpuReadRegister(LcaBX) = &H42 And CpuReadRegister(LcaAX) = &H41, "D: BX = 42h, AX kept"
    Check KeyboardConsumed() = 2 And KeyboardWaits() = 2 And KeyboardCount() = 0, "D: counters 2 consumed / 2 waits"

    ' --- Section E: full queue rejects the ninth key --------------------------------------------
    MemoryPathReset
    KeyboardReset
    For i = 1 To 8
        Check KeyboardPush(i), "E: key " & CStr(i) & " accepted"
    Next i
    Check KeyboardIsFull() And KeyboardTail() = 0, "E: eight keys fill the queue, tail wraps to 0"
    Check Not KeyboardPush(&H99), "E: ninth key rejected"
    Check KeyboardLastStatus() = "KBD_FULL" And KeyboardRejected() = 1, "E: KBD_FULL reported"
    Check KeyboardCount() = 8 And KeyboardHead() = 0 And KeyboardTail() = 0, "E: pointers unchanged by KBD_FULL"
    For i = 0 To 7
        Check RamReadIO(&HE0 + i) = i + 1, "E: slot " & CStr(i) & " unchanged by KBD_FULL"
    Next i

    ' --- Section F: FIFO order across the wrap-around ----------------------------------------
    For i = 1 To 3
        Check KeyboardIn(LcaCX) And CpuReadRegister(LcaCX) = i, "F: IN returns key " & CStr(i)
    Next i
    Check KeyboardHead() = 3 And KeyboardCount() = 5, "F: head=3 count=5"
    Check KeyboardPushText(Chr$(9) & Chr$(10) & Chr$(11)) = 3, "F: three keys after the wrap"
    Check RamReadIO(&HE0) = 9 And RamReadIO(&HE2) = 11 And KeyboardTail() = 3, "F: wrapped keys in E0..E2h"
    Check KeyboardSlotQueued(0) And KeyboardSlotQueued(7) And KeyboardIsFull(), "F: full again after wrap"
    For i = 4 To 11
        Check KeyboardIn(LcaDX) And CpuReadRegister(LcaDX) = i, "F: FIFO order key " & CStr(i)
    Next i
    Check KeyboardIsEmpty() And KeyboardHead() = 3, "F: empty, head followed tail"
    Check KeyboardConsumed() = 11 And KeyboardAccepted() = 11, "F: 11 accepted, 11 consumed"

    ' --- Section G: rejected inputs change nothing -------------------------------------------
    Call KeyboardPush(&H5A)
    failure = 0
    On Error Resume Next
    Call KeyboardIn(9)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_REGISTER And KeyboardCount() = 1 And KeyboardPeek() = &H5A, "G: invalid register keeps the key"
    failure = 0
    On Error Resume Next
    Call KeyboardPush(256)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_BYTE And KeyboardCount() = 1, "G: byte 256 rejected"
    failure = 0
    On Error Resume Next
    Call KeyboardPushText("X" & ChrW(&H3A9))
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_KEYBOARD And KeyboardCount() = 1, "G: non-ASCII text rejected before queueing"

    ' --- Section H: IN never interrupts a pending bus sequence ----------------------------------
    ControlStartFetch 0, 0
    traceBefore = BusTraceCount()
    failure = 0
    On Error Resume Next
    Call KeyboardIn(LcaAX)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_SEQUENCE, "H: IN during a fetch is SEQUENCE_BUSY"
    Check KeyboardCount() = 1 And KeyboardPeek() = &H5A And BusTraceCount() = traceBefore, _
        "H: queue and bus unchanged after SEQUENCE_BUSY"
    MemoryPathReset
    Check KeyboardIn(LcaAX) And CpuReadRegister(LcaAX) = &H5A, "H: the same key is delivered after the bus is free"

    ' --- Section I: the LCD buffer is never touched by the keyboard ------------------------------
    For i = 0 To 7
        Check RamReadIO(&HE8 + i) = &H20, "I: LCD byte " & CStr(i) & " still a space"
    Next i
    Check LCA_KBD_SIZE = 8 And LCA_KBD_BASE_PA = &HE0, "I: queue is the eight bytes at E0h"

    TecladoPanelActualizar
    LcaReport "P13 PASS: " & mChecks & " comprobaciones / assertions." & vbCrLf & _
           "Cola FIFO E0-E7h, KBD_FULL, IN_EMPTY con reanudacion y ciclo IN con En_KBD=1. P12 decide el avance de PC.", True
    Exit Sub
failed:
    LcaReport "P13 FAIL after " & mChecks & " checks: " & Err.Description, False
End Sub

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    mChecks = mChecks + 1
    If Not condition Then Err.Raise vbObjectError + 2132, "LcaKeyboardTests", message
End Sub
