Attribute VB_Name = "LcaKeyboard"
Option Explicit

' P13: keyboard controller. Eight-byte circular FIFO stored in the reserved RAM bytes E0..E7h (PFN 7).
' head/tail/count (0..8) are controller state; the queued bytes live only in RAM (RamReadIO/RamWriteIO).
'   Push: full queue -> KBD_FULL, the byte is rejected and nothing changes.
'   IN R: empty queue -> IN_EMPTY wait (no bus cycle, head/count/register unchanged; P12 keeps PC).
'         with data -> one external cycle IO=1 R=1 En_KBD=1 reads E0h+head; only after the register
'         is written do head advance and count decrement, so each byte is consumed exactly once.
' No worksheet dependency.
Public Const LCA_KBD_BASE_PA As Long = &HE0
Public Const LCA_KBD_BASE_VA As Long = &H1E0
Public Const LCA_KBD_SIZE As Long = 8
Public Const LCA_ERR_KEYBOARD As Long = vbObjectError + 2131
Private Const TRACE_MAX As Long = 64

Private mReady As Boolean
Private mHead As Long
Private mTail As Long
Private mCount As Long
Private mAccepted As Long
Private mRejected As Long
Private mConsumed As Long
Private mWaits As Long
Private mLastStatus As String
Private mTrace(0 To TRACE_MAX - 1) As String
Private mTraceCount As Long

' Empties the queue and clears E0..E7h. The LCD bytes E8..EFh belong to P14 and are not touched.
Public Sub KeyboardReset()
    Dim i As Long
    For i = 0 To LCA_KBD_SIZE - 1
        RamWriteIO LCA_KBD_BASE_PA + i, 0
    Next i
    mHead = 0
    mTail = 0
    mCount = 0
    mAccepted = 0
    mRejected = 0
    mConsumed = 0
    mWaits = 0
    mTraceCount = 0
    For i = 0 To TRACE_MAX - 1
        mTrace(i) = vbNullString
    Next i
    mReady = True
    mLastStatus = "RESET"
    AddTrace "RESET cola vacia / empty queue (head=0 tail=0 count=0)"
End Sub

' A key press. Returns False (KBD_FULL) when the eight slots are in use; the queue is unchanged.
Public Function KeyboardPush(ByVal value As Long) As Boolean
    Dim pa As Long
    RequireKeyboard
    CpuRequireByte value
    If mCount = LCA_KBD_SIZE Then
        mRejected = mRejected + 1
        mLastStatus = "KBD_FULL"
        AddTrace "KBD_FULL rechaza / rejects " & ByteText(value) & " (count=8)"
        KeyboardPush = False
        Exit Function
    End If
    pa = LCA_KBD_BASE_PA + mTail
    RamWriteIO pa, value
    mTail = (mTail + 1) Mod LCA_KBD_SIZE
    mCount = mCount + 1
    mAccepted = mAccepted + 1
    mLastStatus = "KEY"
    AddTrace "KEY " & ByteText(value) & " -> " & PaText(pa) & " (tail=" & CStr(mTail) & _
             " count=" & CStr(mCount) & ")"
    KeyboardPush = True
End Function

' Types every character of text (ASCII 0..255). Stops counting at the first KBD_FULL; returns accepted keys.
Public Function KeyboardPushText(ByVal text As String) As Long
    Dim i As Long, code As Long
    For i = 1 To Len(text)
        code = AscW(Mid$(text, i, 1))
        If code < 0 Or code > 255 Then Err.Raise LCA_ERR_KEYBOARD, "LcaKeyboard", "KBD_NOT_ASCII"
    Next i
    For i = 1 To Len(text)
        If KeyboardPush(AscW(Mid$(text, i, 1))) Then KeyboardPushText = KeyboardPushText + 1
    Next i
End Function

' IN R. True when a byte was loaded into R; False on IN_EMPTY (a wait, not an error).
Public Function KeyboardIn(ByVal targetRegister As Long) As Boolean
    Dim pa As Long, value As Byte
    RequireKeyboard
    Call CpuReadRegister(targetRegister)    ' INVALID_REGISTER even when the queue is empty
    If mCount = 0 Then
        mWaits = mWaits + 1
        mLastStatus = "IN_EMPTY"
        AddTrace "IN " & RegisterText(targetRegister) & " -> IN_EMPTY espera / wait (sin ciclo de bus / no bus cycle)"
        KeyboardIn = False
        Exit Function
    End If
    pa = LCA_KBD_BASE_PA + mHead
    value = ControlKeyboardRead(pa, targetRegister, RamReadIO(pa))
    CpuWriteRegister targetRegister, value
    ' Commit point: the register holds the byte, so the slot is consumed exactly once.
    mHead = (mHead + 1) Mod LCA_KBD_SIZE
    mCount = mCount - 1
    mConsumed = mConsumed + 1
    mLastStatus = "IN"
    AddTrace "IN " & RegisterText(targetRegister) & " <- " & ByteText(value) & " de / from " & PaText(pa) & _
             " En_KBD=1 (head=" & CStr(mHead) & " count=" & CStr(mCount) & ")"
    KeyboardIn = True
End Function

' --- Inspection (no state change) ---------------------------------------------------------

Public Function KeyboardReady() As Boolean
    KeyboardReady = mReady
End Function

Public Function KeyboardCount() As Long
    RequireKeyboard
    KeyboardCount = mCount
End Function

Public Function KeyboardHead() As Long
    RequireKeyboard
    KeyboardHead = mHead
End Function

Public Function KeyboardTail() As Long
    RequireKeyboard
    KeyboardTail = mTail
End Function

Public Function KeyboardIsEmpty() As Boolean
    KeyboardIsEmpty = (KeyboardCount() = 0)
End Function

Public Function KeyboardIsFull() As Boolean
    KeyboardIsFull = (KeyboardCount() = LCA_KBD_SIZE)
End Function

' Raw slot 0..7 (E0h + slot), queued or not.
Public Function KeyboardSlot(ByVal slot As Long) As Byte
    RequireKeyboard
    If slot < 0 Or slot >= LCA_KBD_SIZE Then Err.Raise LCA_ERR_KEYBOARD, "LcaKeyboard", "KBD_SLOT_RANGE"
    KeyboardSlot = RamReadIO(LCA_KBD_BASE_PA + slot)
End Function

' True when the slot holds a byte waiting to be read.
Public Function KeyboardSlotQueued(ByVal slot As Long) As Boolean
    RequireKeyboard
    If slot < 0 Or slot >= LCA_KBD_SIZE Then Err.Raise LCA_ERR_KEYBOARD, "LcaKeyboard", "KBD_SLOT_RANGE"
    KeyboardSlotQueued = ((slot - mHead + LCA_KBD_SIZE) Mod LCA_KBD_SIZE) < mCount
End Function

' Next byte IN would read, without consuming it.
Public Function KeyboardPeek() As Byte
    RequireKeyboard
    If mCount = 0 Then Err.Raise LCA_ERR_KEYBOARD, "LcaKeyboard", "IN_EMPTY"
    KeyboardPeek = RamReadIO(LCA_KBD_BASE_PA + mHead)
End Function

Public Function KeyboardAccepted() As Long
    KeyboardAccepted = mAccepted
End Function

Public Function KeyboardRejected() As Long
    KeyboardRejected = mRejected
End Function

Public Function KeyboardConsumed() As Long
    KeyboardConsumed = mConsumed
End Function

Public Function KeyboardWaits() As Long
    KeyboardWaits = mWaits
End Function

' "RESET", "KEY", "KBD_FULL", "IN" or "IN_EMPTY".
Public Function KeyboardLastStatus() As String
    KeyboardLastStatus = mLastStatus
End Function

Public Function KeyboardTraceCount() As Long
    KeyboardTraceCount = mTraceCount
    If KeyboardTraceCount > TRACE_MAX Then KeyboardTraceCount = TRACE_MAX
End Function

' index 0 = newest retained line.
Public Function KeyboardTraceLine(ByVal index As Long) As String
    If index < 0 Or index >= KeyboardTraceCount() Then _
        Err.Raise LCA_ERR_KEYBOARD, "LcaKeyboard", "KBD_TRACE_INDEX"
    KeyboardTraceLine = mTrace((mTraceCount - 1 - index) Mod TRACE_MAX)
End Function

' --- Internals -----------------------------------------------------------------------------

Private Sub RequireKeyboard()
    If Not mReady Then Err.Raise LCA_ERR_KEYBOARD, "LcaKeyboard", "KBD_RESET_REQUIRED"
End Sub

Private Sub AddTrace(ByVal text As String)
    mTrace(mTraceCount Mod TRACE_MAX) = text
    mTraceCount = mTraceCount + 1
End Sub

Private Function ByteText(ByVal value As Long) As String
    ByteText = Right$("00" & Hex$(value), 2) & "h"
    If value >= 32 And value <= 126 Then ByteText = ByteText & " '" & Chr$(value) & "'"
End Function

Private Function PaText(ByVal pa As Long) As String
    PaText = Right$("00" & Hex$(pa), 2) & "h"
End Function

Private Function RegisterText(ByVal code As Long) As String
    Select Case code
        Case LcaAX: RegisterText = "AX"
        Case LcaBX: RegisterText = "BX"
        Case LcaCX: RegisterText = "CX"
        Case LcaDX: RegisterText = "DX"
        Case Else: RegisterText = "R?"
    End Select
End Function
