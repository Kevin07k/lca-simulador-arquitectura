Attribute VB_Name = "LcaControl"
Option Explicit

' P05 microsequencer. P06 supplies read bytes and commits RAM writes after a successful bus edge.
Public Const LCA_ERR_SEQUENCE As Long = vbObjectError + 2053
Private mFetchStage As Long     ' 0 = none, 1 = read external, 2 = latch internal
Private mFetchIndex As Long
Private mFetchVA As Long
Private mFetchPA As Long
Private mFetched(0 To 3) As Byte
Private mStoreStage As Long     ' 0 = none, 1 = ready for external commit
Private mStoreVA As Long
Private mStorePA As Long
Private mLastBus As Long

Public Sub ControlReset()
    Dim i As Long
    BusReset
    mFetchStage = 0
    mFetchIndex = 0
    mStoreStage = 0
    mLastBus = LCA_BUS_EXTERNAL
    For i = 0 To 3
        mFetched(i) = 0
    Next i
End Sub

Public Sub ControlStartFetch(ByVal va As Long, ByVal pa As Long)
    Dim i As Long
    If mFetchStage <> 0 Or mStoreStage <> 0 Then Err.Raise LCA_ERR_SEQUENCE, "LcaControl", "SEQUENCE_BUSY"
    If va < 0 Or va > 252 Or (va Mod 4) <> 0 Then Err.Raise LCA_ERR_SEQUENCE, "LcaControl", "INVALID_FETCH_VA"
    If pa < 0 Or pa > 220 Or (pa Mod 4) <> 0 Then Err.Raise LCA_ERR_SEQUENCE, "LcaControl", "INVALID_FETCH_PA"
    ' The MMU/CPU validates PC, protection and PTE before calling this transport operation.
    mFetchVA = va
    mFetchPA = pa
    mFetchIndex = 0
    For i = 0 To 3
        mFetched(i) = 0
    Next i
    mFetchStage = 1
End Sub

' For each byte: first external RAM -> cache/MDR, then internal MDR -> IR/buffer.
' A real cache hit can call ControlLatchFetchByte after ControlReadCacheByte instead.
Public Sub ControlReadFetchByte(ByVal memoryByte As Long)
    If mFetchStage <> 1 Then Err.Raise LCA_ERR_SEQUENCE, "LcaControl", "FETCH_READ_ORDER"
    CpuRequireByte memoryByte
    BusFloat LCA_BUS_INTERNAL
    BusBegin LCA_BUS_EXTERNAL, "FETCH[" & CStr(mFetchIndex) & "] READ", _
             mFetchVA + mFetchIndex, mFetchPA + mFetchIndex, True, False, False
    BusDrive LCA_BUS_EXTERNAL, "RAM", memoryByte
    BusReceive LCA_BUS_EXTERNAL, "CACHE/MDR"
    CpuWriteControl LcaMDR, BusClock(LCA_BUS_EXTERNAL)
    CpuWriteControl LcaMAR_V, mFetchVA + mFetchIndex
    CpuWriteControl LcaMAR_P, mFetchPA + mFetchIndex
    mLastBus = LCA_BUS_EXTERNAL
    mFetchStage = 2
End Sub

' A cache hit has no external RAM cycle; that bus remains High-Z.
Public Sub ControlReadCacheByte(ByVal cachedByte As Long)
    If mFetchStage <> 1 Then Err.Raise LCA_ERR_SEQUENCE, "LcaControl", "FETCH_READ_ORDER"
    CpuRequireByte cachedByte
    BusFloat LCA_BUS_EXTERNAL
    BusFloat LCA_BUS_INTERNAL
    BusBegin LCA_BUS_INTERNAL, "FETCH[" & CStr(mFetchIndex) & "] CACHE", _
             mFetchVA + mFetchIndex, mFetchPA + mFetchIndex, False, False, False
    BusDrive LCA_BUS_INTERNAL, "L1I", cachedByte
    BusReceive LCA_BUS_INTERNAL, "MDR"
    CpuWriteControl LcaMDR, BusClock(LCA_BUS_INTERNAL)
    CpuWriteControl LcaMAR_V, mFetchVA + mFetchIndex
    CpuWriteControl LcaMAR_P, mFetchPA + mFetchIndex
    mLastBus = LCA_BUS_INTERNAL
    mFetchStage = 2
End Sub

Public Sub ControlLatchFetchByte()
    If mFetchStage <> 2 Then Err.Raise LCA_ERR_SEQUENCE, "LcaControl", "FETCH_LATCH_ORDER"
    BusFloat LCA_BUS_EXTERNAL
    BusBegin LCA_BUS_INTERNAL, "FETCH[" & CStr(mFetchIndex) & "] LATCH", _
             mFetchVA + mFetchIndex, mFetchPA + mFetchIndex, False, False, False
    BusDrive LCA_BUS_INTERNAL, "MDR", CpuSnapshotMDR()
    If mFetchIndex = 0 Then
        BusReceive LCA_BUS_INTERNAL, "IR"
    Else
        BusReceive LCA_BUS_INTERNAL, "FETCH_BUFFER[" & CStr(mFetchIndex) & "]"
    End If
    mFetched(mFetchIndex) = CByte(BusClock(LCA_BUS_INTERNAL))
    If mFetchIndex = 0 Then CpuWriteControl LcaIR, CLng(mFetched(0))
    mLastBus = LCA_BUS_INTERNAL
    mFetchIndex = mFetchIndex + 1
    If mFetchIndex = 4 Then
        mFetchStage = 0
    Else
        mFetchStage = 1
    End If
End Sub

Public Function ControlFetched(ByVal index As Long) As Byte
    If index < 0 Or index > 3 Then Err.Raise LCA_ERR_SEQUENCE, "LcaControl", "FETCH_INDEX_RANGE"
    ControlFetched = mFetched(index)
End Function

' Only the data path of ST is implemented here. P12 owns instruction decode/PC commit.
Public Sub ControlPrepareStore(ByVal va As Long, ByVal pa As Long, ByVal sourceRegister As Long)
    Dim value As Long
    If mFetchStage <> 0 Or mStoreStage <> 0 Then Err.Raise LCA_ERR_SEQUENCE, "LcaControl", "SEQUENCE_BUSY"
    If va < 256 Or va > 479 Then Err.Raise LCA_ERR_SEQUENCE, "LcaControl", "INVALID_STORE_VA"
    If pa < 0 Or pa > 223 Then Err.Raise LCA_ERR_SEQUENCE, "LcaControl", "INVALID_STORE_PA"
    value = CpuReadRegister(sourceRegister)
    BusFloat LCA_BUS_EXTERNAL
    BusBegin LCA_BUS_INTERNAL, "ST PREPARE", va, pa, False, False, False
    BusDrive LCA_BUS_INTERNAL, RegisterLabel(sourceRegister), value
    BusReceive LCA_BUS_INTERNAL, "MDR"
    CpuWriteControl LcaMDR, BusClock(LCA_BUS_INTERNAL)
    CpuWriteControl LcaMAR_V, va
    CpuWriteControl LcaMAR_P, pa
    mLastBus = LCA_BUS_INTERNAL
    mStoreVA = va
    mStorePA = pa
    mStoreStage = 1
End Sub

' The caller must write this returned byte to its RAM/cache only after this returns.
Public Function ControlCommitStore() As Byte
    Dim value As Long
    If mStoreStage <> 1 Then Err.Raise LCA_ERR_SEQUENCE, "LcaControl", "STORE_COMMIT_ORDER"
    BusFloat LCA_BUS_INTERNAL
    BusBegin LCA_BUS_EXTERNAL, "ST WRITE", mStoreVA, mStorePA, False, True, False
    BusDrive LCA_BUS_EXTERNAL, "CPU/CACHE", CpuSnapshotMDR()
    BusReceive LCA_BUS_EXTERNAL, "RAM"
    value = BusClock(LCA_BUS_EXTERNAL)
    ControlCommitStore = CByte(value)
    mStoreStage = 0
    mLastBus = LCA_BUS_EXTERNAL
End Function

Public Function ControlLastWord() As String
    Dim cycle As LcaBusCycle
    cycle = BusSnapshot(mLastBus)
    ControlLastWord = cycle.Microstep & "  [" & cycle.BusName & "]  " & _
                      "R=" & Abs(CLng(cycle.ReadSignal)) & " W=" & Abs(CLng(cycle.WriteSignal)) & _
                      " C=" & Abs(CLng(cycle.ClockSignal)) & " IO=" & Abs(CLng(cycle.IOSignal)) & _
                      " EN=" & Abs(CLng(cycle.EnableSignal)) & " LD=" & Abs(CLng(cycle.LoadSignal)) & _
                      " KBD=" & Abs(CLng(cycle.EnKBD)) & " LCD=" & Abs(CLng(cycle.LdLCD))
End Function

Private Function CpuSnapshotMDR() As Long
    Dim cpu As LcaCpuState
    cpu = CpuSnapshot()
    CpuSnapshotMDR = cpu.MDR
End Function

Private Function RegisterLabel(ByVal code As Long) As String
    Select Case code
        Case LcaAX: RegisterLabel = "AX"
        Case LcaBX: RegisterLabel = "BX"
        Case LcaCX: RegisterLabel = "CX"
        Case LcaDX: RegisterLabel = "DX"
    End Select
End Function
