Attribute VB_Name = "LcaMemoryPath"
Option Explicit

' P06 bridge: actual RAM drives P05 cycles. The CPU only receives data after BusClock.
Public Const LCA_ERR_MEMORY_PATH As Long = vbObjectError + 2063
Private mFetchPhase As Long       ' 0 = none, 1 = read, 2 = latch
Private mFetchPA As Long
Private mFetchIndex As Long
Private mStoreReady As Boolean
Private mStorePA As Long

Public Sub MemoryPathReset()
    ControlReset
    mFetchPhase = 0
    mFetchIndex = 0
    mStoreReady = False
End Sub

' VA->PA is supplied by the caller for this P06 example. P07 owns translation/PTE.
Public Sub MemoryFetchStart(ByVal va As Long, ByVal pa As Long)
    Dim i As Long
    If mFetchPhase <> 0 Or mStoreReady Then Err.Raise LCA_ERR_MEMORY_PATH, "LcaMemoryPath", "MEMORY_PATH_BUSY"
    ' Validate the whole four-byte physical span before starting any bus cycle.
    If pa < 0 Or pa > &HDC Then Err.Raise LCA_ERR_RAM_PROTECTION, "LcaMemoryPath", "FETCH_PHYSICAL_SPAN"
    For i = 0 To 3
        RamRequireUserAddress pa + i
    Next i
    ControlStartFetch va, pa
    mFetchPA = pa
    mFetchIndex = 0
    mFetchPhase = 1
End Sub

Public Sub MemoryFetchReadNext()
    Dim value As Long
    If mFetchPhase <> 1 Then Err.Raise LCA_ERR_MEMORY_PATH, "LcaMemoryPath", "FETCH_READ_ORDER"
    value = RamReadUser(mFetchPA + mFetchIndex)
    ControlReadFetchByte value
    mFetchPhase = 2
End Sub

Public Sub MemoryFetchLatchNext()
    If mFetchPhase <> 2 Then Err.Raise LCA_ERR_MEMORY_PATH, "LcaMemoryPath", "FETCH_LATCH_ORDER"
    ControlLatchFetchByte
    mFetchIndex = mFetchIndex + 1
    If mFetchIndex = 4 Then
        mFetchPhase = 0
    Else
        mFetchPhase = 1
    End If
End Sub

' Convenience for tests/other modules; the panel calls the two distinct microsteps instead.
Public Sub MemoryFetchFour(ByVal va As Long, ByVal pa As Long)
    Dim i As Long
    MemoryFetchStart va, pa
    For i = 0 To 3
        MemoryFetchReadNext
        MemoryFetchLatchNext
    Next i
End Sub

Public Sub MemoryStorePrepare(ByVal va As Long, ByVal pa As Long, ByVal sourceRegister As Long)
    If mFetchPhase <> 0 Or mStoreReady Then Err.Raise LCA_ERR_MEMORY_PATH, "LcaMemoryPath", "MEMORY_PATH_BUSY"
    RamRequireUserAddress pa
    ControlPrepareStore va, pa, sourceRegister
    mStorePA = pa
    mStoreReady = True
End Sub

Public Sub MemoryStoreCommit()
    Dim transferred As Long
    If Not mStoreReady Then Err.Raise LCA_ERR_MEMORY_PATH, "LcaMemoryPath", "STORE_COMMIT_ORDER"
    transferred = ControlCommitStore()   ' detects any bus fault before the RAM mutation
    RamWriteUser mStorePA, transferred
    mStoreReady = False
End Sub
