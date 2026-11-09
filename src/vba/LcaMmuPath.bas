Attribute VB_Name = "LcaMmuPath"
Option Explicit

' Synchronous P07 adapter. P08 will catch PAGE_FAULT, load a page and retry this same request.
Public Sub MmuFetchFour(ByVal va As Long)
    Dim pa As Long, i As Long
    pa = MmuTranslate(va, MMU_FETCH)
    ' Validate all four translations before starting bus cycles.
    For i = 1 To 3
        If MmuTranslate(va + i, MMU_FETCH) <> pa + i Then _
            Err.Raise LCA_ERR_MMU_STATE, "LcaMmuPath", "FETCH_NOT_CONTIGUOUS"
    Next i
    MemoryFetchStart va, pa
    For i = 0 To 3
        MemoryFetchReadNext
        MemoryFetchLatchNext
        MmuCommitAccess va + i, MMU_FETCH, pa + i
    Next i
End Sub

Public Sub MmuStoreByte(ByVal va As Long, ByVal sourceRegister As Long)
    Dim pa As Long
    pa = MmuTranslate(va, MMU_WRITE)
    MemoryStorePrepare va, pa, sourceRegister
    MemoryStoreCommit
    MmuCommitAccess va, MMU_WRITE, pa
End Sub
