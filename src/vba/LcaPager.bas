Attribute VB_Name = "LcaPager"
Option Explicit

' P08 pager. P09/P10 must invalidate all cache lines of the selected PFN
' before the RAM frame is reused (see docs/P08_PAGE_FAULT_SWAP.md).
Public Const LCA_ERR_PAGER_STATE As Long = vbObjectError + 2082

Private mReady As Boolean
Private mHand As Long
Private mFaultCount As Long
Private mEvictionCount As Long
Private mWritebackCount As Long
Private mLoadByteCount As Long
Private mWritebackByteCount As Long
Private mLastVA As Long
Private mLastVPN As Long
Private mLastVictim As Long
Private mTrace(0 To 127) As String
Private mTraceCount As Long

' Reset memory/image/page table together; callers reset CPU and bus separately.
Public Sub PagerReset()
    Dim i As Long
    RamReset
    MmuReset
    SwapReset
    mReady = True
    mHand = 0
    mFaultCount = 0
    mEvictionCount = 0
    mWritebackCount = 0
    mLoadByteCount = 0
    mWritebackByteCount = 0
    mLastVA = -1
    mLastVPN = -1
    mLastVictim = -1
    mTraceCount = 0
    For i = 0 To 127
        mTrace(i) = vbNullString
    Next i
End Sub

' Validates before touching RAM. Page-fault resolution is the only retry.
' Protection and bus failures remain fatal; a failed bus cycle is never retried.
Public Function PagerResolve(ByVal va As Long, ByVal accessType As Long) As Long
    Dim pa As Long, vpn As Long
    RequirePager
    If MmuTryTranslate(va, accessType, pa) Then
        PagerResolve = pa
        Exit Function
    End If
    vpn = MmuLastFaultVPN()
    PagerHandleFault va, vpn
    PagerResolve = MmuTranslate(va, accessType)
    AddTrace "RESUME VA=" & CStr(va) & " VPN=" & CStr(vpn) & " PA=" & CStr(PagerResolve)
End Function

' Four translations complete before MemoryFetchStart or any bus side effect.
' P02 fixed/aligned four-byte instructions cannot cross a 32-byte page.
Public Sub PagerFetchFour(ByVal va As Long)
    Dim pa As Long, i As Long, nextPA As Long
    RequirePager
    If va < 0 Or va > 252 Or (va Mod 4) <> 0 Then _
        Err.Raise LCA_ERR_PAGER_STATE, "LcaPager", "INVALID_FETCH_PC"
    pa = PagerResolve(va, MMU_FETCH)
    For i = 1 To 3
        nextPA = PagerResolve(va + i, MMU_FETCH)
        If nextPA <> pa + i Then _
            Err.Raise LCA_ERR_PAGER_STATE, "LcaPager", "FETCH_NOT_CONTIGUOUS"
    Next i
    MemoryFetchStart va, pa
    For i = 0 To 3
        MemoryFetchReadNext
        MemoryFetchLatchNext
        MmuCommitAccess va + i, MMU_FETCH, pa + i
    Next i
End Sub

Public Sub PagerStoreByte(ByVal va As Long, ByVal sourceRegister As Long)
    Dim pa As Long, value As Byte
    RequirePager
    value = CpuReadRegister(sourceRegister)  ' Reject a bad register before a page fault can mutate memory.
    pa = PagerResolve(va, MMU_WRITE)
    MemoryStorePrepare va, pa, sourceRegister
    MemoryStoreCommit
    MmuCommitAccess va, MMU_WRITE, pa
End Sub

' P09/P12 can replace this direct byte path with their cache/bus data-read path.
Public Function PagerReadByte(ByVal va As Long) As Byte
    Dim pa As Long
    RequirePager
    pa = PagerResolve(va, MMU_READ)
    PagerReadByte = RamReadUser(pa)
    MmuCommitAccess va, MMU_READ, pa
End Function

' Explicit handler for an absent, access-validated VPN; normally called by PagerResolve.
Private Sub PagerHandleFault(ByVal va As Long, ByVal vpn As Long)
    Dim pfn As Long, victim As Long, offset As Long, pa As Long
    Dim entry As LcaPageEntry
    RequirePager
    If vpn < 0 Or vpn > 14 Or MmuVpnOf(va) <> vpn Then _
        Err.Raise LCA_ERR_PAGER_STATE, "LcaPager", "INVALID_FAULT_VPN"
    entry = MmuPage(vpn)
    If entry.Present Then Err.Raise LCA_ERR_PAGER_STATE, "LcaPager", "PAGE_ALREADY_PRESENT"

    mLastVA = va
    mLastVPN = vpn
    mLastVictim = -1
    mFaultCount = mFaultCount + 1
    AddTrace "PAGE_FAULT VA=" & CStr(va) & " VPN=" & CStr(vpn)
    pfn = ChooseFrame()
    victim = FrameOwner(pfn)
    If victim >= 0 Then
        mLastVictim = victim
        entry = MmuPage(victim)
        ' P09/P10 cache invalidation must occur here, before reading dirty RAM
        ' and before unmapping or overwriting the old frame.
        If entry.Dirty Then
            For offset = 0 To 31
                pa = pfn * 32 + offset
                SwapWriteByte victim, offset, RamReadUser(pa)
                mWritebackByteCount = mWritebackByteCount + 1
            Next offset
            mWritebackCount = mWritebackCount + 1
            MmuClearDirty victim
            AddTrace "WRITEBACK VPN=" & CStr(victim) & " PFN=" & CStr(pfn) & " bytes=32"
        End If
        MmuUnmap victim
        mEvictionCount = mEvictionCount + 1
        AddTrace "EVICT VPN=" & CStr(victim) & " PFN=" & CStr(pfn)
    End If
    For offset = 0 To 31
        RamWriteUser pfn * 32 + offset, SwapReadByte(vpn, offset)
        mLoadByteCount = mLoadByteCount + 1
    Next offset
    MmuMap vpn, pfn
    AddTrace "LOAD VPN=" & CStr(vpn) & " PFN=" & CStr(pfn) & " bytes=32"
End Sub

Private Function ChooseFrame() As Long
    Dim pfn As Long, owner As Long, entry As LcaPageEntry
    For pfn = 0 To 6
        If FrameOwner(pfn) = -1 Then
            ChooseFrame = pfn
            Exit Function
        End If
    Next pfn
    Do
        owner = FrameOwner(mHand)
        If owner < 0 Then Err.Raise LCA_ERR_PAGER_STATE, "LcaPager", "CLOCK_EMPTY_FRAME"
        entry = MmuPage(owner)
        If Not entry.Referenced Then
            ChooseFrame = mHand
            mHand = (mHand + 1) Mod 7
            Exit Function
        End If
        MmuClearReference owner
        AddTrace "CLOCK clear ref VPN=" & CStr(owner) & " PFN=" & CStr(mHand)
        mHand = (mHand + 1) Mod 7
    Loop
End Function

Public Function PagerFrameOwner(ByVal pfn As Long) As Long
    RequirePager
    If pfn < 0 Or pfn > 6 Then Err.Raise LCA_ERR_PAGER_STATE, "LcaPager", "PFN_RESERVED_OR_INVALID"
    PagerFrameOwner = FrameOwner(pfn)
End Function

Private Function FrameOwner(ByVal pfn As Long) As Long
    Dim vpn As Long, entry As LcaPageEntry
    FrameOwner = -1
    For vpn = 0 To 14
        entry = MmuPage(vpn)
        If entry.Present And entry.PFN = pfn Then
            FrameOwner = vpn
            Exit Function
        End If
    Next vpn
End Function

Public Function PagerHand() As Long
    RequirePager
    PagerHand = mHand
End Function

Public Function PagerFaultCount() As Long
    RequirePager
    PagerFaultCount = mFaultCount
End Function

Public Function PagerEvictionCount() As Long
    RequirePager
    PagerEvictionCount = mEvictionCount
End Function

Public Function PagerWritebackCount() As Long
    RequirePager
    PagerWritebackCount = mWritebackCount
End Function

Public Function PagerLoadedBytes() As Long
    RequirePager
    PagerLoadedBytes = mLoadByteCount
End Function

Public Function PagerWrittenBytes() As Long
    RequirePager
    PagerWrittenBytes = mWritebackByteCount
End Function

Public Function PagerLastFaultVA() As Long
    RequirePager
    PagerLastFaultVA = mLastVA
End Function

Public Function PagerLastFaultVPN() As Long
    RequirePager
    PagerLastFaultVPN = mLastVPN
End Function

Public Function PagerLastVictimVPN() As Long
    RequirePager
    PagerLastVictimVPN = mLastVictim
End Function

Public Function PagerTraceCount() As Long
    RequirePager
    PagerTraceCount = mTraceCount
End Function

Public Function PagerTraceLine(ByVal index As Long) As String
    RequirePager
    If index < 0 Or index >= mTraceCount Or index < mTraceCount - 128 Then _
        Err.Raise LCA_ERR_PAGER_STATE, "LcaPager", "TRACE_INDEX_OUT_OF_RANGE"
    PagerTraceLine = mTrace(index Mod 128)
End Function

Private Sub AddTrace(ByVal line As String)
    mTrace(mTraceCount Mod 128) = line
    mTraceCount = mTraceCount + 1
End Sub

Private Sub RequirePager()
    If Not mReady Then Err.Raise LCA_ERR_PAGER_STATE, "LcaPager", "PAGER_RESET_REQUIRED"
End Sub
