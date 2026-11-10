Attribute VB_Name = "LcaPagerTests"
Option Explicit

Private mChecks As Long

Public Sub VerificarPagerP08()
    Dim vpn As Long, offset As Long, code As LcaCpuState
    Dim entry As LcaPageEntry, value As Byte, failure As Long
    On Error GoTo failed
    mChecks = 0
    CpuReset
    MemoryPathReset
    PagerReset
    For vpn = 8 To 14
        For offset = 0 To 31
            SwapWriteByte vpn, offset, vpn + offset
        Next offset
    Next vpn
    SwapWriteByte 0, 0, &H10
    SwapWriteByte 0, 1, 0
    SwapWriteByte 0, 2, &H2A
    SwapWriteByte 0, 3, 0
    CpuWriteRegister LcaAX, &HA5
    PagerFetchFour 0
    code = CpuSnapshot()
    Check code.PC = 0 And code.IR = &H10 And ControlFetched(2) = &H2A, "faulting FETCH resumes once; PC unchanged"
    Check PagerFaultCount() = 1 And PagerLoadedBytes() = 32, "one page fault loads exactly 32 bytes"
    Check BusTraceCount() = 8, "one fetch has four reads and four latches"
    Check PagerFrameOwner(0) = 0 And PagerHand() = 0, "lowest free PFN 0 and untouched CLOCK hand"

    PagerStoreByte 256, LcaAX
    entry = MmuPage(8)
    Check entry.Present And entry.PFN = 1 And entry.Dirty And entry.Referenced, "faulting ST commits once"
    Check RamReadUser(32) = &HA5 And SwapReadByte(8, 0) = 8, "dirty byte remains only in RAM before eviction"
    Check PagerFaultCount() = 2 And BusTraceCount() = 10, "one ST produces exactly two bus microsteps"
    For vpn = 9 To 13
        value = PagerReadByte(vpn * 32)
        Check value = vpn, "backing image loaded at first access"
    Next vpn
    Check PagerFaultCount() = 7 And PagerLoadedBytes() = 224, "seven pages fill seven frames"
    Check PagerFrameOwner(6) = 13 And PagerFrameOwner(1) = 8, "frame ownership before replacement"
    value = PagerReadByte(14 * 32)
    Check value = 14 And PagerLastVictimVPN() = 0, "eighth page evicts clean VPN 0"
    Check PagerHand() = 1 And PagerFrameOwner(0) = 14, "CLOCK hand advances past selected PFN 0"
    entry = MmuPage(0)
    Check Not entry.Present And entry.PFN = -1, "clean victim has no stale mapping"

    PagerFetchFour 0
    Check PagerLastVictimVPN() = 8 And PagerFrameOwner(1) = 0, "CLOCK next chooses dirty VPN 8"
    Check PagerWritebackCount() = 1 And PagerWrittenBytes() = 32, "dirty victim writes exactly 32 bytes"
    Check SwapReadByte(8, 0) = &HA5 And SwapReadByte(8, 31) = 39, "both ends of dirty image preserved"
    Check PagerHand() = 2 And BusTraceCount() = 18, "fault processing does not duplicate bus effects"
    Check RamReadPhysical(&HE8) = &H20, "reserved LCD frame remains unchanged"
    value = PagerReadByte(256)
    entry = MmuPage(8)
    Check value = &HA5 And entry.Present And entry.PFN = 2, "evicted dirty data recovers on refault"
    Check Not entry.Dirty And entry.Referenced, "reloaded page is clean and referenced"
    Check PagerFaultCount() = 10 And PagerLoadedBytes() = 320, "ten faults load 320 bytes exactly"
    Check PagerEvictionCount() = 3 And PagerWritebackCount() = 1, "three replacements, one writeback"
    value = PagerReadByte(256)
    Check value = &HA5 And PagerFaultCount() = 10, "hit does not refault or duplicate effects"

    failure = 0
    On Error Resume Next
    value = PagerReadByte(511)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_MMU_PROTECTION And PagerFaultCount() = 10, "I/O VPN 15 is never paged"
    failure = 0
    On Error Resume Next
    value = PagerReadByte(512)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_MMU_RANGE And PagerFaultCount() = 10, "invalid VA never faults into memory"
    failure = 0
    On Error Resume Next
    value = SwapReadByte(15, 0)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_SWAP_ADDRESS, "reserved VPN has no swap image"
    failure = 0
    On Error Resume Next
    PagerStoreByte 256, 99
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_REGISTER And PagerFaultCount() = 10, "invalid register rejected before any retry"
    code = CpuSnapshot()
    Check code.PC = 0 And CpuReadRegister(LcaAX) = &HA5, "demo does not execute opcodes"
    PagerPanelRefresh
    LcaReport "P08 PASS: " & CStr(mChecks) & " comprobaciones / assertions.", True
    Exit Sub
failed:
    LcaReport "P08 FAIL after " & CStr(mChecks) & " checks: " & Err.Description, False
End Sub

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    mChecks = mChecks + 1
    If Not condition Then Err.Raise vbObjectError + 2083, "LcaPagerTests", message
End Sub
