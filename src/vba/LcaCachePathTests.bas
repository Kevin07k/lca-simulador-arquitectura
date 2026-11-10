Attribute VB_Name = "LcaCachePathTests"
Option Explicit

Private mChecks As Long

' P10 regression: residency before physical access, frame invalidation, swap recovery,
' uncached I/O and no stale cache line. Expected numbers: docs/P10_CACHE_PAGING.md.
Public Sub VerificarCachePaginacionP10()
    Dim s As LcaPathStep, entry As LcaPageEntry, value As Byte, failure As Long
    Dim cpu As LcaCpuState, level As Long, i As Long
    On Error GoTo failed
    mChecks = 0

    ' --- A: the 13-step panel scenario ------------------------------------------------------
    CachePathScenarioReset
    Check BusTraceCount() = 0 And CacheRamFills() = 0 And PagerFaultCount() = 0, "A0: clean reset"
    CachePathScenarioStep 0
    s = CachePathLast()
    cpu = CpuSnapshot()
    Check s.PageFault And s.PFN = 0 And s.VictimVPN = -1, "A1: fetch faults into free PFN0"
    Check cpu.IR = &H10 And cpu.PC = 0 And ControlFetched(2) = &H2A, "A1: IR=10h, PC unchanged"
    Check s.L1Hits = 3 And s.L1Misses = 1 And s.L2Misses = 1 And s.RamCycles = 4, "A1: L1I 3H/1M, one fill"
    Check BusTraceCount() = 12, "A1: 4 fill READ + 4 cache->MDR + 4 latch"
    Check InStr(1, BusTraceLine(0), "FILL[0] READ", vbBinaryCompare) > 0 And _
          InStr(1, BusTraceLine(0), "RAM -> L2", vbBinaryCompare) > 0, "A1: fill cycle is RAM -> L2"
    Check InStr(1, BusTraceLine(4), "FETCH[0] CACHE", vbBinaryCompare) > 0, "A1: byte delivered from L1I"

    CachePathScenarioStep 1
    s = CachePathLast()
    entry = MmuPage(8)
    Check s.PageFault And s.PFN = 1 And s.Value = &HA5, "A2: faulting ST resolved before the write"
    Check entry.Present And entry.PFN = 1 And entry.Dirty And entry.Referenced, "A2: dirty/ref after commit"
    Check RamReadUser(&H20) = &HA5 And SwapReadByte(8, 0) = 8, "A2: write-through RAM, swap still old"
    Check CacheContains(LCA_CACHE_L1D, &H20) And CacheContains(LCA_CACHE_L2, &H20), "A2: write-allocate"
    Check s.RamCycles = 5 And BusTraceCount() = 18, "A2: prepare + 4 fill + 1 write"

    CachePathScenarioStep 2
    s = CachePathLast()
    Check s.Value = &HA5 And s.L1Hits = 1 And s.RamCycles = 0 And Not s.PageFault, "A3: L1D hit"
    Check BusTraceCount() = 19, "A3: a hit adds only the internal L1D -> MDR cycle"

    For i = 3 To 5
        CachePathScenarioStep i
        s = CachePathLast()
        Check s.PageFault And s.PFN = i - 1 And s.RamCycles = 4, "A" & CStr(i + 1) & ": fault into free frame, one fill"
    Next i
    CachePathScenarioStep 6
    s = CachePathLast()
    Check s.Value = 13 And s.L1Misses = 1 And s.L2Hits = 1 And s.RamCycles = 0, "A7: L1D conflict served by L2"
    CachePathScenarioStep 7
    CachePathScenarioStep 8
    Check PagerFrameOwner(6) = 13 And PagerFaultCount() = 7, "A9: seven frames full"
    Check CacheFrameLines(LCA_CACHE_L1I, 0) = 1, "A9: VPN0 code line still in L1I"

    CachePathScenarioStep 9
    s = CachePathLast()
    Check s.Value = 38 And s.VictimVPN = 0 And s.PFN = 0, "A10: eighth page evicts clean VPN0"
    Check s.InvalidatedLines = 1 And Not s.Writeback, "A10: reused PFN0 loses its L1I line"
    Check CacheFrameLines(LCA_CACHE_L1I, 0) = 0, "A10: no L1I line of PFN0 survives"
    Check CacheFrameLines(LCA_CACHE_L2, 1) = 1 And CacheLineByte(LCA_CACHE_L2, 0, 0) = &HA5, _
          "A10: VPN8 line (A5h) still cached in L2 index 0 before step 11"

    CachePathScenarioStep 10
    s = CachePathLast()
    cpu = CpuSnapshot()
    Check s.VictimVPN = 8 And s.PFN = 1 And s.Writeback And s.InvalidatedLines = 1, _
          "A11: dirty VPN8 evicted, its L2 line invalidated, 32 B written back"
    Check cpu.IR = &H10 And ControlFetched(2) = &H2A, "A11: fetch reads new code 10h, never the stale A5h"
    Check s.L2Misses = 1 And s.L2Hits = 0 And CacheLineByte(LCA_CACHE_L2, 0, 0) = &H10, _
          "A11: L2 refilled from RAM after invalidation"
    Check SwapReadByte(8, 0) = &HA5 And SwapReadByte(8, 31) = 39, "A11: swap holds the modified page"

    value = CachePathReadByte(&H100)   ' scenario step 11 (same request as the panel)
    s = CachePathLast()
    entry = MmuPage(8)
    Check value = &HA5 And s.VictimVPN = 9 And s.PFN = 2, "A12: modified data recovered after swap"
    Check entry.Present And Not entry.Dirty And entry.Referenced, "A12: reloaded page clean and referenced"

    Check PagerFaultCount() = 10 And PagerEvictionCount() = 3 And PagerWritebackCount() = 1, "A: paging totals"
    Check PagerInvalidationCount() = 3 And PagerInvalidatedLines() = 3, "A: three reuses, three lines dropped"
    Check CacheHits(LCA_CACHE_L1I) = 6 And CacheMisses(LCA_CACHE_L1I) = 2, "A: L1I 6/2"
    Check CacheHits(LCA_CACHE_L1D) = 1 And CacheMisses(LCA_CACHE_L1D) = 9, "A: L1D 1/9"
    Check CacheHits(LCA_CACHE_L2) = 1 And CacheMisses(LCA_CACHE_L2) = 10, "A: L2 1/10"
    Check CacheRamFills() = 10 And CachePathRamCycles() = 41 And BusTraceCount() = 67, "A: 10 fills, 41 RAM cycles"
    Check PagerHand() = 3, "A: CLOCK hand"

    failure = 0
    On Error Resume Next
    CachePathScenarioStep 12
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    s = CachePathLast()
    Check failure = LCA_ERR_MMU_PROTECTION And s.Failed And s.PA = -1, "A13: LCD address is PROTECTION"
    Check PagerFaultCount() = 10 And CacheMisses(LCA_CACHE_L1D) = 9 And BusTraceCount() = 67, _
          "A13: I/O request touches neither pager, cache nor bus"

    ' --- B: dirty line cached in L1D and L2 when its frame is reused ----------------------
    CachePathScenarioReset
    CachePathStoreByte &H100, LcaAX                 ' VPN8 -> PFN0, PA 00h
    value = CachePathReadByte(&H124)                ' offsets 4,12,20,28 keep L1D index 0 untouched
    value = CachePathReadByte(&H14C)
    value = CachePathReadByte(&H174)
    value = CachePathReadByte(&H19C)
    value = CachePathReadByte(&H1A4)
    value = CachePathReadByte(&H1CC)
    Check CacheFrameLines(LCA_CACHE_L1D, 0) = 1 And CacheLineByte(LCA_CACHE_L1D, 0, 0) = &HA5, _
          "B: dirty byte cached in L1D"
    CachePathFetchFour 0                            ' eighth page reuses PFN0
    s = CachePathLast()
    Check s.VictimVPN = 8 And s.PFN = 0 And s.Writeback And s.InvalidatedLines = 2, "B: L1D + L2 lines dropped"
    Check Not CacheContains(LCA_CACHE_L1D, 0) And CacheLineByte(LCA_CACHE_L2, 0, 0) = &H10, _
          "B: no stale A5h left for PA 00h"
    value = CachePathReadByte(&H100)
    Check value = &HA5 And CachePathLast().VictimVPN = 9, "B: written value survives the round trip"

    ' --- C: P08 uncached ST keeps resident lines coherent (snoop) -----------------------------
    CachePathScenarioReset
    Check CachePathReadByte(&H100) = 8, "C: line cached"
    CpuWriteRegister LcaBX, &H5A
    PagerStoreByte &H100, LcaBX
    Check CachePathReadByte(&H100) = &H5A And CachePathLast().L1Hits = 1, "C: L1D hit returns the new byte"

    ' --- D: I/O and errors never reach the cache ---------------------------------------------
    CachePathScenarioReset
    failure = TryRead(&H1E8)
    Check failure = LCA_ERR_MMU_PROTECTION, "D: LCD read rejected"
    failure = TryStore(&H1E0, LcaAX)
    Check failure = LCA_ERR_MMU_PROTECTION, "D: keyboard write rejected"
    failure = TryFetch(&H100)
    Check failure = LCA_ERR_CACHE_PATH, "D: PC outside code rejected before translation"
    failure = TryRead(512)
    Check failure = LCA_ERR_MMU_RANGE, "D: VA 512 out of range"
    failure = TryStore(&H100, 99)
    Check failure = LCA_ERR_REGISTER, "D: bad register before any fault"
    failure = TryFetch(2)
    Check failure = LCA_ERR_CACHE_PATH, "D: unaligned PC"
    Check PagerFaultCount() = 0 And CacheRamFills() = 0 And BusTraceCount() = 0, "D: no side effects"
    Check RamReadPhysical(&HE8) = &H20 And RamReadPhysical(&HE0) = 0, "D: I/O frame untouched"
    Check CachePathReadByte(&H100) = 8, "D: path still usable after errors"
    For level = 0 To 2
        Check CacheFrameLines(level, 7) = 0, "D: no " & CacheLevelName(level) & " line for PFN7"
    Next level
    Check CachePathAccessCount() = 7, "D: failed requests are traced"

    CachePathPanelRefresh
    LcaReport "P10 PASS: " & CStr(mChecks) & " comprobaciones / assertions." & vbCrLf & _
              "Residencia, invalidacion por marco, swap y E/S sin cache verificados.", True
    Exit Sub
failed:
    LcaReport "P10 FAIL after " & CStr(mChecks) & " checks: " & Err.Description, False
End Sub

Private Function TryRead(ByVal va As Long) As Long
    Dim value As Byte
    On Error Resume Next
    value = CachePathReadByte(va)
    TryRead = Err.Number
    Err.Clear
End Function

Private Function TryStore(ByVal va As Long, ByVal reg As Long) As Long
    On Error Resume Next
    CachePathStoreByte va, reg
    TryStore = Err.Number
    Err.Clear
End Function

Private Function TryFetch(ByVal va As Long) As Long
    On Error Resume Next
    CachePathFetchFour va
    TryFetch = Err.Number
    Err.Clear
End Function

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    mChecks = mChecks + 1
    If Not condition Then Err.Raise vbObjectError + 2102, "LcaCachePathTests", message
End Sub
