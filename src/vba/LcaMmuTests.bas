Attribute VB_Name = "LcaMmuTests"
Option Explicit

Private mChecks As Long

Public Sub VerificarMmuP07()
    Dim vpn As Long, pa As Long, failure As Long
    Dim entry As LcaPageEntry
    Dim cpu As LcaCpuState
    Dim cycle As LcaBusCycle
    On Error GoTo failed
    mChecks = 0
    MmuReset
    RamReset
    MemoryPathReset
    CpuReset

    For vpn = 0 To 15
        entry = MmuPage(vpn)
        Check entry.VPN = vpn And Not entry.Present And entry.PFN = -1, "initial PTE absent"
        Check Not entry.Dirty And Not entry.Referenced, "initial PTE flags clear"
        Check entry.Reserved = (vpn = 15), "only VPN 15 reserved"
    Next vpn
    Check MmuVpnOf(0) = 0 And MmuOffsetOf(0) = 0, "VA 0 split"
    Check MmuVpnOf(31) = 0 And MmuOffsetOf(31) = 31, "VA 31 split"
    Check MmuVpnOf(32) = 1 And MmuOffsetOf(32) = 0, "VA 32 split"
    Check MmuVpnOf(255) = 7 And MmuOffsetOf(255) = 31, "VA 255 split"
    Check MmuVpnOf(256) = 8 And MmuOffsetOf(256) = 0, "VA 256 split"
    Check MmuVpnOf(511) = 15 And MmuOffsetOf(511) = 31, "VA 511 is in reserved VPN 15"

    pa = 123
    Check Not MmuTryTranslate(0, MMU_FETCH, pa), "absent VPN 0 returns False"
    Check pa = -1 And MmuLastFaultVPN() = 0, "absent page never looks like PA zero"
    MmuMap 0, 0
    MmuMap 1, 1
    MmuMap 7, 3
    MmuMap 8, 4
    MmuMap 9, 2
    Check MmuTryTranslate(0, MMU_FETCH, pa) And pa = 0, "resident PFN 0 is valid PA zero"
    Check MmuTranslate(31, MMU_FETCH) = 31, "VA 31 -> PA 31"
    Check MmuTranslate(32, MMU_FETCH) = 32, "VA 32 -> PA 32"
    Check MmuTranslate(255, MMU_FETCH) = 127, "VA 255 -> PA 127"
    Check MmuTranslate(256, MMU_READ) = 128, "VA 256 -> PA 128"
    Check MmuTranslate(300, MMU_READ) = 76, "VA 300, PFN 2 -> PA 76"
    entry = MmuPage(9)
    Check Not entry.Referenced And Not entry.Dirty, "translation alone does not commit ref/dirty"

    failure = 0
    On Error Resume Next
    pa = MmuTranslate(-1, MMU_FETCH)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_MMU_RANGE, "negative VA rejected"
    failure = 0
    On Error Resume Next
    pa = MmuTranslate(512, MMU_READ)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_MMU_RANGE, "VA 512 rejected"
    failure = 0
    On Error Resume Next
    pa = MmuTranslate(511, MMU_READ)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_MMU_PROTECTION, "reserved VA 511 denied"
    failure = 0
    On Error Resume Next
    pa = MmuTranslate(&H1E8, MMU_WRITE)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_MMU_PROTECTION, "LCD address cannot be ST"
    failure = 0
    On Error Resume Next
    pa = MmuTranslate(256, MMU_FETCH)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_MMU_PROTECTION, "cannot fetch from data"
    failure = 0
    On Error Resume Next
    MmuMap 15, 6
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_MMU_PROTECTION, "VPN 15 cannot map"
    failure = 0
    On Error Resume Next
    MmuMap 10, 7
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_MMU_MAPPING, "PFN 7 cannot map"
    failure = 0
    On Error Resume Next
    MmuMap 10, 2
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_MMU_MAPPING, "two VPNs cannot own PFN 2"
    entry = MmuPage(15)
    Check entry.Reserved And Not entry.Present And entry.PFN = -1, "reserved PTE unchanged"

    failure = 0
    On Error Resume Next
    MmuFetchFour 64
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_PAGE_FAULT And MmuLastFaultVPN() = 2, "absent fetch raises recoverable PAGE_FAULT"
    Check BusTraceCount() = 0, "page fault precedes bus transfer"

    CpuWriteControl LcaPC, 32
    CpuWriteRegister LcaAX, &H2A
    RamWriteUser 32, &H10
    RamWriteUser 33, 0
    RamWriteUser 34, &H2A
    RamWriteUser 35, 0
    MmuFetchFour 32
    cpu = CpuSnapshot()
    Check cpu.IR = &H10 And cpu.MAR_V = 35 And cpu.MAR_P = 35, "translated RAM fetch reaches IR"
    Check cpu.PC = 32 And ControlFetched(2) = &H2A, "P12 still owns PC/decoding"
    entry = MmuPage(1)
    Check entry.Referenced And Not entry.Dirty, "completed fetch sets ref only"
    Check BusTraceCount() = 8, "four RAM reads and four internal latches"
    Check RamReadUser(76) = 0, "store target starts empty"
    failure = 0
    On Error Resume Next
    MmuCommitAccess 300, MMU_WRITE, 75
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    entry = MmuPage(9)
    Check failure = LCA_ERR_MMU_STATE And Not entry.Dirty, "stale PA cannot mark a page dirty"
    MmuStoreByte 300, LcaAX
    entry = MmuPage(9)
    cycle = BusSnapshot(LCA_BUS_EXTERNAL)
    Check RamReadUser(76) = &H2A And entry.Dirty And entry.Referenced, "translated ST commits RAM and PTE bits"
    Check cycle.VA = 300 And cycle.PA = 76 And cycle.WriteSignal And cycle.ClockSignal, "ST uses translated PA and bus"
    Check BusTraceCount() = 10, "ST adds two bus microsteps"

    failure = 0
    On Error Resume Next
    MmuUnmap 9
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_MMU_STATE, "dirty page cannot be evicted without writeback"
    MmuClearDirty 9
    MmuUnmap 9
    entry = MmuPage(9)
    Check Not entry.Present And entry.PFN = -1 And Not entry.Referenced, "unmap removes stale PFN/flags"
    MmuMap 10, 2
    Check MmuTranslate(&H140, MMU_READ) = 64, "freed PFN 2 reused by VPN 10"

    MmuPanelRefresh
    MsgBox "P07 PASS: " & mChecks & " comprobaciones / assertions." & vbCrLf & _
           "P08 atiende PAGE_FAULT; P09 integra cache y P12 ejecuta instrucciones.", vbInformation
    Exit Sub
failed:
    MsgBox "P07 FAIL after " & mChecks & " checks: " & Err.Description, vbCritical
End Sub

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    mChecks = mChecks + 1
    If Not condition Then Err.Raise vbObjectError + 2076, "LcaMmuTests", message
End Sub
