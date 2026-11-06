Attribute VB_Name = "LcaRamTests"
Option Explicit

Private mChecks As Long

' Desktop Excel verification: physical boundaries and actual RAM -> P05 bus -> CPU path.
Public Sub VerificarRamP06()
    Dim cycle As LcaBusCycle
    Dim cpu As LcaCpuState
    Dim i As Long, value As Byte, failure As Long
    On Error GoTo failed
    mChecks = 0
    CpuReset
    RamReset
    MemoryPathReset

    RamWritePhysical 0, &H11
    RamWritePhysical 31, &H22
    RamWritePhysical 32, &H33
    RamWritePhysical 255, &H44
    Check RamReadPhysical(0) = &H11 And RamFrameOf(0) = 0 And RamOffsetOf(0) = 0, "PA 0 -> frame 0 offset 0"
    Check RamReadPhysical(31) = &H22 And RamFrameOf(31) = 0 And RamOffsetOf(31) = 31, "PA 31 -> frame 0 offset 31"
    Check RamReadPhysical(32) = &H33 And RamFrameOf(32) = 1 And RamOffsetOf(32) = 0, "PA 32 -> frame 1 offset 0"
    Check RamReadPhysical(255) = &H44 And RamFrameOf(255) = 7 And RamOffsetOf(255) = 31, "PA 255 -> frame 7 offset 31 (service only)"
    Check RamReadPhysical(1) = 0 And RamReadPhysical(33) = 0, "adjacent bytes remain independent"
    Check RamReadIO(&HE8) = &H20 And RamReadIO(&HEF) = &H20, "LCD buffer reset to eight spaces"

    failure = 0
    On Error Resume Next
    value = RamReadPhysical(-1)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_RAM_ADDRESS, "PA -1 rejected"
    failure = 0
    On Error Resume Next
    RamWritePhysical 256, &H55
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_RAM_ADDRESS And RamReadPhysical(255) = &H44, "PA 256 rejected without corruption"
    failure = 0
    On Error Resume Next
    value = RamReadUser(&HFF)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_RAM_PROTECTION, "reserved FFh cannot enter CPU path"
    failure = 0
    On Error Resume Next
    RamWriteUser &HE0, &H55
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_RAM_PROTECTION And RamReadIO(&HE0) = 0, "I/O frame cannot enter CPU data path"

    RamReset
    MemoryPathReset
    CpuWriteControl LcaPC, &H20
    CpuWriteRegister LcaAX, &H2A
    RamWriteUser &H40, &H10
    RamWriteUser &H41, 0
    RamWriteUser &H42, &H2A
    RamWriteUser &H43, 0
    MemoryFetchStart &H20, &H40   ' supplied sample PTE[1] = PFN 2; P07 will implement translation
    For i = 0 To 3
        MemoryFetchReadNext
        cycle = BusSnapshot(LCA_BUS_EXTERNAL)
        Check cycle.ReadSignal And cycle.ClockSignal And Not cycle.IOSignal, "fetch external READ"
        Check cycle.DriverCount = 1 And cycle.Driver = "RAM" And cycle.PA = &H40 + i, "actual RAM drives bus"
        MemoryFetchLatchNext
        cycle = BusSnapshot(LCA_BUS_EXTERNAL)
        Check cycle.DriverCount = 0 And cycle.Driver = "High-Z", "external bus floats for latch"
        Check ControlFetched(i) = RamReadUser(&H40 + i), "CPU fetch buffer equals RAM"
    Next i
    cpu = CpuSnapshot()
    Check cpu.IR = &H10 And cpu.MDR = 0 And cpu.MAR_V = &H23 And cpu.MAR_P = &H43, "IR and CPU controls loaded via P05"
    Check cpu.PC = &H20, "P12, not P06, advances PC"
    Check BusTraceCount() = 8 And InStr(1, BusTraceLine(0), "PA=40h", vbBinaryCompare) > 0, "fetch trace starts at 40h"

    Check RamReadUser(&H6D) = 0, "store destination initially zero"
    MemoryStorePrepare &H14D, &H6D, LcaAX
    Check RamReadUser(&H6D) = 0, "prepare edge cannot write RAM"
    cycle = BusSnapshot(LCA_BUS_INTERNAL)
    Check cycle.Driver = "AX" And cycle.Receiver = "MDR" And cycle.Data = &H2A, "internal write preparation"
    MemoryStoreCommit
    cycle = BusSnapshot(LCA_BUS_EXTERNAL)
    Check RamReadUser(&H6D) = &H2A And cycle.WriteSignal And cycle.ClockSignal, "RAM changed after external WRITE"
    Check cycle.PA = &H6D And cycle.VA = &H14D And cycle.DriverCount = 1, "store address, value, arbitration"
    Check BusTraceCount() = 10 And InStr(1, BusTraceLine(9), "ST WRITE", vbBinaryCompare) > 0, "write recorded in trace"

    BusBegin LCA_BUS_EXTERNAL, "CONTENTION", &H14D, &H6D, False, True, False
    BusDrive LCA_BUS_EXTERNAL, "CPU", &H77
    BusDrive LCA_BUS_EXTERNAL, "OTHER_OUTPUT", &H88
    BusReceive LCA_BUS_EXTERNAL, "RAM"
    failure = 0
    On Error Resume Next
    Call BusClock(LCA_BUS_EXTERNAL)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_CONTENTION, "two drivers rejected before RAM commit"
    Check RamReadUser(&H6D) = &H2A, "contention leaves physical RAM unchanged"
    Check InStr(1, BusTraceLine(BusTraceCount() - 1), "BUS_CONTENTION", vbBinaryCompare) > 0, "failed cycle traced"

    BusPanelActualizar
    MsgBox "P06 PASS: " & mChecks & " comprobaciones / assertions." & vbCrLf & _
           "RAM real y bus P05 comprobados. P07/P09/P12 aun no forman parte de esta prueba.", vbInformation
    Exit Sub
failed:
    MsgBox "P06 FAIL after " & mChecks & " checks: " & Err.Description, vbCritical
End Sub

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    mChecks = mChecks + 1
    If Not condition Then Err.Raise vbObjectError + 2064, "LcaRamTests", message
End Sub
