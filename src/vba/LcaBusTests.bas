Attribute VB_Name = "LcaBusTests"
Option Explicit

Private mChecks As Long

' Run in desktop Excel. Tests use a small local fixture, never the future P06 RAM.
Public Sub VerificarBusesP05()
    Dim i As Long
    Dim bytes(0 To 3) As Byte
    Dim fixture(0 To 255) As Byte
    Dim state As LcaCpuState
    Dim cycle As LcaBusCycle
    Dim failure As Long
    On Error GoTo failed
    mChecks = 0
    CpuReset
    ControlReset
    CpuWriteControl LcaPC, &H20
    CpuWriteRegister LcaAX, &H2A
    fixture(&H40) = &H10
    fixture(&H41) = 0
    fixture(&H42) = &H2A
    fixture(&H43) = 0
    ControlStartFetch &H20, &H40
    For i = 0 To 3
        ControlReadFetchByte fixture(&H40 + i)
        cycle = BusSnapshot(LCA_BUS_EXTERNAL)
        Check cycle.DriverCount = 1 And cycle.Driver = "RAM" And cycle.LoadSignal, "fetch: one RAM driver"
        Check cycle.ReadSignal And Not cycle.WriteSignal And cycle.ClockSignal And Not cycle.IOSignal, "fetch: read word"
        Check cycle.PA = &H40 + i And cycle.VA = &H20 + i, "fetch: addresses"
        ControlLatchFetchByte
        cycle = BusSnapshot(LCA_BUS_EXTERNAL)
        Check cycle.DriverCount = 0 And cycle.Driver = "High-Z", "latch: external High-Z"
        cycle = BusSnapshot(LCA_BUS_INTERNAL)
        Check cycle.DriverCount = 1 And cycle.Driver = "MDR" And cycle.ClockSignal, "latch: MDR driver"
        bytes(i) = ControlFetched(i)
    Next i
    state = CpuSnapshot()
    Check bytes(0) = &H10 And bytes(1) = 0 And bytes(2) = &H2A And bytes(3) = 0, "four fetched bytes"
    Check state.IR = &H10 And state.MDR = 0 And state.MAR_V = &H23 And state.MAR_P = &H43, "CPU control registers"
    Check state.PC = &H20, "fetch does not commit PC"

    ControlPrepareStore &H14D, &H6D, LcaAX
    cycle = BusSnapshot(LCA_BUS_INTERNAL)
    Check cycle.Driver = "AX" And cycle.Receiver = "MDR" And cycle.Data = &H2A, "ST prepare"
    Check fixture(&H6D) = 0, "ST has not written fixture before clock"
    fixture(&H6D) = ControlCommitStore()
    cycle = BusSnapshot(LCA_BUS_EXTERNAL)
    Check fixture(&H6D) = &H2A And cycle.WriteSignal And Not cycle.ReadSignal, "ST commits after clock"
    Check cycle.PA = &H6D And cycle.VA = &H14D And cycle.DriverCount = 1, "ST destination and arbitration"
    Check BusTraceCount() = 10, "8 fetch steps plus 2 write steps"
    Check InStr(1, BusTraceLine(9), "ST WRITE", vbBinaryCompare) > 0, "trace contains final write"

    BusBegin LCA_BUS_EXTERNAL, "IDLE", -1, -1, False, False, False
    Check BusClock(LCA_BUS_EXTERNAL) = -1, "idle clock returns High-Z"
    cycle = BusSnapshot(LCA_BUS_EXTERNAL)
    Check cycle.DriverCount = 0 And cycle.Driver = "High-Z" And Not cycle.LoadSignal, "zero idle drivers"

    BusBegin LCA_BUS_EXTERNAL, "IN SIGNALS", &H1E0, &HE0, True, False, True, True, False
    BusDrive LCA_BUS_EXTERNAL, "KBD", &H41
    BusReceive LCA_BUS_EXTERNAL, "AX"
    Check BusClock(LCA_BUS_EXTERNAL) = &H41, "keyboard bus delivers byte"
    cycle = BusSnapshot(LCA_BUS_EXTERNAL)
    Check cycle.IOSignal And cycle.ReadSignal And cycle.EnKBD And Not cycle.LdLCD, "keyboard control word"
    BusBegin LCA_BUS_EXTERNAL, "OUT SIGNALS", &H1E8, &HE8, False, True, True, False, True
    BusDrive LCA_BUS_EXTERNAL, "AX", &H41
    BusReceive LCA_BUS_EXTERNAL, "LCD"
    Check BusClock(LCA_BUS_EXTERNAL) = &H41, "LCD bus delivers byte"
    cycle = BusSnapshot(LCA_BUS_EXTERNAL)
    Check cycle.IOSignal And cycle.WriteSignal And cycle.LdLCD And Not cycle.EnKBD, "LCD control word"

    BusBegin LCA_BUS_EXTERNAL, "CONTENTION", &H14D, &H6D, False, True, False
    BusDrive LCA_BUS_EXTERNAL, "CPU", &H2A
    BusDrive LCA_BUS_EXTERNAL, "RAM_OUTPUT", &H55
    BusReceive LCA_BUS_EXTERNAL, "RAM"
    failure = 0
    On Error Resume Next
    Call BusClock(LCA_BUS_EXTERNAL)
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_CONTENTION, "competing outputs raise BUS_CONTENTION"
    Check fixture(&H6D) = &H2A, "contention cannot commit write"
    Check InStr(1, BusTraceLine(BusTraceCount() - 1), "BUS_CONTENTION", vbBinaryCompare) > 0, "contention traced"

    BusPanelActualizar
    LcaReport "P05 PASS: " & mChecks & " comprobaciones / assertions." & vbCrLf & _
           "La prueba usa bytes inyectados; no prueba RAM/cache P06 ni ejecucion P12.", True
    Exit Sub
failed:
    LcaReport "P05 FAIL after " & mChecks & " checks: " & Err.Description, False
End Sub

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    mChecks = mChecks + 1
    If Not condition Then Err.Raise vbObjectError + 2054, "LcaBusTests", message
End Sub
