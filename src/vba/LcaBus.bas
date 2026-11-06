Attribute VB_Name = "LcaBus"
Option Explicit

' P05 transport layer. A cycle is inspected before its receiver commits a write.
Public Const LCA_BUS_INTERNAL As Long = 0
Public Const LCA_BUS_EXTERNAL As Long = 1
Public Const LCA_ERR_BUS As Long = vbObjectError + 2051
Public Const LCA_ERR_CONTENTION As Long = vbObjectError + 2052
Private Const TRACE_CAPACITY As Long = 128

Public Type LcaBusCycle
    BusName As String
    Microstep As String
    VA As Long                    ' -1 if the cycle has no virtual address
    PA As Long                    ' -1 if the cycle has no physical address
    Data As Long                  ' meaningful only when DriverCount = 1
    Driver As String
    Receiver As String
    DriverCount As Long
    ReadSignal As Boolean
    WriteSignal As Boolean
    ClockSignal As Boolean
    IOSignal As Boolean
    EnableSignal As Boolean
    LoadSignal As Boolean
    EnKBD As Boolean
    LdLCD As Boolean
    Fault As String
End Type

Private mCycle(0 To 1) As LcaBusCycle
Private mArmed(0 To 1) As Boolean
Private mTrace(0 To TRACE_CAPACITY - 1) As String
Private mTraceCount As Long
Private mTraceNext As Long
Private mInitialized As Boolean

Public Sub BusReset()
    Dim i As Long
    For i = 0 To 1
        mArmed(i) = False
        BusFloat i
    Next i
    mTraceCount = 0
    mTraceNext = 0
    For i = 0 To TRACE_CAPACITY - 1
        mTrace(i) = vbNullString
    Next i
    mInitialized = True
End Sub

' Float the inactive bus without manufacturing an extra clock cycle.
Public Sub BusFloat(ByVal busId As Long)
    Dim emptyCycle As LcaBusCycle
    RequireBus busId
    If mArmed(busId) Then Err.Raise LCA_ERR_BUS, "LcaBus", "BUS_CYCLE_PENDING"
    mCycle(busId) = emptyCycle
    mCycle(busId).BusName = BusLabel(busId)
    mCycle(busId).VA = -1
    mCycle(busId).PA = -1
    mCycle(busId).Driver = "High-Z"
    mCycle(busId).Data = -1
    mArmed(busId) = False
End Sub

Public Sub BusBegin(ByVal busId As Long, ByVal microstep As String, ByVal va As Long, _
                    ByVal pa As Long, ByVal readSignal As Boolean, ByVal writeSignal As Boolean, _
                    ByVal ioSignal As Boolean, Optional ByVal enableKeyboard As Boolean = False, _
                    Optional ByVal loadLcd As Boolean = False)
    If Not mInitialized Then BusReset
    RequireBus busId
    If mArmed(busId) Then Err.Raise LCA_ERR_BUS, "LcaBus", "BUS_CYCLE_PENDING"
    If va < -1 Or va > 511 Or pa < -1 Or pa > 255 Then Err.Raise LCA_ERR_BUS, "LcaBus", "BUS_ADDRESS_RANGE"
    If readSignal And writeSignal Then Err.Raise LCA_ERR_BUS, "LcaBus", "READ_WRITE_CONFLICT"
    If enableKeyboard And (Not ioSignal Or Not readSignal Or writeSignal) Then _
        Err.Raise LCA_ERR_BUS, "LcaBus", "INVALID_KBD_SIGNALS"
    If loadLcd And (Not ioSignal Or Not writeSignal Or readSignal) Then _
        Err.Raise LCA_ERR_BUS, "LcaBus", "INVALID_LCD_SIGNALS"
    If Len(Trim$(microstep)) = 0 Then Err.Raise LCA_ERR_BUS, "LcaBus", "MISSING_MICROSTEP"
    BusFloat busId
    With mCycle(busId)
        .Microstep = microstep
        .VA = va
        .PA = pa
        .ReadSignal = readSignal
        .WriteSignal = writeSignal
        .IOSignal = ioSignal
        .EnKBD = enableKeyboard
        .LdLCD = loadLcd
    End With
    mArmed(busId) = True
End Sub

Public Sub BusDrive(ByVal busId As Long, ByVal sourceName As String, ByVal value As Long)
    RequireArmed busId
    If Len(Trim$(sourceName)) = 0 Then Err.Raise LCA_ERR_BUS, "LcaBus", "MISSING_DRIVER"
    CpuRequireByte value
    With mCycle(busId)
        .DriverCount = .DriverCount + 1
        If .DriverCount = 1 Then
            .Driver = sourceName
            .Data = value
            .EnableSignal = True
        Else
            .Driver = .Driver & " + " & sourceName
            .Data = -1
            .Fault = "BUS_CONTENTION"
        End If
    End With
End Sub

Public Sub BusReceive(ByVal busId As Long, ByVal targetName As String)
    RequireArmed busId
    If Len(Trim$(targetName)) = 0 Then Err.Raise LCA_ERR_BUS, "LcaBus", "MISSING_RECEIVER"
    If mCycle(busId).LoadSignal Then Err.Raise LCA_ERR_BUS, "LcaBus", "MULTIPLE_RECEIVERS"
    mCycle(busId).Receiver = targetName
    mCycle(busId).LoadSignal = True
End Sub

' Returns the transferred byte. The caller performs the actual register/RAM write AFTER this succeeds.
Public Function BusClock(ByVal busId As Long) As Long
    Dim fault As String
    RequireArmed busId
    With mCycle(busId)
        If .DriverCount > 1 Then
            fault = "BUS_CONTENTION"
        ElseIf .DriverCount = 0 Then
            If .ReadSignal Or .WriteSignal Or .IOSignal Or .LoadSignal Then fault = "BUS_NO_DRIVER"
        ElseIf Not .LoadSignal Then
            fault = "BUS_NO_RECEIVER"
        End If
        If Len(fault) > 0 Then .Fault = fault
        .ClockSignal = True
    End With
    mArmed(busId) = False
    AppendTrace BusFormat(mCycle(busId))
    If Len(fault) > 0 Then
        If fault = "BUS_CONTENTION" Then Err.Raise LCA_ERR_CONTENTION, "LcaBus", fault
        Err.Raise LCA_ERR_BUS, "LcaBus", fault
    End If
    BusClock = mCycle(busId).Data    ' -1 for a valid idle cycle
End Function

Public Function BusSnapshot(ByVal busId As Long) As LcaBusCycle
    If Not mInitialized Then BusReset
    RequireBus busId
    BusSnapshot = mCycle(busId)
End Function

Public Function BusTraceCount() As Long
    If Not mInitialized Then BusReset
    BusTraceCount = mTraceCount
End Function

' Index 0 is the oldest retained entry; max 128 cycles.
Public Function BusTraceLine(ByVal index As Long) As String
    Dim first As Long
    If Not mInitialized Then BusReset
    If index < 0 Or index >= mTraceCount Then Err.Raise LCA_ERR_BUS, "LcaBus", "TRACE_INDEX_RANGE"
    first = (mTraceNext - mTraceCount + TRACE_CAPACITY) Mod TRACE_CAPACITY
    BusTraceLine = mTrace((first + index) Mod TRACE_CAPACITY)
End Function

Public Function BusFormat(ByRef cycle As LcaBusCycle) As String
    Dim bits As String
    bits = "R" & CStr(Abs(CLng(cycle.ReadSignal))) & " W" & CStr(Abs(CLng(cycle.WriteSignal))) & _
           " C" & CStr(Abs(CLng(cycle.ClockSignal))) & " I" & CStr(Abs(CLng(cycle.IOSignal))) & _
           " E" & CStr(Abs(CLng(cycle.EnableSignal))) & " L" & CStr(Abs(CLng(cycle.LoadSignal))) & _
           " KBD" & CStr(Abs(CLng(cycle.EnKBD))) & " LCD" & CStr(Abs(CLng(cycle.LdLCD)))
    BusFormat = cycle.Microstep & " | " & cycle.BusName & " | VA=" & AddressText(cycle.VA, 3) & _
                " PA=" & AddressText(cycle.PA, 2) & " | " & cycle.Driver & " -> " & _
                IIf(cycle.LoadSignal, cycle.Receiver, "-") & " | " & AddressText(cycle.Data, 2) & _
                " | " & bits & " | drivers=" & CStr(cycle.DriverCount)
    If Len(cycle.Fault) > 0 Then BusFormat = BusFormat & " | " & cycle.Fault
End Function

Private Function AddressText(ByVal value As Long, ByVal digits As Long) As String
    If value < 0 Then
        AddressText = "High-Z"
    Else
        AddressText = Right$(String$(digits, "0") & Hex$(value), digits) & "h"
    End If
End Function

Private Function BusLabel(ByVal busId As Long) As String
    If busId = LCA_BUS_INTERNAL Then BusLabel = "INTERNAL" Else BusLabel = "EXTERNAL"
End Function

Private Sub AppendTrace(ByVal line As String)
    mTrace(mTraceNext) = line
    mTraceNext = (mTraceNext + 1) Mod TRACE_CAPACITY
    If mTraceCount < TRACE_CAPACITY Then mTraceCount = mTraceCount + 1
End Sub

Private Sub RequireBus(ByVal busId As Long)
    If busId <> LCA_BUS_INTERNAL And busId <> LCA_BUS_EXTERNAL Then _
        Err.Raise LCA_ERR_BUS, "LcaBus", "INVALID_BUS"
End Sub

Private Sub RequireArmed(ByVal busId As Long)
    RequireBus busId
    If Not mArmed(busId) Then Err.Raise LCA_ERR_BUS, "LcaBus", "BUS_CYCLE_NOT_STARTED"
End Sub
