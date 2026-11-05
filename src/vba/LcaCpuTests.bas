Attribute VB_Name = "LcaCpuTests"
Option Explicit

Private mAssertions As Long
Private mCases As Long

' Reproducible Excel test entry point. Engine modules never access worksheets.
Public Sub VerificarCpuAlu()
    Dim report As Worksheet
    Dim errorNumber As Long
    Dim errorText As String
    On Error GoTo Failed
    Set report = ReportSheet()
    report.Cells.ClearContents
    report.Range("A1:B1").value = Array("P04 CPU / ALU", "RUNNING")
    mAssertions = 0
    mCases = 0
    CheckStateAndValidation
    CheckAllArithmetic
    CpuReset
    report.Range("A1:B1").value = Array("P04 CPU / ALU", "PASS")
    report.Range("A2:B2").value = Array("Arithmetic cases / Casos", mCases)
    report.Range("A3:B3").value = Array("Assertions / Comprobaciones", mAssertions)
    report.Range("A4:B4").value = Array("Result / Resultado", "All checks passed / Sin fallos")
    report.Range("A5:B5").value = Array("Final CPU / CPU final", "RESET")
    report.Columns("A:B").AutoFit
    Debug.Print "P04 PASS; cases=" & mCases & "; assertions=" & mAssertions
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    CpuReset
    If Not report Is Nothing Then
        report.Range("B1").value = "FAIL"
        report.Range("A4:B4").value = Array("Failure / Fallo", errorText)
    End If
    Err.Raise errorNumber, "LcaCpuTests", errorText
End Sub

Private Function ReportSheet() As Worksheet
    Dim sheet As Worksheet
    For Each sheet In ThisWorkbook.Worksheets
        If sheet.Name = "P04 Tests" Then
            Set ReportSheet = sheet
            Exit Function
        End If
    Next sheet
    Set sheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    sheet.Name = "P04 Tests"
    Set ReportSheet = sheet
End Function

Private Sub AssertEqual(ByVal actual As Long, ByVal expected As Long, ByVal label As String)
    mAssertions = mAssertions + 1
    If actual <> expected Then
        Err.Raise vbObjectError + 2100, "LcaCpuTests", _
                  label & ": expected=" & expected & ", actual=" & actual
    End If
End Sub

Private Sub CheckStateAndValidation()
    Dim state As LcaCpuState
    Dim index As Long
    Dim actual As Byte
    CpuReset
    For index = 0 To 3
        AssertEqual CpuReadRegister(index), 0, "initial register"
        CpuWriteRegister index, 255 - index
        AssertEqual CpuReadRegister(index), 255 - index, "register mapping"
    Next index
    CpuWriteControl LcaPC, 256
    CpuWriteControl LcaMAR_V, 511
    CpuWriteControl LcaMAR_P, 255
    CpuWriteControl LcaMDR, 128
    CpuWriteControl LcaIR, 255
    state = CpuSnapshot()
    AssertEqual state.PC, 256, "9-bit PC"
    AssertEqual state.MAR_V, 511, "9-bit MAR_V"
    AssertEqual state.MAR_P, 255, "8-bit MAR_P"
    AssertEqual state.MDR, 128, "MDR"
    AssertEqual state.IR, 255, "IR"
    state.AX = 0
    AssertEqual CpuReadRegister(LcaAX), 255, "snapshot is a copy"
    AssertEqual ExpectedError(1), LCA_ERR_REGISTER, "invalid register"
    AssertEqual ExpectedError(2), LCA_ERR_BYTE, "negative byte"
    AssertEqual ExpectedError(3), LCA_ERR_BYTE, "byte 256"
    AssertEqual ExpectedError(4), LCA_ERR_ADDRESS, "negative VA"
    AssertEqual ExpectedError(5), LCA_ERR_ADDRESS, "VA 512"
    AssertEqual ExpectedError(6), LCA_ERR_BYTE, "physical address 256"
    AssertEqual ExpectedError(7), LCA_ERR_CONTROL, "invalid control"
    actual = AluExecute(LcaCMP, LcaAX, LcaAX)
    AssertEqual ExpectedError(8), LCA_ERR_REGISTER, "ALU invalid source"
    AssertEqual ExpectedError(9), LCA_ERR_ALU, "ALU invalid operation"
    AssertEqual ExpectedError(10), LCA_ERR_REGISTER, "ALU invalid destination"
    state = CpuSnapshot()
    AssertEqual state.AX, 255, "errors preserve register"
    AssertEqual state.flags, 1, "errors preserve flags"
    AssertEqual state.PC, 256, "errors preserve PC"
    AssertEqual state.MAR_V, 511, "errors preserve MAR_V"
    AssertEqual state.MAR_P, 255, "errors preserve MAR_P"
    CpuWriteRegister LcaAX, 0
    actual = AluExecute(LcaDEC, LcaAX)
    AssertEqual CLng(CpuFlagIsSet(LCA_FLAG_CF)), -1, "borrow flag query"
    AssertEqual CLng(CpuFlagIsSet(LCA_FLAG_SF)), -1, "sign flag query"
    AssertEqual CLng(CpuFlagIsSet(LCA_FLAG_ZF)), 0, "zero flag query"
    CpuWriteRegister LcaDX, 42
    state = CpuSnapshot()
    AssertEqual state.flags, 6, "MOV-style write preserves flags"
    CpuReset
    state = CpuSnapshot()
    For index = 0 To 3
        AssertEqual CpuReadRegister(index), 0, "reset general register"
    Next index
    AssertEqual state.PC, 0, "reset PC"
    AssertEqual state.MAR_V, 0, "reset MAR_V"
    AssertEqual state.MAR_P, 0, "reset MAR_P"
    AssertEqual state.MDR, 0, "reset MDR"
    AssertEqual state.IR, 0, "reset IR"
    AssertEqual state.flags, 0, "reset FLAGS"
End Sub

Private Function ExpectedError(ByVal scenario As Long) As Long
    Dim result As Byte
    On Error GoTo Caught
    Select Case scenario
        Case 1: CpuWriteRegister 4, 0
        Case 2: CpuWriteRegister LcaAX, -1
        Case 3: CpuWriteRegister LcaAX, 256
        Case 4: CpuWriteControl LcaPC, -1
        Case 5: CpuWriteControl LcaMAR_V, 512
        Case 6: CpuWriteControl LcaMAR_P, 256
        Case 7: CpuWriteControl 99, 0
        Case 8: result = AluExecute(LcaADD, LcaAX, 4)
        Case 9: result = AluExecute(999, LcaAX)
        Case 10: result = AluExecute(LcaINC, -1)
    End Select
    Exit Function
Caught:
    ExpectedError = Err.Number
End Function

Private Sub CheckAllArithmetic()
    Dim leftValue As Long
    Dim rightValue As Long
    Dim operation As Long
    Dim index As Long
    Dim actual As Byte
    For leftValue = 0 To 255
        For rightValue = 0 To 255
            For operation = LcaADD To LcaCMP
                CheckCase operation, leftValue, rightValue
            Next operation
        Next rightValue
        CheckCase LcaINC, leftValue, 1
        CheckCase LcaDEC, leftValue, 1
    Next leftValue
    ' Destination/source alias and all four register codes.
    For index = LcaAX To LcaDX
        CpuWriteRegister index, 128
        actual = AluExecute(LcaADD, index, index)
        AssertEqual actual, 0, "same-register ADD"
        AssertEqual CpuReadRegister(index), 0, "destination mapping"
        AssertEqual CpuSnapshotFlags(), 3, "same-register flags"
    Next index
End Sub

Private Function CpuSnapshotFlags() As Long
    Dim state As LcaCpuState
    state = CpuSnapshot()
    CpuSnapshotFlags = state.flags
End Function

Private Sub CheckCase(ByVal operation As Long, ByVal leftValue As Long, ByVal rightValue As Long)
    Dim raw As Long
    Dim expected As Long
    Dim flags As Long
    Dim actual As Byte
    Dim state As LcaCpuState
    CpuWriteRegister LcaAX, leftValue
    CpuWriteRegister LcaBX, rightValue
    CpuWriteRegister LcaCX, 127
    CpuWriteRegister LcaDX, 128
    If operation = LcaADD Or operation = LcaINC Then
        raw = leftValue + rightValue
    Else
        raw = leftValue - rightValue
    End If
    ' Independent oracle: normalize by addition/subtraction, not bit masking.
    expected = raw
    If expected < 0 Then expected = expected + 256
    If expected > 255 Then expected = expected - 256
    If expected = 0 Then flags = flags + 1
    If raw < 0 Or raw > 255 Then flags = flags + 2
    If expected >= 128 Then flags = flags + 4
    actual = AluExecute(operation, LcaAX, LcaBX)
    mCases = mCases + 1
    AssertEqual actual, expected, "arithmetic return"
    state = CpuSnapshot()
    If operation = LcaCMP Then
        AssertEqual state.AX, leftValue, "CMP preserves destination"
    Else
        AssertEqual state.AX, expected, "arithmetic destination"
    End If
    AssertEqual state.flags, flags, "ZF/CF/SF"
    AssertEqual state.BX, rightValue, "source preserved"
    AssertEqual state.CX, 127, "CX preserved"
    AssertEqual state.DX, 128, "DX preserved"
    AssertEqual state.PC, 0, "ALU does not advance PC"
End Sub
