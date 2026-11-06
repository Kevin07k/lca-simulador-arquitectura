Attribute VB_Name = "LcaBusPanel"
Option Explicit

Private Const SHEET_NAME As String = "P05 Buses"
Private mDemoMemory(0 To 255) As Byte     ' Only the isolated panel fixture; P06 owns real RAM.
Private mDemoStep As Long
Private mDemoStarted As Boolean
Private mDemoStatus As String

Public Sub ConstruirPanelBuses()
    Dim ws As Worksheet
    Dim i As Long
    Set ws = BusSheet(True)
    ws.Cells.UnMerge
    ws.Cells.Clear
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B:R").ColumnWidth = 12
    ws.Rows("1:32").RowHeight = 23
    ws.Range("A1:S34").Interior.Color = RGB(15, 23, 42)
    ws.Range("A1:S34").Font.Name = "Consolas"
    ws.Range("A1:S34").Font.Color = RGB(226, 232, 240)
    ws.Range("B2:R3").Merge
    With ws.Range("B2")
        .Value = "P05  /  BUSES + CONTROL"
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 22
        .Font.Color = RGB(125, 211, 252)
    End With
    ws.Range("B4:R4").Merge
    ws.Range("B4").Value = "Ciclos del motor / Engine cycles  |  R W C IO EN LD"
    AddBusButton ws, "B6:F7", "REINICIAR / RESET", "BusDemoReset"
    AddBusButton ws, "G6:L7", "SIGUIENTE / NEXT", "BusDemoNext"
    AddBusButton ws, "M6:R7", "CONTENCION / CONTENTION", "BusDemoContention"
    ws.Range("B9:R9").Merge
    ws.Range("B9").Value = "MICROPALABRA / CONTROL WORD"
    ws.Range("B10:R11").Merge
    ws.Range("B10:R11").Interior.Color = RGB(30, 41, 59)
    ws.Range("B13:R13").Merge
    ws.Range("B13").Value = "BUSES EN ESTE MICRO-PASO / BUSES IN THIS MICROSTEP"
    ws.Range("B14:R15").Merge
    ws.Range("B16:R17").Merge
    ws.Range("B14:R17").Interior.Color = RGB(30, 41, 59)
    ws.Range("B19:R19").Merge
    ws.Range("B19").Value = "TRAZA REAL DEL MOTOR / LIVE ENGINE TRACE (ultimos 10 / last 10)"
    For i = 20 To 29
        ws.Range(ws.Cells(i, 2), ws.Cells(i, 18)).Merge
        ws.Rows(i).RowHeight = 22
    Next i
    ws.Range("B20:R29").Interior.Color = RGB(30, 41, 59)
    ws.Range("B31:R32").Merge
    ws.Range("B31:R32").WrapText = True
    ws.Range("B31").Font.Color = RGB(125, 211, 252)
    BusDemoReset
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
End Sub

Public Sub MostrarPanelBuses()
    Dim ws As Worksheet
    Set ws = BusSheet(False)
    If ws Is Nothing Then
        ConstruirPanelBuses
    Else
        ws.Activate
        BusPanelActualizar
    End If
End Sub

' Use this after cycles triggered by other modules (P06/P12), not a scripted animation.
Public Sub BusPanelActualizar()
    Dim ws As Worksheet
    Dim internalCycle As LcaBusCycle
    Dim externalCycle As LcaBusCycle
    Dim n As Long, i As Long
    Set ws = BusSheet(False)
    If ws Is Nothing Then Exit Sub
    internalCycle = BusSnapshot(LCA_BUS_INTERNAL)
    externalCycle = BusSnapshot(LCA_BUS_EXTERNAL)
    ws.Range("B10").Value = ControlLastWord()
    ws.Range("B14").Value = BusFormat(internalCycle)
    ws.Range("B16").Value = BusFormat(externalCycle)
    n = BusTraceCount()
    For i = 0 To 9
        ws.Cells(20 + i, 2).Value = vbNullString
        If n > i Then ws.Cells(20 + i, 2).Value = BusTraceLine(n - 1 - i)
    Next i
    ws.Range("B31").Value = mDemoStatus
End Sub

Public Sub BusDemoReset()
    Dim i As Long
    CpuReset
    ControlReset
    For i = 0 To 255
        mDemoMemory(i) = 0
    Next i
    mDemoMemory(&H40) = &H10
    mDemoMemory(&H41) = &H0
    mDemoMemory(&H42) = &H2A
    mDemoMemory(&H43) = &H0
    CpuWriteControl LcaPC, &H20
    CpuWriteRegister LcaAX, &H2A
    mDemoStep = 0
    mDemoStarted = True
    mDemoStatus = "Fixture: PC=020h, 040-043h=10 00 2A 00; AX=2Ah. NEXT recorre fetch y ST 14Dh -> 6Dh."
    BusPanelActualizar
End Sub

' Exactly one engine clock edge per click. The fixture commits its RAM write only after arbitration.
Public Sub BusDemoNext()
    Dim i As Long
    On Error GoTo failed
    If Not mDemoStarted Then BusDemoReset
    If mDemoStep = 0 Then ControlStartFetch &H20, &H40
    If mDemoStep < 8 Then
        i = mDemoStep \ 2
        If (mDemoStep Mod 2) = 0 Then
            ControlReadFetchByte mDemoMemory(&H40 + i)
        Else
            ControlLatchFetchByte
        End If
    ElseIf mDemoStep = 8 Then
        ControlPrepareStore &H14D, &H6D, LcaAX
    ElseIf mDemoStep = 9 Then
        mDemoMemory(&H6D) = ControlCommitStore()
    Else
        mDemoStatus = "Fin / Done. 6Dh=" & Right$("00" & Hex$(mDemoMemory(&H6D)), 2) & "h. RESET para repetir."
        BusPanelActualizar
        Exit Sub
    End If
    mDemoStep = mDemoStep + 1
    mDemoStatus = "Paso / step " & CStr(mDemoStep) & "/10 | IR=" & Right$("00" & Hex$(CpuSnapshotIR()), 2) & _
                  "h | 6Dh=" & Right$("00" & Hex$(mDemoMemory(&H6D)), 2) & _
                  "h | PC=020h (P12 decide avance / commits PC)"
    BusPanelActualizar
    Exit Sub
failed:
    mDemoStatus = "ERROR: " & Err.Description
    BusPanelActualizar
End Sub

Public Sub BusDemoContention()
    Dim failure As String
    ControlReset
    mDemoStep = 10
    mDemoStarted = True
    BusBegin LCA_BUS_EXTERNAL, "ST CONTENTION", &H14D, &H6D, False, True, False
    BusDrive LCA_BUS_EXTERNAL, "CPU", &H2A
    BusDrive LCA_BUS_EXTERNAL, "OTHER_OUTPUT", &H55
    BusReceive LCA_BUS_EXTERNAL, "RAM"
    On Error Resume Next
    Call BusClock(LCA_BUS_EXTERNAL)
    failure = Err.Description
    Err.Clear
    On Error GoTo 0
    mDemoStatus = "BUS_CONTENTION: " & failure & "; 6Dh=" & _
                  Right$("00" & Hex$(mDemoMemory(&H6D)), 2) & "h (sin escritura / no write). RESET para repetir."
    BusPanelActualizar
End Sub

Private Function CpuSnapshotIR() As Long
    Dim state As LcaCpuState
    state = CpuSnapshot()
    CpuSnapshotIR = state.IR
End Function

Private Function BusSheet(ByVal create As Boolean) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SHEET_NAME Then
            Set BusSheet = ws
            Exit Function
        End If
    Next ws
    If create Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_NAME
        Set BusSheet = ws
    End If
End Function

Private Sub AddBusButton(ByVal ws As Worksheet, ByVal address As String, _
                         ByVal caption As String, ByVal macroName As String)
    Dim area As Range
    Dim button As Shape
    Set area = ws.Range(address)
    Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left, area.Top, area.Width - 4, area.Height - 2)
    button.Name = "P05_" & macroName
    button.Fill.ForeColor.RGB = RGB(37, 99, 235)
    button.Line.Visible = msoFalse
    button.TextFrame2.TextRange.Text = caption
    button.TextFrame2.TextRange.Font.Name = "Segoe UI"
    button.TextFrame2.TextRange.Font.Size = 12
    button.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    button.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    button.TextFrame2.VerticalAnchor = msoAnchorMiddle
    button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Sub
