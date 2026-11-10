Attribute VB_Name = "LcaPagerPanel"
Option Explicit

Private Const ROW_BASE As Long = LCA_BASE_PAGER
Private mDemoStep As Long
Private mStatus As String

Public Sub ConstruirPanelSwap()
    Dim ws As Worksheet, i As Long
    Set ws = PagerSheet(True)
    LcaPrepareBand ws, ROW_BASE, 35, "P08_"
    ws.Rows(RA("1:35")).RowHeight = 23
    ws.Range(RA("A1:S35")).Interior.Color = RGB(15, 23, 42)
    With ws.Range(RA("A1:S35")).Font
        .Name = "Consolas"
        .Color = RGB(226, 232, 240)
        .Size = 11
    End With
    ws.Range(RA("B2:S3")).Merge
    With ws.Range(RA("B2"))
        .Value = "P08  /  PAGE FAULT + CLOCK + SWAP"
        .Font.Name = "Segoe UI"
        .Font.Size = 21
        .Font.Bold = True
        .Font.Color = RGB(125, 211, 252)
    End With
    ws.Range(RA("B4:S4")).Merge
    ws.Range(RA("B4")).Value = "15 paginas x 32 bytes de respaldo / 7 marcos reemplazables / PFN 7 reservado E/S"
    AddPagerButton ws, "B6:E7", "REINICIAR", "PagerDemoReset"
    AddPagerButton ws, "F6:I7", "SIGUIENTE", "PagerDemoNext"
    AddPagerButton ws, "J6:M7", "EJECUTAR 10", "PagerDemoRun"
    ws.Range(RA("B9:S10")).Merge
    ws.Range(RA("B9:S10")).WrapText = True
    ws.Range(RA("B12")).Value = "VPN"
    ws.Range(RA("C12")).Value = "P"
    ws.Range(RA("D12")).Value = "PFN"
    ws.Range(RA("E12")).Value = "DIRTY"
    ws.Range(RA("F12")).Value = "REF"
    ws.Range(RA("G12")).Value = "BACKING[0]"
    ws.Range(RA("H12")).Value = "ESTADO"
    ws.Range(RA("B12:H12")).Interior.Color = RGB(30, 64, 105)
    ws.Range(RA("J12:S12")).Merge
    ws.Range(RA("J12")).Value = "PAGE_FAULT / CLOCK / ULTIMOS EVENTOS"
    ws.Range(RA("J12:S12")).Interior.Color = RGB(30, 64, 105)
    For i = 13 To 28
        ws.Range(RA("J" & CStr(i) & ":S" & CStr(i))).Merge
    Next i
    ws.Range(RA("B30:S30")).Merge
    ws.Range(RA("B30")).Value = "El paso 8 reemplaza una pagina limpia; el paso 9 preserva la pagina dirty; el paso 10 recupera A5h."
    PagerDemoReset
    LcaGoto ws, ROW_BASE
End Sub

Public Sub MostrarPanelSwap()
    Dim ws As Worksheet
    Set ws = PagerSheet(False)
    If ws Is Nothing Then
        ConstruirPanelSwap
    Else
        LcaGoto ws, ROW_BASE
        PagerPanelRefresh
    End If
End Sub

' Keep the P07 demo and test isolated; this is P08's separate reproducible image.
Public Sub PagerDemoReset()
    Dim vpn As Long, offset As Long
    On Error GoTo failed
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
    mDemoStep = 0
    mStatus = "Listo: respaldo cargado; pulsa SIGUIENTE diez veces para observar el reemplazo y el dato dirty."
    PagerPanelRefresh
    Exit Sub
failed:
    mStatus = "ERROR RESET: " & Err.Description
    PagerPanelRefresh
End Sub

Public Sub PagerDemoNext()
    Dim vpn As Long, value As Byte
    On Error GoTo failed
    If mDemoStep >= 10 Then
        mStatus = "Fin de la demostracion; pulsa REINICIAR para repetir."
        PagerPanelRefresh
        Exit Sub
    End If
    Select Case mDemoStep
        Case 0, 8
            PagerFetchFour 0
            mStatus = "FETCH VA=0; IR=10h, PC=0."
        Case 1
            PagerStoreByte 256, LcaAX
            mStatus = "ST VA=256: AX=A5h; VPN8 dirty=1."
        Case 2 To 7
            vpn = mDemoStep + 7
            value = PagerReadByte(vpn * 32)
            mStatus = "READ VPN=" & CStr(vpn) & ": dato=" & CStr(value) & "."
        Case 9
            value = PagerReadByte(256)
            mStatus = "READ VPN8 recuperada: " & Hex2(value) & " (esperado A5h)."
    End Select
    mDemoStep = mDemoStep + 1
    PagerPanelRefresh
    Exit Sub
failed:
    mStatus = "ERROR paso " & CStr(mDemoStep + 1) & ": " & Err.Description
    PagerPanelRefresh
End Sub

Public Sub PagerDemoRun()
    Dim i As Long
    PagerDemoReset
    For i = 1 To 10
        PagerDemoNext
        If mDemoStep <> i Then Exit For
    Next i
End Sub

Public Sub PagerPanelRefresh()
    Dim ws As Worksheet, vpn As Long, row As Long, entry As LcaPageEntry
    Dim i As Long, index As Long, first As Long
    Set ws = PagerSheet(False)
    If ws Is Nothing Then Exit Sub
    For vpn = 0 To 15
        row = 13 + vpn
        entry = MmuPage(vpn)
        ws.Cells(ROW_BASE + row, 2).Value = vpn
        ws.Cells(ROW_BASE + row, 3).Value = Abs(CLng(entry.Present))
        ws.Cells(ROW_BASE + row, 4).Value = IIf(entry.Present, CStr(entry.PFN), "-")
        ws.Cells(ROW_BASE + row, 5).Value = Abs(CLng(entry.Dirty))
        ws.Cells(ROW_BASE + row, 6).Value = Abs(CLng(entry.Referenced))
        If vpn = 15 Then
            ws.Cells(ROW_BASE + row, 7).Value = "-"
            ws.Cells(ROW_BASE + row, 8).Value = "E/S RESERVADA"
        Else
            ws.Cells(ROW_BASE + row, 7).Value = Hex2(SwapReadByte(vpn, 0))
            ws.Cells(ROW_BASE + row, 8).Value = IIf(entry.Present, "RESIDENTE", "SWAP")
        End If
        ws.Range(ws.Cells(ROW_BASE + row, 2), ws.Cells(ROW_BASE + row, 8)).Interior.Color = _
            IIf(vpn = PagerLastFaultVPN(), RGB(153, 78, 24), _
            IIf(vpn = 15, RGB(90, 45, 57), IIf(entry.Present, RGB(30, 65, 84), RGB(30, 41, 59))))
    Next vpn
    ws.Range(RA("B9")).Value = mStatus & " Paso " & CStr(mDemoStep) & "/10; faults=" & _
        CStr(PagerFaultCount()) & "; evicciones=" & CStr(PagerEvictionCount()) & _
        "; writebacks=" & CStr(PagerWritebackCount()) & "; cursor CLOCK=" & CStr(PagerHand())
    For i = 13 To 28
        ws.Range(RA("J" & CStr(i))).Value = vbNullString
        ws.Range(RA("J" & CStr(i) & ":S" & CStr(i))).Interior.Color = RGB(30, 41, 59)
    Next i
    first = PagerTraceCount() - 16
    If first < 0 Then first = 0
    For index = first To PagerTraceCount() - 1
        row = 13 + index - first
        ws.Range(RA("J" & CStr(row))).Value = PagerTraceLine(index)
        If InStr(1, PagerTraceLine(index), "PAGE_FAULT", vbTextCompare) > 0 Then _
            ws.Range(RA("J" & CStr(row) & ":S" & CStr(row))).Interior.Color = RGB(153, 78, 24)
    Next index
End Sub

Private Function PagerSheet(ByVal create As Boolean) As Worksheet
    Set PagerSheet = LcaBand("P08_PagerDemoReset", create)
End Function

' Row shift of this panel inside the shared LCA sheet.
Private Function RA(ByVal address As String) As String
    RA = LcaRows(address, ROW_BASE)
End Function

Private Sub AddPagerButton(ByVal ws As Worksheet, ByVal address As String, _
                           ByVal caption As String, ByVal macroName As String)
    Dim area As Range, button As Shape
    Set area = ws.Range(RA(address))
    Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left, area.Top, area.Width - 4, area.Height - 2)
    button.Name = "P08_" & macroName
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

Private Function Hex2(ByVal value As Long) As String
    Hex2 = Right$("00" & Hex$(value), 2) & "h"
End Function
