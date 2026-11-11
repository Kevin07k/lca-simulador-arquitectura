Attribute VB_Name = "LcaCpuDetail"
Option Explicit

' Detailed, graphical part of the P04 CPU block (rows 46..72 of the block): every register as eight
' bit lamps with their weights, a live column chart of AX..DX and the three flags explained.
Private Const ROW_BASE As Long = LCA_BASE_CPU
Private Const COL_BASE As Long = LCA_COL_CPU
Private Const FIRST_ROW As Long = 45
Private Const LAST_ROW As Long = 72
Private Const DATA_ROW As Long = 64          ' chart data, hidden under the chart
Private Const BG As Long = 2758415           ' RGB(15, 23, 42)
Private Const CARD As Long = 3877150         ' RGB(30, 41, 59)
Private Const HEAD As Long = 6897694         ' RGB(30, 64, 105)
Private Const INK As Long = 15788258         ' RGB(226, 232, 240)
Private Const MUTED As Long = 12100500       ' RGB(148, 163, 184)
Private Const ACCENT As Long = 16569213      ' RGB(125, 211, 252)
Private Const LAMP_ON As Long = 2477306      ' RGB(250, 204, 37)
Private Const EDGE As Long = 5587251         ' RGB(51, 65, 85)

Public Sub ConstruirCpuDetalle()
    Dim ws As Worksheet, r As Long, b As Long, x0 As Double, y0 As Double, s As Shape
    Dim names As Variant, colors As Variant, co As ChartObject, area As Range, i As Long
    Set ws = LcaSheet(True)
    ClearArea ws
    LcaRowHeight ws.Rows(LcaShift(CStr(FIRST_ROW) & ":" & CStr(LAST_ROW), ROW_BASE, COL_BASE)), 20
    With ws.Range(RA("A" & CStr(FIRST_ROW) & ":T" & CStr(LAST_ROW)))
        .Interior.Color = BG
        .Font.Name = "Segoe UI"
        .Font.Color = INK
    End With
    Title ws, "B46:S46", "05   VISTA DETALLADA  |  bits de cada registro, grafica de valores y flags (se actualiza con cada operacion)"

    ' Bit lamps: weight header + one row per register.
    names = Array("AX", "BX", "CX", "DX")
    colors = Array(RGB(56, 189, 248), RGB(52, 211, 153), RGB(251, 146, 60), RGB(232, 121, 249))
    x0 = CL(ws, ROW_BASE + 48, 2).Left
    y0 = CL(ws, ROW_BASE + 48, 2).Top
    For b = 0 To 7
        Set s = Box(ws, "w" & CStr(b), x0 + 110 + b * 52, y0, 46, 20, BG, CStr(2 ^ (7 - b)), 9)
        s.Line.Visible = msoFalse
        s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = MUTED
    Next b
    Set s = Box(ws, "whex", x0 + 110 + 8 * 52 + 8, y0, 120, 20, BG, "HEX   DEC", 9)
    s.Line.Visible = msoFalse
    s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = MUTED
    For r = 0 To 3
        Set s = Box(ws, "name" & CStr(r), x0, y0 + 28 + r * 52, 96, 44, CLng(colors(r)), CStr(names(r)), 16, True)
        s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = BG
        For b = 0 To 7
            Set s = Box(ws, "bit" & CStr(r) & CStr(b), x0 + 110 + b * 52 + 3, y0 + 28 + r * 52 + 2, 40, 40, CARD, "0", 14, True, msoShapeOval)
            s.Line.ForeColor.RGB = CLng(colors(r))
            s.Line.Weight = 2
        Next b
        Box ws, "val" & CStr(r), x0 + 110 + 8 * 52 + 8, y0 + 28 + r * 52, 120, 44, CARD, "00h   0", 13, True
    Next r

    ' Flags explained.
    Set area = CL(ws, ROW_BASE + 62, 2)
    Box ws, "fZ", area.Left, area.Top, 196, 52, CARD, "ZF  zero" & vbCr & "resultado = 0", 11, True
    Box ws, "fC", area.Left + 206, area.Top, 196, 52, CARD, "CF  carry" & vbCr & "acarreo / prestamo", 11, True
    Box ws, "fS", area.Left + 412, area.Top, 196, 52, CARD, "SF  sign" & vbCr & "bit 7 del resultado", 11, True

    ' Column chart of the four registers (data cells sit under the chart).
    For i = 0 To 3
        CL(ws, ROW_BASE + DATA_ROW, 13 + i).Value = names(i)
        CL(ws, ROW_BASE + DATA_ROW + 1, 13 + i).Value = 0
    Next i
    Set area = ws.Range(RA("L48:S70"))
    Set co = ws.ChartObjects.Add(area.Left + 6, area.Top, area.Width - 12, area.Height)
    co.Name = "P04D_chart"
    co.Placement = xlMove
    With co.Chart
        .ChartType = xlColumnClustered
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        With .SeriesCollection.NewSeries
            .Name = "valor"
            .Values = ws.Range(CL(ws, ROW_BASE + DATA_ROW + 1, 13), CL(ws, ROW_BASE + DATA_ROW + 1, 16))
            .XValues = ws.Range(CL(ws, ROW_BASE + DATA_ROW, 13), CL(ws, ROW_BASE + DATA_ROW, 16))
            .HasDataLabels = True
            .DataLabels.Font.Color = INK
            .DataLabels.Font.Size = 12
            .DataLabels.Font.Bold = True
            For i = 1 To 4
                .Points(i).Format.Fill.ForeColor.RGB = CLng(colors(i - 1))
            Next i
        End With
        .ChartGroups(1).GapWidth = 60
        .HasLegend = False
        .HasTitle = True
        .ChartTitle.Text = "REGISTROS EN VIVO (0 - 255)"
        .ChartTitle.Font.Color = ACCENT
        .ChartTitle.Font.Size = 13
        .ChartTitle.Font.Bold = True
        .ChartArea.Format.Fill.ForeColor.RGB = CARD
        .ChartArea.Format.Line.ForeColor.RGB = ACCENT
        .PlotArea.Format.Fill.ForeColor.RGB = CARD
        With .Axes(xlValue)
            .MinimumScale = 0
            .MaximumScale = 255
            .MajorUnit = 64
            .TickLabels.Font.Color = MUTED
            .MajorGridlines.Format.Line.ForeColor.RGB = EDGE
        End With
        .Axes(xlCategory).TickLabels.Font.Color = INK
        .Axes(xlCategory).TickLabels.Font.Size = 12
        .Axes(xlCategory).TickLabels.Font.Bold = True
    End With
    CpuDetalleActualizar
End Sub

' Called by PanelActualizar: lamps, values, flags and chart data.
Public Sub CpuDetalleActualizar()
    Dim ws As Worksheet, r As Long, b As Long, v As Long, isOn As Boolean, colors As Variant
    Set ws = LcaBand("P04D_name0", False)
    If ws Is Nothing Then Exit Sub
    colors = Array(RGB(56, 189, 248), RGB(52, 211, 153), RGB(251, 146, 60), RGB(232, 121, 249))
    For r = 0 To 3
        v = CpuReadRegister(r)
        For b = 0 To 7
            isOn = ((v And (2 ^ (7 - b))) <> 0)
            With ws.Shapes("P04D_bit" & CStr(r) & CStr(b))
                .Fill.ForeColor.RGB = IIf(isOn, CLng(colors(r)), CARD)
                .TextFrame2.TextRange.Text = IIf(isOn, "1", "0")
                .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = IIf(isOn, BG, MUTED)
            End With
        Next b
        ws.Shapes("P04D_val" & CStr(r)).TextFrame2.TextRange.Text = AsmHex2(v) & "   " & CStr(v)
        CL(ws, ROW_BASE + DATA_ROW + 1, 13 + r).Value = v
    Next r
    FlagBox ws, "fZ", CpuFlagIsSet(LCA_FLAG_ZF)
    FlagBox ws, "fC", CpuFlagIsSet(LCA_FLAG_CF)
    FlagBox ws, "fS", CpuFlagIsSet(LCA_FLAG_SF)
End Sub

Private Sub FlagBox(ByVal ws As Worksheet, ByVal name As String, ByVal isOn As Boolean)
    With ws.Shapes("P04D_" & name)
        .Fill.ForeColor.RGB = IIf(isOn, LAMP_ON, CARD)
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = IIf(isOn, BG, INK)
    End With
End Sub

Private Sub ClearArea(ByVal ws As Worksheet)
    Dim i As Long, band As Range
    Set band = ws.Range(RA("A" & CStr(FIRST_ROW) & ":T" & CStr(LAST_ROW)))
    band.UnMerge
    band.Clear
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 5) = "P04D_" Then ws.Shapes(i).Delete
    Next i
End Sub

Private Function Box(ByVal ws As Worksheet, ByVal name As String, ByVal left As Double, ByVal top As Double, _
                     ByVal w As Double, ByVal h As Double, ByVal fill As Long, ByVal text As String, _
                     ByVal size As Single, Optional ByVal bold As Boolean = False, _
                     Optional ByVal kind As Long = msoShapeRoundedRectangle) As Shape
    Dim s As Shape
    Set s = ws.Shapes.AddShape(kind, left, top, w, h)
    s.Name = "P04D_" & name
    s.Fill.ForeColor.RGB = fill
    s.Line.ForeColor.RGB = EDGE
    With s.TextFrame2
        .MarginLeft = 1
        .MarginRight = 1
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoTrue
        .TextRange.Text = text
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = size
        .TextRange.Font.Bold = IIf(bold, msoTrue, msoFalse)
        .TextRange.Font.Fill.ForeColor.RGB = INK
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    s.Placement = xlMove
    Set Box = s
End Function

Private Sub Title(ByVal ws As Worksheet, ByVal address As String, ByVal text As String)
    With ws.Range(RA(address))
        .Merge
        .Value = text
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = ACCENT
        .Borders(xlEdgeBottom).Color = HEAD
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
End Sub

Private Function RA(ByVal address As String) As String
    RA = LcaShift(address, ROW_BASE, COL_BASE)
End Function

Private Function CL(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Range
    Set CL = ws.Cells(r, c + COL_BASE)
End Function
