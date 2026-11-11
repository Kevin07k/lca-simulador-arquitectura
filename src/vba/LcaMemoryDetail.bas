Attribute VB_Name = "LcaMemoryDetail"
Option Explicit

' Graphical part of the P08 block: the CLOCK replacement algorithm drawn as a clock (7 frames on a
' circle, R/D bits, the hand on the next candidate) and a map of the 256 physical RAM bytes
' (8 frames x 32 bytes) coloured by what each frame holds. Refreshed by PagerPanelRefresh and P12.
Private Const ROW_BASE As Long = LCA_BASE_PAGER
Private Const COL_BASE As Long = LCA_COL_PAGER
Private Const MAP_TITLE_ROW As Long = 33
Private Const MAP_ROW As Long = 35
Private Const LAST_ROW As Long = 70
Private Const PI As Double = 3.14159265358979
Private Const BG As Long = 2758415           ' RGB(15, 23, 42)
Private Const CARD As Long = 3877150         ' RGB(30, 41, 59)
Private Const HEAD As Long = 6897694         ' RGB(30, 64, 105)
Private Const INK As Long = 15788258         ' RGB(226, 232, 240)
Private Const MUTED As Long = 12100500       ' RGB(148, 163, 184)
Private Const ACCENT As Long = 16569213      ' RGB(125, 211, 252)
Private Const LAMP_ON As Long = 2477306      ' RGB(250, 204, 37)
Private Const EDGE As Long = 5587251         ' RGB(51, 65, 85)
Private Const WARN As Long = 1592985         ' RGB(153, 78, 24)
Private Const CODE_FILL As Long = 9116461    ' RGB(45, 27, 139)
Private Const DATA_FILL As Long = 7954958    ' RGB(14, 98, 121)
Private Const IO_FILL As Long = 3747162      ' RGB(90, 45, 57)
Private Const REF_FILL As Long = 8508206     ' RGB(46, 211, 129)
Private Const KBD_FILL As Long = 10040319    ' RGB(255, 51, 153)
Private Const LCD_FILL As Long = 4441661     ' RGB(61, 197, 67)

Public Sub ConstruirMemoriaDetalle()
    Dim ws As Worksheet, area As Range, cx As Double, cy As Double, radius As Double, i As Long
    Dim a As Double, fx As Double, fy As Double, s As Shape, p As Long, o As Long
    Dim left0 As Double, top0 As Double, cellW As Double, labelW As Double
    Set ws = LcaSheet(True)
    ClearArea ws

    ' CLOCK in J12:S28 of the P08 block.
    ws.Range(RA("J12:S12")).Merge
    With ws.Range(RA("J12"))
        .Value = "RELOJ CLOCK  |  la mano apunta al siguiente candidato; R=1 da una segunda oportunidad"
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Color = ACCENT
        .Interior.Color = BG
    End With
    Set area = ws.Range(RA("J13:S29"))
    cx = area.Left + area.Width / 2
    cy = area.Top + area.Height / 2
    radius = area.Height / 2 - 34
    If radius > area.Width / 2 - 70 Then radius = area.Width / 2 - 70
    Set s = Box(ws, "clockface", cx - radius - 46, cy - radius - 30, 2 * radius + 92, 2 * radius + 60, BG, "", 8, False, msoShapeOval)
    s.Line.ForeColor.RGB = EDGE
    s.Line.Weight = 1.5
    s.Line.DashStyle = msoLineDash
    For i = 0 To 6
        a = -PI / 2 + i * 2 * PI / 7
        fx = cx + radius * Cos(a)
        fy = cy + radius * Sin(a)
        Set s = ws.Shapes.AddConnector(msoConnectorStraight, cx, cy, cx + (radius - 34) * Cos(a), cy + (radius - 34) * Sin(a))
        s.Name = "P08D_hand" & CStr(i)
        s.Line.ForeColor.RGB = LAMP_ON
        s.Line.Weight = 5
        s.Line.EndArrowheadStyle = msoArrowheadTriangle
        s.Visible = msoFalse
        Box ws, "fr" & CStr(i), fx - 50, fy - 30, 100, 60, CARD, "PFN " & CStr(i), 9, True
    Next i
    Box ws, "hub", cx - 42, cy - 22, 84, 44, HEAD, "MANO", 10, True, msoShapeOval

    ' Physical memory map.
    ws.Range(RA("B" & CStr(MAP_TITLE_ROW) & ":S" & CStr(MAP_TITLE_ROW))).Merge
    With ws.Range(RA("B" & CStr(MAP_TITLE_ROW)))
        .Value = "MAPA DE LA RAM FISICA  |  256 bytes = 8 marcos x 32 bytes (valores en hex; byte en 00h = color apagado)"
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = ACCENT
        .Borders(xlEdgeBottom).Color = HEAD
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    left0 = CL(ws, ROW_BASE + MAP_ROW, 2).Left
    top0 = CL(ws, ROW_BASE + MAP_ROW, 2).Top
    labelW = 118
    cellW = (CL(ws, ROW_BASE + MAP_ROW, 20).Left - left0 - labelW) / 32
    For p = 0 To 7
        Box ws, "lbl" & CStr(p), left0, top0 + p * 34, labelW - 6, 30, CARD, "PFN " & CStr(p), 9, True
        For o = 0 To 31
            Set s = Box(ws, "m" & CStr(p) & "_" & CStr(o), left0 + labelW + o * cellW, top0 + p * 34, cellW - 2, 30, CARD, "00", 7, False, msoShapeRectangle)
            s.Line.Visible = msoFalse
            s.TextFrame2.TextRange.Font.Name = "Consolas"
        Next o
    Next p
    ' Legend under the map.
    top0 = top0 + 8 * 34 + 8
    LegendChip ws, "lgc", left0, top0, CODE_FILL, "codigo (VPN 0-7)"
    LegendChip ws, "lgd", left0 + 170, top0, DATA_FILL, "datos (VPN 8-14)"
    LegendChip ws, "lgf", left0 + 340, top0, CARD, "marco libre"
    LegendChip ws, "lgk", left0 + 510, top0, KBD_FILL, "cola teclado E0-E7h"
    LegendChip ws, "lgl", left0 + 680, top0, LCD_FILL, "video LCD E8-EFh"
    LegendChip ws, "lgr", left0 + 850, top0, IO_FILL, "reserva E/S F0-FFh"
    MemoriaDetalleActualizar
End Sub

Public Sub MemoriaDetalleActualizar()
    Dim ws As Worksheet, i As Long, owner As Long, e As LcaPageEntry, fill As Long, text As String
    Dim p As Long, o As Long, v As Long, kind As Long, prev As Boolean
    Set ws = LcaBand("P08D_hub", False)
    If ws Is Nothing Then Exit Sub
    prev = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo done
    For i = 0 To 6
        owner = PagerFrameOwner(i)
        With ws.Shapes("P08D_fr" & CStr(i))
            If owner < 0 Then
                text = "PFN " & CStr(i) & vbCr & "libre"
                fill = CARD
                .Line.ForeColor.RGB = EDGE
                .Line.Weight = 1
            Else
                e = MmuPage(owner)
                text = "PFN " & CStr(i) & "  VPN " & CStr(owner) & vbCr & "R=" & Abs(CLng(e.Referenced)) & "   D=" & Abs(CLng(e.Dirty))
                fill = IIf(e.Referenced, REF_FILL, IIf(owner <= 7, CODE_FILL, DATA_FILL))
                .Line.ForeColor.RGB = IIf(e.Dirty, WARN, EDGE)
                .Line.Weight = IIf(e.Dirty, 4, 1)
            End If
            .TextFrame2.TextRange.Text = text
            .Fill.ForeColor.RGB = fill
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = IIf(fill = REF_FILL, BG, INK)
        End With
        ws.Shapes("P08D_hand" & CStr(i)).Visible = IIf(i = PagerHand(), msoTrue, msoFalse)
    Next i
    ws.Shapes("P08D_hub").TextFrame2.TextRange.Text = "MANO" & vbCr & "PFN " & CStr(PagerHand())

    For p = 0 To 7
        If p = 7 Then
            ws.Shapes("P08D_lbl7").TextFrame2.TextRange.Text = "PFN 7  E/S"
            ws.Shapes("P08D_lbl7").Fill.ForeColor.RGB = IO_FILL
        Else
            owner = PagerFrameOwner(p)
            If owner < 0 Then
                ws.Shapes("P08D_lbl" & CStr(p)).TextFrame2.TextRange.Text = "PFN " & CStr(p) & "  libre"
                ws.Shapes("P08D_lbl" & CStr(p)).Fill.ForeColor.RGB = CARD
            Else
                ws.Shapes("P08D_lbl" & CStr(p)).TextFrame2.TextRange.Text = "PFN " & CStr(p) & "  VPN " & CStr(owner)
                ws.Shapes("P08D_lbl" & CStr(p)).Fill.ForeColor.RGB = IIf(owner <= 7, CODE_FILL, DATA_FILL)
            End If
        End If
        For o = 0 To 31
            v = RamReadPhysical(p * 32 + o)
            If p = 7 Then
                If o < 8 Then
                    kind = KBD_FILL
                ElseIf o < 16 Then
                    kind = LCD_FILL
                Else
                    kind = IO_FILL
                End If
            ElseIf owner < 0 Then
                kind = CARD
            Else
                kind = IIf(owner <= 7, CODE_FILL, DATA_FILL)
            End If
            With ws.Shapes("P08D_m" & CStr(p) & "_" & CStr(o))
                .TextFrame2.TextRange.Text = Right$("0" & Hex$(v), 2)
                If v = 0 Or (p = 7 And o >= 8 And o < 16 And v = &H20) Or kind = CARD Then
                    .Fill.ForeColor.RGB = CARD
                    .Fill.Transparency = 0
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = IIf(kind = CARD, EDGE, kind)
                Else
                    .Fill.ForeColor.RGB = kind
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                End If
            End With
        Next o
    Next p
done:
    Application.ScreenUpdating = prev
End Sub

Private Sub LegendChip(ByVal ws As Worksheet, ByVal name As String, ByVal left As Double, ByVal top As Double, _
                       ByVal fill As Long, ByVal text As String)
    Dim s As Shape
    Set s = Box(ws, name, left, top, 160, 22, fill, text, 9, True)
    If fill = CARD Then s.Line.ForeColor.RGB = MUTED
End Sub

Private Sub ClearArea(ByVal ws As Worksheet)
    Dim i As Long, band As Range
    Set band = ws.Range(RA("A" & CStr(MAP_TITLE_ROW - 1) & ":T" & CStr(LAST_ROW)))
    band.UnMerge
    band.Clear
    band.Interior.Color = BG
    ws.Range(RA("J12:S29")).UnMerge
    ws.Range(RA("J12:S29")).Clear
    ws.Range(RA("J12:S29")).Interior.Color = BG
    LcaRowHeight ws.Rows(LcaShift(CStr(MAP_TITLE_ROW - 1) & ":" & CStr(LAST_ROW), ROW_BASE, COL_BASE)), 20
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 5) = "P08D_" Then ws.Shapes(i).Delete
    Next i
End Sub

Private Function Box(ByVal ws As Worksheet, ByVal name As String, ByVal left As Double, ByVal top As Double, _
                     ByVal w As Double, ByVal h As Double, ByVal fill As Long, ByVal text As String, _
                     ByVal size As Single, Optional ByVal bold As Boolean = False, _
                     Optional ByVal kind As Long = msoShapeRoundedRectangle) As Shape
    Dim s As Shape
    Set s = ws.Shapes.AddShape(kind, left, top, w, h)
    s.Name = "P08D_" & name
    s.Fill.ForeColor.RGB = fill
    s.Line.ForeColor.RGB = EDGE
    With s.TextFrame2
        .MarginLeft = 0
        .MarginRight = 0
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

Private Function RA(ByVal address As String) As String
    RA = LcaShift(address, ROW_BASE, COL_BASE)
End Function

Private Function CL(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Range
    Set CL = ws.Cells(r, c + COL_BASE)
End Function
