Attribute VB_Name = "LcaBusDetail"
Option Explicit

' Graphical part of the P05/P06 block: the two tri-state buses drawn live. Every device hangs from the
' bus through a tri-state buffer; the driver lights green, the receiver amber, the transferred byte rides
' a bubble on the bus and the eight control lines of each bus are lamps. Two drivers = contention (red).
Private Const ROW_BASE As Long = LCA_BASE_BUS
Private Const COL_BASE As Long = LCA_COL_BUS
Private Const FIRST_ROW As Long = 35
Private Const LAST_ROW As Long = 72
Private Const DESIGN_W As Double = 1380
Private Const BG As Long = 2758415           ' RGB(15, 23, 42)
Private Const CARD As Long = 3877150         ' RGB(30, 41, 59)
Private Const HEAD As Long = 6897694         ' RGB(30, 64, 105)
Private Const INK As Long = 15788258         ' RGB(226, 232, 240)
Private Const MUTED As Long = 12100500       ' RGB(148, 163, 184)
Private Const ACCENT As Long = 16569213      ' RGB(125, 211, 252)
Private Const LAMP_ON As Long = 2477306      ' RGB(250, 204, 37)
Private Const EDGE As Long = 5587251         ' RGB(51, 65, 85)
Private Const DRIVE As Long = 8508206        ' RGB(46, 211, 129)
Private Const BAD As Long = 3355647         ' RGB(255, 51, 51)
Private Const PINK As Long = 10040319        ' RGB(255, 51, 153)

Private mX0 As Double
Private mY0 As Double
Private mSX As Double

Public Sub ConstruirBusDetalle()
    Dim ws As Worksheet, devices As Variant, i As Long, x As Double, labels As Variant, b As Long
    Set ws = LcaSheet(True)
    ClearArea ws
    ws.Range(RA("B" & CStr(FIRST_ROW) & ":S" & CStr(FIRST_ROW))).Merge
    With ws.Range(RA("B" & CStr(FIRST_ROW)))
        .Value = "DIAGRAMA DE BUSES TRI-ESTADO EN VIVO  |  verde = quien maneja el bus (Enable)   ambar = quien recibe (Load)   rojo = contencion"
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = ACCENT
        .Borders(xlEdgeBottom).Color = HEAD
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    mX0 = CL(ws, ROW_BASE + FIRST_ROW + 2, 2).Left
    mY0 = CL(ws, ROW_BASE + FIRST_ROW + 2, 2).Top
    mSX = (CL(ws, ROW_BASE + FIRST_ROW + 2, 20).Left - mX0) / DESIGN_W

    ' Devices on top, each with a tri-state buffer down to the external bus.
    devices = Array(Array("cpu", "CPU" & vbCr & "AX BX CX DX" & vbCr & "MDR  IR  MAR"), Array("cache", "CACHE" & vbCr & "L1I  L1D  L2"), _
                    Array("ram", "RAM" & vbCr & "8 marcos x 32 B"), Array("kbd", "TECLADO" & vbCr & "cola E0-E7h"), _
                    Array("lcd", "LCD 8x1" & vbCr & "video E8-EFh"))
    For i = 0 To 4
        x = 20 + i * 272
        Box ws, CStr(devices(i)(0)), x, 0, 220, 110, CARD, CStr(devices(i)(1)), 13, True
        Box ws, "buf_" & CStr(devices(i)(0)), x + 85, 132, 50, 44, EDGE, "", 8, False, msoShapeIsoscelesTriangle
        ws.Shapes("P05D_buf_" & CStr(devices(i)(0))).Rotation = 180
        Box ws, "stub_" & CStr(devices(i)(0)), x + 106, 110, 8, 22, EDGE, "", 6, False, msoShapeRectangle
        Box ws, "drop_" & CStr(devices(i)(0)), x + 106, 176, 8, 30, EDGE, "", 6, False, msoShapeRectangle
    Next i
    Box ws, "bus_ext", 0, 206, 1380, 34, HEAD, "BUS EXTERNO  (direcciones / datos / control)", 11, True, msoShapeRectangle
    Box ws, "bub_ext", 640, 196, 110, 54, PINK, "--", 16, True, msoShapeOval
    ' Internal bus inside the CPU side.
    Box ws, "bus_int", 0, 300, 1380, 34, HEAD, "BUS INTERNO  (registros <-> MDR <-> IR)", 11, True, msoShapeRectangle
    Box ws, "bub_int", 640, 290, 110, 54, PINK, "--", 16, True, msoShapeOval
    Box ws, "link_int", 126, 240, 8, 60, EDGE, "", 6, False, msoShapeRectangle

    ' Signal lamps of both buses + the formatted cycle.
    labels = Array("R", "W", "C", "IO", "EN", "LD", "KBD", "LCD")
    For b = 0 To 7
        Box ws, "le" & CStr(b), 20 + b * 52, 360, 44, 44, CARD, CStr(labels(b)), 9, True, msoShapeOval
        Box ws, "li" & CStr(b), 720 + b * 52, 360, 44, 44, CARD, CStr(labels(b)), 9, True, msoShapeOval
    Next b
    Box ws, "tag_e", 20, 410, 420, 22, BG, "senales del bus EXTERNO", 9, True
    Box ws, "tag_i", 720, 410, 420, 22, BG, "senales del bus INTERNO", 9, True
    Box ws, "cyc_e", 0, 440, 680, 64, CARD, "", 9, False, msoShapeRectangle
    Box ws, "cyc_i", 700, 440, 680, 64, CARD, "", 9, False, msoShapeRectangle
    ws.Shapes("P05D_bub_ext").Visible = msoFalse
    ws.Shapes("P05D_bub_int").Visible = msoFalse
    BusDetalleActualizar
End Sub

' Repaints both buses from the live engine state (called by BusPanelActualizar and by P12).
Public Sub BusDetalleActualizar()
    Dim ws As Worksheet, ext As LcaBusCycle, inn As LcaBusCycle
    Set ws = LcaBand("P05D_bus_ext", False)
    If ws Is Nothing Then Exit Sub
    mX0 = CL(ws, ROW_BASE + FIRST_ROW + 2, 2).Left
    mSX = (CL(ws, ROW_BASE + FIRST_ROW + 2, 20).Left - mX0) / DESIGN_W
    ext = BusSnapshot(LCA_BUS_EXTERNAL)
    inn = BusSnapshot(LCA_BUS_INTERNAL)
    ResetDevices ws
    PaintBus ws, "ext", "le", ext, True
    PaintBus ws, "int", "li", inn, False
End Sub

Private Sub ResetDevices(ByVal ws As Worksheet)
    Dim d As Variant
    For Each d In Array("cpu", "cache", "ram", "kbd", "lcd")
        ws.Shapes("P05D_" & CStr(d)).Line.ForeColor.RGB = EDGE
        ws.Shapes("P05D_" & CStr(d)).Line.Weight = 1
        ws.Shapes("P05D_buf_" & CStr(d)).Fill.ForeColor.RGB = EDGE
        ws.Shapes("P05D_drop_" & CStr(d)).Fill.ForeColor.RGB = EDGE
    Next d
End Sub

Private Sub PaintBus(ByVal ws As Worksheet, ByVal busKey As String, ByVal lampKey As String, _
                    ByRef c As LcaBusCycle, ByVal external As Boolean)
    Dim sig As Variant, i As Long, drv As String, rcv As String, active As Boolean, xd As Double, xr As Double
    sig = Array(c.ReadSignal, c.WriteSignal, c.ClockSignal, c.IOSignal, c.EnableSignal, c.LoadSignal, c.EnKBD, c.LdLCD)
    For i = 0 To 7
        With ws.Shapes("P05D_" & lampKey & CStr(i))
            .Fill.ForeColor.RGB = IIf(CBool(sig(i)), LAMP_ON, CARD)
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = IIf(CBool(sig(i)), BG, MUTED)
        End With
    Next i
    active = (c.DriverCount > 0)
    With ws.Shapes("P05D_bus_" & busKey)
        If c.DriverCount > 1 Then
            .Fill.ForeColor.RGB = BAD
        ElseIf active Then
            .Fill.ForeColor.RGB = RGB(30, 90, 150)
        Else
            .Fill.ForeColor.RGB = HEAD
        End If
    End With
    ws.Shapes("P05D_cyc_" & Left$(busKey, 1)).TextFrame2.TextRange.Text = _
        IIf(active, c.Microstep & vbCr & "VA " & AddrText(c.VA, 3) & "   PA " & AddrText(c.PA, 2) & "   " & c.Driver & "  ->  " & c.Receiver, _
            "High-Z: nadie maneja este bus" & vbCr & "(un acierto de cache o un paso interno no usa el bus externo)")
    With ws.Shapes("P05D_bub_" & busKey)
        .Visible = IIf(active And c.DriverCount = 1, msoTrue, msoFalse)
        If active And c.DriverCount = 1 Then
            drv = DeviceOf(c.Driver)
            rcv = DeviceOf(c.Receiver)
            .TextFrame2.TextRange.Text = Right$("0" & Hex$(c.Data), 2) & "h"
            If external Then
                Highlight ws, drv, DRIVE
                Highlight ws, rcv, LAMP_ON
                xd = DeviceCenter(drv)
                xr = DeviceCenter(rcv)
                .Left = mX0 + ((xd + xr) / 2) * mSX - .Width / 2
            Else
                .Left = mX0 + 300 * mSX - .Width / 2
                If drv = "cpu" Or rcv = "cpu" Then ws.Shapes("P05D_cpu").Line.ForeColor.RGB = LAMP_ON
            End If
        End If
    End With
End Sub

Private Sub Highlight(ByVal ws As Worksheet, ByVal device As String, ByVal color As Long)
    With ws.Shapes("P05D_" & device)
        .Line.ForeColor.RGB = color
        .Line.Weight = 4
    End With
    ws.Shapes("P05D_buf_" & device).Fill.ForeColor.RGB = color
    ws.Shapes("P05D_drop_" & device).Fill.ForeColor.RGB = color
End Sub

' Maps an engine driver/receiver name to the drawn device.
Private Function DeviceOf(ByVal name As String) As String
    Dim u As String
    u = UCase$(name)
    If InStr(u, "RAM") > 0 Then
        DeviceOf = "ram"
    ElseIf InStr(u, "KBD") > 0 Then
        DeviceOf = "kbd"
    ElseIf InStr(u, "LCD") > 0 Then
        DeviceOf = "lcd"
    ElseIf InStr(u, "L1") > 0 Or InStr(u, "L2") > 0 Or u = "CACHE" Then
        DeviceOf = "cache"
    Else
        DeviceOf = "cpu"
    End If
End Function

Private Function DeviceCenter(ByVal device As String) As Double
    Dim i As Long
    Select Case device
        Case "cpu": i = 0
        Case "cache": i = 1
        Case "ram": i = 2
        Case "kbd": i = 3
        Case Else: i = 4
    End Select
    DeviceCenter = 20 + i * 272 + 110
End Function

Private Function AddrText(ByVal value As Long, ByVal digits As Long) As String
    If value < 0 Then
        AddrText = "-"
    Else
        AddrText = Right$(String$(digits, "0") & Hex$(value), digits) & "h"
    End If
End Function

Private Sub ClearArea(ByVal ws As Worksheet)
    Dim i As Long, band As Range
    Set band = ws.Range(RA("A" & CStr(FIRST_ROW - 1) & ":T" & CStr(LAST_ROW)))
    band.UnMerge
    band.Clear
    band.Interior.Color = BG
    LcaRowHeight ws.Rows(LcaShift(CStr(FIRST_ROW - 1) & ":" & CStr(LAST_ROW), ROW_BASE, COL_BASE)), 16
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 5) = "P05D_" Then ws.Shapes(i).Delete
    Next i
End Sub

Private Function Box(ByVal ws As Worksheet, ByVal name As String, ByVal x As Double, ByVal y As Double, _
                     ByVal w As Double, ByVal h As Double, ByVal fill As Long, ByVal text As String, _
                     ByVal size As Single, Optional ByVal bold As Boolean = False, _
                     Optional ByVal kind As Long = msoShapeRoundedRectangle) As Shape
    Dim s As Shape
    Set s = ws.Shapes.AddShape(kind, mX0 + x * mSX, mY0 + y, w * mSX, h)
    s.Name = "P05D_" & name
    s.Fill.ForeColor.RGB = fill
    s.Line.ForeColor.RGB = EDGE
    With s.TextFrame2
        .MarginLeft = 2
        .MarginRight = 2
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = msoTrue
        .VerticalAnchor = msoAnchorMiddle
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
