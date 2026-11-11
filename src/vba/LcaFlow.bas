Attribute VB_Name = "LcaFlow"
Option Explicit

' Shared "data in transit" animation. A packet (a rounded bubble with its value) travels between
' anchors: a Range (its centre), a shape name, or an Array(x, y) in points. Used by the P15 arithmetic
' demos and by the P08, P09, P13 and P14 blocks so every part of the simulator shows data moving the
' same way. One speed for everything: 0 = off, 1 = slow, 2 = normal, 3 = fast.
#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const PACKET As Long = 10040319      ' RGB(255, 51, 153)
Private Const TRAIL_COUNT As Long = 3
Private mSpeed As Long
Private mSpeedSet As Boolean

Private Sub EnsureSpeed()
    If Not mSpeedSet Then
        mSpeed = 2
        mSpeedSet = True
    End If
End Sub

Public Function FlowSpeed() As Long
    EnsureSpeed
    FlowSpeed = mSpeed
End Function

Public Sub FlowSetSpeed(ByVal value As Long)
    If value < 0 Or value > 3 Then value = 2
    mSpeed = value
    mSpeedSet = True
End Sub

' normal -> slow -> off -> fast -> normal
Public Sub FlowCycleSpeed()
    EnsureSpeed
    mSpeed = (mSpeed + 3) Mod 4
End Sub

Public Function FlowSpeedText() As String
    EnsureSpeed
    FlowSpeedText = Choose(mSpeed + 1, "SIN ANIMACION", "LENTA", "NORMAL", "RAPIDA")
End Function

Public Function FlowOn() As Boolean
    FlowOn = (FlowSpeed() > 0)
End Function

Public Function FlowPacketColor() As Long
    FlowPacketColor = PACKET
End Function

Private Function FramesPerLeg() As Long
    FramesPerLeg = Choose(FlowSpeed() + 1, 1, 18, 11, 6)
End Function

Private Function TickMs() As Long
    TickMs = Choose(FlowSpeed() + 1, 0, 22, 14, 7)
End Function

Public Sub FlowPause(ByVal ticks As Long)
    If FlowSpeed() = 0 Then Exit Sub
    DoEvents
    Sleep TickMs() * ticks
End Sub

' Centre of an anchor in sheet points.
Public Sub FlowAnchorXY(ByVal ws As Worksheet, ByVal anchor As Variant, ByRef x As Double, ByRef y As Double)
    Dim s As Shape
    ' Objects first: IsArray() is True for a multi-cell Range (its default property Value is an array).
    If IsObject(anchor) Then
        x = anchor.Left + anchor.Width / 2
        y = anchor.Top + anchor.Height / 2
    ElseIf TypeName(anchor) = "String" Then
        Set s = ws.Shapes(CStr(anchor))
        x = s.Left + s.Width / 2
        y = s.Top + s.Height / 2
    ElseIf IsArray(anchor) Then
        x = CDbl(anchor(LBound(anchor)))
        y = CDbl(anchor(LBound(anchor) + 1))
    Else
        Err.Raise vbObjectError + 2171, "LcaFlow", "FLOW_ANCHOR_TYPE"
    End If
End Sub

' Creates a packet shape named FLOW_<name> centred on the anchor.
Public Function FlowSpawn(ByVal ws As Worksheet, ByVal name As String, ByVal text As String, _
                          ByVal color As Long, ByVal anchor As Variant) As Shape
    Dim x As Double, y As Double, w As Double, s As Shape
    FlowAnchorXY ws, anchor, x, y
    w = 34
    If Len(text) > 3 Then w = 34 + 7 * (Len(text) - 3)
    If w > 96 Then w = 96
    On Error Resume Next
    ws.Shapes("FLOW_" & name).Delete
    On Error GoTo 0
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x - w / 2, y - 17, w, 34)
    s.Adjustments.Item(1) = 0.5
    s.Name = "FLOW_" & name
    s.Fill.ForeColor.RGB = color
    s.Line.ForeColor.RGB = RGB(255, 255, 255)
    s.Line.Weight = 2
    On Error Resume Next
    s.Glow.Color.RGB = color
    s.Glow.Radius = 10
    On Error GoTo 0
    With s.TextFrame2
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = msoFalse
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = text
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    s.Placement = xlFreeFloating
    Set FlowSpawn = s
End Function

Public Sub FlowKill(ByVal ws As Worksheet, ByVal name As String)
    On Error Resume Next
    ws.Shapes("FLOW_" & name).Delete
    On Error GoTo 0
End Sub

' Moves several packets at the same time (names and anchors are parallel arrays).
Public Sub FlowMoveTo(ByVal ws As Worksheet, ByVal names As Variant, ByVal anchors As Variant)
    Dim n As Long, i As Long, f As Long, frames As Long, t As Double, e As Double
    Dim sx() As Double, sy() As Double, tx() As Double, ty() As Double, s As Shape
    If Not FlowOn() Then Exit Sub
    Application.ScreenUpdating = True
    n = UBound(names) - LBound(names) + 1
    ReDim sx(0 To n - 1)
    ReDim sy(0 To n - 1)
    ReDim tx(0 To n - 1)
    ReDim ty(0 To n - 1)
    For i = 0 To n - 1
        Set s = ws.Shapes("FLOW_" & CStr(names(LBound(names) + i)))
        sx(i) = s.Left + s.Width / 2
        sy(i) = s.Top + s.Height / 2
        FlowAnchorXY ws, anchors(LBound(anchors) + i), tx(i), ty(i)
    Next i
    frames = FramesPerLeg()
    For f = 1 To frames
        t = f / frames
        e = t * t * (3 - 2 * t)
        For i = 0 To n - 1
            Set s = ws.Shapes("FLOW_" & CStr(names(LBound(names) + i)))
            s.Left = sx(i) + (tx(i) - sx(i)) * e - s.Width / 2
            s.Top = sy(i) + (ty(i) - sy(i)) * e - s.Height / 2
        Next i
        FlowPause 1
    Next f
End Sub

' One packet with a comet trail through every anchor in order, then it disappears.
Public Sub FlowSend(ByVal ws As Worksheet, ByVal text As String, ByVal color As Long, ParamArray anchors() As Variant)
    Dim n As Long, leg As Long, f As Long, frames As Long, total As Long, idx As Long, k As Long, j As Long
    Dim hx() As Double, hy() As Double, ax As Double, ay As Double, tx As Double, ty As Double
    Dim t As Double, e As Double, p As Shape, trail(0 To TRAIL_COUNT - 1) As Shape, tw As Double
    If Not FlowOn() Then Exit Sub
    n = UBound(anchors) - LBound(anchors)
    If n < 1 Then Exit Sub
    Application.ScreenUpdating = True
    frames = FramesPerLeg()
    total = n * frames
    ReDim hx(0 To total)
    ReDim hy(0 To total)
    FlowAnchorXY ws, anchors(LBound(anchors)), ax, ay
    hx(0) = ax
    hy(0) = ay
    idx = 0
    For leg = 1 To n
        FlowAnchorXY ws, anchors(LBound(anchors) + leg), tx, ty
        For f = 1 To frames
            t = f / frames
            e = t * t * (3 - 2 * t)
            idx = idx + 1
            hx(idx) = ax + (tx - ax) * e
            hy(idx) = ay + (ty - ay) * e
        Next f
        ax = tx
        ay = ty
    Next leg
    Set p = FlowSpawn(ws, "send", text, color, Array(hx(0), hy(0)))
    For k = 0 To TRAIL_COUNT - 1
        tw = 16 - 4 * k
        Set trail(k) = ws.Shapes.AddShape(msoShapeOval, hx(0) - tw / 2, hy(0) - tw / 2, tw, tw)
        trail(k).Name = "FLOW_sendt" & CStr(k)
        trail(k).Fill.ForeColor.RGB = color
        trail(k).Fill.Transparency = 0.3 + 0.2 * k
        trail(k).Line.Visible = msoFalse
        trail(k).Placement = xlFreeFloating
    Next k
    p.ZOrder msoBringToFront
    For idx = 1 To total
        p.Left = hx(idx) - p.Width / 2
        p.Top = hy(idx) - p.Height / 2
        For k = 0 To TRAIL_COUNT - 1
            j = idx - 2 * (k + 1)
            If j < 0 Then j = 0
            trail(k).Left = hx(j) - trail(k).Width / 2
            trail(k).Top = hy(j) - trail(k).Height / 2
        Next k
        FlowPause 1
    Next idx
    p.Delete
    For k = 0 To TRAIL_COUNT - 1
        trail(k).Delete
    Next k
End Sub

' Flashes a shape's fill (components "lighting up" when data arrives).
Public Sub FlowPulse(ByVal ws As Worksheet, ByVal shapeName As String, ByVal color As Long, Optional ByVal times As Long = 2)
    Dim s As Shape, orig As Long, i As Long
    If Not FlowOn() Then Exit Sub
    Set s = ws.Shapes(shapeName)
    orig = s.Fill.ForeColor.RGB
    For i = 1 To times
        s.Fill.ForeColor.RGB = color
        FlowPause 3
        s.Fill.ForeColor.RGB = orig
        FlowPause 2
    Next i
End Sub

' Flashes the fill of a cell range (for the cell-based blocks).
Public Sub FlowPulseRange(ByVal area As Range, ByVal color As Long, Optional ByVal times As Long = 2)
    Dim orig As Long, i As Long
    If Not FlowOn() Then Exit Sub
    orig = area.Interior.Color
    For i = 1 To times
        area.Interior.Color = color
        FlowPause 3
        area.Interior.Color = orig
        FlowPause 2
    Next i
End Sub

Public Function FlowHasShape(ByVal ws As Worksheet, ByVal name As String) As Boolean
    Dim s As Shape
    For Each s In ws.Shapes
        If s.Name = name Then
            FlowHasShape = True
            Exit Function
        End If
    Next s
End Function
