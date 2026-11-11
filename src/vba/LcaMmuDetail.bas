Attribute VB_Name = "LcaMmuDetail"
Option Explicit

' Graphical part of the P07 block: the address in the VA input split bit by bit (VPN | offset), the page
' table entry it selects, and the resulting physical address (PFN | same offset). Read-only: it looks at
' the page table with MmuPage and never changes ref/dirty bits.
Private Const ROW_BASE As Long = LCA_BASE_MMU
Private Const COL_BASE As Long = LCA_COL_MMU
Private Const BG As Long = 2758415           ' RGB(15, 23, 42)
Private Const CARD As Long = 3877150         ' RGB(30, 41, 59)
Private Const HEAD As Long = 6897694         ' RGB(30, 64, 105)
Private Const INK As Long = 15788258         ' RGB(226, 232, 240)
Private Const MUTED As Long = 12100500       ' RGB(148, 163, 184)
Private Const ACCENT As Long = 16569213      ' RGB(125, 211, 252)
Private Const LAMP_ON As Long = 2477306      ' RGB(250, 204, 37)
Private Const EDGE As Long = 5587251         ' RGB(51, 65, 85)
Private Const WARN As Long = 1592985         ' RGB(153, 78, 24)
Private Const VPN_FILL As Long = 9116461     ' RGB(45, 27, 139)
Private Const OFF_FILL As Long = 7954958     ' RGB(14, 98, 121)
Private Const PFN_FILL As Long = 3433750     ' RGB(22, 101, 52)
Private Const BIT_W As Double = 34

Public Sub ConstruirMmuDetalle()
    Dim ws As Worksheet, x0 As Double, y0 As Double, i As Long, s As Shape
    Set ws = LcaSheet(True)
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 5) = "P07D_" Then ws.Shapes(i).Delete
    Next i
    With ws.Range(RA("J21:S21"))
        .Merge
        .Value = "TRADUCCION BIT A BIT de la VA escrita arriba  (VA = VPN x 32 + offset;  PA = PFN x 32 + offset)"
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Color = ACCENT
        .Interior.Color = BG
    End With
    x0 = CL(ws, ROW_BASE + 22, 10).Left
    y0 = CL(ws, ROW_BASE + 22, 10).Top + 4
    Box ws, "lva", x0, y0, 90, 30, BG, "VA  9 bits", 10, True
    For i = 0 To 8
        Set s = Box(ws, "va" & CStr(i), x0 + 96 + i * BIT_W, y0, BIT_W - 3, 30, IIf(i < 4, VPN_FILL, OFF_FILL), "0", 12, True)
    Next i
    Box ws, "lvpn", x0 + 96, y0 + 32, 4 * BIT_W - 3, 16, BG, "VPN (4)", 8
    Box ws, "loff", x0 + 96 + 4 * BIT_W, y0 + 32, 5 * BIT_W - 3, 16, BG, "offset (5)", 8
    Box ws, "pte", x0 + 96, y0 + 54, 9 * BIT_W - 3, 32, CARD, "PTE", 10, True
    Box ws, "lpa", x0, y0 + 96, 90, 30, BG, "PA  8 bits", 10, True
    For i = 0 To 7
        Box ws, "pa" & CStr(i), x0 + 96 + BIT_W + i * BIT_W, y0 + 96, BIT_W - 3, 30, IIf(i < 3, PFN_FILL, OFF_FILL), "0", 12, True
    Next i
    Box ws, "lpfn", x0 + 96 + BIT_W, y0 + 128, 3 * BIT_W - 3, 16, BG, "PFN (3)", 8
    Box ws, "explain", x0 + 96 + 9 * BIT_W + 16, y0, CL(ws, ROW_BASE + 22, 20).Left - (x0 + 96 + 9 * BIT_W + 16) - 8, 144, CARD, "", 10
    MmuDetalleActualizar
End Sub

Public Sub MmuDetalleActualizar()
    Dim ws As Worksheet, va As Long, vpn As Long, offset As Long, e As LcaPageEntry, i As Long, pa As Long, raw As Variant
    Set ws = LcaBand("P07D_pte", False)
    If ws Is Nothing Then Exit Sub
    raw = ws.Range(RA("B7")).Value
    If IsNumeric(raw) And Not IsEmpty(raw) Then va = CLng(raw) Else va = -1
    If va < 0 Or va > 511 Then
        SetText ws, "explain", "Escribe una VA entre 0 y 511 en la caja de arriba y pulsa TRADUCIR.", WARN
        Exit Sub
    End If
    vpn = va \ 32
    offset = va Mod 32
    For i = 0 To 8
        SetText ws, "va" & CStr(i), CStr((va \ 2 ^ (8 - i)) And 1), IIf(i < 4, VPN_FILL, OFF_FILL)
    Next i
    e = MmuPage(vpn)
    If vpn = 15 Then
        SetText ws, "pte", "PTE[15]  RESERVADA para E/S  ->  sin traduccion (PROTECTION)", WARN
        BlankPa ws
        SetText ws, "explain", "VA " & AsmHex3(va) & " = VPN 15, offset " & CStr(offset) & vbCr & _
                "La pagina 15 es el teclado y el LCD: solo IN/OUT la usan, nunca LD/ST.", WARN
    ElseIf Not e.Present Then
        SetText ws, "pte", "PTE[" & CStr(vpn) & "]  P=0  ->  PAGE_FAULT (P08 la trae del swap)", WARN
        BlankPa ws
        SetText ws, "explain", "VA " & AsmHex3(va) & " = VPN " & CStr(vpn) & " x 32 + " & CStr(offset) & vbCr & _
                "La pagina no esta en RAM: el acceso causa un fallo de pagina recuperable.", CARD
    Else
        pa = e.PFN * 32 + offset
        SetText ws, "pte", "PTE[" & CStr(vpn) & "]  P=1  PFN=" & CStr(e.PFN) & "  D=" & Abs(CLng(e.Dirty)) & "  R=" & Abs(CLng(e.Referenced)), PFN_FILL
        For i = 0 To 7
            SetText ws, "pa" & CStr(i), CStr((pa \ 2 ^ (7 - i)) And 1), IIf(i < 3, PFN_FILL, OFF_FILL)
        Next i
        SetText ws, "explain", "VA " & AsmHex3(va) & " (" & CStr(va) & ") = VPN " & CStr(vpn) & " x 32 + " & CStr(offset) & vbCr & _
                "PA " & AsmHex2(pa) & " (" & CStr(pa) & ") = PFN " & CStr(e.PFN) & " x 32 + " & CStr(offset) & vbCr & _
                "El offset pasa igual; solo cambia el numero de pagina por el de marco.", CARD
    End If
End Sub

Private Sub BlankPa(ByVal ws As Worksheet)
    Dim i As Long
    For i = 0 To 7
        SetText ws, "pa" & CStr(i), "-", CARD
    Next i
End Sub

Private Sub SetText(ByVal ws As Worksheet, ByVal name As String, ByVal text As String, ByVal fill As Long)
    With ws.Shapes("P07D_" & name)
        .TextFrame2.TextRange.Text = text
        .Fill.ForeColor.RGB = fill
    End With
End Sub

Private Function Box(ByVal ws As Worksheet, ByVal name As String, ByVal left As Double, ByVal top As Double, _
                     ByVal w As Double, ByVal h As Double, ByVal fill As Long, ByVal text As String, _
                     ByVal size As Single, Optional ByVal bold As Boolean = False) As Shape
    Dim s As Shape
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, left, top, w, h)
    s.Name = "P07D_" & name
    s.Fill.ForeColor.RGB = fill
    s.Line.ForeColor.RGB = IIf(fill = BG, BG, EDGE)
    With s.TextFrame2
        .MarginLeft = 2
        .MarginRight = 2
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = msoTrue
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = text
        .TextRange.Font.Name = IIf(Len(text) = 1, "Consolas", "Segoe UI")
        .TextRange.Font.Size = size
        .TextRange.Font.Bold = IIf(bold, msoTrue, msoFalse)
        .TextRange.Font.Fill.ForeColor.RGB = IIf(fill = BG, MUTED, INK)
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
