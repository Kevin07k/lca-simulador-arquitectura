Attribute VB_Name = "LcaCoverPanel"
Option Explicit

' Cover block of the LCA dashboard (replaces the old "Inicio" sheet): logo, project, team, big
' navigation tiles to every part of the simulator and the installation / regression status.
Private Const ROW_BASE As Long = LCA_BASE_COVER
Private Const COL_BASE As Long = LCA_COL_COVER
Private Const BAND_ROWS As Long = 26
Private Const STATUS_ROW As Long = 24
Private Const BG As Long = 2758415          ' RGB(15, 23, 42)
Private Const CARD As Long = 3877150        ' RGB(30, 41, 59)
Private Const HEAD As Long = 6897694        ' RGB(30, 64, 105)
Private Const INK As Long = 15788258        ' RGB(226, 232, 240)
Private Const MUTED As Long = 12100500      ' RGB(148, 163, 184)
Private Const ACCENT As Long = 16569213     ' RGB(125, 211, 252)
Private Const LAMP_ON As Long = 2477306     ' RGB(250, 204, 37)
Private Const BLUE As Long = 15426341       ' RGB(37, 99, 235)
Public Const LCA_LOGO_SHAPE As String = "LOGO_LCA"

Public Sub ConstruirPortada()
    Dim ws As Worksheet, tiles As Variant, i As Long, x As Double, y As Double, w As Double, h As Double
    Dim left0 As Double, top0 As Double, span As Double
    Set ws = LcaSheet(True)
    LcaPrepareBand ws, ROW_BASE, BAND_ROWS, "P00_", COL_BASE
    LcaRowHeight ws.Rows(RA("1:" & CStr(BAND_ROWS))), 22
    With ws.Range(RA("A1:T" & CStr(BAND_ROWS)))
        .Interior.Color = BG
        .Font.Name = "Segoe UI"
        .Font.Color = INK
        .VerticalAlignment = xlCenter
    End With
    ' Accent strip on the left of the title.
    PlaceLogo ws
    ws.Range(RA("F2:S3")).Merge
    With ws.Range(RA("F2"))
        .Value = "LCA  -  SIMULADOR DE ARQUITECTURA DE COMPUTADORAS"
        .Font.Size = 26
        .Font.Bold = True
        .Font.Color = ACCENT
    End With
    ws.Range(RA("F4:S4")).Merge
    With ws.Range(RA("F4"))
        .Value = "Segundo parcial de Arquitectura de Computadoras  |  Equipo LCA:  Kevin (@Kevin07k)  y  Javier (@Javiertorricos)"
        .Font.Size = 13
        .Font.Color = INK
    End With
    ws.Range(RA("F5:S5")).Merge
    With ws.Range(RA("F5"))
        .Value = "Microsoft Excel + VBA  |  todo el simulador en una sola hoja: haz clic en un modulo para ir a el."
        .Font.Size = 11
        .Font.Color = MUTED
    End With
    Chips ws

    ' Navigation tiles: 2 rows x 5.
    tiles = Array( _
        Array("INICIAR SIMULADOR", "P12  maquina completa animada", "LcaIrMaquina", LAMP_ON), _
        Array("DEMOS ARITMETICOS", "P15  suma, producto, Fibonacci...", "LcaIrAritmetica", RGB(236, 72, 153)), _
        Array("EDITOR ASM", "P11  escribir, ensamblar y cargar", "LcaIrEditor", BLUE), _
        Array("CPU + ALU", "P04  registros, bits, flags, ALU", "LcaIrCpu", BLUE), _
        Array("BUSES + RAM", "P05/P06  ciclos y micropalabra", "LcaIrBuses", BLUE), _
        Array("MEMORIA VIRTUAL", "P07/P08  paginas, CLOCK, swap", "LcaIrSwap", BLUE), _
        Array("CACHES", "P09/P10  L1I, L1D, L2 + paginacion", "LcaIrCaches", BLUE), _
        Array("TECLADO + LCD", "P13/P14  E/S con cola FIFO", "LcaIrTeclado", BLUE), _
        Array("TABLERO COMPLETO", "vista general de todo el simulador", "LcaVistaGeneral", HEAD), _
        Array("LOGS", "trazas completas y pruebas", "LcaIrLogs", HEAD))
    left0 = CL(ws, ROW_BASE + 9, 2).Left
    top0 = CL(ws, ROW_BASE + 9, 2).Top
    span = CL(ws, ROW_BASE + 9, 20).Left - left0
    w = span / 5 - 12
    h = CL(ws, ROW_BASE + 15, 2).Top - top0 - 10
    For i = 0 To 9
        x = left0 + (i Mod 5) * (span / 5)
        y = top0 + (i \ 5) * (h + 14)
        Tile ws, "tile" & CStr(i), x, y, w, h, CStr(tiles(i)(0)), CStr(tiles(i)(1)), CStr(tiles(i)(2)), CLng(tiles(i)(3))
    Next i

    ' Status line + actions.
    ws.Range(RA("B" & CStr(STATUS_ROW) & ":J" & CStr(STATUS_ROW))).Merge
    With ws.Range(RA("B" & CStr(STATUS_ROW)))
        .Value = "Estado VBA: pulsa VERIFICAR INSTALACION o VERIFICAR TODO."
        .Interior.Color = CARD
        .Font.Name = "Consolas"
        .Font.Size = 11
        .IndentLevel = 1
    End With
    SmallButton ws, "K" & CStr(STATUS_ROW) & ":M" & CStr(STATUS_ROW), "VERIFICAR INSTALACION", "VerificarInstalacion"
    SmallButton ws, "N" & CStr(STATUS_ROW) & ":P" & CStr(STATUS_ROW), "VERIFICAR TODO (P04-P15)", "VerificarTodo"
    SmallButton ws, "Q" & CStr(STATUS_ROW) & ":S" & CStr(STATUS_ROW), "PROBAR DEMOS P15", "VerificarDemoAritmeticaP15"
End Sub

' Writes a status text in the cover (used by VerificarInstalacion and VerificarTodo).
Public Sub PortadaEstado(ByVal text As String)
    Dim ws As Worksheet
    Set ws = LcaBand("P00_tile0", False)
    If ws Is Nothing Then Exit Sub
    ws.Range(RA("B" & CStr(STATUS_ROW))).Value = text
End Sub

Public Sub LcaIrPortada()
    LcaGoto LcaSheet(True), ROW_BASE, COL_BASE
End Sub

' Logo: embedded picture kept across rebuilds; inserted from assets/lca-chip.png the first time,
' otherwise a drawn chip.
Private Sub PlaceLogo(ByVal ws As Worksheet)
    Dim s As Shape, area As Range, path As String, found As Boolean
    Set area = ws.Range(RA("B2:D8"))
    For Each s In ws.Shapes
        If s.Name = LCA_LOGO_SHAPE Then
            ' Accept only a real picture or the drawn "LCA" chip; anything else is discarded.
            If s.Type = msoPicture Then
                found = True
            ElseIf s.TextFrame2.HasText Then
                If s.TextFrame2.TextRange.Text = "LCA" Then found = True
            End If
            If Not found Then
                s.Delete
                Exit For
            End If
        End If
    Next s
    If Not found Then
        path = ThisWorkbook.Path & Application.PathSeparator & ".." & Application.PathSeparator & "assets" & _
               Application.PathSeparator & "lca-chip.png"
        On Error Resume Next
        If Len(Dir$(path)) > 0 Then
            Set s = ws.Shapes.AddPicture(path, msoFalse, msoTrue, area.Left, area.Top, -1, -1)
            If Not s Is Nothing Then
                s.Name = LCA_LOGO_SHAPE
                found = True
            End If
        End If
        On Error GoTo 0
    End If
    If Not found Then
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left, area.Top, area.Height, area.Height)
        s.Name = LCA_LOGO_SHAPE
        s.Fill.ForeColor.RGB = BLUE
        s.Line.ForeColor.RGB = ACCENT
        s.Line.Weight = 3
        With s.TextFrame2
            .TextRange.Text = "LCA"
            .TextRange.Font.Size = 40
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
        End With
    End If
    Set s = ws.Shapes(LCA_LOGO_SHAPE)
    s.LockAspectRatio = msoTrue
    s.Height = area.Height
    s.Left = area.Left + (area.Width - s.Width) / 2
    s.Top = area.Top
    s.Placement = xlMove
End Sub

Private Sub Chips(ByVal ws As Worksheet)
    Dim labels As Variant, i As Long, x As Double, top As Double, w As Double, s As Shape
    labels = Array("CPU 8 bits", "VA 9 bits", "paginas de 32 B", "CLOCK + swap", "L1I / L1D / L2", "teclado FIFO", "LCD 8x1", "ISA de 4 bytes")
    x = CL(ws, ROW_BASE + 7, 6).Left
    top = CL(ws, ROW_BASE + 7, 6).Top + 2
    w = (CL(ws, ROW_BASE + 7, 20).Left - x) / 8
    For i = 0 To 7
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x + i * w + 3, top, w - 6, 20)
        s.Name = "P00_chip" & CStr(i)
        s.Adjustments.Item(1) = 0.5
        s.Fill.ForeColor.RGB = CARD
        s.Line.ForeColor.RGB = ACCENT
        s.Line.Weight = 0.75
        With s.TextFrame2
            .MarginTop = 0
            .MarginBottom = 0
            .TextRange.Text = CStr(labels(i))
            .TextRange.Font.Name = "Segoe UI"
            .TextRange.Font.Size = 9
            .TextRange.Font.Fill.ForeColor.RGB = ACCENT
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
        End With
    Next i
End Sub

Private Sub Tile(ByVal ws As Worksheet, ByVal name As String, ByVal x As Double, ByVal y As Double, _
                 ByVal w As Double, ByVal h As Double, ByVal title As String, ByVal subtitle As String, _
                 ByVal macroName As String, ByVal fill As Long)
    Dim s As Shape, tileInk As Long
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x + 6, y, w, h)
    s.Name = "P00_" & name
    s.Adjustments.Item(1) = 0.1
    s.Fill.ForeColor.RGB = fill
    s.Line.Visible = msoFalse
    s.Shadow.Visible = msoTrue
    s.Shadow.OffsetX = 3
    s.Shadow.OffsetY = 3
    s.Shadow.Transparency = 0.45
    If fill = LAMP_ON Then tileInk = BG Else tileInk = RGB(255, 255, 255)
    With s.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = title & vbCr & subtitle
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Fill.ForeColor.RGB = tileInk
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextRange.Paragraphs(1).Font.Size = 16
        .TextRange.Paragraphs(1).Font.Bold = msoTrue
        .TextRange.Paragraphs(2).Font.Size = 10
    End With
    s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Sub

Private Sub SmallButton(ByVal ws As Worksheet, ByVal address As String, ByVal caption As String, ByVal macroName As String)
    Dim area As Range, s As Shape
    Set area = ws.Range(RA(address))
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left + 3, area.Top + 1, area.Width - 6, area.Height - 2)
    s.Name = "P00_btn_" & macroName
    s.Fill.ForeColor.RGB = HEAD
    s.Line.ForeColor.RGB = ACCENT
    With s.TextFrame2
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = caption
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Sub

Private Function RA(ByVal address As String) As String
    RA = LcaShift(address, ROW_BASE, COL_BASE)
End Function

Private Function CL(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Range
    Set CL = ws.Cells(r, c + COL_BASE)
End Function
