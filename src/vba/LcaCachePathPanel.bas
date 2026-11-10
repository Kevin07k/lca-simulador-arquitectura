Attribute VB_Name = "LcaCachePathPanel"
Option Explicit

' P10 band of the LCA sheet: last-access pipeline (VA -> PTE -> PA -> L1 -> L2 -> RAM),
' the seven frames with their cached lines, global counters and the integrated trace.
Private Const ROW_BASE As Long = LCA_BASE_CACHE_PATH
Private Const BAND_ROWS As Long = 47
Private Const BG As Long = 2758415          ' RGB(15, 23, 42)
Private Const CARD As Long = 3877150        ' RGB(30, 41, 59)
Private Const HEAD As Long = 6897694        ' RGB(30, 64, 105)
Private Const INK As Long = 15788258        ' RGB(226, 232, 240)
Private Const MUTED As Long = 12100500      ' RGB(148, 163, 184)
Private Const ACCENT As Long = 16569213     ' RGB(125, 211, 252)
Private Const button As Long = 15426341     ' RGB(37, 99, 235)
Private Const OK_GREEN As Long = 3433750    ' RGB(22, 101, 52)
Private Const WARN As Long = 1592985        ' RGB(153, 78, 24)
Private Const BAD As Long = 1908095         ' RGB(127, 29, 29)
Private Const IDLE As Long = 5587251        ' RGB(51, 65, 85)  stage not used
Private Const RESERVED As Long = 3747162    ' RGB(90, 45, 57)
Private Const RESIDENT As Long = 5521694    ' RGB(30, 65, 84)

Private mStep As Long
Private mStatus As String

Public Sub ConstruirPanelCachePaginacion()
    Dim ws As Worksheet, i As Long, Stage As Variant, col As Long
    Set ws = PathSheet(True)
    LcaPrepareBand ws, ROW_BASE, BAND_ROWS, "P10_"
    ws.Rows(RA("1:" & CStr(BAND_ROWS))).RowHeight = 21
    With ws.Range(RA("A1:AZ" & CStr(BAND_ROWS)))
        .Interior.Color = BG
        .Font.Name = "Consolas"
        .Font.Size = 11
        .Font.Color = INK
        .VerticalAlignment = xlCenter
    End With

    ' Header and controls
    ws.Rows(ROW_BASE + 2).RowHeight = 30
    ws.Range(RA("B2:M2")).Merge
    With ws.Range(RA("B2"))
        .Value = "P10  /  CACHE + PAGINACION"
        .Font.Name = "Segoe UI"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = ACCENT
    End With
    ws.Range(RA("B3:S3")).Merge
    With ws.Range(RA("B3"))
        .Value = "VA -> PTE (residencia primero) -> PA -> L1I/L1D -> L2 -> RAM  |  invalidacion por marco  |  write-through: swap recupera el dato"
        .Font.Color = MUTED
    End With
    AddButton ws, "N2:P2", "VER P08", "LcaIrSwap", HEAD
    AddButton ws, "Q2:S2", "VER P09", "LcaIrCaches", HEAD
    AddButton ws, "B5:E6", "REINICIAR / RESET", "CachePathDemoReset", button
    AddButton ws, "F5:I6", "SIGUIENTE / NEXT", "CachePathDemoNext", button
    AddButton ws, "J5:M6", "EJECUTAR TODO / RUN ALL", "CachePathDemoRun", button
    ws.Range(RA("N5:S6")).Merge
    ws.Range(RA("N5")).WrapText = True
    ws.Range(RA("N5")).Font.Color = MUTED
    ws.Range(RA("N5")).Font.Size = 9
    ws.Range(RA("N5")).Value = "13 pasos reproducibles: 7 marcos, 8.a pagina, expulsion dirty, linea vieja invalidada, A5h recuperado, E/S rechazada."
    ws.Range(RA("B8:S9")).Merge
    ws.Range(RA("B8")).WrapText = True
    ws.Range(RA("B8")).Interior.Color = CARD
    ws.Range(RA("B8")).Font.Name = "Segoe UI"

    ' Pipeline of the last access: six stage cards of three columns each
    SectionTitle ws, "B11:S11", "01   RECORRIDO DEL ULTIMO ACCESO / LAST ACCESS PATH"
    col = 2
    For Each Stage In Array("VA", "PTE / PAGINA", "PA", "L1", "L2", "RAM")
        With ws.Range(ws.Cells(ROW_BASE + 12, col), ws.Cells(ROW_BASE + 12, col + 2))
            .Merge
            .Value = CStr(Stage)
            .Font.Bold = True
            .Font.Color = MUTED
            .HorizontalAlignment = xlCenter
        End With
        With ws.Range(ws.Cells(ROW_BASE + 13, col), ws.Cells(ROW_BASE + 14, col + 2))
            .Merge
            .Font.Size = 18
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        With ws.Range(ws.Cells(ROW_BASE + 15, col), ws.Cells(ROW_BASE + 15, col + 2))
            .Merge
            .Font.Size = 9
            .HorizontalAlignment = xlCenter
        End With
        If col < 17 Then
            ' A thick background-coloured edge splits the stages into separate cards.
            With ws.Range(ws.Cells(ROW_BASE + 12, col + 2), ws.Cells(ROW_BASE + 15, col + 2)).Borders(xlEdgeRight)
                .Color = BG
                .Weight = xlThick
            End With
        End If
        col = col + 3
    Next Stage
    ws.Rows(ROW_BASE + 13).RowHeight = 24
    ws.Rows(ROW_BASE + 14).RowHeight = 24

    ' Frames (left) and counters (right)
    SectionTitle ws, "B17:J17", "02   MARCOS / FRAMES  (lineas en cache por nivel)"
    SectionTitle ws, "L17:S17", "03   CONTADORES / COUNTERS"
    HeaderRow ws, 18, Array("PFN", "VPN", "DIRTY", "REF", "L1I", "L1D", "L2", "CLOCK", "ESTADO")
    ws.Range(RA("L18:P18")).Merge
    ws.Range(RA("L18")).Value = "METRICA / METRIC"
    ws.Range(RA("Q18:S18")).Merge
    ws.Range(RA("Q18")).Value = "VALOR / VALUE"
    ws.Range(RA("L18:S18")).Interior.Color = HEAD
    ws.Range(RA("L18:S18")).Font.Bold = True
    For i = 19 To 27
        ws.Range(RA("L" & CStr(i) & ":P" & CStr(i))).Merge
        ws.Range(RA("Q" & CStr(i) & ":S" & CStr(i))).Merge
        ws.Range(RA("L" & CStr(i) & ":S" & CStr(i))).Interior.Color = CARD
    Next i
    ws.Range(RA("B19:J27")).HorizontalAlignment = xlCenter
    ws.Range(RA("Q18:S27")).HorizontalAlignment = xlCenter
    ws.Range(RA("Q19:S27")).Font.Bold = True
    ws.Range(RA("Q19:S27")).Font.Color = ACCENT

    ' Integrated trace
    SectionTitle ws, "B29:S29", "04   TRAZA INTEGRADA / INTEGRATED TRACE  (mas reciente arriba / newest first)"
    For i = 30 To 41
        ws.Range(RA("B" & CStr(i) & ":S" & CStr(i))).Merge
        ws.Range(RA("B" & CStr(i) & ":S" & CStr(i))).Interior.Color = CARD
        ws.Range(RA("B" & CStr(i))).Font.Size = 10
    Next i

    ' Legend
    SectionTitle ws, "B43:S43", "LEYENDA / LEGEND"
    Legend ws, "B44:D44", OK_GREEN, "HIT / OK"
    Legend ws, "E44:G44", WARN, "MISS / FAULT"
    Legend ws, "H44:J44", BAD, "ERROR"
    Legend ws, "K44:M44", IDLE, "NO USADO / UNUSED"
    Legend ws, "N44:P44", RESIDENT, "MARCO OCUPADO"
    Legend ws, "Q44:S44", RESERVED, "PFN7 E/S"
    ws.Range(RA("B46:S46")).Merge
    ws.Range(RA("B46")).Font.Color = MUTED
    ws.Range(RA("B46")).Font.Size = 9
    ws.Range(RA("B46")).Value = "Paso 11: sin invalidar PFN1, L2 devolveria A5h (dato viejo de VPN8) en lugar del codigo 10h. Prueba: VerificarCachePaginacionP10."

    CachePathDemoReset
    LcaGoto ws, ROW_BASE
End Sub

Public Sub MostrarPanelCachePaginacion()
    Dim ws As Worksheet
    Set ws = PathSheet(False)
    If ws Is Nothing Then
        ConstruirPanelCachePaginacion
    Else
        LcaGoto ws, ROW_BASE
        CachePathPanelRefresh
    End If
End Sub

Public Sub CachePathDemoReset()
    On Error GoTo failed
    CachePathScenarioReset
    mStep = 0
    mStatus = "Listo / ready: respaldo cargado (VPN0 = 10 00 2A 00, datos VPN8..14 = VPN+offset, AX = A5h). Pulsa SIGUIENTE."
    CachePathPanelRefresh
    Exit Sub
failed:
    mStatus = "ERROR RESET: " & Err.Description
    CachePathPanelRefresh
End Sub

Public Sub CachePathDemoNext()
    Dim text As String
    If Not CachePathReady() Then CachePathDemoReset
    If mStep >= LCA_P10_STEPS Then
        mStatus = "Fin / done. REINICIAR para repetir / RESET to repeat."
        CachePathPanelRefresh
        Exit Sub
    End If
    On Error GoTo failed
    text = CachePathScenarioStep(mStep)
    mStep = mStep + 1
    mStatus = "Paso " & CStr(mStep) & "/" & CStr(LCA_P10_STEPS) & ": " & text
    CachePathPanelRefresh
    Exit Sub
failed:
    mStep = mStep + 1
    If mStep = LCA_P10_STEPS Then
        mStatus = "Paso " & CStr(mStep) & "/" & CStr(LCA_P10_STEPS) & ": " & CachePathScenarioText(mStep - 1) & _
                  "  [esperado / expected: " & Err.Description & "]"
    Else
        mStatus = "ERROR paso " & CStr(mStep) & ": " & Err.Description
    End If
    CachePathPanelRefresh
End Sub

Public Sub CachePathDemoRun()
    Dim i As Long
    CachePathDemoReset
    Application.ScreenUpdating = False
    For i = 1 To LCA_P10_STEPS
        CachePathDemoNext
    Next i
    Application.ScreenUpdating = True
End Sub

' Redraws the band from the engines; does nothing while the band is not built.
Public Sub CachePathPanelRefresh()
    Dim ws As Worksheet
    Set ws = PathSheet(False)
    If ws Is Nothing Then Exit Sub
    If Not CachePathReady() Then Exit Sub
    ws.Range(RA("B8")).Value = mStatus
    DrawPipeline ws
    DrawFrames ws
    DrawCounters ws
    DrawTrace ws
End Sub

Private Sub DrawPipeline(ByVal ws As Worksheet)
    Dim s As LcaPathStep, l1 As String
    s = CachePathLast()
    If s.Kind = 0 Then
        Stage ws, 0, "-", "sin accesos / no access", IDLE
        Stage ws, 1, "-", "", IDLE
        Stage ws, 2, "-", "", IDLE
        Stage ws, 3, "-", "", IDLE
        Stage ws, 4, "-", "", IDLE
        Stage ws, 5, "-", "", IDLE
        Exit Sub
    End If
    If s.Kind = LCA_PATH_FETCH Then l1 = "L1I" Else l1 = "L1D"
    Stage ws, 0, Hex3(s.VA), CachePathKindName(s.Kind) & "  VPN " & IIf(s.VPN >= 0, CStr(s.VPN), "-"), _
          IIf(s.Failed And s.PA < 0, BAD, OK_GREEN)
    If s.Failed And s.PA < 0 Then
        Stage ws, 1, "ERROR", Left$(s.Outcome, 40), BAD
        Stage ws, 2, "-", "sin traduccion", IDLE
        Stage ws, 3, "-", "sin cache / no cache", IDLE
        Stage ws, 4, "-", "", IDLE
        Stage ws, 5, "-", "sin ciclo / no cycle", IDLE
        Exit Sub
    End If
    If s.PageFault Then
        If s.VictimVPN >= 0 Then
            Stage ws, 1, "FAULT", "expulsa VPN" & CStr(s.VictimVPN) & IIf(s.Writeback, " + WB", "") & _
                  ", inval " & CStr(s.InvalidatedLines), WARN
        Else
            Stage ws, 1, "FAULT", "marco libre / free frame", WARN
        End If
    Else
        Stage ws, 1, "PRESENTE", "PTE hit, sin fallo", OK_GREEN
    End If
    Stage ws, 2, Hex2(s.PA), "PFN " & CStr(s.PFN) & "  offset " & CStr(s.PA Mod 32), OK_GREEN
    Stage ws, 3, l1 & IIf(s.L1Misses > 0, " MISS", " HIT"), CStr(s.L1Hits) & " hit / " & CStr(s.L1Misses) & " miss", _
          IIf(s.L1Misses > 0, WARN, OK_GREEN)
    If s.L2Hits + s.L2Misses = 0 Then
        Stage ws, 4, "-", "no consultada", IDLE
    Else
        Stage ws, 4, IIf(s.L2Misses > 0, "MISS", "HIT"), CStr(s.L2Hits) & " hit / " & CStr(s.L2Misses) & " miss", _
              IIf(s.L2Misses > 0, WARN, OK_GREEN)
    End If
    If s.RamCycles = 0 Then
        Stage ws, 5, "0", "bus externo High-Z", IDLE
    Else
        Stage ws, 5, CStr(s.RamCycles), "ciclos / cycles  byte=" & Hex2(s.Value), WARN
    End If
End Sub

Private Sub Stage(ByVal ws As Worksheet, ByVal index As Long, ByVal big As String, _
                  ByVal detail As String, ByVal color As Long)
    Dim col As Long
    col = 2 + 3 * index
    ws.Cells(ROW_BASE + 13, col).Value = big
    ws.Cells(ROW_BASE + 15, col).Value = detail
    ws.Range(ws.Cells(ROW_BASE + 12, col), ws.Cells(ROW_BASE + 15, col + 2)).Interior.Color = color
End Sub

Private Sub DrawFrames(ByVal ws As Worksheet)
    Dim pfn As Long, row As Long, owner As Long, entry As LcaPageEntry, color As Long
    Dim s As LcaPathStep
    s = CachePathLast()
    For pfn = 0 To 7
        row = ROW_BASE + 19 + pfn
        ws.Cells(row, 2).Value = pfn
        If pfn = 7 Then
            ws.Range(ws.Cells(row, 3), ws.Cells(row, 9)).Value = "-"
            ws.Cells(row, 10).Value = "E/S RESERVADA"
            color = RESERVED
        Else
            owner = PagerFrameOwner(pfn)
            If owner < 0 Then
                ws.Range(ws.Cells(row, 3), ws.Cells(row, 5)).Value = "-"
                ws.Cells(row, 10).Value = "LIBRE / FREE"
                color = CARD
            Else
                entry = MmuPage(owner)
                ws.Cells(row, 3).Value = owner
                ws.Cells(row, 4).Value = Abs(CLng(entry.Dirty))
                ws.Cells(row, 5).Value = Abs(CLng(entry.Referenced))
                ws.Cells(row, 10).Value = IIf(owner <= 7, "CODIGO", "DATOS")
                color = RESIDENT
            End If
            If pfn = s.PFN And Not s.Failed Then color = IIf(s.PageFault, WARN, OK_GREEN)
        End If
        ws.Cells(row, 6).Value = CacheFrameLines(LCA_CACHE_L1I, pfn)
        ws.Cells(row, 7).Value = CacheFrameLines(LCA_CACHE_L1D, pfn)
        ws.Cells(row, 8).Value = CacheFrameLines(LCA_CACHE_L2, pfn)
        If pfn < 7 Then ws.Cells(row, 9).Value = IIf(PagerHand() = pfn, "<- mano", "")
        ws.Range(ws.Cells(row, 2), ws.Cells(row, 10)).Interior.Color = color
    Next pfn
End Sub

Private Sub DrawCounters(ByVal ws As Worksheet)
    Dim names As Variant, values As Variant, i As Long
    names = Array("Accesos / requests", "Page faults", "Expulsiones / evictions", _
                  "Writebacks (32 B)", "Lineas invalidadas / lines dropped", _
                  "L1I hit / miss", "L1D hit / miss", "L2 hit / miss", "Ciclos RAM / RAM cycles")
    values = Array(CStr(CachePathAccessCount()), CStr(PagerFaultCount()), CStr(PagerEvictionCount()), _
                   CStr(PagerWritebackCount()), CStr(PagerInvalidatedLines()), _
                   HitMiss(LCA_CACHE_L1I), HitMiss(LCA_CACHE_L1D), HitMiss(LCA_CACHE_L2), _
                   CStr(CachePathRamCycles()))
    For i = 0 To 8
        ws.Cells(ROW_BASE + 19 + i, 12).Value = names(i)
        ws.Cells(ROW_BASE + 19 + i, 17).Value = "'" & values(i)
    Next i
End Sub

Private Sub DrawTrace(ByVal ws As Worksheet)
    Dim i As Long, row As Long, text As String, color As Long
    For i = 0 To 11
        row = ROW_BASE + 30 + i
        text = vbNullString
        If i < CachePathTraceCount() Then text = CachePathTraceLine(i)
        ws.Cells(row, 2).Value = text
        color = CARD
        If InStr(1, text, "ERROR", vbBinaryCompare) > 0 Then
            color = BAD
        ElseIf InStr(1, text, "evict", vbBinaryCompare) > 0 Then
            color = WARN
        End If
        ws.Range(ws.Cells(row, 2), ws.Cells(row, 19)).Interior.Color = color
    Next i
End Sub

Private Function HitMiss(ByVal level As Long) As String
    HitMiss = CStr(CacheHits(level)) & " / " & CStr(CacheMisses(level))
End Function

Private Sub SectionTitle(ByVal ws As Worksheet, ByVal address As String, ByVal text As String)
    With ws.Range(RA(address))
        .Merge
        .Value = text
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Color = ACCENT
        .Borders(xlEdgeBottom).Color = HEAD
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
End Sub

Private Sub HeaderRow(ByVal ws As Worksheet, ByVal row As Long, ByVal labels As Variant)
    Dim i As Long
    For i = 0 To UBound(labels)
        ws.Cells(ROW_BASE + row, 2 + i).Value = labels(i)
    Next i
    With ws.Range(ws.Cells(ROW_BASE + row, 2), ws.Cells(ROW_BASE + row, 2 + UBound(labels)))
        .Interior.Color = HEAD
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub Legend(ByVal ws As Worksheet, ByVal address As String, ByVal color As Long, ByVal text As String)
    With ws.Range(RA(address))
        .Merge
        .Value = text
        .Interior.Color = color
        .HorizontalAlignment = xlCenter
        .Font.Size = 9
    End With
End Sub

Private Function PathSheet(ByVal create As Boolean) As Worksheet
    Set PathSheet = LcaBand("P10_CachePathDemoReset", create)
End Function

' Row shift of this panel inside the shared LCA sheet.
Private Function RA(ByVal address As String) As String
    RA = LcaRows(address, ROW_BASE)
End Function

Private Sub AddButton(ByVal ws As Worksheet, ByVal address As String, ByVal caption As String, _
                      ByVal macroName As String, ByVal fillColor As Long)
    Dim area As Range, button As Shape
    Set area = ws.Range(RA(address))
    Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left + 2, area.Top + 2, _
                                    area.Width - 4, area.Height - 4)
    button.Name = "P10_" & macroName
    button.fill.ForeColor.RGB = fillColor
    button.line.Visible = msoFalse
    button.TextFrame2.TextRange.text = caption
    button.TextFrame2.TextRange.Font.Name = "Segoe UI"
    button.TextFrame2.TextRange.Font.Size = 11
    button.TextFrame2.TextRange.Font.Bold = msoTrue
    button.TextFrame2.TextRange.Font.fill.ForeColor.RGB = RGB(255, 255, 255)
    button.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    button.TextFrame2.VerticalAnchor = msoAnchorMiddle
    button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Sub

Private Function Hex2(ByVal value As Long) As String
    If value < 0 Then
        Hex2 = "-"
    Else
        Hex2 = Right$("00" & Hex$(value), 2) & "h"
    End If
End Function

Private Function Hex3(ByVal value As Long) As String
    If value < 0 Then
        Hex3 = "-"
    Else
        Hex3 = Right$("000" & Hex$(value), 3) & "h"
    End If
End Function
