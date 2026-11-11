Attribute VB_Name = "LcaUiTests"
Option Explicit

Private mChecks As Long

' Interface regression: block address arithmetic (LcaShift / LcaRows) and, when the dashboard is built,
' the workbook structure (only LCA + Logs), every block marker, block backgrounds and a full refresh
' of the animated machine and the detail views without errors.
Public Sub VerificarInterfazUI()
    Dim ws As Worksheet, markers As Variant, i As Long, blocks As Variant, sh As Worksheet
    Dim names As String, failure As Long
    On Error GoTo failed
    mChecks = 0

    ' --- A: address arithmetic of the 2-D grid ---------------------------------------------------
    Check LcaShift("B2:R3", 10, 20) = "V12:AL13", "A: rows and columns shifted"
    Check LcaShift("A1:AZ43", 0, 40) = "AO1:BH43", "A: columns past T clamped to the block"
    Check LcaShift("1:42", 5, 40) = "6:47", "A: row-only range"
    Check LcaShift("$B$2", 5, 5) = "$B$2", "A: absolute address untouched"
    Check LcaShift("T5", 0, 60) = "CB5", "A: last column of the 4th block"
    Check LcaShift("b7", 0, 0) = "B7", "A: lower case accepted"
    Check LcaRows("B7:C9", 10) = "B17:C19", "A: LcaRows keeps columns"

    ' --- B: dashboard structure (only when the LCA sheet has been built) ---------------------------
    Set ws = LcaBand("P12_btn_MaquinaPaso", False)
    If ws Is Nothing Then
        LcaReport "UI PASS: " & CStr(mChecks) & " comprobaciones (tablero no construido: solo aritmetica de bloques).", True
        Exit Sub
    End If
    For Each sh In ThisWorkbook.Worksheets
        names = names & sh.Name & ";"
    Next sh
    Check names = "LCA;Logs;", "B: the workbook has only the LCA and Logs sheets (got " & names & ")"
    markers = Array("P00_tile0", "P12_btn_MaquinaPaso", "P04_LcaIrDiagrama", "P04D_name0", "D_Chip", "P05_BusDemoReset", _
                    "P05D_bus_ext", "P07_MmuPanelReset", "P07D_pte", "P08_PagerDemoReset", "P08D_hub", _
                    "P09_CacheDemoReset", "P09_chart", "P10_CachePathDemoReset", "P11_EditorEnsamblar", _
                    "P13_TecladoReset", "P14_LcdPanelReset", "P12_chart", "P04D_chart", "LOGO_LCA", "P15_btn_ArithPaso", "P15_chart", "P15_tank3", "P16_btn_EcoPaso", "P16_btn_EcoDemoOcho")
    For i = 0 To UBound(markers)
        Check HasShape(ws, CStr(markers(i))), "B: block element " & CStr(markers(i)) & " present"
    Next i
    blocks = Array(Array(LCA_BASE_COVER, LCA_COL_COVER), Array(LCA_BASE_MACHINE, LCA_COL_MACHINE), _
                   Array(LCA_BASE_EDITOR, LCA_COL_EDITOR), Array(LCA_BASE_KEYBOARD, LCA_COL_KEYBOARD), _
                   Array(LCA_BASE_MMU, LCA_COL_MMU), Array(LCA_BASE_CACHE, LCA_COL_CACHE), Array(LCA_BASE_ARITH, LCA_COL_ARITH), _
                   Array(LCA_BASE_ECHO, LCA_COL_ECHO))
    For i = 0 To UBound(blocks)
        Check ws.Cells(CLng(blocks(i)(0)) + 1, CLng(blocks(i)(1)) + 1).Interior.Color = RGB(15, 23, 42), _
              "B: dark background at the origin of block " & CStr(i)
    Next i

    ' --- C: refresh every live view on a loaded program -------------------------------------------
    Check LoaderLoadSource(LoaderSampleProgram()), "C: sample program loads"
    Call ExecRun
    failure = 0
    On Error Resume Next
    MaquinaRefrescar
    failure = failure + Err.Number: Err.Clear
    PanelActualizar
    failure = failure + Err.Number: Err.Clear
    MemoriaDetalleActualizar
    failure = failure + Err.Number: Err.Clear
    BusDetalleActualizar
    failure = failure + Err.Number: Err.Clear
    MmuDetalleActualizar
    failure = failure + Err.Number: Err.Clear
    LogsActualizar
    failure = failure + Err.Number: Err.Clear
    ArithPanelRefrescar
    failure = failure + Err.Number: Err.Clear
    EcoPanelRefrescar
    failure = failure + Err.Number: Err.Clear
    On Error GoTo failed
    Check failure = 0, "C: machine, CPU, memory, bus, MMU views and Logs refresh without errors"
    Check ws.Shapes("P04D_bit00").TextFrame2.TextRange.Text = "0" And ws.Shapes("P04D_bit07").TextFrame2.TextRange.Text = "1", _
          "C: AX = 0Fh drawn as bits 0000 1111"
    Check InStr(ws.Shapes("P12_kpi0").TextFrame2.TextRange.Text, "20") > 0, "C: indicator shows 20 instructions"
    If Not LogsSheet(False) Is Nothing Then
        Check InStr(CStr(LogsSheet(False).Range("B6").Value), "HLT") > 0, "C: Logs shows the HLT as newest line"
    End If
    ExecReset
    LcaReport "UI PASS: " & CStr(mChecks) & " comprobaciones / assertions." & vbCrLf & _
              "Cuadricula de bloques, hojas LCA + Logs, elementos de cada bloque y refresco de las vistas verificados.", True
    Exit Sub
failed:
    LcaReport "UI FAIL after " & CStr(mChecks) & " checks: " & Err.Description, False
End Sub

Private Function HasShape(ByVal ws As Worksheet, ByVal name As String) As Boolean
    Dim s As Shape
    For Each s In ws.Shapes
        If s.Name = name Then
            HasShape = True
            Exit Function
        End If
    Next s
End Function

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    mChecks = mChecks + 1
    If Not condition Then Err.Raise vbObjectError + 2151, "LcaUiTests", message
End Sub
