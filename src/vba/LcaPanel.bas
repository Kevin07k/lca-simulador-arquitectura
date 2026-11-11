Attribute VB_Name = "LcaPanel"
Option Explicit

Private Const ROW_BASE As Long = LCA_BASE_CPU
Private Const COL_BASE As Long = LCA_COL_CPU
Private Const NAVY As Long = &H25150B
Private Const CARD As Long = &H3A291B
Private Const INK As Long = &HF8F1E7
Private Const MUTED As Long = &HBEA893
Private Const BLUE As Long = &HE67B36
Private Const CYAN As Long = &HCECB49

' Presentation adapter: the CPU/ALU remain independent of Excel cells.
Public Sub ConstruirPanelCpu()
    Dim ws As Worksheet
    Dim index As Long
    Dim startCol As Long
    Dim row As Long
    Dim names As Variant
    Dim labels As Variant
    Dim macroNames As Variant
    Set ws = FindPanel(True)
    LcaPrepareBand ws, ROW_BASE, 44, "P04_", COL_BASE
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    LcaRowHeight ws.Rows(RA("1:40")), 17
    LcaRowHeight ws.Rows(ROW_BASE + 1), 12
    LcaRowHeight ws.Rows(ROW_BASE + 2), 25
    LcaRowHeight ws.Rows(ROW_BASE + 3), 23
    With ws.Range(RA("A1:S40"))
        .Interior.color = NAVY
        .Font.name = "Segoe UI"
        .Font.size = 11
        .Font.color = INK
        .VerticalAlignment = xlCenter
    End With
    ws.Range(RA("T1:AZ44")).Interior.color = NAVY
    ws.Range(RA("A41:S44")).Interior.color = NAVY
    ws.Range(RA("B2:D4")).Merge
    With ws.Range(RA("B2:D4"))
        .value = "LCA"
        .Interior.color = BLUE
        .Font.size = 30
        .Font.bold = True
        .HorizontalAlignment = xlCenter
    End With
    TextLine ws, "E2:M3", "CPU / ALU", 26, INK, True
    TextLine ws, "E4:M4", "Laboratorio interactivo / Interactive lab", 12, MUTED
    AddButton ws, "N2:R3", "DIAGRAMA / DIAGRAM", "LcaIrDiagrama", BLUE
    TextLine ws, "B6:M6", "01   REGISTROS / REGISTERS", 11, CYAN, True
    TextLine ws, "N6:R6", "P04   |   8-bit DATA / 9-bit VA", 10, MUTED
    names = Array("AX", "BX", "CX", "DX")
    For index = 0 To 3
        startCol = 2 + 4 * index
        With ws.Range(CL(ws, ROW_BASE + 8, startCol), CL(ws, ROW_BASE + 11, startCol + 3))
            .Interior.color = CARD
            .Borders(xlEdgeRight).color = NAVY
            .Borders(xlEdgeRight).Weight = xlThick
        End With
        TextLine ws, ws.Range(CL(ws, ROW_BASE + 8, startCol), CL(ws, ROW_BASE + 8, startCol + 3)).address, _
                 CStr(names(index)), 12, CYAN, True
        With ws.Range(CL(ws, ROW_BASE + 9, startCol), CL(ws, ROW_BASE + 10, startCol + 3))
            .Merge
            .Font.size = 29
            .Font.bold = True
            .HorizontalAlignment = xlCenter
        End With
        With ws.Range(CL(ws, ROW_BASE + 11, startCol), CL(ws, ROW_BASE + 11, startCol + 3))
            .Merge
            .Font.name = "Consolas"
            .Font.size = 12
            .Font.color = MUTED
            .HorizontalAlignment = xlCenter
        End With
    Next index
    TextLine ws, "B13:G13", "02   FLAGS", 11, CYAN, True
    TextLine ws, "H13:R13", "CONTROL / CONTROL REGISTERS", 11, CYAN, True
    ws.Range(RA("B14:R16")).Interior.color = CARD
    TextLine ws, "B14:C16", "ZF  0", 16, MUTED, True
    TextLine ws, "D14:E16", "CF  0", 16, MUTED, True
    TextLine ws, "F14:G16", "SF  0", 16, MUTED, True
    names = Array("PC", "MAR_V", "MAR_P", "MDR", "IR")
    For index = 0 To 4
        startCol = 8 + index * 2
        TextLine ws, ws.Range(CL(ws, ROW_BASE + 14, startCol), CL(ws, ROW_BASE + 14, startCol + 1)).address, _
                 CStr(names(index)), 10, MUTED, True
        TextLine ws, ws.Range(CL(ws, ROW_BASE + 15, startCol), CL(ws, ROW_BASE + 16, startCol + 1)).address, _
                 "000h", 15, INK, True
    Next index
    TextLine ws, "B18:R18", "03   OPERAR / OPERATE", 11, CYAN, True
    TextLine ws, "B19:D19", "REGISTRO / REGISTER", 10, MUTED
    TextLine ws, "E19:G19", "VALOR / VALUE  0-255", 10, MUTED
    TextLine ws, "H19:J19", "DESTINO / DESTINATION", 10, MUTED
    TextLine ws, "K19:M19", "ORIGEN / SOURCE", 10, MUTED
    TextLine ws, "B20:D21", "AX", 17, NAVY, True
    TextLine ws, "E20:G21", "0", 17, NAVY, True
    TextLine ws, "H20:J21", "AX", 17, NAVY, True
    TextLine ws, "K20:M21", "BX", 17, NAVY, True
    ws.Range(RA("B20:M21")).Interior.color = RGB(230, 240, 250)
    AddRegisterValidation ws.Range(RA("B20:D21"))
    AddRegisterValidation ws.Range(RA("H20:J21"))
    AddRegisterValidation ws.Range(RA("K20:M21"))
    With ws.Range(RA("E20:G21")).Validation
        .Delete
        .Add xlValidateWholeNumber, xlValidAlertStop, xlBetween, "0", "255"
        .IgnoreBlank = False
        .ErrorTitle = "Byte 0-255"
        .ErrorMessage = "Solo enteros / Integers only: 0-255"
        .ShowError = True
    End With
    AddButton ws, "N20:R21", "CARGAR / LOAD", "PanelCargar", BLUE
    TextLine ws, "B22:R22", "1  Carga un registro.  2  Elige destino/origen.  3  Ejecuta la ALU. / Load, select, execute.", 10, MUTED
    labels = Array("ADD", "SUB", "CMP", "INC", "DEC")
    macroNames = Array("PanelSumar", "PanelRestar", "PanelComparar", "PanelIncrementar", "PanelDecrementar")
    For index = 0 To 4
        startCol = 2 + index * 3
        AddButton ws, ws.Range(CL(ws, ROW_BASE + 24, startCol), CL(ws, ROW_BASE + 25, startCol + 2)).address, _
                  CStr(labels(index)), CStr(macroNames(index)), BLUE
    Next index
    AddButton ws, "Q24:R25", "RESET", "PanelReset", RGB(78, 96, 120)
    TextLine ws, "B27:R27", "LISTO / READY", 11, CYAN, True
    TextLine ws, "B29:R29", "04   ULTIMAS ACCIONES / RECENT ACTIONS", 11, CYAN, True
    names = Array("ACCION / ACTION", "DESTINO / DEST", "ANTES / BEFORE", "RESULTADO / RESULT", "FLAGS  Z / C / S")
    For index = 0 To 4
        startCol = 2 + 3 * index
        TextLine ws, ws.Range(CL(ws, ROW_BASE + 30, startCol), CL(ws, ROW_BASE + 30, IIf(index = 4, 18, startCol + 2))).address, _
                 CStr(names(index)), 9, MUTED, True
        For row = 31 To 35
            TextLine ws, ws.Range(CL(ws, ROW_BASE + row, startCol), CL(ws, ROW_BASE + row, IIf(index = 4, 18, startCol + 2))).address, _
                     "", 10, INK
        Next row
    Next index
    ws.Range(RA("B31:R35")).Interior.color = CARD
    TextLine ws, "B37:L38", "CPU + ALU reales / Live CPU + ALU" & vbLf & _
             "Laboratorio manual. Ciclo de instrucciones: P12 / Instruction cycle: P12", 10, MUTED
    ws.Range(RA("B37:L38")).WrapText = True
    AddButton ws, "M37:R38", "VERIFICAR / RUN TESTS", "PanelVerificar", RGB(78, 96, 120)
    CpuReset
    DiagramaReiniciar
    PanelActualizar
    FitPanel ws
End Sub

Public Sub MostrarPanelCpu()
    Dim ws As Worksheet
    Set ws = FindPanel(False)
    If ws Is Nothing Then
        ConstruirPanelCpu
    Else
        PanelActualizar
        FitPanel ws
    End If
End Sub

Private Sub FitPanel(ByVal ws As Worksheet)
    LcaGoto ws, ROW_BASE, COL_BASE
    ws.Range(RA("B20")).Select
End Sub

Public Sub PanelActualizar()
    Dim ws As Worksheet
    Dim state As LcaCpuState
    Dim index As Long
    Dim value As Long
    Dim values As Variant
    Set ws = FindPanel(False)
    If ws Is Nothing Then Exit Sub
    state = CpuSnapshot()
    For index = 0 To 3
        value = CpuReadRegister(index)
        CL(ws, ROW_BASE + 9, 2 + index * 4).value = value
        CL(ws, ROW_BASE + 11, 2 + index * 4).value = Right$("00" & Hex$(value), 2) & "h   " & BinaryByte(value)
    Next index
    PaintFlag ws.Range(RA("B14:C16")), "ZF", CpuFlagIsSet(LCA_FLAG_ZF)
    PaintFlag ws.Range(RA("D14:E16")), "CF", CpuFlagIsSet(LCA_FLAG_CF)
    PaintFlag ws.Range(RA("F14:G16")), "SF", CpuFlagIsSet(LCA_FLAG_SF)
    values = Array(state.PC, state.MAR_V, state.MAR_P, state.MDR, state.IR)
    For index = 0 To 4
        CL(ws, ROW_BASE + 15, 8 + index * 2).value = Right$("000" & Hex$(values(index)), IIf(index < 2, 3, 2)) & "h"
    Next index
    DiagramaActualizar
    CpuDetalleActualizar
End Sub

Public Sub PanelCargar()
    Dim ws As Worksheet
    Dim code As Long
    Dim number As Double
    Dim previous As Long
    On Error GoTo failed
    Set ws = FindPanel(False)
    code = RegisterCode(ws.Range(RA("B20")).value)
    If IsError(ws.Range(RA("E20")).value) Then Err.Raise 5, , "BYTE: 0-255"
    If Not IsNumeric(ws.Range(RA("E20")).value) Or IsEmpty(ws.Range(RA("E20")).value) Then Err.Raise 5, , "BYTE: 0-255"
    number = CDbl(ws.Range(RA("E20")).value)
    If number < 0 Or number > 255 Or number <> Fix(number) Then Err.Raise 5, , "BYTE: 0-255"
    previous = CpuReadRegister(code)
    CpuWriteRegister code, CLng(number)
    PanelActualizar
    AddTrace "LOAD", RegisterName(code), previous, CLng(number)
    StatusLine "CARGADO / LOADED  " & RegisterName(code) & " = " & CLng(number), False
    Exit Sub
failed:
    StatusLine "ERROR: " & Err.Description, True
End Sub

Public Sub PanelSumar()
    RunOperation LcaADD, "ADD"
End Sub

Public Sub PanelRestar()
    RunOperation LcaSUB, "SUB"
End Sub

Public Sub PanelComparar()
    RunOperation LcaCMP, "CMP"
End Sub

Public Sub PanelIncrementar()
    RunOperation LcaINC, "INC"
End Sub

Public Sub PanelDecrementar()
    RunOperation LcaDEC, "DEC"
End Sub

Private Sub RunOperation(ByVal operation As Long, ByVal label As String)
    Dim ws As Worksheet
    Dim destination As Long
    Dim source As Long
    Dim previous As Long
    Dim result As Long
    Dim rightValue As Long
    On Error GoTo failed
    Set ws = FindPanel(False)
    destination = RegisterCode(ws.Range(RA("H20")).value)
    If operation = LcaADD Or operation = LcaSUB Or operation = LcaCMP Then
        source = RegisterCode(ws.Range(RA("K20")).value)
    End If
    previous = CpuReadRegister(destination)
    If operation = LcaINC Or operation = LcaDEC Then
        rightValue = 1
    Else
        rightValue = CpuReadRegister(source)
    End If
    result = AluExecute(operation, destination, source)
    DiagramaRegistrarOperacion label, destination, source, previous, rightValue, result, (operation <> LcaCMP)
    PanelActualizar
    AddTrace label, RegisterName(destination), previous, result
    If operation = LcaCMP Then
        StatusLine "CMP: resultado / result = " & result & "; registro conservado / register unchanged", False
    Else
        StatusLine label & "  " & RegisterName(destination) & " = " & result & "  |  PC sin cambio / unchanged", False
    End If
    Exit Sub
failed:
    StatusLine "ERROR: " & Err.Description, True
End Sub

Public Sub PanelReset()
    CpuReset
    DiagramaReiniciar
    PanelActualizar
    AddTrace "RESET", "ALL", 0, 0
    StatusLine "RESET  |  CPU y flags en cero / CPU and flags cleared", False
End Sub

Public Sub PanelVerificar()
    On Error GoTo failed
    VerificarCpuAlu
    DiagramaReiniciar
    MostrarPanelCpu
    If LcaLastPassed() Then
        StatusLine "PASS  |  197120 casos / cases  -  1379895 comprobaciones / assertions", False
    Else
        StatusLine LcaLastReport(), True
    End If
    Exit Sub
failed:
    MostrarPanelCpu
    StatusLine "FAIL: " & Err.Description, True
End Sub

Private Function FindPanel(ByVal create As Boolean) As Worksheet
    Set FindPanel = LcaBand("P04_LcaIrDiagrama", create)
End Function

' Row shift of this panel inside the shared LCA sheet.
Private Function RA(ByVal address As String) As String
    RA = LcaShift(address, ROW_BASE, COL_BASE)
End Function

Private Sub TextLine(ByVal ws As Worksheet, ByVal address As String, ByVal text As String, _
                     ByVal size As Long, ByVal color As Long, Optional ByVal bold As Boolean = False)
    With ws.Range(RA(address))
        .Merge
        .value = text
        .Font.size = size * 1.25
        .Font.color = color
        .Font.bold = bold
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal address As String, ByVal text As String, _
                      ByVal macroName As String, ByVal color As Long)
    Dim area As Range
    Dim button As shape
    Set area = ws.Range(RA(address))
    Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left + 3, area.Top + 2, area.Width - 6, area.Height - 4)
    button.name = "P04_" & macroName
    button.fill.ForeColor.RGB = color
    button.Line.Visible = msoFalse
    button.Adjustments.Item(1) = 0.12
    button.OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!" & macroName
    With button.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.text = text
        .TextRange.Font.name = "Segoe UI"
        .TextRange.Font.size = 14
        .TextRange.Font.bold = msoTrue
        .TextRange.Font.fill.ForeColor.RGB = INK
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
End Sub

Private Sub AddRegisterValidation(ByVal area As Range)
    With area.Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, xlBetween, Join(Array("AX", "BX", "CX", "DX"), Application.International(xlListSeparator))
        .InCellDropdown = True
        .IgnoreBlank = False
        .ShowError = True
    End With
End Sub

Private Function RegisterCode(ByVal value As Variant) As Long
    If IsError(value) Then Err.Raise LCA_ERR_REGISTER, , "INVALID_REGISTER"
    Select Case UCase$(Trim$(CStr(value)))
        Case "AX": RegisterCode = LcaAX
        Case "BX": RegisterCode = LcaBX
        Case "CX": RegisterCode = LcaCX
        Case "DX": RegisterCode = LcaDX
        Case Else: Err.Raise LCA_ERR_REGISTER, , "INVALID_REGISTER"
    End Select
End Function

Private Function RegisterName(ByVal code As Long) As String
    Dim names As Variant
    names = Array("AX", "BX", "CX", "DX")
    RegisterName = CStr(names(code))
End Function

Private Function BinaryByte(ByVal value As Long) As String
    Dim bit As Long
    For bit = 7 To 0 Step -1
        BinaryByte = BinaryByte & IIf((value And CLng(2 ^ bit)) <> 0, "1", "0")
    Next bit
End Function

Private Sub PaintFlag(ByVal area As Range, ByVal label As String, ByVal enabled As Boolean)
    area.Cells(1, 1).value = label & "  " & IIf(enabled, "1", "0")
    area.Font.color = IIf(enabled, CYAN, MUTED)
End Sub

Private Sub StatusLine(ByVal text As String, ByVal failed As Boolean)
    Dim ws As Worksheet
    Set ws = FindPanel(False)
    If ws Is Nothing Then Exit Sub
    ws.Range(RA("B27")).value = text
    ws.Range(RA("B27:R27")).Font.color = IIf(failed, RGB(255, 133, 133), CYAN)
End Sub

Private Sub AddTrace(ByVal action As String, ByVal destination As String, ByVal previous As Long, ByVal result As Long)
    Dim ws As Worksheet
    Dim row As Long
    Dim col As Variant
    Set ws = FindPanel(False)
    For row = 35 To 32 Step -1
        For Each col In Array(2, 5, 8, 11, 14)
            CL(ws, ROW_BASE + row, CLng(col)).value = CL(ws, ROW_BASE + row - 1, CLng(col)).value
        Next col
    Next row
    ws.Range(RA("B31")).value = action
    ws.Range(RA("E31")).value = destination
    ws.Range(RA("H31")).value = previous
    ws.Range(RA("K31")).value = result
    ws.Range(RA("N31")).value = "Z=" & Abs(CLng(CpuFlagIsSet(LCA_FLAG_ZF))) & "  C=" & _
                              Abs(CLng(CpuFlagIsSet(LCA_FLAG_CF))) & "  S=" & Abs(CLng(CpuFlagIsSet(LCA_FLAG_SF)))
End Sub





' Cell of this block: row is absolute (ROW_BASE already added), column is relative to the block.
Private Function CL(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Range
    Set CL = ws.Cells(r, c + COL_BASE)
End Function
