Attribute VB_Name = "LcaDiagrama"
Option Explicit

Private Const ROW_BASE As Long = LCA_BASE_DIAGRAM
Private Const COL_BASE As Long = LCA_COL_DIAGRAM
Private mHasOperation As Boolean
Private mOperation As String
Private mDestination As Long
Private mSource As Long
Private mLeft As Long
Private mRight As Long
Private mResult As Long
Private mWrites As Boolean

' Native vector shapes inspired by the local 8085 and bus lecture diagrams.
' Arrows show ALU dependencies, not simulated bus cycles (P05).
Public Sub ConstruirDiagramaCpu()
    Dim ws As Worksheet
    Dim index As Long
    Dim x As Single, y As Single, w As Single, h As Single
    Dim regY As Single
    Dim box As shape
    Dim form As FreeformBuilder
    Dim names As Variant
    Set ws = DiagramSheet(True)
    LcaPrepareBand ws, ROW_BASE, 42, "D_", COL_BASE
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    LcaRowHeight ws.Rows(RA("1:38")), 18
    ws.Range(RA("A1:AZ42")).Interior.color = RGB(242, 246, 251)
    ws.Range(RA("A1:T5")).Interior.color = RGB(15, 23, 42)
    ws.Range(RA("B2:N3")).Merge
    With ws.Range(RA("B2:N3"))
        .value = "LCA  /  ARQUITECTURA DEL CPU"
        .Font.name = "Segoe UI"
        .Font.size = 26
        .Font.bold = True
        .Font.color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
    End With
    ws.Range(RA("B4:S4")).Merge
    With ws.Range(RA("B4:S4"))
        .value = "CPU block diagram  |  Registros, ALU y flags / Registers, ALU and flags  |  P04"
        .Font.name = "Segoe UI"
        .Font.size = 13
        .Font.color = RGB(125, 211, 252)
    End With
    x = ws.Range(RA("B7")).Left
    y = ws.Range(RA("B7")).Top
    w = ws.Range(RA("B7:S29")).Width
    h = ws.Range(RA("B7:S29")).Height
    Set box = DiagramBox(ws, "D_Chip", x, y, w, h, "", RGB(255, 255, 255), RGB(180, 195, 212), 14)
    box.Line.Weight = 2
    DiagramBox ws, "D_ChipTitle", x + w * 0.02, y + h * 0.02, w * 0.26, 24, _
               "CPU LCA  -  8-bit", RGB(239, 246, 255), RGB(239, 246, 255), 17
    ' Short edge pins reinforce the processor outline without adding fake ports.
    For index = 1 To 8
        DiagramLine ws, "D_PinL" & index, x - 9, y + h * (0.1 + index * 0.09), x, y + h * (0.1 + index * 0.09), RGB(148, 163, 184), False
        DiagramLine ws, "D_PinR" & index, x + w, y + h * (0.1 + index * 0.09), x + w + 9, y + h * (0.1 + index * 0.09), RGB(148, 163, 184), False
    Next index
    DiagramBox ws, "D_BankTitle", x + w * 0.04, y + h * 0.13, w * 0.23, h * 0.07, _
               "BANCO / REGISTER BANK", RGB(37, 99, 235), RGB(37, 99, 235), 16, True
    For index = 0 To 3
        regY = y + h * (0.23 + index * 0.15)
        DiagramBox ws, "D_R" & index, x + w * 0.04, regY, w * 0.23, h * 0.12, _
                   "", RGB(239, 246, 255), RGB(147, 197, 253), 19
        DiagramLine ws, "D_RegLink" & index, x + w * 0.27, regY + h * 0.06, x + w * 0.32, regY + h * 0.06, RGB(37, 99, 235), False
    Next index
    DiagramLine ws, "D_OperandRoute", x + w * 0.32, y + h * 0.2, x + w * 0.32, y + h * 0.74, RGB(37, 99, 235), False
    DiagramLine ws, "D_ToA", x + w * 0.32, y + h * 0.2, x + w * 0.4, y + h * 0.2, RGB(37, 99, 235), True
    DiagramLine ws, "D_ToBTop", x + w * 0.32, y + h * 0.2, x + w * 0.32, y + h * 0.08, RGB(8, 145, 178), False
    DiagramLine ws, "D_ToB", x + w * 0.32, y + h * 0.08, x + w * 0.64, y + h * 0.08, RGB(8, 145, 178), False
    DiagramLine ws, "D_ToBDown", x + w * 0.64, y + h * 0.08, x + w * 0.64, y + h * 0.14, RGB(8, 145, 178), True
    DiagramBox ws, "D_A", x + w * 0.4, y + h * 0.14, w * 0.13, h * 0.13, _
               "A / --", RGB(239, 246, 255), RGB(37, 99, 235), 17
    DiagramBox ws, "D_B", x + w * 0.575, y + h * 0.14, w * 0.13, h * 0.13, _
               "B / --", RGB(236, 254, 255), RGB(8, 145, 178), 17
    DiagramLine ws, "D_AIntoAlu", x + w * 0.465, y + h * 0.27, x + w * 0.465, y + h * 0.39, RGB(37, 99, 235), True
    DiagramLine ws, "D_BIntoAlu", x + w * 0.64, y + h * 0.27, x + w * 0.64, y + h * 0.39, RGB(8, 145, 178), True
    ' Classic notched ALU outline, built as an editable freeform.
    Set form = ws.Shapes.BuildFreeform(msoEditingCorner, x + w * 0.4, y + h * 0.4)
    form.AddNodes msoSegmentLine, msoEditingCorner, x + w * 0.505, y + h * 0.4
    form.AddNodes msoSegmentLine, msoEditingCorner, x + w * 0.55, y + h * 0.48
    form.AddNodes msoSegmentLine, msoEditingCorner, x + w * 0.595, y + h * 0.4
    form.AddNodes msoSegmentLine, msoEditingCorner, x + w * 0.7, y + h * 0.4
    form.AddNodes msoSegmentLine, msoEditingCorner, x + w * 0.65, y + h * 0.7
    form.AddNodes msoSegmentLine, msoEditingCorner, x + w * 0.45, y + h * 0.7
    form.AddNodes msoSegmentLine, msoEditingCorner, x + w * 0.4, y + h * 0.4
    Set box = form.ConvertToShape
    box.name = "D_Alu"
    box.fill.ForeColor.RGB = RGB(255, 247, 237)
    box.Line.ForeColor.RGB = RGB(217, 119, 6)
    box.Line.Weight = 2.5
    Set box = DiagramBox(ws, "D_AluLabel", x + w * 0.46, y + h * 0.51, w * 0.18, h * 0.14, _
                         "ALU / READY", RGB(255, 247, 237), RGB(255, 247, 237), 20)
    box.Line.Visible = msoFalse
    DiagramLine ws, "D_ResultLine", x + w * 0.55, y + h * 0.7, x + w * 0.55, y + h * 0.8, RGB(5, 150, 105), True
    DiagramBox ws, "D_Result", x + w * 0.405, y + h * 0.8, w * 0.29, h * 0.15, _
               "RESULTADO / RESULT" & vbLf & "--", RGB(236, 253, 245), RGB(5, 150, 105), 18
    DiagramLine ws, "D_WriteBack1", x + w * 0.405, y + h * 0.875, x + w * 0.155, y + h * 0.875, RGB(5, 150, 105), False
    DiagramLine ws, "D_WriteBack2", x + w * 0.155, y + h * 0.875, x + w * 0.155, y + h * 0.8, RGB(5, 150, 105), True
    DiagramBox ws, "D_WriteLabel", x + w * 0.21, y + h * 0.9, w * 0.18, h * 0.06, _
               "WRITE-BACK", RGB(255, 255, 255), RGB(255, 255, 255), 12
    DiagramBox ws, "D_Control", x + w * 0.77, y + h * 0.14, w * 0.19, h * 0.37, _
               "", RGB(248, 250, 252), RGB(148, 163, 184), 16
    DiagramLine ws, "D_ToFlags1", x + w * 0.7, y + h * 0.56, x + w * 0.735, y + h * 0.56, RGB(217, 119, 6), False
    DiagramLine ws, "D_ToFlags2", x + w * 0.735, y + h * 0.56, x + w * 0.735, y + h * 0.73, RGB(217, 119, 6), False
    DiagramLine ws, "D_ToFlags3", x + w * 0.735, y + h * 0.73, x + w * 0.77, y + h * 0.73, RGB(217, 119, 6), True
    DiagramBox ws, "D_Flags", x + w * 0.77, y + h * 0.63, w * 0.19, h * 0.22, _
               "", RGB(255, 247, 237), RGB(217, 119, 6), 18
    DiagramBox ws, "D_State", x + w * 0.77, y + h * 0.89, w * 0.19, h * 0.06, _
               "P04  /  CPU STATE", RGB(255, 255, 255), RGB(255, 255, 255), 12
    names = Array("ADD", "SUB", "CMP", "INC", "DEC", "RESET")
    For index = 0 To 5
        Set box = DiagramBox(ws, "D_Button" & index, CL(ws, ROW_BASE + 32, 2 + index * 3).Left, ws.Rows(ROW_BASE + 32).Top, _
                             ws.Range(CL(ws, ROW_BASE + 32, 2 + index * 3), CL(ws, ROW_BASE + 33, 4 + index * 3)).Width - 8, _
                             ws.Range(RA("B32:B33")).Height, CStr(names(index)), RGB(37, 99, 235), RGB(37, 99, 235), 16, True)
        box.OnAction = "'" & ThisWorkbook.name & "'!" & CStr(ArrayMacro(index))
    Next index
    Set box = DiagramBox(ws, "D_OpenPanel", ws.Range(RA("O2")).Left, ws.Rows(ROW_BASE + 2).Top, ws.Range(RA("O2:S3")).Width, _
                         ws.Range(RA("O2:O3")).Height, "CONTROLES / CONTROLS", RGB(37, 99, 235), RGB(37, 99, 235), 15, True)
    box.OnAction = "'" & ThisWorkbook.name & "'!LcaIrCpu"
    ws.Range(RA("B35:S36")).Merge
    With ws.Range(RA("B35:S36"))
        .value = "Carga valores y elige registros en CONTROLES. Flechas = dependencias ALU; microciclos de buses: P05." & vbLf & _
                 "Load values and select registers in CONTROLS. Arrows = ALU dependencies; bus microcycles: P05."
        .Font.name = "Segoe UI"
        .Font.size = 12
        .Font.color = RGB(71, 85, 105)
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    DiagramaActualizar
    FitDiagram ws
End Sub

Public Sub MostrarDiagramaCpu()
    Dim ws As Worksheet
    Set ws = DiagramSheet(False)
    If ws Is Nothing Then
        ConstruirDiagramaCpu
    Else
        DiagramaActualizar
        FitDiagram ws
    End If
End Sub

Public Sub DiagramaRegistrarOperacion(ByVal operation As String, ByVal destination As Long, ByVal source As Long, _
                                     ByVal leftValue As Long, ByVal rightValue As Long, ByVal result As Long, ByVal writes As Boolean)
    mHasOperation = True
    mOperation = operation
    mDestination = destination
    mSource = source
    mLeft = leftValue
    mRight = rightValue
    mResult = result
    mWrites = writes
End Sub

Public Sub DiagramaReiniciar()
    mHasOperation = False
End Sub

Public Sub DiagramaActualizar()
    Dim ws As Worksheet
    Dim state As LcaCpuState
    Dim index As Long
    Dim names As Variant
    Dim value As Long
    Set ws = DiagramSheet(False)
    If ws Is Nothing Then Exit Sub
    state = CpuSnapshot()
    names = Array("AX", "BX", "CX", "DX")
    For index = 0 To 3
        value = CpuReadRegister(index)
        ws.Shapes("D_R" & index).TextFrame2.TextRange.text = CStr(names(index)) & "  |  " & value & "  /  " & Right$("00" & Hex$(value), 2) & "h"
        ws.Shapes("D_R" & index).fill.ForeColor.RGB = RGB(239, 246, 255)
    Next index
    ws.Shapes("D_Control").TextFrame2.TextRange.text = "CONTROL / STATE" & vbLf & _
        "PC       " & Right$("000" & Hex$(state.PC), 3) & "h" & vbLf & _
        "MAR_V  " & Right$("000" & Hex$(state.MAR_V), 3) & "h" & vbLf & _
        "MAR_P  " & Right$("00" & Hex$(state.MAR_P), 2) & "h" & vbLf & _
        "MDR      " & Right$("00" & Hex$(state.MDR), 2) & "h" & vbLf & _
        "IR          " & Right$("00" & Hex$(state.IR), 2) & "h"
    ws.Shapes("D_Flags").TextFrame2.TextRange.text = "FLAGS" & vbLf & "ZF " & Abs(CLng(CpuFlagIsSet(LCA_FLAG_ZF))) & _
        "   CF " & Abs(CLng(CpuFlagIsSet(LCA_FLAG_CF))) & "   SF " & Abs(CLng(CpuFlagIsSet(LCA_FLAG_SF)))
    If mHasOperation Then
        ws.Shapes("D_R" & mDestination).fill.ForeColor.RGB = RGB(219, 234, 254)
        ws.Shapes("D_A").TextFrame2.TextRange.text = "A: " & CStr(names(mDestination)) & " = " & CStr(mLeft)
        ws.Shapes("D_B").TextFrame2.TextRange.text = "B: " & IIf(mOperation = "INC" Or mOperation = "DEC", "CONST", CStr(names(mSource))) & " = " & CStr(mRight)
        ws.Shapes("D_AluLabel").TextFrame2.TextRange.text = "ALU / " & mOperation
        ws.Shapes("D_Result").TextFrame2.TextRange.text = "RESULTADO / RESULT" & vbLf & mResult & " / " & Right$("00" & Hex$(mResult), 2) & "h"
        ws.Shapes("D_WriteLabel").TextFrame2.TextRange.text = IIf(mWrites, "WRITE-BACK -> " & CStr(names(mDestination)), "CMP: SOLO FLAGS / FLAGS ONLY")
        ws.Shapes("D_WriteBack1").Visible = IIf(mWrites, msoTrue, msoFalse)
        ws.Shapes("D_WriteBack2").Visible = IIf(mWrites, msoTrue, msoFalse)
    Else
        ws.Shapes("D_A").TextFrame2.TextRange.text = "A / --"
        ws.Shapes("D_B").TextFrame2.TextRange.text = "B / --"
        ws.Shapes("D_AluLabel").TextFrame2.TextRange.text = "ALU / READY"
        ws.Shapes("D_Result").TextFrame2.TextRange.text = "RESULTADO / RESULT" & vbLf & "--"
        ws.Shapes("D_WriteLabel").TextFrame2.TextRange.text = "WRITE-BACK"
        ws.Shapes("D_WriteBack1").Visible = msoFalse
        ws.Shapes("D_WriteBack2").Visible = msoFalse
    End If
End Sub

Private Function ArrayMacro(ByVal index As Long) As String
    Dim names As Variant
    names = Array("PanelSumar", "PanelRestar", "PanelComparar", "PanelIncrementar", "PanelDecrementar", "PanelReset")
    ArrayMacro = CStr(names(index))
End Function

Private Function DiagramSheet(ByVal create As Boolean) As Worksheet
    Set DiagramSheet = LcaBand("D_Chip", create)
End Function

' Row shift of this diagram inside the shared LCA sheet.
Private Function RA(ByVal address As String) As String
    RA = LcaShift(address, ROW_BASE, COL_BASE)
End Function

Private Sub FitDiagram(ByVal ws As Worksheet)
    LcaGoto ws, ROW_BASE, COL_BASE
End Sub

Private Function DiagramBox(ByVal ws As Worksheet, ByVal name As String, ByVal x As Single, ByVal y As Single, _
                            ByVal w As Single, ByVal h As Single, ByVal text As String, ByVal fill As Long, _
                            ByVal edge As Long, ByVal size As Long, Optional ByVal white As Boolean = False) As shape
    Dim box As shape
    Set box = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    box.name = name
    box.Adjustments.Item(1) = 0.08
    box.fill.ForeColor.RGB = fill
    box.Line.ForeColor.RGB = edge
    box.Line.Weight = 1.5
    With box.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .MarginTop = 1
        .MarginBottom = 1
        .MarginLeft = 5
        .MarginRight = 5
        .TextRange.text = text
        .TextRange.Font.name = "Segoe UI"
        .TextRange.Font.size = size
        .TextRange.Font.bold = msoTrue
        .TextRange.Font.fill.ForeColor.RGB = IIf(white, RGB(255, 255, 255), RGB(30, 41, 59))
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextRange.ParagraphFormat.SpaceAfter = 0
        .TextRange.ParagraphFormat.SpaceBefore = 0
    End With
    Set DiagramBox = box
End Function

Private Sub DiagramLine(ByVal ws As Worksheet, ByVal name As String, ByVal x1 As Single, ByVal y1 As Single, _
                        ByVal x2 As Single, ByVal y2 As Single, ByVal color As Long, ByVal arrow As Boolean)
    Dim wire As shape
    Set wire = ws.Shapes.AddConnector(msoConnectorStraight, x1, y1, x2, y2)
    wire.name = name
    wire.Line.ForeColor.RGB = color
    wire.Line.Weight = 2.5
    If arrow Then wire.Line.EndArrowheadStyle = msoArrowheadTriangle
End Sub

' Cell of this block: row is absolute (ROW_BASE already added), column is relative to the block.
Private Function CL(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Range
    Set CL = ws.Cells(r, c + COL_BASE)
End Function
