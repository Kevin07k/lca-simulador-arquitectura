Attribute VB_Name = "LcaMicroPanel"
Option Explicit

' P05b band of the LCA sheet: the control word of the last executed instruction, one clock micro-step at a time.
' Shows the 25 signals (Enable / Load / Control) as lamps, the tri-state buffer of every device (conducting or
' High-Z, never two at once) and the whole micro-sequence. Buttons walk through the micro-steps or animate them.
' Everything shown is read from LcaMicro, which the executor fills after every instruction.
Private Const ROW_BASE As Long = LCA_BASE_MICRO
Private Const COL_BASE As Long = LCA_COL_MICRO
Private Const BAND_ROWS As Long = 48
Private Const FIRST_SEQ_ROW As Long = 31
Private mSel As Long
Private mSeenSerial As Long

Private Const C_PANEL As Long = 3877150        ' RGB(30, 41, 59)
Private Const C_OFF_TXT As Long = 9139300      ' RGB(100, 116, 139)

Public Sub ConstruirPanelMicro()
    Dim ws As Worksheet, i As Long, v As Variant
    Set ws = LcaSheet(True)
    LcaPrepareBand ws, ROW_BASE, BAND_ROWS, "P05B_", COL_BASE
    LcaRowHeight ws.Rows(RA("1:48")), 23
    ws.Range(RA("A1:T48")).Interior.Color = RGB(15, 23, 42)
    ws.Range(RA("A1:T48")).Font.Name = "Consolas"
    ws.Range(RA("A1:T48")).Font.Color = RGB(226, 232, 240)
    ws.Range(RA("B2:S3")).Merge
    With ws.Range(RA("B2"))
        .Value = "P05b  /  MICRO-PALABRA DE CONTROL  /  CONTROL WORD"
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 22
        .Font.Color = RGB(125, 211, 252)
    End With
    ws.Range(RA("B4:S4")).Merge
    ws.Range(RA("B4")).Value = "Un vector de 25 senales por pulso de reloj: Enable (buffer tri-estado hacia el bus), Load (receptor) y Control. Solo un Enable por micro-paso; el resto en High-Z."
    ws.Range(RA("B4")).WrapText = True
    ws.Rows(ROW_BASE + 4).RowHeight = 34
    AddBtn ws, "B6:E7", "< MICRO-PASO", "MicroAnterior"
    AddBtn ws, "F6:J7", "MICRO-PASO >", "MicroSiguiente"
    AddBtn ws, "K6:O7", "ANIMAR / ANIMATE", "MicroAnimar"
    AddBtn ws, "P6:S7", "< MAQUINA", "LcaIrMaquina"

    ws.Range(RA("B9:S9")).Merge
    ws.Range(RA("B10:S10")).Merge
    ws.Range(RA("B9:S10")).Interior.Color = C_PANEL
    ws.Range(RA("B9")).Font.Bold = True
    ws.Range(RA("B9")).Font.Size = 13
    ws.Range(RA("B10")).Font.Color = RGB(125, 211, 252)

    Header ws, 12, "ENABLE  -  buffer tri-estado de salida (1 = conduce el bus, 0 = High-Z)"
    Header ws, 16, "LOAD  -  el receptor captura el dato del bus"
    Header ws, 20, "CONTROL"
    Header ws, 24, "BUFFERS TRI-ESTADO DE CADA DISPOSITIVO  (verde = conduce, gris = High-Z)"
    For i = 0 To MIC_ENABLES - 1
        CL(ws, ROW_BASE + 13, 3 + i).Value = MicroSignalName(i)
        CL(ws, ROW_BASE + 25, 3 + i).Value = MicroDeviceName(i)
    Next i
    For i = MIC_LD_PC To MIC_LD_LCD
        CL(ws, ROW_BASE + 17, 3 + i - MIC_LD_PC).Value = MicroSignalName(i)
    Next i
    For i = MIC_MEM_R To MIC_INC_PC
        CL(ws, ROW_BASE + 21, 3 + i - MIC_MEM_R).Value = MicroSignalName(i)
    Next i
    CL(ws, ROW_BASE + 21, 8).Value = "ALU op"
    ws.Range(RA("B13")).Value = "senal"
    ws.Range(RA("B14")).Value = "valor"
    ws.Range(RA("B17")).Value = "senal"
    ws.Range(RA("B18")).Value = "valor"
    ws.Range(RA("B21")).Value = "senal"
    ws.Range(RA("B22")).Value = "valor"
    ws.Range(RA("B25")).Value = "dispositivo"
    ws.Range(RA("B26")).Value = "estado"
    For Each v In Array(13, 14, 17, 18, 21, 22, 25, 26)
        With ws.Range(CL(ws, ROW_BASE + CLng(v), 2), CL(ws, ROW_BASE + CLng(v), 13))
            .HorizontalAlignment = xlCenter
            .Borders.Color = RGB(51, 65, 85)
            .Borders.Weight = xlThin
        End With
    Next v
    For Each v In Array(13, 17, 21, 25)
        ws.Range(CL(ws, ROW_BASE + CLng(v), 2), CL(ws, ROW_BASE + CLng(v), 13)).Font.Size = 9
        ws.Range(CL(ws, ROW_BASE + CLng(v), 2), CL(ws, ROW_BASE + CLng(v), 13)).Font.Color = RGB(148, 163, 184)
    Next v
    ws.Range(RA("B27:S27")).Merge
    ws.Range(RA("B27")).Interior.Color = C_PANEL
    ws.Range(RA("B27")).Font.Bold = True

    Header ws, 29, "SECUENCIA COMPLETA DE LA ULTIMA INSTRUCCION  (vector = Enable | Load | Control)"
    ws.Range(RA("B30")).Value = "T"
    ws.Range(RA("C30")).Value = "MICRO-OPERACION"
    ws.Range(RA("L30")).Value = "En(10) Ld(11) Ctl(4)"
    ws.Range(RA("B30:S30")).Font.Bold = True
    ws.Range(RA("B30:S30")).Font.Color = RGB(148, 163, 184)
    For i = 0 To MIC_MAX_STEPS - 1
        ws.Range(CL(ws, ROW_BASE + FIRST_SEQ_ROW + i, 3), CL(ws, ROW_BASE + FIRST_SEQ_ROW + i, 11)).Merge
        ws.Range(CL(ws, ROW_BASE + FIRST_SEQ_ROW + i, 12), CL(ws, ROW_BASE + FIRST_SEQ_ROW + i, 19)).Merge
    Next i
    ws.Range(RA("B" & CStr(FIRST_SEQ_ROW) & ":S" & CStr(FIRST_SEQ_ROW + MIC_MAX_STEPS - 1))).Interior.Color = C_PANEL
    ws.Range(RA("B46:S47")).Merge
    ws.Range(RA("B46")).WrapText = True
    ws.Range(RA("B46")).Font.Color = RGB(125, 211, 252)
    mSel = 0
    mSeenSerial = -1
    MicroPanelRefrescar
    LcaGoto ws, ROW_BASE, COL_BASE
End Sub

Public Sub MostrarPanelMicro()
    If LcaBand("P05B_MicroSiguiente", False) Is Nothing Then
        ConstruirPanelMicro
    Else
        LcaGoto LcaSheet(True), ROW_BASE, COL_BASE
        MicroPanelRefrescar
    End If
End Sub

Public Function MicroPanelConstruido() As Boolean
    MicroPanelConstruido = Not LcaBand("P05B_MicroSiguiente", False) Is Nothing
End Function

Public Sub MicroSiguiente()
    If MicroCount() = 0 Then Exit Sub
    mSel = (mSel + 1) Mod MicroCount()
    MicroPanelRefrescar
End Sub

Public Sub MicroAnterior()
    If MicroCount() = 0 Then Exit Sub
    mSel = (mSel + MicroCount() - 1) Mod MicroCount()
    MicroPanelRefrescar
End Sub

' Lights the micro-steps one after another, like the clock ticking through the instruction.
Public Sub MicroAnimar()
    Dim i As Long, t As Double
    If MicroCount() = 0 Then Exit Sub
    For i = 0 To MicroCount() - 1
        mSel = i
        MicroPanelRefrescar
        DoEvents
        t = Timer
        Do While Timer - t < 0.6 And Timer >= t
            DoEvents
        Loop
    Next i
End Sub

Public Function MicroPanelPaso() As Long
    MicroPanelPaso = mSel
End Function

Public Sub MicroPanelSeleccionar(ByVal i As Long)
    If MicroCount() = 0 Then Exit Sub
    If i < 0 Then i = 0
    If i >= MicroCount() Then i = MicroCount() - 1
    mSel = i
    MicroPanelRefrescar
End Sub

' Redraws the band from LcaMicro. A new instruction restarts the selection at T1.
Public Sub MicroPanelRefrescar()
    Dim ws As Worksheet, i As Long, n As Long, s As Long
    Set ws = LcaBand("P05B_MicroSiguiente", False)
    If ws Is Nothing Then Exit Sub
    n = MicroCount()
    If MicroSerial() <> mSeenSerial Then
        mSeenSerial = MicroSerial()
        mSel = 0
    End If
    If mSel >= n Then mSel = 0
    If n = 0 Then
        ws.Range(RA("B9")).Value = "Sin instruccion ejecutada"
        ws.Range(RA("B10")).Value = "Cargue un programa en la barra de la maquina y pulse PASO: aqui aparece su micro-palabra."
        For i = 0 To MIC_ENABLES - 1
            Lamp ws, 14, 3 + i, False, 1
            Buf ws, 3 + i, False
        Next i
        For i = 0 To MIC_LD_LCD - MIC_LD_PC
            Lamp ws, 18, 3 + i, False, 2
        Next i
        For i = 0 To 3
            Lamp ws, 22, 3 + i, False, 3
        Next i
        CL(ws, ROW_BASE + 22, 8).Value = "-"
        ws.Range(RA("B27")).Value = "Bus en reposo: ningun buffer conduce (High-Z)."
        For i = 0 To MIC_MAX_STEPS - 1
            ws.Range(CL(ws, ROW_BASE + FIRST_SEQ_ROW + i, 2), CL(ws, ROW_BASE + FIRST_SEQ_ROW + i, 19)).Value = vbNullString
        Next i
        ws.Range(RA("B46")).Value = "Regla fisica: solo un Enable por micro-paso."
        Exit Sub
    End If

    ws.Range(RA("B9")).Value = "#" & CStr(ExecSteps()) & "  PC=" & Right$("000" & Hex$(MicroInstructionPC()), 3) & "h   " & MicroInstruction() & _
                               "     (" & CStr(n) & " micro-pasos / micro-steps)"
    ws.Range(RA("B10")).Value = "micro-paso " & CStr(mSel + 1) & " de " & CStr(n) & ":  " & MicroLabel(mSel) & _
                                IIf(Len(MicroBusData(mSel)) > 0, "    dato / data: " & MicroBusData(mSel), vbNullString)
    For i = 0 To MIC_ENABLES - 1
        Lamp ws, 14, 3 + i, MicroSignal(mSel, i), 1
        Buf ws, 3 + i, MicroSignal(mSel, i)
    Next i
    For i = MIC_LD_PC To MIC_LD_LCD
        Lamp ws, 18, 3 + i - MIC_LD_PC, MicroSignal(mSel, i), 2
    Next i
    For i = MIC_MEM_R To MIC_INC_PC
        Lamp ws, 22, 3 + i - MIC_MEM_R, MicroSignal(mSel, i), 3
    Next i
    CL(ws, ROW_BASE + 22, 8).Value = IIf(Len(MicroAluOp(mSel)) > 0, MicroAluOp(mSel), "-")
    With CL(ws, ROW_BASE + 22, 8)
        .Interior.Color = IIf(Len(MicroAluOp(mSel)) > 0, RGB(168, 85, 247), C_PANEL)
        .Font.Color = IIf(Len(MicroAluOp(mSel)) > 0, RGB(255, 255, 255), C_OFF_TXT)
        .Font.Bold = True
    End With
    If MicroEnableCount(mSel) = 0 Then
        ws.Range(RA("B27")).Value = "Bus: High-Z (ningun buffer conduce)   ->   receptores: " & MicroReceivers(mSel)
    Else
        ws.Range(RA("B27")).Value = "Bus: " & MicroDriver(mSel) & " conduce (los otros 9 en High-Z)   ->   receptores: " & MicroReceivers(mSel)
    End If
    For i = 0 To MIC_MAX_STEPS - 1
        If i < n Then
            CL(ws, ROW_BASE + FIRST_SEQ_ROW + i, 2).Value = "T" & CStr(i + 1)
            CL(ws, ROW_BASE + FIRST_SEQ_ROW + i, 3).Value = Mid$(MicroLabel(i), InStr(MicroLabel(i), " ") + 2)
            CL(ws, ROW_BASE + FIRST_SEQ_ROW + i, 12).Value = "'" & MicroWord(i)
        Else
            CL(ws, ROW_BASE + FIRST_SEQ_ROW + i, 2).Value = vbNullString
            CL(ws, ROW_BASE + FIRST_SEQ_ROW + i, 3).Value = vbNullString
            CL(ws, ROW_BASE + FIRST_SEQ_ROW + i, 12).Value = vbNullString
        End If
        s = IIf(i = mSel And i < n, 1, 0)
        With ws.Range(CL(ws, ROW_BASE + FIRST_SEQ_ROW + i, 2), CL(ws, ROW_BASE + FIRST_SEQ_ROW + i, 19))
            .Interior.Color = IIf(s = 1, RGB(30, 64, 175), C_PANEL)
            .Font.Bold = (s = 1)
        End With
    Next i
    If Len(MicroCheckRules()) = 0 Then
        ws.Range(RA("B46")).Value = "Regla verificada en toda la secuencia: nunca hay mas de un Enable activo (sin contencion); sin Enable el bus flota en High-Z."
    Else
        ws.Range(RA("B46")).Value = "VIOLACION: " & MicroCheckRules()
    End If
End Sub

' kind: 1 = enable (green), 2 = load (amber), 3 = control (cyan).
Private Sub Lamp(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal isOn As Boolean, ByVal kind As Long)
    Dim onColor As Long
    Select Case kind
        Case 1: onColor = RGB(34, 197, 94)
        Case 2: onColor = RGB(250, 204, 21)
        Case Else: onColor = RGB(56, 189, 248)
    End Select
    With CL(ws, ROW_BASE + r, c)
        .Value = IIf(isOn, 1, 0)
        .Interior.Color = IIf(isOn, onColor, C_PANEL)
        .Font.Color = IIf(isOn, RGB(15, 23, 42), C_OFF_TXT)
        .Font.Bold = isOn
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub Buf(ByVal ws As Worksheet, ByVal c As Long, ByVal conducting As Boolean)
    With CL(ws, ROW_BASE + 26, c)
        .Value = IIf(conducting, "1 CONDUCE", "High-Z")
        .Interior.Color = IIf(conducting, RGB(34, 197, 94), C_PANEL)
        .Font.Color = IIf(conducting, RGB(15, 23, 42), C_OFF_TXT)
        .Font.Bold = conducting
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub Header(ByVal ws As Worksheet, ByVal r As Long, ByVal text As String)
    ws.Range(CL(ws, ROW_BASE + r, 2), CL(ws, ROW_BASE + r, 19)).Merge
    With CL(ws, ROW_BASE + r, 2)
        .Value = text
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(125, 211, 252)
    End With
End Sub

Private Sub AddBtn(ByVal ws As Worksheet, ByVal address As String, ByVal caption As String, ByVal macroName As String)
    Dim area As Range, button As Shape
    Set area = ws.Range(RA(address))
    Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left, area.Top, area.Width - 4, area.Height - 2)
    button.Name = "P05B_" & macroName
    button.Fill.ForeColor.RGB = RGB(37, 99, 235)
    button.Line.Visible = msoFalse
    button.TextFrame2.TextRange.Text = caption
    button.TextFrame2.TextRange.Font.Name = "Segoe UI"
    button.TextFrame2.TextRange.Font.Size = 12
    button.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    button.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    button.TextFrame2.VerticalAnchor = msoAnchorMiddle
    button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Sub

Private Function RA(ByVal address As String) As String
    RA = LcaShift(address, ROW_BASE, COL_BASE)
End Function

Private Function CL(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Range
    Set CL = ws.Cells(r, c + COL_BASE)
End Function
