Attribute VB_Name = "LcaEditorPanel"
Option Explicit

' P11 band of the LCA sheet: assembly editor in cells, listing, located errors and load summary.
Private Const ROW_BASE As Long = LCA_BASE_EDITOR
Private Const COL_BASE As Long = LCA_COL_EDITOR
Private Const BAND_ROWS As Long = 98
Private Const EDIT_FIRST As Long = 13        ' band row of editor line 1
Private Const EDIT_LINES As Long = 64
Private Const ERR_FIRST As Long = 80
Private Const ERR_ROWS As Long = 8
Private Const BG As Long = 2758415          ' RGB(15, 23, 42)
Private Const CARD As Long = 3877150        ' RGB(30, 41, 59)
Private Const HEAD As Long = 6897694        ' RGB(30, 64, 105)
Private Const INK As Long = 15788258        ' RGB(226, 232, 240)
Private Const MUTED As Long = 12100500      ' RGB(148, 163, 184)
Private Const ACCENT As Long = 16569213     ' RGB(125, 211, 252)
Private Const BUTTON As Long = 15426341     ' RGB(37, 99, 235)
Private Const OK_GREEN As Long = 3433750    ' RGB(22, 101, 52)
Private Const BAD As Long = 1908095         ' RGB(127, 29, 29)
Private Const EDITOR_BG As Long = 4402462   ' RGB(30, 45, 67)

Private mStatus As String
Private mStatusOk As Boolean

Public Sub ConstruirPanelEditor()
    Dim ws As Worksheet, i As Long, r As Long
    Set ws = EditorSheet(True)
    LcaPrepareBand ws, ROW_BASE, BAND_ROWS, "P11_", COL_BASE
    LcaRowHeight ws.Rows(RA("1:" & CStr(BAND_ROWS))), 19
    With ws.Range(RA("A1:AZ" & CStr(BAND_ROWS)))
        .Interior.Color = BG
        .Font.Name = "Consolas"
        .Font.Size = 11
        .Font.Color = INK
        .VerticalAlignment = xlCenter
    End With

    LcaRowHeight ws.Rows(ROW_BASE + 2), 30
    ws.Range(RA("B2:M2")).Merge
    With ws.Range(RA("B2"))
        .Value = "P11  /  EDITOR + ENSAMBLADOR + CARGADOR"
        .Font.Name = "Segoe UI"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = ACCENT
    End With
    ws.Range(RA("B3:S3")).Merge
    With ws.Range(RA("B3"))
        .Value = "Escribe ASM en las celdas -> ENSAMBLAR valida todo -> CARGAR copia el codigo a VPN 0..7 (swap) y reinicia la ejecucion (PC = 000h)"
        .Font.Color = MUTED
    End With
    AddButton ws, "Q2:S2", "VER P10", "LcaIrCachePaginacion", HEAD
    AddButton ws, "B5:D6", "ENSAMBLAR", "EditorEnsamblar", BUTTON
    AddButton ws, "E5:G6", "CARGAR", "EditorCargar", OK_GREEN
    AddButton ws, "H5:J6", "PROBAR FETCH", "EditorProbarFetch", HEAD
    AddButton ws, "K5:M6", "EJEMPLO", "EditorEjemplo", HEAD
    AddButton ws, "N5:P6", "LIMPIAR", "EditorLimpiar", HEAD
    ws.Range(RA("B8:S9")).Merge
    ws.Range(RA("B8")).WrapText = True
    ws.Range(RA("B8")).Font.Name = "Segoe UI"

    ' 01 editor (B:I) and 02 listing (K:S)
    SectionTitle ws, "B11:I11", "01   EDITOR ASM  (64 lineas / lines)"
    SectionTitle ws, "K11:S11", "02   LISTADO / LISTING  (direccion, bytes, decodificado)"
    ws.Range(RA("B12")).Value = "LIN"
    ws.Range(RA("C12:I12")).Merge
    ws.Range(RA("C12")).Value = "CODIGO  /  CODE     [etiqueta:] MNEMONICO op1, op2   ; comentario"
    ws.Range(RA("K12")).Value = "ADDR"
    ws.Range(RA("L12:M12")).Merge
    ws.Range(RA("L12")).Value = "BYTES"
    ws.Range(RA("N12:Q12")).Merge
    ws.Range(RA("N12")).Value = "INSTRUCCION / INSTRUCTION"
    ws.Range(RA("R12")).Value = "LIN"
    ws.Range(RA("S12")).Value = "VPN"
    With ws.Range(RA("B12:I12"))
        .Interior.Color = HEAD
        .Font.Bold = True
    End With
    With ws.Range(RA("K12:S12"))
        .Interior.Color = HEAD
        .Font.Bold = True
    End With
    For i = 0 To EDIT_LINES - 1
        r = ROW_BASE + EDIT_FIRST + i
        CL(ws, r, 2).Value = i + 1
        CL(ws, r, 2).Font.Color = MUTED
        CL(ws, r, 2).HorizontalAlignment = xlCenter
        With ws.Range(CL(ws, r, 3), CL(ws, r, 9))
            .Merge
            .NumberFormat = "@"
            .Interior.Color = EDITOR_BG
            .Locked = False
            .HorizontalAlignment = xlLeft
        End With
        ws.Range(CL(ws, r, 12), CL(ws, r, 13)).Merge
        ws.Range(CL(ws, r, 14), CL(ws, r, 17)).Merge
        ws.Range(CL(ws, r, 11), CL(ws, r, 19)).Interior.Color = CARD
    Next i
    ws.Range(CL(ws, ROW_BASE + 12, 11), CL(ws, ROW_BASE + EDIT_FIRST + EDIT_LINES - 1, 13)).HorizontalAlignment = xlCenter
    ws.Range(CL(ws, ROW_BASE + 12, 18), CL(ws, ROW_BASE + EDIT_FIRST + EDIT_LINES - 1, 19)).HorizontalAlignment = xlCenter

    ' 03 errors
    SectionTitle ws, "B" & CStr(ERR_FIRST - 2) & ":S" & CStr(ERR_FIRST - 2), "03   ERRORES / ERRORS  (linea, columna, mensaje)"
    For i = 0 To ERR_ROWS - 1
        r = ROW_BASE + ERR_FIRST - 1 + i
        ws.Range(CL(ws, r, 4), CL(ws, r, 19)).Merge
        ws.Range(CL(ws, r, 2), CL(ws, r, 19)).Interior.Color = CARD
    Next i

    ' 04 load summary and syntax reference
    SectionTitle ws, "B89:S89", "04   CARGA / LOAD"
    ws.Range(RA("B90:S91")).Merge
    ws.Range(RA("B90")).WrapText = True
    ws.Range(RA("B90")).Interior.Color = CARD
    SectionTitle ws, "B93:S93", "SINTAXIS / SYNTAX"
    ws.Range(RA("B94:S97")).Merge
    With ws.Range(RA("B94"))
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Size = 10
        .Font.Color = MUTED
        .Value = "NOP | HLT | MOV R,#imm | MOV R,R2 | LD R,[VA] | ST [VA],R | ADD/SUB/CMP R,R2 | INC/DEC R | " & _
                 "JMP/JZ/JNZ/JC/JS destino | IN R | OUT R.   R = AX BX CX DX.   Numeros hex con sufijo h y 1-3 digitos, " & _
                 "empezando por 0-9 (2Ah, 0FFh, 104h).   LD/ST: 100h-1DFh.   Saltos: 000h-0FCh, multiplo de 4, o etiqueta.   " & _
                 "Maximo 64 instrucciones de 4 bytes; el resto del codigo se llena con HLT.   Guia: docs/P11_ASSEMBLER_LOADER.md"
    End With

    EditorWrite LoaderSampleProgram()
    mStatus = "Ejemplo cargado en el editor. Pulsa ENSAMBLAR o CARGAR."
    mStatusOk = True
    EditorRefresh
    LcaGoto ws, ROW_BASE, COL_BASE
End Sub

Public Sub MostrarPanelEditor()
    Dim ws As Worksheet
    Set ws = EditorSheet(False)
    If ws Is Nothing Then
        ConstruirPanelEditor
    Else
        LcaGoto ws, ROW_BASE, COL_BASE
    End If
End Sub

Public Sub EditorEnsamblar()
    If AsmAssemble(EditorRead()) Then
        mStatus = "ENSAMBLADO OK: " & CStr(AsmInstructionCount()) & " instrucciones, " & _
                  CStr(AsmInstructionCount() * 4) & " bytes. Aun no cargado: pulsa CARGAR."
        mStatusOk = True
    Else
        mStatus = "ERRORES: " & CStr(AsmErrorCount()) & ". Corrige las lineas marcadas en rojo; el programa cargado no cambio."
        mStatusOk = False
    End If
    EditorRefresh
End Sub

Public Sub EditorCargar()
    On Error GoTo failed
    If LoaderLoadSource(EditorRead()) Then
        mStatus = "CARGADO: " & CStr(LoaderInstructionCount()) & " instrucciones en " & PagesText() & ". CPU, buses, RAM, caches y tabla de paginas reiniciados; PC = 000h."
        mStatusOk = True
        PanelActualizar
        CachePathPanelRefresh
    Else
        mStatus = "NO CARGADO: " & CStr(AsmErrorCount()) & " error(es). El programa y el estado anteriores se conservan."
        mStatusOk = False
    End If
    EditorRefresh
    Exit Sub
failed:
    mStatus = "ERROR: " & Err.Description
    mStatusOk = False
    EditorRefresh
End Sub

' Fetches the instruction at PC through the P10 path to show that the code lives in virtual pages.
Public Sub EditorProbarFetch()
    Dim cpu As LcaCpuState, s As LcaPathStep
    On Error GoTo failed
    If Not LoaderLoaded() Then
        mStatus = "Primero CARGAR un programa."
        mStatusOk = False
        EditorRefresh
        Exit Sub
    End If
    cpu = CpuSnapshot()
    CachePathFetchFour cpu.PC
    s = CachePathLast()
    mStatus = "FETCH PC=" & AsmHex3(cpu.PC) & ": " & AsmDisassemble(ControlFetched(0), ControlFetched(1), _
              ControlFetched(2), ControlFetched(3)) & "  |  " & IIf(s.PageFault, "PAGE_FAULT -> PFN ", "PFN ") & _
              CStr(s.PFN) & ", L1I " & CStr(s.L1Hits) & "H/" & CStr(s.L1Misses) & "M. (P12 ejecutara la instruccion.)"
    mStatusOk = True
    PanelActualizar
    CachePathPanelRefresh
    EditorRefresh
    Exit Sub
failed:
    mStatus = "ERROR FETCH: " & Err.Description
    mStatusOk = False
    EditorRefresh
End Sub

Public Sub EditorEjemplo()
    EditorWrite LoaderSampleProgram()
    mStatus = "Ejemplo escrito en el editor. Pulsa CARGAR."
    mStatusOk = True
    AsmAssemble EditorRead()
    EditorRefresh
End Sub

Public Sub EditorLimpiar()
    EditorWrite Array()
    mStatus = "Editor vacio. El programa cargado no cambia hasta pulsar CARGAR."
    mStatusOk = True
    AsmAssemble EditorRead()
    EditorRefresh
End Sub

' Editor lines 1..64 as a 1-based string array.
Public Function EditorRead() As Variant
    Dim ws As Worksheet, i As Long, lines() As String
    Set ws = EditorSheet(False)
    ReDim lines(1 To EDIT_LINES)
    If Not ws Is Nothing Then
        For i = 1 To EDIT_LINES
            lines(i) = CStr(CL(ws, ROW_BASE + EDIT_FIRST + i - 1, 3).Value)
        Next i
    End If
    EditorRead = lines
End Function

Public Sub EditorWrite(ByVal source As Variant)
    Dim ws As Worksheet, i As Long, n As Long
    Set ws = EditorSheet(False)
    If ws Is Nothing Then Exit Sub
    n = 0
    If IsArray(source) Then
        If UBound(source) >= LBound(source) Then n = UBound(source) - LBound(source) + 1
    End If
    For i = 0 To EDIT_LINES - 1
        If i < n Then
            CL(ws, ROW_BASE + EDIT_FIRST + i, 3).Value = CStr(source(LBound(source) + i))
        Else
            CL(ws, ROW_BASE + EDIT_FIRST + i, 3).Value = vbNullString
        End If
    Next i
End Sub

' Redraws listing, error marks and load summary from the assembler/loader state.
Public Sub EditorRefresh()
    Dim ws As Worksheet, i As Long, r As Long, addr As Long, k As Long, bytesText As String
    Set ws = EditorSheet(False)
    If ws Is Nothing Then Exit Sub
    ws.Range(RA("B8")).Value = mStatus
    ws.Range(RA("B8")).Interior.Color = IIf(mStatusOk, OK_GREEN, BAD)

    For i = 0 To EDIT_LINES - 1
        r = ROW_BASE + EDIT_FIRST + i
        ws.Range(CL(ws, r, 3), CL(ws, r, 9)).Interior.Color = EDITOR_BG
        ws.Range(CL(ws, r, 11), CL(ws, r, 19)).Value = vbNullString
        CL(ws, r, 14).Font.Color = INK
        ws.Range(CL(ws, r, 11), CL(ws, r, 19)).Interior.Color = CARD
        If i < AsmInstructionCount() Then
            addr = i * 4
            CL(ws, r, 11).Value = AsmHex3(addr)
            If AsmInstructionValid(i) Then
                bytesText = vbNullString
                For k = 0 To 3
                    bytesText = bytesText & Right$("00" & Hex$(AsmByte(addr + k)), 2) & " "
                Next k
                CL(ws, r, 12).Value = RTrim$(bytesText)
                CL(ws, r, 14).Value = AsmDisassemble(AsmByte(addr), AsmByte(addr + 1), AsmByte(addr + 2), AsmByte(addr + 3))
            Else
                CL(ws, r, 12).Value = "-- -- -- --"
                CL(ws, r, 14).Value = "ERROR (ver seccion 03 / see section 03)"
                ws.Range(CL(ws, r, 11), CL(ws, r, 19)).Interior.Color = BAD
            End If
            CL(ws, r, 18).Value = AsmInstructionLine(i)
            CL(ws, r, 19).Value = addr \ 32
        End If
    Next i
    If AsmInstructionCount() < EDIT_LINES And AsmSucceeded() Then
        r = ROW_BASE + EDIT_FIRST + AsmInstructionCount()
        CL(ws, r, 11).Value = AsmHex3(AsmInstructionCount() * 4)
        CL(ws, r, 12).Value = "01 00 00 00"
        CL(ws, r, 14).Value = "HLT ... (relleno hasta 0FCh / fill)"
        CL(ws, r, 14).Font.Color = MUTED
    End If

    For i = 0 To ERR_ROWS - 1
        r = ROW_BASE + ERR_FIRST - 1 + i
        ws.Range(CL(ws, r, 2), CL(ws, r, 4)).Value = vbNullString
        ws.Range(CL(ws, r, 2), CL(ws, r, 19)).Interior.Color = CARD
        If i < AsmErrorCount() Then
            CL(ws, r, 2).Value = "L" & CStr(AsmErrorLine(i))
            CL(ws, r, 3).Value = "C" & CStr(AsmErrorColumn(i))
            CL(ws, r, 4).Value = AsmErrorText(i)
            ws.Range(CL(ws, r, 2), CL(ws, r, 19)).Interior.Color = BAD
        End If
    Next i
    If AsmErrorCount() > ERR_ROWS Then _
        CL(ws, ROW_BASE + ERR_FIRST - 1 + ERR_ROWS - 1, 4).Value = "... y " & CStr(AsmErrorCount() - ERR_ROWS + 1) & " mas / more"
    For i = 0 To AsmErrorCount() - 1
        If AsmErrorLine(i) >= 1 And AsmErrorLine(i) <= EDIT_LINES Then _
            ws.Range(CL(ws, ROW_BASE + EDIT_FIRST + AsmErrorLine(i) - 1, 3), _
                     CL(ws, ROW_BASE + EDIT_FIRST + AsmErrorLine(i) - 1, 9)).Interior.Color = BAD
    Next i

    If LoaderLoaded() Then
        ws.Range(RA("B90")).Value = "Programa cargado #" & CStr(LoaderLoadCount()) & ": " & _
            CStr(LoaderInstructionCount()) & " instrucciones (" & CStr(LoaderInstructionCount() * 4) & " B) en " & _
            PagesText() & " del respaldo; VPN 0..7 = 256 B de codigo (resto HLT), VPN 8..14 = datos en 00h. " & _
            "Primeros bytes de VPN0: " & FirstBytes() & ". Las paginas entran en RAM por fallo de pagina (P08/P10)."
    Else
        ws.Range(RA("B90")).Value = "Ningun programa cargado / no program loaded."
    End If
End Sub

Private Function PagesText() As String
    If LoaderCodePages() <= 1 Then
        PagesText = "VPN 0"
    Else
        PagesText = "VPN 0.." & CStr(LoaderCodePages() - 1)
    End If
End Function

Private Function FirstBytes() As String
    Dim i As Long
    For i = 0 To 7
        FirstBytes = FirstBytes & Right$("00" & Hex$(SwapReadByte(0, i)), 2) & " "
    Next i
    FirstBytes = RTrim$(FirstBytes)
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

Private Function EditorSheet(ByVal create As Boolean) As Worksheet
    Set EditorSheet = LcaBand("P11_EditorEnsamblar", create)
End Function

' Row shift of this panel inside the shared LCA sheet.
Private Function RA(ByVal address As String) As String
    RA = LcaShift(address, ROW_BASE, COL_BASE)
End Function

Private Sub AddButton(ByVal ws As Worksheet, ByVal address As String, ByVal caption As String, _
                      ByVal macroName As String, ByVal fillColor As Long)
    Dim area As Range, button As Shape
    Set area = ws.Range(RA(address))
    Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left + 2, area.Top + 2, _
                                    area.Width - 4, area.Height - 4)
    button.Name = "P11_" & macroName
    button.Fill.ForeColor.RGB = fillColor
    button.Line.Visible = msoFalse
    button.TextFrame2.TextRange.Text = caption
    button.TextFrame2.TextRange.Font.Name = "Segoe UI"
    button.TextFrame2.TextRange.Font.Size = 11
    button.TextFrame2.TextRange.Font.Bold = msoTrue
    button.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    button.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    button.TextFrame2.VerticalAnchor = msoAnchorMiddle
    button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Sub

' Cell of this block: row is absolute (ROW_BASE already added), column is relative to the block.
Private Function CL(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Range
    Set CL = ws.Cells(r, c + COL_BASE)
End Function
