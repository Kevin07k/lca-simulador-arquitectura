Attribute VB_Name = "LcaControls"
Option Explicit

' P18: one set of controls for the whole dashboard.
'   - LOAD: a catalogue of every program of the simulator (P11 editor, P12 sample, P15 arithmetic demos,
'     P16 echo, P17 memory demo); loading one goes through the loader of its own block so that block
'     shows it too;
'   - PAUSE: a request flag read by the animated RUN loops (button PAUSA or the Esc key);
'   - SYNC: after every user action, every view of the shared machine is refreshed (machine, CPU, diagram,
'     buses, pages, swap/CLOCK, caches, cache+paging, keyboard, LCD, demo blocks and Logs), so registers,
'     buses, tables, caches and I/O always show the same state;
'   - team identity shown by the machine bar and the cover.
' No worksheet dependency except the refresh calls (each one is skipped when its block is not built).
Public Const LCA_TEAM_TEXT As String = "Equipo LCA:  Kevin (@Kevin07k)  y  Javier (@Javiertorricos)"
Public Const LCA_ERR_CONTROLS As Long = vbObjectError + 2191
Public Const LCA_PROGRAM_COUNT As Long = 9

Private mPause As Boolean
Private mPauseAt As Long
Private mCurrent As Long
Private mLastSyncErrors As Long
Private mLastSyncViews As Long

' --- Program catalogue (LOAD) ------------------------------------------------------------------

Public Function LcaProgramaNombre(ByVal i As Long) As String
    Select Case i
        Case 0: LcaProgramaNombre = "EDITOR (P11)"
        Case 1: LcaProgramaNombre = "EJEMPLO SUMA (P12)"
        Case 2 To 6: LcaProgramaNombre = ArithName(i - 2) & " (P15)"
        Case 7: LcaProgramaNombre = "ECO TECLADO-LCD (P16)"
        Case 8: LcaProgramaNombre = "MEMORIA + CACHE (P17)"
        Case Else: Err.Raise LCA_ERR_CONTROLS, "LcaControls", "PROGRAM_INDEX"
    End Select
End Function

' Comma-separated names for the data validation list of the machine bar.
Public Function LcaProgramaLista() As String
    Dim i As Long
    For i = 0 To LCA_PROGRAM_COUNT - 1
        If i > 0 Then LcaProgramaLista = LcaProgramaLista & ","
        LcaProgramaLista = LcaProgramaLista & LcaProgramaNombre(i)
    Next i
End Function

' Index of a catalogue name, -1 when unknown.
Public Function LcaProgramaIndice(ByVal name As String) As Long
    Dim i As Long
    LcaProgramaIndice = -1
    For i = 0 To LCA_PROGRAM_COUNT - 1
        If StrComp(Trim$(name), LcaProgramaNombre(i), vbTextCompare) = 0 Then
            LcaProgramaIndice = i
            Exit Function
        End If
    Next i
End Function

' Source lines of a catalogue entry (the editor entry reads the P11 editor).
Public Function LcaProgramaFuente(ByVal i As Long) As Variant
    Select Case i
        Case 0: LcaProgramaFuente = EditorRead()
        Case 1: LcaProgramaFuente = LoaderSampleProgram()
        Case 2 To 6: LcaProgramaFuente = ArithSource(i - 2)
        Case 7: LcaProgramaFuente = EchoSource()
        Case 8: LcaProgramaFuente = MemDemoSource()
        Case Else: Err.Raise LCA_ERR_CONTROLS, "LcaControls", "PROGRAM_INDEX"
    End Select
End Function

' LOAD: assembles and loads entry i through its own block, resets the whole machine and shows the program
' in the P11 editor. Returns False when it does not assemble (the machine keeps the previous program).
Public Function LcaCargarPrograma(ByVal i As Long) As Boolean
    Dim source As Variant
    source = LcaProgramaFuente(i)
    LcaPausaLimpiar
    Select Case i
        Case 0, 1
            LcaCargarPrograma = LoaderLoadSource(source)
        Case 2 To 6
            If ArithPanelBuilt() Then
                ArithSeleccionar i - 2
                LcaCargarPrograma = LoaderLoaded() And ExecSteps() = 0
            Else
                LcaCargarPrograma = ArithLoad(i - 2)
            End If
        Case 7
            LcaCargarPrograma = EchoLoad()
        Case 8
            If MemPanelConstruido() Then
                MemCargar
                LcaCargarPrograma = MemDemoLoaded()
            Else
                LcaCargarPrograma = MemDemoLoad()
            End If
    End Select
    If Not LcaCargarPrograma Then Exit Function
    mCurrent = i
    If i > 0 Then
        On Error Resume Next
        EditorWrite source
        EditorRefresh
        On Error GoTo 0
    End If
End Function

' Catalogue entry loaded last through LcaCargarPrograma.
Public Function LcaProgramaActual() As Long
    LcaProgramaActual = mCurrent
End Function

' --- PAUSE --------------------------------------------------------------------------------------

' Button PAUSA: the running RUN loop stops after the instruction in progress.
Public Sub LcaPausar()
    mPause = True
End Sub

' Test / demo hook: pause automatically after n completed instructions of the next RUN (0 = off).
Public Sub LcaPausaTrasPasos(ByVal n As Long)
    If n < 0 Then n = 0
    mPauseAt = n
End Sub

' Read by the RUN loops after every instruction (stepsDone = instructions completed in this RUN).
Public Function LcaPausaPedida(ByVal stepsDone As Long) As Boolean
    DoEvents
    If mPause Then LcaPausaPedida = True
    If mPauseAt > 0 And stepsDone >= mPauseAt Then LcaPausaPedida = True
End Function

' A RUN that starts now forgets an old PAUSA click (the test hook stays).
Public Sub LcaPausaOlvidarClic()
    mPause = False
End Sub

Public Sub LcaPausaLimpiar()
    mPause = False
    mPauseAt = 0
End Sub

' --- SYNC ---------------------------------------------------------------------------------------

' Refreshes every view of the shared machine. A view whose block is not built returns at once. Returns
' the number of views that raised an error (0 = all views in step). skipMachine avoids redrawing the
' machine block when its own action already did.
Public Function LcaSincronizar(Optional ByVal skipMachine As Boolean = False) As Long
    Dim views As Variant, i As Long, failures As Long, prev As Boolean
    prev = Application.ScreenUpdating
    Application.ScreenUpdating = False
    views = Array("MaquinaRefrescar", "PanelActualizar", "BusPanelActualizar", "BusDetalleActualizar", _
                  "MmuPanelRefresh", "MmuDetalleActualizar", "PagerPanelRefresh", "MemoriaDetalleActualizar", _
                  "CachePanelActualizar", "CachePathPanelRefresh", "TecladoPanelActualizar", "LcdPanelActualizar", _
                  "ArithPanelRefrescar", "EcoPanelRefrescar", "MemPanelRefrescar", "LogsActualizar")
    For i = 0 To UBound(views)
        If Not (skipMachine And i = 0) Then
            On Error Resume Next
            Application.Run "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & CStr(views(i))
            If Err.Number <> 0 Then failures = failures + 1
            Err.Clear
            On Error GoTo 0
        End If
    Next i
    Application.ScreenUpdating = prev
    mLastSyncErrors = failures
    mLastSyncViews = UBound(views) + 1
    LcaSincronizar = failures
End Function

Public Sub LcaSincronizarTodo()
    Call LcaSincronizar
End Sub

Public Function LcaUltimaSincronizacionErrores() As Long
    LcaUltimaSincronizacionErrores = mLastSyncErrors
End Function

Public Function LcaVistasSincronizadas() As Long
    LcaVistasSincronizadas = mLastSyncViews
End Function

' True when the views of the dashboard show the engine state: machine registers/PC, P04 registers and
' PC, LCD band = RAM, the demo blocks agree with RAM, and the newest Logs line is the newest trace line.
' Blocks that are not built are not checked.
Public Function LcaVistasCoherentes() As Boolean
    Dim ws As Worksheet, i As Long, v As Long, t As String, s As Shape
    Set ws = LcaSheet(False)
    If ws Is Nothing Then
        LcaVistasCoherentes = True
        Exit Function
    End If
    For i = 0 To 3
        v = CpuReadRegister(i)
        Set s = FindShape(ws, "P12_reg" & CStr(i))
        If Not s Is Nothing Then
            t = s.TextFrame2.TextRange.Text
            If InStr(t, AsmHex2(v)) = 0 Then Exit Function
        End If
        If Len(CStr(ws.Cells(LCA_BASE_CPU + 9, LCA_COL_CPU + 2 + i * 4).Value)) > 0 Then
            If CLng(ws.Cells(LCA_BASE_CPU + 9, LCA_COL_CPU + 2 + i * 4).Value) <> v Then Exit Function
        End If
    Next i
    If Not FindShape(ws, "P07_MmuPanelReset") Is Nothing Then
        For i = 0 To 15
            If CLng(ws.Cells(LCA_BASE_MMU + 13 + i, LCA_COL_MMU + 4).Value) <> Abs(CLng(MmuPage(i).Present)) Then Exit Function
        Next i
    End If
    Set s = FindShape(ws, "P12_pc")
    If Not s Is Nothing Then
        If InStr(s.TextFrame2.TextRange.Text, AsmHex3(CpuSnapshot().PC)) = 0 Then Exit Function
    End If
    If Not FindShape(ws, "P14_LcdPanelReset") Is Nothing Then
        If Not LcdPanelCoherente() Then Exit Function
    End If
    If EchoLoaded() And EcoPanelConstruido() Then
        If Not EcoPanelCoherente() Then Exit Function
    End If
    If MemDemoLoaded() And MemPanelConstruido() Then
        If Not MemPanelCoherente() Then Exit Function
    End If
    If Not LogsSheet(False) Is Nothing And ExecTraceCount() > 0 Then
        If InStr(CStr(LogsSheet(False).Range("B6").Value), ExecTraceLine(0)) = 0 Then Exit Function
    End If
    LcaVistasCoherentes = True
End Function

Private Function FindShape(ByVal ws As Worksheet, ByVal name As String) As Shape
    Dim s As Shape
    For Each s In ws.Shapes
        If s.Name = name Then
            Set FindShape = s
            Exit Function
        End If
    Next s
End Function

Private Function ArithPanelBuilt() As Boolean
    Dim ws As Worksheet
    Set ws = LcaSheet(False)
    If ws Is Nothing Then Exit Function
    ArithPanelBuilt = Not FindShape(ws, "P15_btn_ArithPaso") Is Nothing
End Function
