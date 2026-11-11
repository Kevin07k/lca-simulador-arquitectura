Attribute VB_Name = "LcaControlsTests"
Option Explicit

Private mChecks As Long

' P18 regression: the program catalogue loads every program through its own block (LOAD), the machine bar
' has STEP / RUN / PAUSE / RESET / LOAD / speed and the team identity, PAUSE stops RUN between instructions
' and RUN continues to the same result, and after every action all views show the same machine state
' (registers, PC, page table, LCD, demo blocks, Logs).
Public Sub VerificarControlesP18()
    Dim i As Long, n As Long, ws As Worksheet, prevSpeed As Long, names As String, built As Boolean, item As Variant
    On Error GoTo failed
    mChecks = 0
    prevSpeed = FlowSpeed()

    ' --- A: program catalogue ---------------------------------------------------------------------
    Check LCA_PROGRAM_COUNT = 9, "A: nine programs in the catalogue"
    For i = 0 To LCA_PROGRAM_COUNT - 1
        Check Len(LcaProgramaNombre(i)) > 0 And InStr(LcaProgramaNombre(i), ",") = 0, "A: name " & CStr(i) & " usable in a list"
        Check LcaProgramaIndice(LcaProgramaNombre(i)) = i, "A: " & LcaProgramaNombre(i) & " found by name"
        names = names & LcaProgramaNombre(i) & ";"
    Next i
    Check Len(LcaProgramaLista()) <= 255, "A: the list fits a data validation (255 characters)"
    Check LcaProgramaIndice("NO EXISTE") = -1, "A: an unknown name is rejected"
    For i = 1 To LCA_PROGRAM_COUNT - 1
        Check LcaCargarPrograma(i), "A: " & LcaProgramaNombre(i) & " loads"
        Check LoaderLoaded() And ExecSteps() = 0 And CpuSnapshot().PC = 0 And LcaProgramaActual() = i, _
              "A: " & LcaProgramaNombre(i) & " leaves the machine reset (PC = 000h)"
        Check AsmAssembleText(Join(LcaProgramaFuente(i), vbLf)) And AsmInstructionCount() = LoaderInstructionCount(), _
              "A: " & LcaProgramaNombre(i) & " is the program of its source"
    Next i
    Check LcaCargarPrograma(7) And EchoLoaded(), "A: ECO goes through the P16 loader (its block recognises it)"
    Check LcaCargarPrograma(8) And MemDemoLoaded(), "A: MEMORIA goes through the P17 loader"
    Check LcaCargarPrograma(3) And LoaderInstructionCount() = UBound(ArithSource(1)), "A: MULTIPLICAR (P15) loaded"
    Check Not EchoLoaded() And Not MemDemoLoaded(), "A: loading another program replaces the shared one"

    Set ws = LcaBand("P12_btn_MaquinaPaso", False)
    built = Not ws Is Nothing
    If Not built Then GoTo done

    ' --- B: the machine bar has every control and the team identity -------------------------------------
    For Each item In Array("P12_btn_MaquinaCargarPrograma", "P12_btn_MaquinaPaso", "P12_btn_MaquinaEjecutar", _
                           "P12_btn_MaquinaPausa", "P12_btn_MaquinaReset", "P12_btn_MaquinaVelocidad")
        Check HasShape(ws, CStr(item)), "B: control " & CStr(item) & " present"
    Next item
    Check InStr(ws.Range(LcaShift("M2", LCA_BASE_MACHINE, LCA_COL_MACHINE)).Value, "Kevin") > 0 And _
          InStr(ws.Range(LcaShift("M2", LCA_BASE_MACHINE, LCA_COL_MACHINE)).Value, "Javier") > 0, "B: team identity on the machine bar"
    Check ws.Range(LcaShift("B7", LCA_BASE_MACHINE, LCA_COL_MACHINE)).Validation.Type = xlValidateList, "B: program selector is a list"
    Check HasShape(ws, "NAVM_0") And HasShape(ws, "P12_nav_0"), "B: navigation to every block (MAQUINA buttons and bar)"

    ' --- C: LOAD from the selector, sync after the action ----------------------------------------------
    FlowSetSpeed 0
    MaquinaElegirPrograma 1
    MaquinaCargarPrograma
    Check MaquinaProgramaElegido() = LcaProgramaNombre(1) And LcaProgramaActual() = 1 And ExecSteps() = 0, "C: CARGAR loads the chosen program"
    Check LcaUltimaSincronizacionErrores() = 0 And LcaVistasSincronizadas() = 17, "C: 17 views refreshed without errors"
    Check LcaVistasCoherentes(), "C: every view shows the reset machine"
    MaquinaPaso
    Check ExecSteps() = 1 And LcaVistasCoherentes(), "C: PASO runs one instruction and every view follows"
    MaquinaEjecutar
    Check ExecStatus() = EXEC_HALTED And CpuReadRegister(0) = &HF And ExecSteps() = 20, "C: EJECUTAR runs SUMA to HLT (AX = 0Fh, 20 instructions)"
    Check LcaVistasCoherentes(), "C: registers, PC, page table, LCD and Logs agree after RUN"
    Check InStr(MaquinaEstadoTexto(), "HLT") > 0, "C: status shows HLT"

    ' --- D: PAUSE stops RUN between instructions; RUN continues to the same result --------------------
    MaquinaReset
    Check ExecSteps() = 0 And CpuSnapshot().PC = 0 And LcaVistasCoherentes(), "D: RESET restarts and every view follows"
    LcaPausaTrasPasos 7
    MaquinaEjecutar
    Check ExecSteps() = 7 And ExecStatus() = EXEC_READY And MaquinaEnPausa(), "D: PAUSA after 7 instructions, machine still READY"
    Check InStr(MaquinaEstadoTexto(), "PAUSA") > 0, "D: status shows EN PAUSA"
    Check LcaVistasCoherentes(), "D: every view shows the paused machine"
    MaquinaPaso
    Check ExecSteps() = 8 And Not MaquinaEnPausa(), "D: PASO advances one instruction from the pause"
    MaquinaEjecutar
    Check ExecStatus() = EXEC_HALTED And ExecSteps() = 20 And CpuReadRegister(0) = &HF And CpuReadRegister(1) = &HF, _
          "D: EJECUTAR continues to the same HLT state as an uninterrupted run"
    MaquinaReset
    LcaPausar
    MaquinaEjecutar
    Check ExecStatus() = EXEC_HALTED And ExecSteps() = 20, "D: an old PAUSA click does not stop the next RUN"
    FlowSetSpeed 3
    MaquinaReset
    LcaPausaTrasPasos 2
    MaquinaEjecutar
    Check ExecSteps() = 2 And MaquinaEnPausa() And LcaVistasCoherentes(), "D: PAUSA also stops the animated RUN"
    FlowSetSpeed 0

    ' --- E: I/O stays in step: echo through the machine controls -----------------------------------------
    MaquinaElegirPrograma 7
    MaquinaCargarPrograma
    MaquinaEjecutar
    Check ExecStatus() = EXEC_WAIT And LcaVistasCoherentes(), "E: ECO waits for a key; every view agrees"
    Call KeyboardPushText("LCA.")
    MaquinaEjecutar
    Check ExecStatus() = EXEC_HALTED And Left$(LcdText(), 3) = "LCA" And LcaVistasCoherentes(), _
          "E: keys typed, LCD = LCA, the P14 and P16 blocks and Logs agree"
    Check EcoPanelEstado() = "TERMINADO (HLT)", "E: the P16 block shows the run made from the machine bar"

    ' --- F: MEMORIA through the machine: page table and caches in step ----------------------------------
    MaquinaElegirPrograma 8
    MaquinaCargarPrograma
    MaquinaEjecutar
    Check ExecStatus() = EXEC_HALTED And ExecSteps() = MEMDEMO_STEPS, "F: MEMORIA runs to HLT from the machine bar"
    Check MemDemoCountersMatch() And LcaVistasCoherentes(), "F: counters as documented; P07 table, P17 block and Logs agree"

done:
    FlowSetSpeed prevSpeed
    LcaPausaLimpiar
    LcaReport "P18 PASS: " & CStr(mChecks) & " comprobaciones / assertions." & vbCrLf & _
              "Catalogo CARGAR, PASO/EJECUTAR/PAUSA/RESET/VELOCIDAD, identidad del equipo y vistas sincronizadas" & _
              IIf(built, ".", " (tablero no construido: solo catalogo)."), True
    Exit Sub
failed:
    FlowSetSpeed prevSpeed
    LcaPausaLimpiar
    LcaReport "P18 FAIL after " & CStr(mChecks) & " checks: " & Err.Description, False
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
    If Not condition Then Err.Raise vbObjectError + 2192, "LcaControlsTests", message
End Sub
