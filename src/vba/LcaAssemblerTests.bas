Attribute VB_Name = "LcaAssemblerTests"
Option Explicit

Private mChecks As Long

' P11 regression: encoding of every form, located errors, limits, round trip and loader contract.
Public Sub VerificarEnsambladorP11()
    Dim expected As Variant, i As Long, k As Long, cpu As LcaCpuState, entry As LcaPageEntry
    Dim src() As String, failure As Long, text As String, level As Long
    On Error GoTo failed
    mChecks = 0

    ' --- A: sample program -------------------------------------------------------------------
    Check AsmAssemble(LoaderSampleProgram()), "A: sample assembles"
    Check AsmInstructionCount() = 8 And AsmErrorCount() = 0, "A: eight instructions"
    expected = Array(&H10, 0, 0, 0, &H10, 2, 5, 0, &H20, 0, 2, 0, &H24, 2, 0, 0, _
                     &H32, 0, 8, 0, &H13, 0, 0, 1, &H12, 1, 0, 1, 1, 0, 0, 0)
    For i = 0 To 31
        Check AsmByte(i) = expected(i), "A: byte " & AsmHex2(i)
    Next i
    Check AsmByte(&H20) = 1 And AsmByte(&H21) = 0 And AsmByte(&HFC) = 1 And AsmByte(&HFF) = 0, "A: HLT fill to 0FCh"
    Check AsmLabelAddress("bucle") = 8 And AsmLabelAddress("FIN") = &H1C, "A: labels, case-insensitive"
    Check AsmInstructionLine(0) = 2 And AsmInstructionLine(2) = 4 And AsmInstructionLine(7) = 9, "A: source lines"

    ' --- B: every mnemonic and operand form --------------------------------------------------
    Check AsmAssembleText(Join(Array("NOP", "HLT", "MOV DX, #0FFh", "MOV BX, CX", "LD AX, [1DFh]", _
        "ST [100h], DX", "ADD AX, BX", "SUB BX, AX", "CMP CX, DX", "INC DX", "DEC AX", "JMP 0FCh", _
        "JZ 000h", "JNZ 004h", "JC 008h", "JS 00Ch", "IN BX", "OUT CX"), vbLf)), "B: all forms assemble"
    expected = Array(0, 0, 0, 0, 1, 0, 0, 0, &H10, 3, &HFF, 0, &H11, 1, 2, 0, &H12, 0, &HDF, 1, _
                     &H13, 3, 0, 1, &H20, 0, 1, 0, &H21, 1, 0, 0, &H22, 2, 3, 0, &H23, 3, 0, 0, _
                     &H24, 0, 0, 0, &H30, 0, &HFC, 0, &H31, 0, 0, 0, &H32, 0, 4, 0, &H33, 0, 8, 0, _
                     &H34, 0, &HC, 0, &H40, 1, 0, 0, &H41, 2, 0, 0)
    Check AsmInstructionCount() = 18, "B: eighteen instructions"
    For i = 0 To 71
        Check AsmByte(i) = expected(i), "B: byte " & AsmHex2(i)
    Next i
    ' Round trip: the disassembled text assembles back to the same four bytes.
    For i = 0 To 17
        text = AsmDisassemble(expected(4 * i), expected(4 * i + 1), expected(4 * i + 2), expected(4 * i + 3))
        Check AsmAssembleText(text), "B: round trip assembles '" & text & "'"
        For k = 0 To 3
            Check AsmByte(k) = expected(4 * i + k), "B: round trip bytes '" & text & "'"
        Next k
    Next i
    Check AsmAssembleText("  mov ax,#2ah   ; comentario: con dos puntos") And AsmByte(2) = &H2A, "B: lower case, comments"
    Check AsmDisassemble(&H77, 0, 0, 0) = "INVALID_OPCODE 77h", "B: decoder rejects opcode"
    Check Left$(AsmDisassemble(&H10, 4, 0, 0), 16) = "INVALID_ENCODING", "B: decoder rejects register 4"
    Check Left$(AsmDisassemble(&H30, 0, 2, 0), 16) = "INVALID_ENCODING", "B: decoder rejects unaligned jump"

    ' --- C: errors with line and column, all collected ---------------------------------------
    ReDim src(1 To 13)
    src(1) = "start:  MOV EX, #01h"
    src(2) = "        FOO AX"
    src(3) = "        MOV AX, #100h"
    src(4) = "        LD AX, [0FFh]"
    src(5) = "        ST [1E8h], AX"
    src(6) = "        JMP 002h"
    src(7) = "        JNZ nowhere"
    src(8) = "start:  NOP"
    src(9) = "        MOV AX, #10"
    src(10) = "        ADD AX"
    src(11) = "        INC AX, BX"
    src(12) = "9bad:   NOP"
    src(13) = "        MOV AX,"
    Check Not AsmAssemble(src) And Not AsmSucceeded(), "C: invalid program rejected"
    Check AsmErrorCount() = 13, "C: thirteen errors, one per line (got " & CStr(AsmErrorCount()) & ")"
    For i = 1 To 13
        Check ErrorAt(i) >= 0, "C: error reported on line " & CStr(i)
    Next i
    Check AsmErrorColumn(ErrorAt(1)) = 13 And InStr(AsmErrorText(ErrorAt(1)), "Registro") > 0, "C: L1 C13 register"
    Check AsmErrorColumn(ErrorAt(2)) = 9 And InStr(AsmErrorText(ErrorAt(2)), "FOO") > 0, "C: L2 C9 mnemonic"
    Check AsmErrorColumn(ErrorAt(3)) = 18 And InStr(AsmErrorText(ErrorAt(3)), "rango") > 0, "C: L3 C18 immediate range"
    Check InStr(AsmErrorText(ErrorAt(4)), "100h-1DFh") > 0, "C: L4 data region"
    Check InStr(AsmErrorText(ErrorAt(5)), "E/S") > 0, "C: L5 I/O address needs IN/OUT"
    Check InStr(AsmErrorText(ErrorAt(6)), "multiplo de 4") > 0, "C: L6 unaligned target"
    Check InStr(AsmErrorText(ErrorAt(7)), "no definida") > 0, "C: L7 undefined label"
    Check InStr(AsmErrorText(ErrorAt(8)), "duplicada") > 0 And AsmErrorColumn(ErrorAt(8)) = 1, "C: L8 duplicate label"
    Check InStr(AsmErrorText(ErrorAt(9)), "sufijo h") > 0, "C: L9 decimal needs h"
    Check InStr(AsmErrorText(ErrorAt(10)), "requiere 2") > 0, "C: L10 operand count"
    Check InStr(AsmErrorText(ErrorAt(11)), "requiere 1") > 0, "C: L11 operand count"
    Check InStr(AsmErrorText(ErrorAt(12)), "Etiqueta invalida") > 0 And AsmErrorColumn(ErrorAt(12)) = 1, "C: L12 bad label"
    Check InStr(AsmErrorText(ErrorAt(13)), "vacio") > 0 And AsmErrorColumn(ErrorAt(13)) = 16, "C: L13 C16 empty operand"
    Check AsmErrorSummary(ErrorAt(1)) = "L1 C13: " & AsmErrorText(ErrorAt(1)), "C: summary format"
    Check AsmInstructionCount() = 12 And AsmInstructionLine(11) = 13, "C: bad-label line holds no instruction"
    Check Not AsmInstructionValid(0) And AsmInstructionValid(7) And AsmInstructionLine(7) = 8, "C: per-instruction validity"

    ' --- D: limits ---------------------------------------------------------------------------
    ReDim src(1 To 66)
    src(1) = "; 64 instrucciones exactas"
    For i = 2 To 65
        src(i) = "NOP"
    Next i
    src(66) = "fin:"
    Check AsmAssemble(src) And AsmInstructionCount() = 64, "D: 64 instructions fit"
    Check AsmLabelAddress("fin") = 256, "D: label after the last slot"
    src(66) = "HLT"
    Check Not AsmAssemble(src) And AsmErrorLine(0) = 66, "D: 65th instruction rejected on its line"
    Check AsmAssembleText("lazo: JMP lazo") And AsmByte(0) = &H30 And AsmByte(2) = 0, "D: label and instruction on one line"
    Check Not AsmAssembleText("; solo comentario" & vbLf & vbLf) And AsmErrorLine(0) = 0, "D: empty program rejected"
    Check Not AsmAssembleText("MOV AX, #FFh"), "D: FFh must be written 0FFh"
    Check Not AsmAssembleText("JMP 100h"), "D: jump outside code"

    ' --- E: loader contract ------------------------------------------------------------------
    CpuWriteRegister LcaDX, &H55
    Check LoaderLoadSource(LoaderSampleProgram()), "E: sample loads"
    cpu = CpuSnapshot()
    Check cpu.PC = 0 And cpu.AX = 0 And cpu.DX = 0 And cpu.flags = 0 And cpu.IR = 0, "E: CPU reset, PC = 000h"
    Check SwapReadByte(0, 0) = &H10 And SwapReadByte(0, 5) = 2 And SwapReadByte(0, &H1C) = 1, "E: code in VPN0 backing"
    Check SwapReadByte(7, 28) = 1 And SwapReadByte(7, 31) = 0, "E: HLT fill up to VPN7"
    Check SwapReadByte(8, 0) = 0 And SwapReadByte(14, 31) = 0, "E: data pages zero"
    For i = 0 To 15
        entry = MmuPage(i)
        Check Not entry.Present, "E: PTE " & CStr(i) & " not present"
    Next i
    For level = 0 To 2
        Check CacheHits(level) = 0 And CacheMisses(level) = 0, "E: cache counters zero"
    Next level
    Check PagerFaultCount() = 0 And BusTraceCount() = 0 And RamReadPhysical(0) = 0, "E: paging, bus and RAM reset"
    Check RamReadPhysical(&HE8) = &H20, "E: LCD buffer eight spaces"
    Check LoaderInstructionCount() = 8 And LoaderCodePages() = 1, "E: one code page used"

    CachePathFetchFour 0
    cpu = CpuSnapshot()
    Check cpu.IR = &H10 And ControlFetched(1) = 0 And PagerFaultCount() = 1, "E: first fetch faults VPN0 in, IR = 10h"
    CachePathFetchFour 4
    Check ControlFetched(1) = 2 And ControlFetched(2) = 5, "E: second instruction MOV CX, #05h"

    ' A failed load keeps the previous image and does not reset execution.
    Check Not LoaderLoadText("MOV EX, #01h"), "E: invalid program not loaded"
    Check LoaderImageByte(4) = &H10 And LoaderImageByte(6) = 5 And PagerFaultCount() = 1, "E: previous image and state kept"
    failure = 0
    On Error Resume Next
    LoaderLoadAssembled
    failure = Err.Number
    Err.Clear
    On Error GoTo failed
    Check failure = LCA_ERR_LOADER, "E: cannot load a failed assembly"

    ' Reset restores the backing image (data back to zero).
    CpuWriteRegister LcaAX, &H77
    CachePathStoreByte &H100, LcaAX
    Check CachePathReadByte(&H100) = &H77, "E: program data written"
    LoaderResetExecution
    Check CachePathReadByte(&H100) = 0 And CpuReadRegister(LcaAX) = 0, "E: reset restores zero data and registers"

    Check LoaderLoadText("MOV BX, #2Ah" & vbCrLf & "HLT"), "E: second program loads"
    Check LoaderImageByte(0) = &H10 And LoaderImageByte(1) = 1 And LoaderImageByte(2) = &H2A, "E: image replaced"
    Check LoaderImageByte(8) = 1 And LoaderInstructionCount() = 2, "E: new HLT fill"

    LcaReport "P11 PASS: " & CStr(mChecks) & " comprobaciones / assertions." & vbCrLf & _
              "Codificacion, errores con linea/columna, limites y contrato del cargador verificados.", True
    Exit Sub
failed:
    LcaReport "P11 FAIL after " & CStr(mChecks) & " checks: " & Err.Description, False
End Sub

Private Function ErrorAt(ByVal lineNo As Long) As Long
    Dim i As Long, found As Long
    ErrorAt = -1
    For i = 0 To AsmErrorCount() - 1
        If AsmErrorLine(i) = lineNo Then
            found = found + 1
            ErrorAt = i
        End If
    Next i
    If found > 1 Then ErrorAt = -1
End Function

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    mChecks = mChecks + 1
    If Not condition Then Err.Raise vbObjectError + 2113, "LcaAssemblerTests", message
End Sub
