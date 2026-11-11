Attribute VB_Name = "LcaAssembler"
Option Explicit

' P11 two-pass assembler for the fixed P02 ISA (docs/ARCHITECTURE.md section 2).
' Every instruction is four bytes [opcode, R, X, H]; operand address = 256*H + X.
' Syntax (case-insensitive):  [label:] MNEMONIC [op1[, op2]]   ; comment
'   registers AX BX CX DX | immediate #2Ah | data address [104h] | jump target 008h or label
'   numbers: 1..3 hex digits + h suffix, first character 0-9 (write 0FFh, not FFh)
' All errors are collected (not only the first) with source line and column. No worksheet access.
Public Const LCA_ASM_MAX_INSTRUCTIONS As Long = 64
Public Const LCA_ASM_CODE_BYTES As Long = 256
Private Const MAX_ERRORS As Long = 200

Private Type AsmLine
    Label As String
    LabelCol As Long
    Mnemonic As String
    MnemonicCol As Long
    OpCount As Long
    Op(0 To 3) As String
    OpCol(0 To 3) As Long
    ErrText As String
    ErrCol As Long
End Type

Private mCode(0 To 255) As Byte
Private mInstrCount As Long
Private mInstrLine(0 To 63) As Long
Private mInstrOk(0 To 63) As Boolean
Private mErrLine(0 To MAX_ERRORS - 1) As Long
Private mErrCol(0 To MAX_ERRORS - 1) As Long
Private mErrText(0 To MAX_ERRORS - 1) As String
Private mErrCount As Long
Private mLabelName() As String
Private mLabelAddr() As Long
Private mLabelCount As Long
Private mSucceeded As Boolean
Private mLineCount As Long

' Assembles an array of source lines (any lower bound; line 1 = first element).
' Returns True only when there is at least one instruction and no error.
Public Function AsmAssemble(ByVal source As Variant) As Boolean
    Dim first As Long, last As Long, i As Long, lineNo As Long, addr As Long
    Dim ln As AsmLine, count As Long, b(0 To 3) As Long, k As Long
    ClearState
    If Not IsArray(source) Then
        AddError 0, 0, "Fuente vacia / empty source"
        Exit Function
    End If
    first = LBound(source)
    last = UBound(source)
    mLineCount = last - first + 1

    ' Pass 1: syntax shape and label addresses.
    count = 0
    For i = first To last
        lineNo = i - first + 1
        ln = ParseLine(CStr(source(i)))
        If Len(ln.ErrText) > 0 Then AddError lineNo, ln.ErrCol, ln.ErrText
        If Len(ln.Label) > 0 Then DefineLabel ln.Label, count * 4, lineNo, ln.LabelCol
        If Len(ln.Mnemonic) > 0 Then
            count = count + 1
            If count = LCA_ASM_MAX_INSTRUCTIONS + 1 Then _
                AddError lineNo, ln.MnemonicCol, "Mas de 64 instrucciones (256 bytes de codigo) / more than 64 instructions"
        End If
    Next i

    ' Pass 2: encode.
    count = 0
    For i = first To last
        lineNo = i - first + 1
        ln = ParseLine(CStr(source(i)))
        If Len(ln.Mnemonic) > 0 Then
            If count < LCA_ASM_MAX_INSTRUCTIONS Then
                mInstrLine(count) = lineNo
                If Len(ln.ErrText) = 0 Then
                    If EncodeLine(ln, lineNo, b) Then
                        addr = count * 4
                        For k = 0 To 3
                            mCode(addr + k) = CByte(b(k))
                        Next k
                        mInstrOk(count) = True
                    End If
                End If
            End If
            count = count + 1
        End If
    Next i
    If count = 0 And mErrCount = 0 Then AddError 0, 0, "El programa no tiene instrucciones / no instructions"
    If count > LCA_ASM_MAX_INSTRUCTIONS Then count = LCA_ASM_MAX_INSTRUCTIONS
    mInstrCount = count
    ' P02: after the last instruction the 256-byte image is filled with HLT (01 00 00 00).
    For addr = mInstrCount * 4 To LCA_ASM_CODE_BYTES - 4 Step 4
        mCode(addr) = 1
        mCode(addr + 1) = 0
        mCode(addr + 2) = 0
        mCode(addr + 3) = 0
    Next addr
    mSucceeded = (mErrCount = 0)
    AsmAssemble = mSucceeded
End Function

' Convenience: splits text on line breaks and assembles it.
Public Function AsmAssembleText(ByVal text As String) As Boolean
    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbCr, vbLf)
    AsmAssembleText = AsmAssemble(Split(text, vbLf))
End Function

' --- Results ----------------------------------------------------------------------------------

Public Function AsmSucceeded() As Boolean
    AsmSucceeded = mSucceeded
End Function

Public Function AsmInstructionCount() As Long
    AsmInstructionCount = mInstrCount
End Function

' Source line (1-based) of instruction index 0..count-1.
Public Function AsmInstructionLine(ByVal index As Long) As Long
    If index < 0 Or index >= mInstrCount Then Err.Raise LCA_ERR_LOADER, "LcaAssembler", "ASM_INDEX"
    AsmInstructionLine = mInstrLine(index)
End Function

' False when that instruction had an error (its four bytes are not meaningful).
Public Function AsmInstructionValid(ByVal index As Long) As Boolean
    If index < 0 Or index >= mInstrCount Then Err.Raise LCA_ERR_LOADER, "LcaAssembler", "ASM_INDEX"
    AsmInstructionValid = mInstrOk(index)
End Function

Public Function AsmByte(ByVal address As Long) As Byte
    If address < 0 Or address >= LCA_ASM_CODE_BYTES Then Err.Raise LCA_ERR_LOADER, "LcaAssembler", "ASM_ADDRESS"
    AsmByte = mCode(address)
End Function

Public Function AsmErrorCount() As Long
    AsmErrorCount = mErrCount
End Function

Public Function AsmErrorLine(ByVal index As Long) As Long
    RequireError index
    AsmErrorLine = mErrLine(index)
End Function

Public Function AsmErrorColumn(ByVal index As Long) As Long
    RequireError index
    AsmErrorColumn = mErrCol(index)
End Function

Public Function AsmErrorText(ByVal index As Long) As String
    RequireError index
    AsmErrorText = mErrText(index)
End Function

' "L5 C9: message" (line 0 = whole program).
Public Function AsmErrorSummary(ByVal index As Long) As String
    RequireError index
    AsmErrorSummary = "L" & CStr(mErrLine(index)) & " C" & CStr(mErrCol(index)) & ": " & mErrText(index)
End Function

' Address of a label, or -1 when undefined.
Public Function AsmLabelAddress(ByVal name As String) As Long
    Dim i As Long
    AsmLabelAddress = -1
    For i = 0 To mLabelCount - 1
        If mLabelName(i) = UCase$(name) Then
            AsmLabelAddress = mLabelAddr(i)
            Exit Function
        End If
    Next i
End Function

' Canonical text of one four-byte instruction, or "INVALID_OPCODE xxh" / "INVALID_ENCODING ...".
' Used for the listing and as a decoder reference for P12.
Public Function AsmDisassemble(ByVal op As Long, ByVal r As Long, ByVal x As Long, ByVal h As Long) As String
    Dim reg As String, target As Long
    If r >= 0 And r <= 3 Then reg = RegName(r)
    Select Case op
        Case &H0, &H1
            If r <> 0 Or x <> 0 Or h <> 0 Then GoTo bad
            If op = 0 Then AsmDisassemble = "NOP" Else AsmDisassemble = "HLT"
        Case &H10
            If r > 3 Or h <> 0 Then GoTo bad
            AsmDisassemble = "MOV " & reg & ", #" & AsmNumber(x)
        Case &H11, &H20, &H21, &H22
            If r > 3 Or x > 3 Or h <> 0 Then GoTo bad
            AsmDisassemble = MnemonicOf(op) & " " & reg & ", " & RegName(x)
        Case &H12, &H13
            target = 256 * h + x
            If r > 3 Or h > 1 Or target < &H100 Or target > &H1DF Then GoTo bad
            If op = &H12 Then
                AsmDisassemble = "LD " & reg & ", [" & Hex3(target) & "]"
            Else
                AsmDisassemble = "ST [" & Hex3(target) & "], " & reg
            End If
        Case &H23, &H24, &H40, &H41
            If r > 3 Or x <> 0 Or h <> 0 Then GoTo bad
            AsmDisassemble = MnemonicOf(op) & " " & reg
        Case &H30 To &H34
            target = 256 * h + x
            If r <> 0 Or h > 1 Or target > &HFC Or (target Mod 4) <> 0 Then GoTo bad
            AsmDisassemble = MnemonicOf(op) & " " & Hex3(target)
        Case Else
            AsmDisassemble = "INVALID_OPCODE " & Hex2(op)
    End Select
    Exit Function
bad:
    AsmDisassemble = "INVALID_ENCODING " & Hex2(op) & " " & Hex2(r) & " " & Hex2(x) & " " & Hex2(h)
End Function

' Assembler-valid hex literal: 2Ah, 0FFh (a leading 0 when the first digit is a letter).
Public Function AsmNumber(ByVal value As Long) As String
    AsmNumber = Hex$(value)
    If Len(AsmNumber) < 2 Then AsmNumber = "0" & AsmNumber
    If Left$(AsmNumber, 1) >= "A" Then AsmNumber = "0" & AsmNumber
    AsmNumber = AsmNumber & "h"
End Function

Public Function AsmHex2(ByVal value As Long) As String
    AsmHex2 = Hex2(value)
End Function

Public Function AsmHex3(ByVal value As Long) As String
    AsmHex3 = Hex3(value)
End Function

' --- Pass helpers -------------------------------------------------------------------------------

Private Sub ClearState()
    Dim i As Long
    For i = 0 To LCA_ASM_CODE_BYTES - 1
        mCode(i) = 0
    Next i
    For i = 0 To LCA_ASM_MAX_INSTRUCTIONS - 1
        mInstrLine(i) = 0
        mInstrOk(i) = False
    Next i
    mInstrCount = 0
    mErrCount = 0
    mLabelCount = 0
    ReDim mLabelName(0 To 0)
    ReDim mLabelAddr(0 To 0)
    mSucceeded = False
    mLineCount = 0
End Sub

Private Sub DefineLabel(ByVal name As String, ByVal address As Long, ByVal lineNo As Long, ByVal col As Long)
    If AsmLabelAddress(name) >= 0 Then
        AddError lineNo, col, "Etiqueta duplicada '" & name & "' / duplicate label"
        Exit Sub
    End If
    ReDim Preserve mLabelName(0 To mLabelCount)
    ReDim Preserve mLabelAddr(0 To mLabelCount)
    mLabelName(mLabelCount) = UCase$(name)
    mLabelAddr(mLabelCount) = address
    mLabelCount = mLabelCount + 1
End Sub

Private Sub AddError(ByVal lineNo As Long, ByVal col As Long, ByVal text As String)
    If mErrCount >= MAX_ERRORS Then Exit Sub
    mErrLine(mErrCount) = lineNo
    mErrCol(mErrCount) = col
    mErrText(mErrCount) = text
    mErrCount = mErrCount + 1
End Sub

Private Sub RequireError(ByVal index As Long)
    If index < 0 Or index >= mErrCount Then Err.Raise LCA_ERR_LOADER, "LcaAssembler", "ASM_ERROR_INDEX"
End Sub

' Splits one line into label, mnemonic and comma-separated operands, keeping 1-based columns.
Private Function ParseLine(ByVal text As String) As AsmLine
    Dim ln As AsmLine, s As String, p As Long, q As Long, n As Long
    Dim candidate As String, startPos As Long, piece As String, lead As Long
    s = Replace(text, vbTab, " ")
    p = InStr(1, s, ";")
    If p > 0 Then s = Left$(s, p - 1)
    n = Len(s)
    p = SkipSpaces(s, 1)
    If p > n Then
        ParseLine = ln
        Exit Function
    End If
    q = InStr(p, s, ":")
    If q > 0 Then
        candidate = RTrim$(Mid$(s, p, q - p))
        If IsIdentifier(candidate) Then
            ln.Label = candidate
            ln.LabelCol = p
        Else
            ln.ErrText = "Etiqueta invalida '" & candidate & "' (letra o _ y luego letras, digitos o _; no registro ni mnemonico) / invalid label"
            ln.ErrCol = p
            ParseLine = ln
            Exit Function
        End If
        p = SkipSpaces(s, q + 1)
        If p > n Then
            ParseLine = ln
            Exit Function
        End If
    End If
    startPos = p
    Do While p <= n
        If Mid$(s, p, 1) = " " Then Exit Do
        p = p + 1
    Loop
    ln.Mnemonic = UCase$(Mid$(s, startPos, p - startPos))
    ln.MnemonicCol = startPos
    p = SkipSpaces(s, p)
    If p > n Then
        ParseLine = ln
        Exit Function
    End If
    ' Operands: split the remainder on commas.
    startPos = p
    Do
        q = InStr(startPos, s, ",")
        If q = 0 Then
            piece = Mid$(s, startPos)
        Else
            piece = Mid$(s, startPos, q - startPos)
        End If
        lead = Len(piece) - Len(LTrim$(piece))
        If Len(Trim$(piece)) = 0 Then
            ln.ErrText = "Operando vacio / empty operand"
            ln.ErrCol = startPos
            Exit Do
        End If
        If ln.OpCount = 4 Then
            ln.ErrText = "Demasiados operandos / too many operands"
            ln.ErrCol = startPos + lead
            Exit Do
        End If
        ln.Op(ln.OpCount) = Trim$(piece)
        ln.OpCol(ln.OpCount) = startPos + lead
        ln.OpCount = ln.OpCount + 1
        If q = 0 Then Exit Do
        startPos = q + 1
    Loop
    ParseLine = ln
End Function

Private Function SkipSpaces(ByVal s As String, ByVal p As Long) As Long
    Do While p <= Len(s)
        If Mid$(s, p, 1) <> " " Then Exit Do
        p = p + 1
    Loop
    SkipSpaces = p
End Function

' Validates operand forms and fields; on success fills b(0..3).
Private Function EncodeLine(ByRef ln As AsmLine, ByVal lineNo As Long, ByRef b() As Long) As Boolean
    Dim op As Long, r As Long, v As Long, need As Long
    b(0) = 0: b(1) = 0: b(2) = 0: b(3) = 0
    op = OpcodeOf(ln.Mnemonic)
    If op < 0 Then
        AddError lineNo, ln.MnemonicCol, "Mnemonico desconocido '" & ln.Mnemonic & "' / unknown mnemonic"
        Exit Function
    End If
    Select Case op
        Case &H0, &H1: need = 0
        Case &H23, &H24, &H40, &H41, &H30 To &H34: need = 1
        Case Else: need = 2
    End Select
    If ln.OpCount <> need Then
        AddError lineNo, ln.MnemonicCol, ln.Mnemonic & " requiere " & CStr(need) & " operando(s), tiene " & _
                 CStr(ln.OpCount) & " / expects " & CStr(need) & " operand(s)"
        Exit Function
    End If
    b(0) = op
    Select Case op
        Case &H0, &H1
        Case &H10                                  ' MOV R,#imm | MOV R,R2
            If Not NeedRegister(ln, 0, lineNo, r) Then Exit Function
            b(1) = r
            If Left$(ln.Op(1), 1) = "#" Then
                If Not NeedNumber(Mid$(ln.Op(1), 2), ln.OpCol(1) + 1, lineNo, &HFF, "inmediato 00h-0FFh / immediate", v) Then Exit Function
                b(2) = v
            ElseIf RegCode(ln.Op(1)) >= 0 Then
                b(0) = &H11
                b(2) = RegCode(ln.Op(1))
            Else
                AddError lineNo, ln.OpCol(1), "Se esperaba #inmediato o registro, no '" & ln.Op(1) & "' / expected #imm or register"
                Exit Function
            End If
        Case &H12                                  ' LD R,[VA]
            If Not NeedRegister(ln, 0, lineNo, r) Then Exit Function
            If Not NeedDataAddress(ln, 1, lineNo, v) Then Exit Function
            b(1) = r: b(2) = v Mod 256: b(3) = v \ 256
        Case &H13                                  ' ST [VA],R
            If Not NeedDataAddress(ln, 0, lineNo, v) Then Exit Function
            If Not NeedRegister(ln, 1, lineNo, r) Then Exit Function
            b(1) = r: b(2) = v Mod 256: b(3) = v \ 256
        Case &H20, &H21, &H22                      ' ADD/SUB/CMP R,R2
            If Not NeedRegister(ln, 0, lineNo, r) Then Exit Function
            b(1) = r
            If Not NeedRegister(ln, 1, lineNo, r) Then Exit Function
            b(2) = r
        Case &H23, &H24, &H40, &H41                ' INC/DEC/IN/OUT R
            If Not NeedRegister(ln, 0, lineNo, r) Then Exit Function
            b(1) = r
        Case &H30 To &H34                          ' jumps
            If Not NeedTarget(ln, lineNo, v) Then Exit Function
            b(2) = v Mod 256: b(3) = v \ 256
    End Select
    EncodeLine = True
End Function

Private Function NeedRegister(ByRef ln As AsmLine, ByVal i As Long, ByVal lineNo As Long, ByRef r As Long) As Boolean
    r = RegCode(ln.Op(i))
    If r < 0 Then
        AddError lineNo, ln.OpCol(i), "Registro invalido '" & ln.Op(i) & "' (AX, BX, CX, DX) / invalid register"
        Exit Function
    End If
    NeedRegister = True
End Function

Private Function NeedDataAddress(ByRef ln As AsmLine, ByVal i As Long, ByVal lineNo As Long, ByRef v As Long) As Boolean
    Dim t As String
    t = ln.Op(i)
    If Left$(t, 1) <> "[" Or Right$(t, 1) <> "]" Then
        AddError lineNo, ln.OpCol(i), "Se esperaba [direccion] de datos, ej. [104h] / expected [address]"
        Exit Function
    End If
    If Not NeedNumber(Trim$(Mid$(t, 2, Len(t) - 2)), ln.OpCol(i) + 1, lineNo, &H1FF, "direccion / address", v) Then Exit Function
    If v >= &H1E0 Then
        AddError lineNo, ln.OpCol(i), "Direccion " & Hex3(v) & " es E/S: usar IN/OUT / I/O address, use IN/OUT"
        Exit Function
    End If
    If v < &H100 Then
        AddError lineNo, ln.OpCol(i), "Direccion " & Hex3(v) & " fuera de datos 100h-1DFh / outside data region"
        Exit Function
    End If
    NeedDataAddress = True
End Function

Private Function NeedTarget(ByRef ln As AsmLine, ByVal lineNo As Long, ByRef v As Long) As Boolean
    Dim t As String
    t = ln.Op(0)
    If IsIdentifier(t) Then
        v = AsmLabelAddress(t)
        If v < 0 Then
            AddError lineNo, ln.OpCol(0), "Etiqueta no definida '" & t & "' / undefined label"
            Exit Function
        End If
    Else
        If Not NeedNumber(t, ln.OpCol(0), lineNo, &H1FF, "destino / target", v) Then Exit Function
    End If
    If v > &HFC Or (v Mod 4) <> 0 Then
        AddError lineNo, ln.OpCol(0), "Destino " & Hex3(v) & " invalido: 000h-0FCh y multiplo de 4 / invalid branch target"
        Exit Function
    End If
    NeedTarget = True
End Function

' Hex number with h suffix, 1..3 digits, first character 0-9.
Private Function NeedNumber(ByVal t As String, ByVal col As Long, ByVal lineNo As Long, _
                            ByVal maxValue As Long, ByVal what As String, ByRef v As Long) As Boolean
    Dim body As String, i As Long, c As String
    t = UCase$(Trim$(t))
    If Len(t) < 2 Or Right$(t, 1) <> "H" Then
        If Len(t) > 0 And IsAllDigits(t) Then
            AddError lineNo, col, "Numero '" & t & "' sin sufijo h: usar hexadecimal, ej. " & t & "h / use hex with h suffix"
        Else
            AddError lineNo, col, "Numero invalido '" & t & "' (hex con sufijo h, ej. 2Ah, 0FFh) / invalid number"
        End If
        Exit Function
    End If
    body = Left$(t, Len(t) - 1)
    If Len(body) > 3 Or Not (Left$(body, 1) >= "0" And Left$(body, 1) <= "9") Then
        AddError lineNo, col, "Numero invalido '" & t & "' (1-3 digitos hex, empieza por 0-9: 0FFh) / invalid number"
        Exit Function
    End If
    For i = 1 To Len(body)
        c = Mid$(body, i, 1)
        If Not ((c >= "0" And c <= "9") Or (c >= "A" And c <= "F")) Then
            AddError lineNo, col, "Numero invalido '" & t & "' / invalid hex digit"
            Exit Function
        End If
    Next i
    v = CLng("&H" & body)
    If v > maxValue Then
        AddError lineNo, col, "Valor " & t & " fuera de rango para " & what & " (max " & Hex3(maxValue) & ") / out of range"
        Exit Function
    End If
    NeedNumber = True
End Function

Private Function IsAllDigits(ByVal t As String) As Boolean
    Dim i As Long
    For i = 1 To Len(t)
        If Mid$(t, i, 1) < "0" Or Mid$(t, i, 1) > "9" Then Exit Function
    Next i
    IsAllDigits = True
End Function

Private Function IsIdentifier(ByVal t As String) As Boolean
    Dim i As Long, c As String
    If Len(t) = 0 Or Len(t) > 32 Then Exit Function
    For i = 1 To Len(t)
        c = UCase$(Mid$(t, i, 1))
        If Not ((c >= "A" And c <= "Z") Or c = "_" Or (i > 1 And c >= "0" And c <= "9")) Then Exit Function
    Next i
    If RegCode(t) >= 0 Or OpcodeOf(UCase$(t)) >= 0 Or UCase$(t) = "MOV" Then Exit Function
    IsIdentifier = True
End Function

Private Function RegCode(ByVal t As String) As Long
    Select Case UCase$(Trim$(t))
        Case "AX": RegCode = 0
        Case "BX": RegCode = 1
        Case "CX": RegCode = 2
        Case "DX": RegCode = 3
        Case Else: RegCode = -1
    End Select
End Function

Private Function RegName(ByVal code As Long) As String
    RegName = Choose(code + 1, "AX", "BX", "CX", "DX")
End Function

' MOV maps to 10h here; EncodeLine switches to 11h for a register source.
Private Function OpcodeOf(ByVal m As String) As Long
    Select Case m
        Case "NOP": OpcodeOf = &H0
        Case "HLT": OpcodeOf = &H1
        Case "MOV": OpcodeOf = &H10
        Case "LD": OpcodeOf = &H12
        Case "ST": OpcodeOf = &H13
        Case "ADD": OpcodeOf = &H20
        Case "SUB": OpcodeOf = &H21
        Case "CMP": OpcodeOf = &H22
        Case "INC": OpcodeOf = &H23
        Case "DEC": OpcodeOf = &H24
        Case "JMP": OpcodeOf = &H30
        Case "JZ": OpcodeOf = &H31
        Case "JNZ": OpcodeOf = &H32
        Case "JC": OpcodeOf = &H33
        Case "JS": OpcodeOf = &H34
        Case "IN": OpcodeOf = &H40
        Case "OUT": OpcodeOf = &H41
        Case Else: OpcodeOf = -1
    End Select
End Function

Private Function MnemonicOf(ByVal op As Long) As String
    Select Case op
        Case &H11: MnemonicOf = "MOV"
        Case &H20: MnemonicOf = "ADD"
        Case &H21: MnemonicOf = "SUB"
        Case &H22: MnemonicOf = "CMP"
        Case &H23: MnemonicOf = "INC"
        Case &H24: MnemonicOf = "DEC"
        Case &H30: MnemonicOf = "JMP"
        Case &H31: MnemonicOf = "JZ"
        Case &H32: MnemonicOf = "JNZ"
        Case &H33: MnemonicOf = "JC"
        Case &H34: MnemonicOf = "JS"
        Case &H40: MnemonicOf = "IN"
        Case &H41: MnemonicOf = "OUT"
    End Select
End Function

Private Function Hex2(ByVal value As Long) As String
    Hex2 = Right$("00" & Hex$(value), 2) & "h"
End Function

Private Function Hex3(ByVal value As Long) As String
    Hex3 = Right$("000" & Hex$(value), 3) & "h"
End Function
