Attribute VB_Name = "LcaArith"
Option Explicit

' P15: catalogue of arithmetic demo programs written in the P11 assembly. Every program uses AX/BX/CX/DX
' and a loop, stores its result in data memory (VPN 8: [100h], [101h]) and ends with HLT.
' Expected results, instruction counts and loop figures are derived by hand and asserted by
' VerificarDemoAritmeticaP15. No worksheet dependency.
'   0 SUMA         AX = 5+4+3+2+1                       -> AX = 15
'   1 MULTIPLICAR  AX = 6 x 7 by repeated addition      -> AX = 42
'   2 FIBONACCI    F(10) and F(11)                      -> AX = 55, BX = 89
'   3 DIVIDIR      47 / 5 by repeated subtraction       -> CX = 9 (quotient), AX = 2 (remainder)
'   4 MCD          gcd(48, 18) by Euclid's subtraction  -> AX = 6
Public Const ARITH_COUNT As Long = 5

' Number of programs (callable from scripts, which cannot read the Const).
Public Function ArithCountValue() As Long
    ArithCountValue = ARITH_COUNT
End Function

Public Function ArithName(ByVal i As Long) As String
    ArithName = Choose(i + 1, "SUMA", "MULTIPLICAR", "FIBONACCI", "DIVIDIR", "MCD")
End Function

Public Function ArithTitle(ByVal i As Long) As String
    ArithTitle = Choose(i + 1, "Suma 5+4+3+2+1 con un bucle", "Multiplicar 6 x 7 sumando", _
                        "Serie de Fibonacci (10 terminos)", "Dividir 47 / 5 restando", "Maximo comun divisor de 48 y 18")
End Function

Public Function ArithFile(ByVal i As Long) As String
    ArithFile = Choose(i + 1, "01_suma.asm", "02_multiplicacion.asm", "03_fibonacci.asm", "04_division.asm", "05_mcd.asm")
End Function

Public Function ArithSource(ByVal i As Long) As Variant
    Select Case i
        Case 0
            ArithSource = Array( _
                "; SUMA: AX = 5+4+3+2+1.  Entrada: CX = 5 (cuantos sumandos).", _
                "        MOV AX, #00h      ; AX acumula la suma", _
                "        MOV CX, #05h      ; CX cuenta las vueltas", _
                "bucle:  ADD AX, CX        ; AX = AX + CX", _
                "        DEC CX            ; CX = CX - 1", _
                "        JNZ bucle         ; repite mientras CX <> 0", _
                "        ST [100h], AX     ; guarda el resultado en memoria", _
                "        LD BX, [100h]     ; y lo lee de vuelta en BX", _
                "        HLT")
        Case 1
            ArithSource = Array( _
                "; MULTIPLICAR: AX = 6 x 7 sumando 6 siete veces.", _
                "        MOV AX, #00h      ; AX = producto", _
                "        MOV BX, #06h      ; BX = multiplicando", _
                "        MOV CX, #07h      ; CX = multiplicador (cuenta vueltas)", _
                "suma:   ADD AX, BX        ; producto = producto + 6", _
                "        DEC CX", _
                "        JNZ suma", _
                "        ST [100h], AX     ; guarda el producto", _
                "        HLT")
        Case 2
            ArithSource = Array( _
                "; FIBONACCI: 10 terminos. AX = F(n), BX = F(n+1), DX = temporal.", _
                "        MOV AX, #00h      ; F(0)", _
                "        MOV BX, #01h      ; F(1)", _
                "        MOV CX, #0Ah      ; 10 vueltas", _
                "fib:    MOV DX, AX", _
                "        ADD DX, BX        ; DX = F(n) + F(n+1)", _
                "        MOV AX, BX        ; AX = F(n+1)", _
                "        MOV BX, DX        ; BX = F(n+2)", _
                "        DEC CX", _
                "        JNZ fib", _
                "        ST [100h], AX     ; F(10)", _
                "        HLT")
        Case 3
            ArithSource = Array( _
                "; DIVIDIR: 47 / 5 restando 5 hasta que ya no cabe.", _
                "        MOV AX, #2Fh      ; AX = dividendo (47), al final el resto", _
                "        MOV BX, #05h      ; BX = divisor", _
                "        MOV CX, #00h      ; CX = cociente", _
                "rest:   CMP AX, BX        ; compara sin modificar", _
                "        JC fin            ; AX < BX: ya no cabe otro divisor", _
                "        SUB AX, BX", _
                "        INC CX            ; un divisor mas cabe", _
                "        JMP rest", _
                "fin:    ST [100h], CX     ; cociente", _
                "        ST [101h], AX     ; resto", _
                "        HLT")
        Case Else
            ArithSource = Array( _
                "; MCD de 48 y 18 por el metodo de Euclides con restas.", _
                "        MOV AX, #30h      ; AX = 48", _
                "        MOV BX, #12h      ; BX = 18", _
                "mcd:    CMP AX, BX", _
                "        JZ fin            ; iguales: ese es el MCD", _
                "        JC menor          ; AX < BX", _
                "        SUB AX, BX        ; AX = AX - BX", _
                "        JMP mcd", _
                "menor:  SUB BX, AX        ; BX = BX - AX", _
                "        JMP mcd", _
                "fin:    ST [100h], AX     ; guarda el MCD", _
                "        HLT")
    End Select
End Function

' Source text exactly as stored in programs/arith/<file> (LF line ends).
Public Function ArithAsmText(ByVal i As Long) As String
    ArithAsmText = Join(ArithSource(i), vbLf)
End Function

Public Function ArithInputText(ByVal i As Long) As String
    ArithInputText = Choose(i + 1, _
        "CX = 5 (cuantos sumandos). AX empieza en 0.", _
        "BX = 6 (multiplicando) y CX = 7 (multiplicador). AX empieza en 0.", _
        "AX = F(0) = 0, BX = F(1) = 1 y CX = 10 vueltas.", _
        "AX = 47 (dividendo) y BX = 5 (divisor). CX empieza en 0.", _
        "AX = 48 y BX = 18.")
End Function

Public Function ArithExpectedText(ByVal i As Long) As String
    ArithExpectedText = Choose(i + 1, _
        "AX = 15 (0Fh); BX = 15; [100h] = 15", _
        "AX = 42 (2Ah); BX = 6; [100h] = 42", _
        "AX = 55 (37h), BX = 89 (59h); [100h] = 55", _
        "CX = 9 (cociente), AX = 2 (resto); [100h] = 9, [101h] = 2", _
        "AX = 6 y BX = 6; [100h] = 6")
End Function

Public Function ArithAlgorithm(ByVal i As Long) As String
    ArithAlgorithm = Choose(i + 1, _
        "Un acumulador en AX y un contador en CX: cada vuelta suma CX a AX y baja CX; JNZ repite hasta CX = 0. Cinco vueltas: 5+4+3+2+1.", _
        "Multiplicar es sumar repetido: cada vuelta suma el multiplicando (BX) al producto (AX) y descuenta una vuelta (CX).", _
        "AX y BX guardan dos terminos seguidos; DX calcula el siguiente. Los registros se desplazan: AX <- BX, BX <- DX. Cuatro registros trabajando juntos.", _
        "Dividir es restar repetido: mientras el divisor cabe, se resta y se cuenta un cociente. CMP compara sin destruir y el acarreo (CF) avisa cuando AX < BX.", _
        "Euclides: se resta el menor del mayor hasta que ambos son iguales. CMP + JZ detecta el final y JC elige cual de los dos restar.")
End Function

' Expected register after HLT (reg 0..3 = AX..DX).
Public Function ArithExpectedReg(ByVal i As Long, ByVal reg As Long) As Long
    Select Case i
        Case 0: ArithExpectedReg = Choose(reg + 1, 15, 15, 0, 0)
        Case 1: ArithExpectedReg = Choose(reg + 1, 42, 6, 0, 0)
        Case 2: ArithExpectedReg = Choose(reg + 1, 55, 89, 0, 89)
        Case 3: ArithExpectedReg = Choose(reg + 1, 2, 5, 9, 0)
        Case Else: ArithExpectedReg = Choose(reg + 1, 6, 6, 0, 0)
    End Select
End Function

' Expected data byte at [100h + offset], or -1 when the program does not write it.
Public Function ArithExpectedMem(ByVal i As Long, ByVal offset As Long) As Long
    ArithExpectedMem = -1
    If offset = 0 Then ArithExpectedMem = Choose(i + 1, 15, 42, 55, 9, 6)
    If offset = 1 And i = 3 Then ArithExpectedMem = 2
End Function

' Instructions executed from the first MOV up to and including HLT.
Public Function ArithExpectedSteps(ByVal i As Long) As Long
    ArithExpectedSteps = Choose(i + 1, 20, 26, 65, 53, 26)
End Function

' Address of the instruction that starts every loop pass, and how many times it runs.
Public Function ArithLoopHead(ByVal i As Long) As Long
    ArithLoopHead = Choose(i + 1, &H8, &HC, &HC, &HC, &H8)
End Function

Public Function ArithLoopTotal(ByVal i As Long) As Long
    ArithLoopTotal = Choose(i + 1, 5, 7, 10, 10, 5)
End Function

' Top of the register-trace chart axis.
Public Function ArithChartMax(ByVal i As Long) As Long
    ArithChartMax = Choose(i + 1, 20, 50, 100, 50, 50)
End Function

' Final flags after HLT: bit0 = ZF, bit1 = CF, bit2 = SF.
Public Function ArithExpectedFlags(ByVal i As Long) As Long
    ArithExpectedFlags = Choose(i + 1, 1, 1, 1, 6, 1)
End Function

' Loads and resets: the program lands in the backing store of VPN 0.., the whole machine is reset.
Public Function ArithLoad(ByVal i As Long) As Boolean
    ArithLoad = LoaderLoadSource(ArithSource(i))
End Function

' Data byte at [100h + offset] without touching caches or counters: RAM when the page is resident,
' otherwise the swap image.
Public Function ArithMemByte(ByVal offset As Long) As Long
    Dim e As LcaPageEntry
    e = MmuPage(8)
    If e.Present Then
        ArithMemByte = RamReadUser(e.PFN * 32 + offset)
    Else
        ArithMemByte = SwapReadByte(8, offset)
    End If
End Function

' True when the machine halted with exactly the expected registers, flags and memory.
Public Function ArithMatches(ByVal i As Long) As Boolean
    Dim r As Long
    If ExecStatus() <> EXEC_HALTED Then Exit Function
    For r = 0 To 3
        If CpuReadRegister(r) <> ArithExpectedReg(i, r) Then Exit Function
    Next r
    If ArithExpectedMem(i, 0) >= 0 And ArithMemByte(0) <> ArithExpectedMem(i, 0) Then Exit Function
    If ArithExpectedMem(i, 1) >= 0 And ArithMemByte(1) <> ArithExpectedMem(i, 1) Then Exit Function
    ArithMatches = True
End Function

' Short live result of the running machine, e.g. "AX = 42 (2Ah)".
Public Function ArithResultText(ByVal i As Long) As String
    Select Case i
        Case 2
            ArithResultText = "AX = " & CStr(CpuReadRegister(0)) & ",  BX = " & CStr(CpuReadRegister(1))
        Case 3
            ArithResultText = "CX = " & CStr(CpuReadRegister(2)) & ",  AX = " & CStr(CpuReadRegister(0))
        Case Else
            ArithResultText = "AX = " & CStr(CpuReadRegister(0)) & " (" & AsmHex2(CpuReadRegister(0)) & ")"
    End Select
End Function
