; FIBONACCI: 10 terminos. AX = F(n), BX = F(n+1), DX = temporal.
        MOV AX, #00h      ; F(0)
        MOV BX, #01h      ; F(1)
        MOV CX, #0Ah      ; 10 vueltas
fib:    MOV DX, AX
        ADD DX, BX        ; DX = F(n) + F(n+1)
        MOV AX, BX        ; AX = F(n+1)
        MOV BX, DX        ; BX = F(n+2)
        DEC CX
        JNZ fib
        ST [100h], AX     ; F(10)
        HLT
