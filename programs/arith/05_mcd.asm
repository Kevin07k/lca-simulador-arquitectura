; MCD de 48 y 18 por el metodo de Euclides con restas.
        MOV AX, #30h      ; AX = 48
        MOV BX, #12h      ; BX = 18
mcd:    CMP AX, BX
        JZ fin            ; iguales: ese es el MCD
        JC menor          ; AX < BX
        SUB AX, BX        ; AX = AX - BX
        JMP mcd
menor:  SUB BX, AX        ; BX = BX - AX
        JMP mcd
fin:    ST [100h], AX     ; guarda el MCD
        HLT
