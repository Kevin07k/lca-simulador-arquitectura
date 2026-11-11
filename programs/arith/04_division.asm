; DIVIDIR: 47 / 5 restando 5 hasta que ya no cabe.
        MOV AX, #2Fh      ; AX = dividendo (47), al final el resto
        MOV BX, #05h      ; BX = divisor
        MOV CX, #00h      ; CX = cociente
rest:   CMP AX, BX        ; compara sin modificar
        JC fin            ; AX < BX: ya no cabe otro divisor
        SUB AX, BX
        INC CX            ; un divisor mas cabe
        JMP rest
fin:    ST [100h], CX     ; cociente
        ST [101h], AX     ; resto
        HLT
