; MULTIPLICAR: AX = 6 x 7 sumando 6 siete veces.
        MOV AX, #00h      ; AX = producto
        MOV BX, #06h      ; BX = multiplicando
        MOV CX, #07h      ; CX = multiplicador (cuenta vueltas)
suma:   ADD AX, BX        ; producto = producto + 6
        DEC CX
        JNZ suma
        ST [100h], AX     ; guarda el producto
        HLT
