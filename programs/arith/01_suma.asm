; SUMA: AX = 5+4+3+2+1.  Entrada: CX = 5 (cuantos sumandos).
        MOV AX, #00h      ; AX acumula la suma
        MOV CX, #05h      ; CX cuenta las vueltas
bucle:  ADD AX, CX        ; AX = AX + CX
        DEC CX            ; CX = CX - 1
        JNZ bucle         ; repite mientras CX <> 0
        ST [100h], AX     ; guarda el resultado en memoria
        LD BX, [100h]     ; y lo lee de vuelta en BX
        HLT
