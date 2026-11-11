; ECO: cada tecla del teclado aparece en el LCD. La tecla '.' termina.
        MOV BX, #2Eh      ; BX = '.' (tecla de fin)
        MOV CX, #00h      ; CX cuenta los caracteres mostrados
eco:    IN AX             ; toma la tecla mas antigua; cola vacia = espera (WAIT_IN)
        CMP AX, BX        ; es el punto?
        JZ fin            ; si: termina sin mostrarlo
        OUT AX            ; no: la escribe en la celda del cursor del LCD
        INC CX            ; un caracter mas
        JMP eco
fin:    ST [100h], CX     ; guarda cuantos caracteres se mostraron
        HLT
