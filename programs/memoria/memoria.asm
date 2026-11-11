; MEMORIA: dispersion (faults y misses), repeticion (hits), mas paginas que marcos y recuperar un dato.
        MOV AX, #0A1h     ; --- fase 1: DISPERSION, siete paginas distintas ---
        ST [100h], AX     ; pagina 8   (este dato se recupera en la fase 2)
        MOV AX, #0A2h
        ST [140h], AX     ; pagina 10
        MOV AX, #0A3h
        ST [180h], AX     ; pagina 12
        MOV AX, #0A4h
        ST [1C0h], AX     ; pagina 14
        MOV AX, #0A5h
        ST [120h], AX     ; pagina 9
        MOV AX, #0A6h
        ST [160h], AX     ; pagina 11
        MOV AX, #0A7h
        ST [1A0h], AX     ; pagina 13   (la ultima: sigue residente)
        MOV CX, #04h      ; --- fase 2: REPETICION, cuatro lecturas del mismo byte ---
rep:    LD BX, [100h]     ; 1a vez: recupera A1h del swap; luego acierta en cache
        DEC CX
        JNZ rep
        LD DX, [1A0h]     ; --- fase 3: pagina residente, sin falta ---
        ST [101h], BX     ; --- fase 4: copia el dato recuperado ---
        HLT
