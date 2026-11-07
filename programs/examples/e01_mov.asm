; E01. MOV: copy values into registers
; Input : the immediate values 12h and 34h
; Output: AX = 12h, BX = 34h, CX = 12h, DX = 34h
; Watch : sheet CPU, registers change one by one (MOV does not change flags)
        MOV AX, 12h       ; register <- immediate
        MOV BX, 34h
        MOV CX, AX        ; register <- register
        MOV DX, BX
        HLT
