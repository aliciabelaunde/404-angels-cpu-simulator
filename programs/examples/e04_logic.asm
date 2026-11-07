; E04. AND, OR, XOR and NOT on bits
; Input : AX = 0B4h (1011 0100)
; Output: AND 0Fh -> 04h (keeps the low 4 bits)
;         OR 0F0h -> 0F4h (turns the high 4 bits on)
;         XOR 0FFh -> 4Bh (flips every bit)
;         NOT     -> 4Bh (also flips every bit, but keeps the flags)
;         BX = 04h, CX = 0F4h, DX = 4Bh, AX = 4Bh
; Watch : sheet CPU, CF is always 0 after AND, OR and XOR
        MOV AX, 0B4h
        MOV BX, AX
        AND BX, 0Fh
        MOV CX, AX
        OR CX, 0F0h
        MOV DX, AX
        XOR DX, 0FFh
        NOT AX
        HLT
