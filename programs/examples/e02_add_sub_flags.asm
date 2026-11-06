; E02. ADD and SUB with the carry flag (CF) and zero flag (ZF)
; Input : AX = 0F0h, BX = 20h
; Output: 0F0h + 20h = 110h -> AX = 10h with CF = 1 (carry)
;         10h - 20h -> AX = 0F0h with CF = 1 (borrow) and SF = 1
;         0F0h - 0F0h -> AX = 0 with ZF = 1
; Watch : sheet CPU, flags ZF, CF and SF after every instruction
        MOV AX, 0F0h
        MOV BX, 20h
        ADD AX, BX        ; carry: CF = 1, AX = 10h
        SUB AX, BX        ; borrow: CF = 1, SF = 1, AX = 0F0h
        SUB AX, 0F0h      ; result 0: ZF = 1, CF = 0
        HLT
