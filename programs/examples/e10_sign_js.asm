; E10. JS and JNS: is the result negative?
; Input : 5 - 9
; Output: SF = 1 (the result 0FCh is -4 in two's complement),
;         so DX = 1 means "negative"
; Watch : sheet CPU, SF and the jump taken
        MOV AX, 5
        SUB AX, 9
        MOV DX, 0
        JNS DONE          ; not negative -> DX stays 0
        MOV DX, 1         ; negative
DONE:   HLT
