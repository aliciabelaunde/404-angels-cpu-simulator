; E09. Conditional jumps JC and JNC: the bigger of two numbers
; Input : X = 47 at 100h, Y = 83 at 101h
; Output: MAX = 83 = 53h at 102h
; Watch : sheet CPU, CMP sets CF = 1 because X < Y, so JC jumps
        LOAD AX, [X]
        LOAD BX, [Y]
        CMP AX, BX
        JC YBIG           ; X < Y -> Y is the max
        JMP SAVE
YBIG:   MOV AX, BX
SAVE:   STORE [MAX], AX
        HLT
        ORG 100h
X:      DB 47
Y:      DB 83
MAX:    DB 0
