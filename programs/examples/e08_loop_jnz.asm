; E08. Loop with DEC and JNZ: sum 1 + 2 + ... + N
; Input : N = 10 at 100h
; Output: AX = 55 = 37h
; Watch : sheet CPU, the jump goes back while ZF = 0
        LOAD CX, [N]
        MOV AX, 0
LOOP:   ADD AX, CX
        DEC CX
        JNZ LOOP          ; repeat while CX is not 0
        HLT
        ORG 100h
N:      DB 10
