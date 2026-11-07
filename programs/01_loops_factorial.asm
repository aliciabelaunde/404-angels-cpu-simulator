; 1. Loops: factorial with AX, BX, CX, DX
; RES = N! using nested loops (multiplication by repeated addition)
;   CX = outer counter (n, n-1, ... 2)
;   AX = result     BX = copy of the result     DX = inner counter
        LOAD CX, [N]      ; CX = n
        MOV AX, 1         ; AX = 1
OUTER:  CMP CX, 1
        JZ DONE           ; CX = 1 -> finished
        JC DONE           ; CX = 0 -> finished (0! = 1)
        MOV BX, AX        ; BX = result so far
        MOV DX, CX
        DEC DX            ; DX = CX - 1 additions
INNER:  ADD AX, BX        ; AX = AX + BX
        DEC DX
        JNZ INNER
        DEC CX
        JMP OUTER
DONE:   STORE [RES], AX   ; RES = n!
        HLT
        ORG 100h          ; data area (page 8)
N:      DB 5              ; change N in RAM to try other values (max 5)
RES:    DB 0              ; 5! = 120 = 78h
