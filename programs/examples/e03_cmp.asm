; E03. CMP: compare without changing the register
; Input : AX = 7, then compare with 7, 9 and 3
; Output: AX stays 7. Flags after each CMP:
;         CMP AX, 7 -> ZF = 1 (equal)
;         CMP AX, 9 -> CF = 1 (AX is smaller)
;         CMP AX, 3 -> ZF = 0, CF = 0 (AX is bigger)
; Watch : sheet CPU, only the flags change
        MOV AX, 7
        CMP AX, 7
        CMP AX, 9
        CMP AX, 3
        HLT
