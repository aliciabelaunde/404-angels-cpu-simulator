; E05. INC and DEC: they change ZF and SF but keep CF
; Input : AX = 0FFh, CF = 1 (set by the first ADD)
; Output: INC AX -> AX = 0 and ZF = 1, CF still 1
;         DEC AX -> AX = 0FFh and SF = 1, CF still 1
; Watch : sheet CPU, CF does not move during INC and DEC
        MOV AX, 0FFh
        ADD AX, 1         ; AX = 0, CF = 1
        MOV AX, 0FFh
        INC AX            ; AX = 0, ZF = 1, CF = 1
        DEC AX            ; AX = 0FFh, SF = 1, CF = 1
        HLT
