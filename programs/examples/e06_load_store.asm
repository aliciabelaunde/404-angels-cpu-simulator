; E06. LOAD and STORE with a direct address
; Input : A = 25 at 100h, B = 17 at 101h
; Output: SUM = 42 = 2Ah at 102h
; Watch : sheet Memory (page 8 is loaded with a page fault)
;         and sheet Cache (L1D misses, then hits on the same line)
        LOAD AX, [A]
        LOAD BX, [B]
        ADD AX, BX
        STORE [SUM], AX
        HLT
        ORG 100h
A:      DB 25
B:      DB 17
SUM:    DB 0
