; E15. Page faults: demand paging (no replacement yet)
; Input : one byte read from pages 8, 9 and 10
; Output: 4 page faults (code page 0 + 3 data pages); AX = 1, BX = 2, CX = 3
; Watch : sheet Memory. Each page goes from gray (P = 0) to green (P = 1)
;         and the PAGE IN arrow turns red at every fault
        LOAD AX, [100h]   ; page 8  -> page fault
        LOAD BX, [120h]   ; page 9  -> page fault
        LOAD CX, [140h]   ; page 10 -> page fault
        LOAD DX, [101h]   ; page 8 again -> no fault (already in RAM)
        HLT
        ORG 100h
        DB 1, 9
        ORG 120h
        DB 2
        ORG 140h
        DB 3
