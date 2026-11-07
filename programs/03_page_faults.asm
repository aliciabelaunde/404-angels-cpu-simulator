; 3. Scattered accesses: page faults, FIFO and cache misses
; Visits the first byte of the 7 data pages (8 to 14), one page apart.
; Code page 0 + 7 data pages = 8 pages, but only 7 frames are free,
; so FIFO must replace pages and write back the dirty ones.
        MOV CX, 7         ; 7 data pages
        MOV BX, 0         ; [BX] = 100h + BX -> page 8
PASS1:  LOAD AX, [BX]     ; read  -> page fault the first time
        INC AX
        STORE [BX], AX    ; write -> Dirty = 1
        ADD BX, 20h       ; next page (+32 bytes)
        DEC CX
        JNZ PASS1
        MOV CX, 7         ; second pass: add the 7 new values
        MOV BX, 0
        MOV DX, 0
PASS2:  LOAD AX, [BX]     ; replaced pages come back from swap
        ADD DX, AX
        ADD BX, 20h
        DEC CX
        JNZ PASS2
        STORE [SUM], DX   ; SUM = 2+3+4+5+6+7+8 = 35 = 23h
        HLT
        ORG 100h
        DB 1              ; page 8
        ORG 120h
        DB 2              ; page 9
        ORG 140h
        DB 3              ; page 10
        ORG 160h
        DB 4              ; page 11
        ORG 180h
        DB 5              ; page 12
        ORG 1A0h
        DB 6              ; page 13
        ORG 1C0h
        DB 7              ; page 14
        ORG 1DFh
SUM:    DB 0              ; last byte of page 14
