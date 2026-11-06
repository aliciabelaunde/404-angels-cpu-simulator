; E18. Cache: hits on the same line vs misses on different lines
; Input : 8 reads of 100h-107h (one 8-byte line), then 4 reads 8 bytes apart
; Output: the first part has 1 L1D miss and 7 hits;
;         the second part misses on every read (different lines)
; Watch : sheet Cache, the hit rate of L1D and the path in green or red
        MOV BX, 0
        MOV CX, 8
SAME:   LOAD AX, [BX]     ; 100h + BX, same line
        INC BX
        DEC CX
        JNZ SAME
        MOV BX, 8
        MOV CX, 4
FAR:    LOAD AX, [BX]     ; 108h, 110h, 118h, 120h: a new line each time
        ADD BX, 8
        DEC CX
        JNZ FAR
        HLT
