; E17. Swap with write back: a modified page is saved before it leaves
; Input : 77h is written in page 8 (Dirty = 1), then 6 other pages are read
; Output: page 8 is taken out by FIFO -> WRITE BACK (it goes to swap with 77h).
;         Reading 100h again brings it back from swap: BX = 77h,
;         and the LCD shows OK if the value survived the swap
; Watch : sheet Memory, Dirty = 1 on page 8, the WRITE BACK arrow in orange,
;         then the PAGE IN arrow in red when page 8 comes back
        MOV AX, 77h
        STORE [100h], AX  ; page 8 -> frame 1, Dirty = 1
        LOAD CX, [120h]   ; page 9
        LOAD CX, [140h]   ; page 10
        LOAD CX, [160h]   ; page 11
        LOAD CX, [180h]   ; page 12
        LOAD CX, [1A0h]   ; page 13 (RAM full)
        LOAD CX, [1C0h]   ; page 14 -> page 0 (code) leaves, it was not modified
        NOP               ; fetch: page 0 comes back -> page 8 leaves WITH write back
        LOAD BX, [100h]   ; page 8 comes back from swap
        CMP BX, 77h
        JNZ FAIL
        MOV AX, 4Fh       ; 'O'
        OUT [LCD], AX
        MOV AX, 4Bh       ; 'K'
        OUT [LCD], AX
        HLT
FAIL:   MOV AX, 58h       ; 'X'
        OUT [LCD], AX
        HLT
