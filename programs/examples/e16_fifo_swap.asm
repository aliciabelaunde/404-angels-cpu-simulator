; E16. FIFO replacement: RAM is full, the oldest page leaves
; Input : reads from pages 8 to 14 (7 data pages) + code page 0 = 8 pages,
;         but only 7 frames can be used
; Output: page 14 needs a frame, so FIFO takes out page 0 (the code, the oldest).
;         The next instruction fetch brings page 0 back and takes out page 8.
;         9 page faults in total, no write back (no page was modified)
; Watch : sheet Memory, the FIFO victim in orange and "Next FIFO victim"
        LOAD AX, [100h]   ; page 8  -> frame 1
        LOAD AX, [120h]   ; page 9  -> frame 2
        LOAD AX, [140h]   ; page 10 -> frame 3
        LOAD AX, [160h]   ; page 11 -> frame 4
        LOAD AX, [180h]   ; page 12 -> frame 5
        LOAD AX, [1A0h]   ; page 13 -> frame 6 (RAM full)
        LOAD AX, [1C0h]   ; page 14 -> FIFO takes out page 0
        HLT               ; fetch -> page 0 comes back, page 8 leaves
