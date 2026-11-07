; E11. OUT: write a word on the LCD
; Input : the ASCII codes of "HI 404"
; Output: the LCD shows HI 404
; Watch : sheet IO (screen and cursor) and the Ld_LCD signal on sheet CPU
        MOV AX, 48h       ; 'H'
        OUT [LCD], AX
        MOV AX, 49h       ; 'I'
        OUT [LCD], AX
        MOV AX, 20h       ; ' '
        OUT [LCD], AX
        MOV AX, 34h       ; '4'
        OUT [LCD], AX
        MOV AX, 30h       ; '0'
        OUT [LCD], AX
        MOV AX, 34h       ; '4'
        OUT [LCD], AX
        HLT
