; E14. LCD scroll and backspace
; Input : the digits 0 to 9 (10 chars for an 8-char screen), then one backspace
; Output: the LCD scrolls and ends with "2345678 "
; Watch : sheet IO, after the 8th char the text moves one place to the left
        MOV AX, 30h       ; '0'
        MOV CX, 10
NEXT:   OUT [LCD], AX
        INC AX
        DEC CX
        JNZ NEXT
        MOV AX, 8         ; backspace
        OUT [LCD], AX
        HLT
