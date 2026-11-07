; 2. Keyboard echo to the LCD 8x1
; Type text in sheet IO and press SEND. Each char read from the
; keyboard queue (IN) is shown on the LCD (OUT). Stops after 8 chars.
        MOV CX, 8         ; CX = chars left to show
WAIT:   IN AX, [KBD]      ; AX = next char of the queue (0 = empty)
        CMP AX, 0
        JZ WAIT           ; queue empty -> keep waiting (polling)
        OUT [LCD], AX     ; echo the char on the LCD
        DEC CX
        JNZ WAIT
        HLT
