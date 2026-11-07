; E12. IN: read one char from the keyboard queue
; Input : type one letter in sheet IO and press SEND (before RUN)
; Output: AX = ASCII code of the letter, and the letter on the LCD
; Watch : sheet IO, HEAD moves and COUNT goes down by 1; En_KBD on sheet CPU
WAIT:   IN AX, [KBD]      ; 0 if the queue is empty
        CMP AX, 0
        JZ WAIT
        OUT [LCD], AX
        HLT
