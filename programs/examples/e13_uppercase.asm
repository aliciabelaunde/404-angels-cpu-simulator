; E13. Keyboard -> process -> LCD: lower case to upper case
; Input : type "hola" (lower case) in sheet IO and press SEND
; Output: the LCD shows HOLA
; How   : a lower-case letter is 20h more than the upper-case one,
;         so SUB 20h converts it. Chars below 'a' (61h) are shown as they are.
; Watch : sheet IO (queue going down, LCD filling up)
        MOV CX, 4         ; 4 chars
WAIT:   IN AX, [KBD]
        CMP AX, 0
        JZ WAIT           ; queue empty -> wait
        CMP AX, 61h
        JC SHOW           ; below 'a' -> no change
        SUB AX, 20h       ; 'a' -> 'A'
SHOW:   OUT [LCD], AX
        DEC CX
        JNZ WAIT
        HLT
