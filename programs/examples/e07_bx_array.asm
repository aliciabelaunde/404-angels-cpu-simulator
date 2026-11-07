; E07. [BX] addressing: add an array of 5 bytes and show the sum on the LCD
; Input : the array 3, 1, 4, 1, 5 at 100h-104h ([BX] = 100h + BX)
; Output: DX = 14 = 0Eh, stored at 105h, and the LCD shows SUM=14
; How   : the sum is split into two decimal digits by subtracting 10
;         until it is below 10 (CX counts the tens), then 30h turns
;         each digit into its ASCII code
; Watch : sheet CPU (BX moves through the array), sheet Cache (one line,
;         many L1D hits) and sheet IO (the LCD)
        MOV BX, 0         ; BX = index of the first element
        MOV CX, 5         ; 5 elements
        MOV DX, 0         ; DX = sum
NEXT:   LOAD AX, [BX]     ; AX = array[BX]
        ADD DX, AX
        INC BX
        DEC CX
        JNZ NEXT
        STORE [BX], DX    ; BX = 5 -> 105h
        MOV AX, 53h       ; 'S'
        OUT [LCD], AX
        MOV AX, 55h       ; 'U'
        OUT [LCD], AX
        MOV AX, 4Dh       ; 'M'
        OUT [LCD], AX
        MOV AX, 3Dh       ; '='
        OUT [LCD], AX
        MOV AX, DX        ; AX = sum
        MOV CX, 0         ; CX = tens
TENS:   CMP AX, 10
        JC DIGITS         ; AX < 10 -> AX is the units digit
        SUB AX, 10
        INC CX
        JMP TENS
DIGITS: ADD CX, 30h       ; tens  -> ASCII
        OUT [LCD], CX
        ADD AX, 30h       ; units -> ASCII
        OUT [LCD], AX
        HLT
        ORG 100h
        DB 3, 1, 4, 1, 5
