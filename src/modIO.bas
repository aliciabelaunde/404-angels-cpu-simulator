Attribute VB_Name = "modIO"
'==============================================================================
'  404 Angels - CPU Simulator
'  modIO : keyboard queue (8 B) and LCD 8x1 video buffer (8 B), memory-mapped
'  OWNER : Paola
'
'  Both live in RAM frame 7 (page 15), so the CPU, the bus and the sheets
'  always see the same bytes:
'     E0h-E7h  keyboard queue  (virtual 1E0h-1E7h)
'     E8h-EFh  LCD video buffer (virtual 1E8h-1EFh)
'
'  Keyboard queue = circular FIFO, as in the class lab:
'     insert  (keyboard): if COUNT < 8 -> RAM[E0h + TAIL] = char
'                                         TAIL = (TAIL + 1) mod 8, COUNT + 1
'     extract (IN)      : if COUNT > 0 -> char = RAM[E0h + HEAD]
'                                         HEAD = (HEAD + 1) mod 8, COUNT - 1
'     COUNT = 8 -> queue full, the new char is discarded (overflow)
'==============================================================================
Option Explicit

Private Const KBD_PA As Long = &HE0     ' physical address of the keyboard queue
Private Const LCD_PA As Long = &HE8     ' physical address of the LCD buffer

' --- Keyboard queue pointers (the chars themselves are in RAM) ----------------
Private KbdHead As Long                 ' next char to read
Private KbdTail As Long                 ' next free position
Private KbdCount As Long                ' chars waiting in the queue

' --- LCD ------------------------------------------------------------------------
Private LcdCursor As Long               ' next LCD position (0..8)

' --- Last event (shown on sheet IO) -------------------------------------------
Private LastRead As Long                ' queue position read by IN (-1 = none)
Private LastWrite As Long               ' queue position written by the keyboard (-1 = none)
Private LastLcd As Long                 ' LCD position written by OUT (-1 = none)
Private IoEvent As String

'==============================================================================
'  1. RESET
'==============================================================================
Public Sub IO_Reset()
    Dim i As Long
    For i = 0 To 7
        RAM_Write KBD_PA + i, 0
        RAM_Write LCD_PA + i, 32        ' 32 = space: the LCD starts blank
    Next i
    KbdHead = 0
    KbdTail = 0
    KbdCount = 0
    LcdCursor = 0
    LastRead = -1
    LastWrite = -1
    LastLcd = -1
    IoEvent = "Keyboard queue empty, LCD clear."
End Sub

'==============================================================================
'  2. KEYBOARD QUEUE
'==============================================================================
' The keyboard puts one char in the queue (producer)
Private Sub KeyboardPush(ByVal ch As Long)
    If KbdCount = 8 Then
        IoEvent = "Queue FULL (COUNT = 8): '" & Chr(ch) & "' was discarded."
        Exit Sub
    End If
    RAM_Write KBD_PA + KbdTail, ch
    LastWrite = KbdTail
    IoEvent = "Keyboard: '" & Chr(ch) & "' stored at " & Hex2(KBD_PA + KbdTail) & "h (TAIL = " & KbdTail & ")."
    KbdTail = (KbdTail + 1) Mod 8
    KbdCount = KbdCount + 1
End Sub

' IN reg, [KBD]: the CPU takes one char from the queue (consumer). 0 = empty
Public Function IO_KeyboardRead() As Long
    Dim ch As Long
    LastWrite = -1
    If KbdCount = 0 Then
        LastRead = -1
        IoEvent = "IN: queue empty (COUNT = 0), returns 0."
        IO_KeyboardRead = 0
        Exit Function
    End If
    ch = RAM_Read(KBD_PA + KbdHead)
    RAM_Write KBD_PA + KbdHead, 0       ' the position is free again
    LastRead = KbdHead
    IoEvent = "IN: '" & Chr(ch) & "' read from " & Hex2(KBD_PA + KbdHead) & "h (HEAD = " & KbdHead & ")."
    KbdHead = (KbdHead + 1) Mod 8
    KbdCount = KbdCount - 1
    IO_KeyboardRead = ch
End Function

' Button SEND: every char typed in cell B5 of sheet IO goes to the queue
Public Sub BtnKeyboardSend()
    Dim ws As Worksheet, txt As String, i As Long
    Set ws = GetSheet(SH_IO)
    txt = CStr(ws.Range("B5").Value)
    LastRead = -1
    For i = 1 To Len(txt)
        KeyboardPush Asc(Mid(txt, i, 1))
    Next i
    ws.Range("B5").Value = ""
    IO_Paint
End Sub

'==============================================================================
'  3. LCD 8x1  (temporary: writes at the cursor, no scroll yet)
'==============================================================================
Public Sub IO_LcdWrite(ByVal value As Long)
    If LcdCursor < 8 Then
        RAM_Write LCD_PA + LcdCursor, value
        LastLcd = LcdCursor
        LcdCursor = LcdCursor + 1
    End If
End Sub

'==============================================================================
'  4. SHEET "IO"  (temporary)
'==============================================================================
Public Sub IO_SetupSheet()
    GetSheet(SH_IO).Range("A1").Value = TEAM_NAME & " - Keyboard and LCD 8x1"
End Sub

Public Sub IO_Paint()
End Sub
