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
'  3. LCD 8x1
'     OUT [LCD], reg writes the char at the cursor and moves the cursor.
'     When the line is full, the text scrolls one position to the left.
'     Char 08h (backspace) erases the previous char.
'==============================================================================
Public Sub IO_LcdWrite(ByVal value As Long)
    Dim i As Long
    value = value And &HFF
    LastRead = -1
    LastWrite = -1

    ' backspace: go back one position and write a space
    If value = 8 Then
        If LcdCursor > 0 Then
            LcdCursor = LcdCursor - 1
            RAM_Write LCD_PA + LcdCursor, 32
            LastLcd = LcdCursor
            IoEvent = "OUT: backspace, cursor back to " & LcdCursor & "."
        Else
            IoEvent = "OUT: backspace ignored (cursor at 0)."
        End If
        Exit Sub
    End If

    ' line full: scroll left (each cell takes the value of the next one)
    If LcdCursor = 8 Then
        For i = 0 To 6
            RAM_Write LCD_PA + i, RAM_Read(LCD_PA + i + 1)
        Next i
        LcdCursor = 7
    End If

    RAM_Write LCD_PA + LcdCursor, value
    LastLcd = LcdCursor
    IoEvent = "OUT: '" & Chr(value) & "' written at " & Hex2(LCD_PA + LcdCursor) & "h (cursor " & LcdCursor & ")."
    LcdCursor = LcdCursor + 1
End Sub

'==============================================================================
'  4. SHEET "IO"
'     A4:I6    keyboard: type in B5 and press SEND
'     A8:I17   keyboard queue E0h-E7h with HEAD / TAIL / COUNT
'     A19:I23  LCD 8x1 E8h-EFh with cursor
'     A25:C25  bus signals En_KBD and Ld_LCD (read from modBus)
'     A27      last event
'
'  Colors: yellow = read now (IN) or written now (OUT), green = just typed,
'          red = queue full
'==============================================================================
Private Sub IoHeader(ByVal c As Range, ByVal txt As String)
    c.Value = txt
    c.Font.Bold = True
    c.Font.Color = RGB(255, 255, 255)
    c.Interior.Color = RGB(64, 64, 64)
    c.HorizontalAlignment = xlCenter
End Sub

Private Function CharText(ByVal v As Long) As String
    If v >= 32 And v <= 126 Then
        CharText = Chr(v)
    Else
        CharText = ""
    End If
End Function

Public Sub IO_SetupSheet()
    Dim ws As Worksheet, i As Long, shp As Shape
    Set ws = GetSheet(SH_IO)
    ws.Cells.Clear
    ws.Cells.Font.Name = "Consolas"
    ws.Cells.Font.Size = 10
    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B:I").ColumnWidth = 8

    ws.Range("A1").Value = TEAM_NAME & " - Keyboard and LCD 8x1"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Memory-mapped I/O in RAM frame 7: keyboard queue E0h-E7h, LCD buffer E8h-EFh"

    ' keyboard input
    IoHeader ws.Range("A4:I4"), "KEYBOARD"
    ws.Range("A5").Value = "Type text:"
    ws.Range("B5:D5").Interior.Color = RGB(255, 242, 204)
    ws.Range("B5").NumberFormat = "@"
    ws.Range("A6").Value = "Write in B5 and press SEND. Each char goes to the queue (max 8)."
    ws.Range("A6").Font.Size = 8

    ' SEND button (same style as the CPU buttons)
    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        If Left(ws.Shapes(i).Name, 5) = "btnIO" Then ws.Shapes(i).Delete
    Next i
    Set shp = ws.Shapes.AddShape(5, ws.Range("F5").Left, ws.Range("F5").Top - 2, ws.Range("F5:G5").Width, ws.Range("F5").Height + 4)
    shp.Name = "btnIOSend"
    shp.Fill.ForeColor.RGB = RGB(84, 130, 53)
    shp.Line.Visible = msoFalse
    shp.OnAction = "BtnKeyboardSend"
    shp.TextFrame.Characters.Text = "SEND"
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp.TextFrame.VerticalAlignment = xlVAlignCenter
    On Error GoTo 0

    ' keyboard queue
    IoHeader ws.Range("A8:I8"), "KEYBOARD QUEUE (circular FIFO, 8 bytes)"
    ws.Range("A9").Value = "Address"
    ws.Range("A10").Value = "Char"
    ws.Range("A11").Value = "Hex"
    ws.Range("A12").Value = "Pointer"
    For i = 0 To 7
        ws.Cells(9, 2 + i).Value = Hex2(KBD_PA + i) & "h"
        ws.Cells(9, 2 + i).Font.Bold = True
    Next i
    ws.Range("A14").Value = "HEAD (read)"
    ws.Range("A15").Value = "TAIL (write)"
    ws.Range("A16").Value = "COUNT"
    ws.Range("A17").Value = "Status"

    ' LCD
    IoHeader ws.Range("A19:I19"), "LCD 8x1 (video buffer)"
    ws.Range("A20").Value = "Address"
    ws.Range("A21").Value = "Screen"
    ws.Range("A22").Value = "Hex"
    ws.Range("A23").Value = "Cursor"
    For i = 0 To 7
        ws.Cells(20, 2 + i).Value = Hex2(LCD_PA + i) & "h"
        ws.Cells(20, 2 + i).Font.Bold = True
    Next i
    ws.Range("B21:I21").Font.Size = 20
    ws.Range("B21:I21").Font.Bold = True
    ws.Rows(21).RowHeight = 34

    ' bus signals
    ws.Range("A25").Value = "Bus signals"
    ws.Range("B25").Value = "En_KBD"
    ws.Range("C25").Value = "Ld_LCD"
    ws.Range("A27").Value = "Last event"

    ws.Range("B9:I23").HorizontalAlignment = xlCenter
    ws.Range("B10:I11").NumberFormat = "@"
    ws.Range("B21:I22").NumberFormat = "@"
    ws.Range("B14:B17").HorizontalAlignment = xlLeft
    ws.Range("B25:C25").HorizontalAlignment = xlCenter

    IO_Paint
End Sub

Public Sub IO_Paint()
    Dim ws As Worksheet, i As Long, v As Long, color As Long
    Set ws = GetSheet(SH_IO)

    ' keyboard queue
    For i = 0 To 7
        v = RAM_Read(KBD_PA + i)
        ws.Cells(10, 2 + i).Value = CharText(v)
        ws.Cells(11, 2 + i).Value = Hex2(v)
        ws.Cells(12, 2 + i).Value = ""
        color = RGB(242, 242, 242)
        If i = LastWrite Then color = RGB(198, 239, 206)   ' green: just typed
        If i = LastRead Then color = RGB(255, 235, 156)    ' yellow: read by IN
        ws.Range(ws.Cells(10, 2 + i), ws.Cells(11, 2 + i)).Interior.Color = color
    Next i
    If KbdHead = KbdTail Then
        ws.Cells(12, 2 + KbdHead).Value = "H/T"
    Else
        ws.Cells(12, 2 + KbdHead).Value = "HEAD"
        ws.Cells(12, 2 + KbdTail).Value = "TAIL"
    End If
    ws.Range("B14").Value = KbdHead
    ws.Range("B15").Value = KbdTail
    ws.Range("B16").Value = KbdCount
    If KbdCount = 0 Then
        ws.Range("B17").Value = "EMPTY"
        ws.Range("B17:D17").Interior.Color = RGB(217, 217, 217)
    ElseIf KbdCount = 8 Then
        ws.Range("B17").Value = "FULL"
        ws.Range("B17:D17").Interior.Color = RGB(255, 199, 206)
    Else
        ws.Range("B17").Value = KbdCount & " char(s) waiting"
        ws.Range("B17:D17").Interior.Color = RGB(198, 239, 206)
    End If

    ' LCD
    For i = 0 To 7
        v = RAM_Read(LCD_PA + i)
        ws.Cells(21, 2 + i).Value = CharText(v)
        ws.Cells(22, 2 + i).Value = Hex2(v)
        If i = LastLcd Then
            ws.Cells(21, 2 + i).Interior.Color = RGB(60, 90, 40)
        Else
            ws.Cells(21, 2 + i).Interior.Color = RGB(30, 50, 20)
        End If
        ws.Cells(21, 2 + i).Font.Color = RGB(150, 255, 120)
    Next i
    ws.Range("B23").Value = LcdCursor

    ' bus signals: green when the keyboard drives the bus or the LCD loads from it
    If InStr(ControlWord, "En_KBD") > 0 Then
        ws.Range("B25").Interior.Color = RGB(0, 176, 80)
    Else
        ws.Range("B25").Interior.Color = RGB(217, 217, 217)
    End If
    If InStr(ControlWord, "Ld_LCD") > 0 Then
        ws.Range("C25").Interior.Color = RGB(0, 112, 192)
    Else
        ws.Range("C25").Interior.Color = RGB(217, 217, 217)
    End If

    ws.Range("B27").Value = IoEvent
End Sub
