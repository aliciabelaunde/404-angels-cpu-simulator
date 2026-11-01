Attribute VB_Name = "modIO"
'==============================================================================
'  404 Angels - CPU Simulator
'  modIO : keyboard queue (8B) and LCD 8x1 video buffer (8B), memory-mapped
'  OWNER : Paola
'==============================================================================
Option Explicit

Private KbdQueue As String
Private LcdText As String

Public Sub IO_Reset()
    KbdQueue = ""
    LcdText = ""
End Sub

Public Function IO_KeyboardRead() As Long
    ' STUB: pop the first char of the queue (0 if empty)
    If Len(KbdQueue) = 0 Then
        IO_KeyboardRead = 0
    Else
        IO_KeyboardRead = Asc(Left$(KbdQueue, 1))
        KbdQueue = Mid$(KbdQueue, 2)
    End If
End Function

Public Sub IO_LcdWrite(ByVal value As Long)
    ' STUB: append to the 8-char display
    LcdText = Right$(LcdText & Chr$(value And &HFF), IO_SIZE)
End Sub

Public Sub BtnKeyboardSend()
    ' STUB: read the typed text from sheet IO into the queue (max 8 chars)
End Sub

Public Sub IO_SetupSheet()
    GetSheet(SH_IO).Range("A1").Value = TEAM_NAME & " - Keyboard and LCD 8x1"
End Sub

Public Sub IO_Paint()
    ' STUB
End Sub
