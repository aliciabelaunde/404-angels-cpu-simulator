Attribute VB_Name = "modVirtualMemory"
'==============================================================================
'  404 Angels - CPU Simulator
'  modVirtualMemory : physical RAM (8 frames x 32B), swap (16 pages x 32B),
'                     page table (P, frame, Dirty, Ref), page fault, FIFO
'  OWNER            : Paola
'
'  On page replacement call Cache_InvalidateFrame(frame) BEFORE overwriting it.
'==============================================================================
Option Explicit

Private RAM(0 To 255) As Long
Private Swap(0 To 511) As Long

Public Sub VM_Reset()
    Dim i As Long
    For i = 0 To 255: RAM(i) = 0: Next i
    For i = 0 To 511: Swap(i) = 0: Next i
End Sub

Public Function VM_Translate(ByVal vaddr As Long, ByVal isWrite As Boolean) As Long
    ' STUB: identity mapping on the first 256 bytes (no paging yet)
    VM_Translate = vaddr And &HFF
End Function

Public Sub VM_LoadImage(ByRef bytes() As Long, ByVal length As Long)
    ' STUB: copy the program to swap and to RAM (no paging yet)
    Dim i As Long
    For i = 0 To length - 1
        Swap(i) = bytes(i) And &HFF
        If i < 256 Then RAM(i) = Swap(i)
    Next i
End Sub

Public Function RAM_Read(ByVal pa As Long) As Long
    RAM_Read = RAM(pa And &HFF)
End Function

Public Sub RAM_Write(ByVal pa As Long, ByVal value As Long)
    RAM(pa And &HFF) = value And &HFF
End Sub

Public Sub VM_SetupSheet()
    GetSheet(SH_MEMORY).Range("A1").Value = TEAM_NAME & " - Virtual memory and page table"
End Sub

Public Sub VM_Paint()
    ' STUB
End Sub
