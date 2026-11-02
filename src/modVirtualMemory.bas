Attribute VB_Name = "modVirtualMemory"
'==============================================================================
'  404 Angels - CPU Simulator
'  modVirtualMemory : virtual memory with paging
'  OWNER            : Paola
'
'  Virtual memory : 512 bytes = 16 pages of 32 bytes  (VPN 0..15)
'  Physical RAM   : 256 bytes =  8 frames of 32 bytes (PFN 0..7)
'  Swap (disk)    : keeps a copy of the 16 pages
'
'  Virtual address (9 bits) = VPN (4 bits) + Offset (5 bits)
'     VPN    = vaddr \ 32
'     Offset = vaddr Mod 32
'     Physical address = PFN * 32 + Offset
'
'  Page 15 (keyboard queue and LCD) is always in frame 7.
'  When RAM is full, FIFO replaces the oldest page of frames 0..6.
'==============================================================================
Option Explicit

' --- Memories ------------------------------------------------------------------
Private RAM(0 To 255) As Long           ' physical RAM: 8 frames x 32 bytes
Private Swap(0 To 511) As Long          ' disk: 16 pages x 32 bytes

' --- Page table (one entry per virtual page) -----------------------------------
Private PageP(0 To 15) As Long          ' P: 1 = page in RAM, 0 = page on disk
Private PageFrame(0 To 15) As Long      ' PFN: frame that holds the page (-1 = none)
Private PageDirty(0 To 15) As Long      ' Dirty: 1 = page was modified in RAM
Private PageRef(0 To 15) As Long        ' Ref: 1 = page was used

' --- Frames and FIFO -----------------------------------------------------------
Private FrameOwner(0 To 7) As Long      ' page stored in each frame (-1 = free)
Private FramesUsed As Long              ' frames 0..6 already used
Private NextVictim As Long              ' FIFO: frame with the oldest page

' --- Counters and last access (shown on sheet Memory) --------------------------
Public VM_Accesses As Long
Public VM_PageFaults As Long
Private LastVaddr As Long
Private LastVPN As Long
Private LastOffset As Long
Private LastPA As Long
Private LastWasWrite As Boolean
Private LastWasFault As Boolean
Private VictimPage As Long              ' page taken out of RAM (-1 = none)
Private VictimSaved As Boolean          ' True if the victim was saved to swap
Private EventText As String
Private MemChanged As Boolean           ' True when RAM or swap must be repainted

'==============================================================================
'  1. RESET
'==============================================================================
Public Sub VM_Reset()
    Dim i As Long
    For i = 0 To 255
        RAM(i) = 0
    Next i
    For i = 0 To 511
        Swap(i) = 0
    Next i
    For i = 0 To 15
        PageP(i) = 0
        PageFrame(i) = -1
        PageDirty(i) = 0
        PageRef(i) = 0
    Next i
    For i = 0 To 7
        FrameOwner(i) = -1
    Next i
    FramesUsed = 0
    NextVictim = 0

    ' page 15 (keyboard and LCD) is always in frame 7
    PageP(15) = 1
    PageFrame(15) = 7
    FrameOwner(7) = 15

    VM_Accesses = 0
    VM_PageFaults = 0
    LastVPN = -1
    LastPA = -1
    LastWasWrite = False
    LastWasFault = False
    VictimPage = -1
    VictimSaved = False
    EventText = "Memory reset. Page 15 (I/O) is in frame 7."
    MemChanged = True
End Sub

'==============================================================================
'  2. ADDRESS TRANSLATION  (temporary: no paging yet)
'==============================================================================
Public Function VM_Translate(ByVal vaddr As Long, ByVal isWrite As Boolean) As Long
    VM_Translate = vaddr And &HFF
End Function

'==============================================================================
'  3. PROGRAM LOADER  (temporary: copies the program straight into RAM)
'==============================================================================
Public Sub VM_LoadImage(ByRef bytes() As Long, ByVal length As Long)
    Dim i As Long
    For i = 0 To length - 1
        Swap(i) = bytes(i) And &HFF
        If i < 256 Then RAM(i) = Swap(i)
    Next i
    MemChanged = True
End Sub

'==============================================================================
'  4. PHYSICAL RAM (used by the cache and the I/O module)
'==============================================================================
Public Function RAM_Read(ByVal pa As Long) As Long
    RAM_Read = RAM(pa And &HFF)
End Function

Public Sub RAM_Write(ByVal pa As Long, ByVal value As Long)
    RAM(pa And &HFF) = value And &HFF
    MemChanged = True
End Sub

'==============================================================================
'  5. SHEET "Memory"  (temporary)
'==============================================================================
Public Sub VM_SetupSheet()
    GetSheet(SH_MEMORY).Range("A1").Value = TEAM_NAME & " - Virtual memory and page table"
End Sub

Public Sub VM_Paint()
End Sub
