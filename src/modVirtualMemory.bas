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
'  2. ADDRESS TRANSLATION AND PAGE FAULT
'==============================================================================
Public Function VM_Translate(ByVal vaddr As Long, ByVal isWrite As Boolean) As Long
    Dim vpn As Long, offset As Long

    vaddr = vaddr And &H1FF             ' keep 9 bits (000h..1FFh)
    vpn = vaddr \ 32
    offset = vaddr Mod 32

    VM_Accesses = VM_Accesses + 1
    LastWasFault = False
    VictimPage = -1
    VictimSaved = False

    ' P = 0: the page is on disk -> page fault
    If PageP(vpn) = 0 Then
        PageFault vpn
    End If

    PageRef(vpn) = 1
    If isWrite Then
        PageDirty(vpn) = 1
    End If

    LastVaddr = vaddr
    LastVPN = vpn
    LastOffset = offset
    LastWasWrite = isWrite
    LastPA = PageFrame(vpn) * 32 + offset
    If LastWasFault = False Then
        EventText = "Page " & vpn & " is in frame " & PageFrame(vpn) & " (P = 1)."
    End If

    VM_Translate = LastPA
End Function

' Page fault: same 4 steps as in class
'   1. the CPU tries to use a page with P = 0
'   2. look for the page in swap
'   3. take a free frame (or the oldest one with FIFO) and copy the 32 bytes
'   4. update the table (P = 1) and continue the instruction
Private Sub PageFault(ByVal vpn As Long)
    Dim frame As Long, i As Long

    VM_PageFaults = VM_PageFaults + 1
    LastWasFault = True

    ' choose a frame
    If FramesUsed < 7 Then
        frame = FramesUsed              ' there is still a free frame
        FramesUsed = FramesUsed + 1
    Else
        frame = NextVictim              ' FIFO: the oldest page leaves
        NextVictim = (NextVictim + 1) Mod 7
        VictimPage = FrameOwner(frame)

        ' if the victim was modified, save it to swap first
        If PageDirty(VictimPage) = 1 Then
            For i = 0 To 31
                Swap(VictimPage * 32 + i) = RAM(frame * 32 + i)
            Next i
            VictimSaved = True
        End If

        PageP(VictimPage) = 0
        PageFrame(VictimPage) = -1
        PageDirty(VictimPage) = 0
        PageRef(VictimPage) = 0
    End If

    ' the cache must forget what it had from this frame
    Cache_InvalidateFrame frame

    ' copy the page from swap to RAM
    For i = 0 To 31
        RAM(frame * 32 + i) = Swap(vpn * 32 + i)
    Next i

    ' update the page table
    PageP(vpn) = 1
    PageFrame(vpn) = frame
    PageDirty(vpn) = 0
    FrameOwner(frame) = vpn

    EventText = "PAGE FAULT: page " & vpn & " copied from swap to frame " & frame & "."
    If VictimPage >= 0 Then
        EventText = EventText & " FIFO took out page " & VictimPage
        If VictimSaved Then
            EventText = EventText & " (Dirty = 1, saved to swap)."
        Else
            EventText = EventText & " (Dirty = 0, not saved)."
        End If
    End If
    MemChanged = True
End Sub

'==============================================================================
'  3. PROGRAM LOADER
'  The program goes to swap. Pages 0..14 start with P = 0, so the first
'  access to each page is a page fault.
'==============================================================================
Public Sub VM_LoadImage(ByRef bytes() As Long, ByVal length As Long)
    Dim i As Long

    If length > 512 Then length = 512
    For i = 0 To length - 1
        Swap(i) = bytes(i) And &HFF
    Next i

    ' pages 0..14 are on disk
    For i = 0 To 14
        PageP(i) = 0
        PageFrame(i) = -1
        PageDirty(i) = 0
        PageRef(i) = 0
    Next i

    ' frames 0..6 are empty
    For i = 0 To 6 * 32 + 31
        RAM(i) = 0
    Next i
    For i = 0 To 6
        FrameOwner(i) = -1
    Next i
    FramesUsed = 0
    NextVictim = 0

    ' page 15 is always in frame 7
    For i = 0 To 31
        RAM(7 * 32 + i) = Swap(15 * 32 + i)
    Next i

    LastVPN = -1
    LastPA = -1
    LastWasFault = False
    VictimPage = -1
    EventText = "Program loaded in swap (" & length & " bytes). Pages 0-14 have P = 0."
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
