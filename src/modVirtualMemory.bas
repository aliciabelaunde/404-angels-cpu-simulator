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
'  5. SHEET "Memory"
'     A4:B9     address translation       D4:F7   status (PAGE FAULT / PAGE HIT)
'     A11:B15   counters
'     A17:F33   page table (VPN, area, P, PFN, Dirty, Ref)
'     H17:AO25  physical RAM, one row per frame (32 bytes)
'     H29:AO45  swap, one row per page (32 bytes)
'
'  Colors: green = P 1    gray = P 0    blue = page 15 (I/O)
'          yellow = used now    red = page fault    orange = FIFO victim
'==============================================================================
Private Sub HeaderCell(ByVal c As Range, ByVal txt As String)
    c.Value = txt
    c.Font.Bold = True
    c.Font.Color = RGB(255, 255, 255)
    c.Interior.Color = RGB(64, 64, 64)
    c.HorizontalAlignment = xlCenter
End Sub

Private Function AreaName(ByVal vpn As Long) As String
    If vpn <= 7 Then
        AreaName = "Code"
    ElseIf vpn <= 14 Then
        AreaName = "Data"
    Else
        AreaName = "I/O"
    End If
End Function

Public Sub VM_SetupSheet()
    Dim ws As Worksheet, i As Long
    Set ws = GetSheet(SH_MEMORY)
    ws.Cells.Clear
    ws.Cells.Font.Name = "Consolas"
    ws.Cells.Font.Size = 10

    ws.Range("A1").Value = TEAM_NAME & " - Virtual memory and page table"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Virtual: 16 pages x 32 B = 512 B  |  RAM: 8 frames x 32 B = 256 B  |  Replacement: FIFO (frames 0-6)"

    ' address translation
    HeaderCell ws.Range("A4"), "ADDRESS TRANSLATION"
    HeaderCell ws.Range("B4"), "Value"
    ws.Range("A5").Value = "Virtual address"
    ws.Range("A6").Value = "VPN = vaddr \ 32"
    ws.Range("A7").Value = "Offset = vaddr mod 32"
    ws.Range("A8").Value = "PFN (page table)"
    ws.Range("A9").Value = "Physical = PFN * 32 + Offset"
    HeaderCell ws.Range("D4"), "STATUS"
    ws.Range("D5").Font.Size = 14
    ws.Range("D5").Font.Bold = True

    ' counters
    HeaderCell ws.Range("A11"), "COUNTERS"
    HeaderCell ws.Range("B11"), "Value"
    ws.Range("A12").Value = "Memory accesses"
    ws.Range("A13").Value = "Page faults"
    ws.Range("A14").Value = "Page fault rate"
    ws.Range("A15").Value = "Next FIFO victim"

    ' page table
    ws.Range("A16").Value = "PAGE TABLE"
    ws.Range("A16").Font.Bold = True
    HeaderCell ws.Range("A17"), "VPN"
    HeaderCell ws.Range("B17"), "Area"
    HeaderCell ws.Range("C17"), "P"
    HeaderCell ws.Range("D17"), "PFN"
    HeaderCell ws.Range("E17"), "Dirty"
    HeaderCell ws.Range("F17"), "Ref"

    ' physical RAM and swap
    ws.Range("H16").Value = "PHYSICAL RAM (8 frames x 32 bytes)"
    ws.Range("H16").Font.Bold = True
    HeaderCell ws.Range("H17"), "PFN"
    HeaderCell ws.Range("I17"), "Page"
    ws.Range("H28").Value = "SWAP / DISK (16 pages x 32 bytes)"
    ws.Range("H28").Font.Bold = True
    HeaderCell ws.Range("H29"), "VPN"
    HeaderCell ws.Range("I29"), "Area"
    For i = 0 To 31
        HeaderCell ws.Cells(17, 10 + i), Hex2(i)
        HeaderCell ws.Cells(29, 10 + i), Hex2(i)
    Next i

    ws.Columns("A").ColumnWidth = 30
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C:F").ColumnWidth = 7
    ws.Columns("G").ColumnWidth = 2
    ws.Columns("H:I").ColumnWidth = 6
    ws.Range("J1:AO1").EntireColumn.ColumnWidth = 3.3
    ws.Range("B5:B15").NumberFormat = "@"
    ws.Range("H18:AO25").NumberFormat = "@"
    ws.Range("H30:AO45").NumberFormat = "@"
    ws.Range("B5:F15").HorizontalAlignment = xlCenter
    ws.Range("A18:F33").HorizontalAlignment = xlCenter
    ws.Range("H18:AO45").HorizontalAlignment = xlCenter
    ws.Range("C9:F9").HorizontalAlignment = xlLeft
    ws.Range("D5:F7").HorizontalAlignment = xlLeft
    ws.Range("H16").HorizontalAlignment = xlLeft
    ws.Range("H28").HorizontalAlignment = xlLeft

    DrawVmDiagram ws
    MemChanged = True
    VM_Paint
End Sub

Public Sub VM_Paint()
    Dim ws As Worksheet
    Set ws = GetSheet(SH_MEMORY)
    PaintTranslation ws
    PaintPageTable ws
    If MemChanged Then
        PaintBytes ws                   ' only when RAM or swap changed (faster)
        MemChanged = False
    End If
    PaintFramesAndSwap ws
    PaintVmDiagram ws
End Sub

Private Sub PaintTranslation(ByVal ws As Worksheet)
    If LastVPN >= 0 Then
        ws.Range("B5").Value = Hex3(LastVaddr) & "h"
        ws.Range("B6").Value = CStr(LastVPN)
        ws.Range("B7").Value = CStr(LastOffset)
        ws.Range("B8").Value = CStr(PageFrame(LastVPN))
        ws.Range("B9").Value = Hex2(LastPA) & "h"
        ws.Range("C9").Value = PageFrame(LastVPN) & " * 32 + " & LastOffset & " = " & LastPA
    Else
        ws.Range("B5:B9").Value = "-"
        ws.Range("C9").Value = ""
    End If

    ws.Range("B12").Value = CStr(VM_Accesses)
    ws.Range("B13").Value = CStr(VM_PageFaults)
    If VM_Accesses > 0 Then
        ws.Range("B14").Value = Round(VM_PageFaults * 100 / VM_Accesses, 1) & " %"
    Else
        ws.Range("B14").Value = "0 %"
    End If
    If FramesUsed < 7 Then
        ws.Range("B15").Value = "free frame " & FramesUsed
    Else
        ws.Range("B15").Value = "frame " & NextVictim & " (page " & FrameOwner(NextVictim) & ")"
    End If

    ' status box
    If LastWasFault Then
        ws.Range("D5").Value = "PAGE FAULT"
        ws.Range("D5:F5").Interior.Color = RGB(192, 0, 0)
        ws.Range("D5:F5").Font.Color = RGB(255, 255, 255)
    ElseIf LastVPN >= 0 Then
        ws.Range("D5").Value = "PAGE HIT"
        ws.Range("D5:F5").Interior.Color = RGB(0, 128, 0)
        ws.Range("D5:F5").Font.Color = RGB(255, 255, 255)
    Else
        ws.Range("D5").Value = "-"
        ws.Range("D5:F5").Interior.Color = RGB(217, 217, 217)
        ws.Range("D5:F5").Font.Color = RGB(0, 0, 0)
    End If
    If LastVPN < 0 Then
        ws.Range("D6").Value = ""
    ElseIf LastWasWrite Then
        ws.Range("D6").Value = "write"
    Else
        ws.Range("D6").Value = "read"
    End If
    ws.Range("D7").Value = EventText
End Sub

Private Sub PaintPageTable(ByVal ws As Worksheet)
    Dim vpn As Long, r As Long, color As Long
    For vpn = 0 To 15
        r = 18 + vpn
        ws.Cells(r, 1).Value = vpn
        ws.Cells(r, 2).Value = AreaName(vpn)
        ws.Cells(r, 3).Value = PageP(vpn)
        If PageFrame(vpn) >= 0 Then
            ws.Cells(r, 4).Value = PageFrame(vpn)
        Else
            ws.Cells(r, 4).Value = "-"
        End If
        ws.Cells(r, 5).Value = PageDirty(vpn)
        ws.Cells(r, 6).Value = PageRef(vpn)

        If vpn = 15 Then
            color = RGB(221, 235, 247)  ' blue: pinned I/O page
        ElseIf PageP(vpn) = 1 Then
            color = RGB(198, 239, 206)  ' green: in RAM
        Else
            color = RGB(217, 217, 217)  ' gray: on disk
        End If
        If vpn = VictimPage Then color = RGB(248, 203, 173)   ' orange: FIFO victim
        If vpn = LastVPN Then
            If LastWasFault Then
                color = RGB(255, 199, 206)  ' red: page fault
            Else
                color = RGB(255, 235, 156)  ' yellow: used now
            End If
        End If
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Interior.Color = color
    Next vpn
End Sub

Private Sub PaintBytes(ByVal ws As Worksheet)
    Dim f As Long, p As Long, i As Long
    For f = 0 To 7
        ws.Cells(18 + f, 8).Value = f
        If FrameOwner(f) >= 0 Then
            ws.Cells(18 + f, 9).Value = FrameOwner(f)
        Else
            ws.Cells(18 + f, 9).Value = "free"
        End If
        For i = 0 To 31
            ws.Cells(18 + f, 10 + i).Value = Hex2(RAM(f * 32 + i))
        Next i
    Next f
    For p = 0 To 15
        ws.Cells(30 + p, 8).Value = p
        ws.Cells(30 + p, 9).Value = AreaName(p)
        For i = 0 To 31
            ws.Cells(30 + p, 10 + i).Value = Hex2(Swap(p * 32 + i))
        Next i
    Next p
End Sub

Private Sub PaintFramesAndSwap(ByVal ws As Worksheet)
    Dim f As Long, p As Long, color As Long
    ' frames
    For f = 0 To 7
        If f = 7 Then
            color = RGB(221, 235, 247)
        ElseIf FrameOwner(f) >= 0 Then
            color = RGB(198, 239, 206)
        Else
            color = RGB(217, 217, 217)
        End If
        If LastWasFault And LastVPN >= 0 Then
            If f = PageFrame(LastVPN) Then color = RGB(255, 199, 206)
        End If
        ws.Range(ws.Cells(18 + f, 8), ws.Cells(18 + f, 41)).Interior.Color = color
    Next f
    ' the byte used now
    If LastPA >= 0 Then
        ws.Cells(18 + LastPA \ 32, 10 + LastPA Mod 32).Interior.Color = RGB(255, 235, 156)
    End If
    ' swap
    For p = 0 To 15
        If LastWasFault And p = LastVPN Then
            color = RGB(255, 199, 206)  ' red: read from disk now
        ElseIf VictimSaved And p = VictimPage Then
            color = RGB(248, 203, 173)  ' orange: written back now
        Else
            color = RGB(242, 242, 242)
        End If
        ws.Range(ws.Cells(30 + p, 8), ws.Cells(30 + p, 41)).Interior.Color = color
    Next p
End Sub

'==============================================================================
'  6. PAGING DIAGRAM (shapes on sheet Memory, under the page table)
'     CPU -> MMU (page table) -> RAM frame
'     PAGE IN arrow    = red when a page comes from swap (page fault)
'     WRITE BACK arrow = orange when a dirty page goes back to swap
'==============================================================================
Private Const VMDG As String = "vmdg_"       ' prefix of every shape of this diagram

Private Function VmShape(ByVal ws As Worksheet, ByVal nm As String, ByVal kind As Long, _
                         ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                         ByVal txt As String) As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(kind, ws.Range("A36").Left + x, ws.Range("A36").Top + y, w, h)
    shp.Name = VMDG & nm
    shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Line.ForeColor.RGB = RGB(31, 42, 74)
    shp.Line.Weight = 1
    If txt <> "" Then
        shp.TextFrame.Characters.Text = txt
        shp.TextFrame.Characters.Font.Size = 8
        shp.TextFrame.Characters.Font.Bold = True
        shp.TextFrame.Characters.Font.Color = RGB(31, 42, 74)
        shp.TextFrame.HorizontalAlignment = xlHAlignCenter
        shp.TextFrame.VerticalAlignment = xlVAlignCenter
    End If
    Set VmShape = shp
End Function

Private Sub VmText(ByVal ws As Worksheet, ByVal nm As String, ByVal txt As String)
    ws.Shapes(VMDG & nm).TextFrame.Characters.Text = txt
End Sub

Private Sub VmFill(ByVal ws As Worksheet, ByVal nm As String, ByVal color As Long)
    ws.Shapes(VMDG & nm).Fill.ForeColor.RGB = color
End Sub

Private Sub DeleteVmDiagram(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left(ws.Shapes(i).Name, Len(VMDG)) = VMDG Then ws.Shapes(i).Delete
    Next i
End Sub

Private Sub DrawVmDiagram(ByVal ws As Worksheet)
    Dim shp As Shape, i As Long
    On Error Resume Next                     ' if a shape fails, the sheet still works
    DeleteVmDiagram ws

    ' frame and title
    Set shp = VmShape(ws, "frame", 1, 0, 0, 430, 300, "")
    shp.Fill.ForeColor.RGB = RGB(244, 247, 251)
    shp.Line.ForeColor.RGB = RGB(200, 208, 222)
    Set shp = VmShape(ws, "title", 1, 0, 0, 430, 18, "PAGING DIAGRAM - live view of every memory access")
    shp.Fill.ForeColor.RGB = RGB(31, 42, 74)
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)

    ' CPU -> MMU
    Call VmShape(ws, "cpu", 5, 10, 28, 120, 36, "CPU" & vbLf & "vaddr -")
    Call VmShape(ws, "arrow_mmu", 33, 135, 38, 30, 16, "")
    Call VmShape(ws, "mmu", 5, 170, 28, 250, 36, "MMU / PAGE TABLE" & vbLf & "-")

    ' RAM: 8 frames
    Set shp = VmShape(ws, "ram_label", 1, 10, 74, 150, 12, "RAM (8 frames)")
    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse
    For i = 0 To 7
        Call VmShape(ws, "frame" & i, 5, 10, 90 + i * 25, 150, 21, "Frame " & i & ": free")
    Next i

    ' SWAP: 16 pages in 2 columns
    Set shp = VmShape(ws, "swap_label", 1, 290, 74, 130, 12, "SWAP / DISK (16 pages)")
    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse
    For i = 0 To 15
        Call VmShape(ws, "page" & i, 1, 290 + (i \ 8) * 68, 90 + (i Mod 8) * 25, 62, 21, "Page " & i)
    Next i

    ' arrows between swap and RAM, and the status box
    Call VmShape(ws, "pagein", 34, 172, 105, 108, 36, "PAGE IN")
    Call VmShape(ws, "writeback", 33, 172, 160, 108, 36, "WRITE BACK")
    Call VmShape(ws, "status", 5, 180, 225, 92, 40, "-")
End Sub

Private Sub PaintVmDiagram(ByVal ws As Worksheet)
    Dim i As Long, color As Long
    On Error GoTo NoDiagram                  ' the diagram does not exist yet

    ' CPU and MMU boxes
    If LastVPN >= 0 Then
        VmText ws, "cpu", "CPU" & vbLf & "vaddr " & Hex3(LastVaddr) & "h"
        VmText ws, "mmu", "MMU / PAGE TABLE" & vbLf & "VPN " & LastVPN & " | Offset " & LastOffset & "  ->  PFN " & PageFrame(LastVPN)
    Else
        VmText ws, "cpu", "CPU" & vbLf & "vaddr -"
        VmText ws, "mmu", "MMU / PAGE TABLE" & vbLf & "-"
    End If

    ' frames
    For i = 0 To 7
        If FrameOwner(i) >= 0 Then
            VmText ws, "frame" & i, "Frame " & i & ": page " & FrameOwner(i)
        Else
            VmText ws, "frame" & i, "Frame " & i & ": free"
        End If
        If i = 7 Then
            color = RGB(221, 235, 247)
        ElseIf FrameOwner(i) >= 0 Then
            color = RGB(198, 239, 206)
        Else
            color = RGB(217, 217, 217)
        End If
        If LastVPN >= 0 Then
            If i = PageFrame(LastVPN) Then
                If LastWasFault Then
                    color = RGB(255, 199, 206)
                Else
                    color = RGB(255, 235, 156)
                End If
            End If
        End If
        VmFill ws, "frame" & i, color
    Next i

    ' swap pages
    For i = 0 To 15
        If LastWasFault And i = LastVPN Then
            color = RGB(255, 199, 206)
        ElseIf VictimSaved And i = VictimPage Then
            color = RGB(248, 203, 173)
        ElseIf PageP(i) = 1 Then
            color = RGB(198, 239, 206)
        Else
            color = RGB(255, 255, 255)
        End If
        VmFill ws, "page" & i, color
    Next i

    ' arrows and status
    If LastWasFault Then
        VmFill ws, "pagein", RGB(192, 0, 0)
        VmText ws, "pagein", "PAGE IN" & vbLf & "page " & LastVPN & " -> frame " & PageFrame(LastVPN)
        VmFill ws, "status", RGB(192, 0, 0)
        VmText ws, "status", "PAGE FAULT"
    Else
        VmFill ws, "pagein", RGB(230, 230, 230)
        VmText ws, "pagein", "PAGE IN"
        If LastVPN >= 0 Then
            VmFill ws, "status", RGB(0, 128, 0)
            VmText ws, "status", "PAGE HIT"
        Else
            VmFill ws, "status", RGB(230, 230, 230)
            VmText ws, "status", "-"
        End If
    End If
    If VictimSaved Then
        VmFill ws, "writeback", RGB(237, 125, 49)
        VmText ws, "writeback", "WRITE BACK" & vbLf & "page " & VictimPage
    Else
        VmFill ws, "writeback", RGB(230, 230, 230)
        VmText ws, "writeback", "WRITE BACK"
    End If
NoDiagram:
End Sub