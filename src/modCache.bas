Attribute VB_Name = "modCache"
'==============================================================================
'  404 Angels - CPU Simulator
'  modCache : L1I (8B), L1D (8B), L2 (32B), hit/miss counters, hit rate
'  OWNER    : Alicia
'
'  Organization (line = 8 bytes, tag = physical line number = address \ 8)
'    L1I : 1 line, instructions only (fetch and operand bytes)
'    L1D : 1 line, data only (LOAD / STORE)
'    L2  : 4 lines (32 bytes), unified, fully associative, LRU replacement
'  Write policy : write-through, no write-allocate (RAM is always up to date)
'  Non-cacheable: physical E0h-EFh (keyboard queue and LCD buffer)
'
'  Sheet "Cache": tables (A:K) + live diagram (shapes from column N)
'==============================================================================
Option Explicit

Public L1I_Hits As Long, L1I_Misses As Long
Public L1D_Hits As Long, L1D_Misses As Long
Public L2_Hits As Long, L2_Misses As Long
Public Uncached As Long

Private Const LINE_SIZE As Long = 8
Private Const L2_LINES As Long = 4
Private Const IO_LO As Long = &HE0
Private Const IO_HI As Long = &HEF
Private Const CD As String = "cd_"           ' prefix of the diagram shapes

Private L1I_Valid As Boolean, L1I_Tag As Long, L1I_Data(0 To 7) As Long
Private L1D_Valid As Boolean, L1D_Tag As Long, L1D_Data(0 To 7) As Long
Private L2_Valid(0 To 3) As Boolean, L2_Tag(0 To 3) As Long
Private L2_Data(0 To 3, 0 To 7) As Long, L2_Used(0 To 3) As Long
Private UseClock As Long

' --- last access (for the animation) ---
Private JustAccessed As Boolean
Private LastAddr As Long
Private LastWrite As Boolean
Private LastL1 As String         ' "L1I" or "L1D"
Private LastL1Hit As Boolean
Private LastL2State As String    ' "HIT", "MISS" or ""
Private LastL2Line As Long       ' L2 line used (-1 = none)
Private LastFromRAM As Boolean
Private LastUncached As Boolean
Private LastText As String

'==============================================================================
'  PUBLIC INTERFACE (docs/INTERFACES.md)
'==============================================================================
Public Sub Cache_Reset()
    Dim i As Long, j As Long
    L1I_Hits = 0: L1I_Misses = 0
    L1D_Hits = 0: L1D_Misses = 0
    L2_Hits = 0: L2_Misses = 0
    Uncached = 0
    L1I_Valid = False: L1D_Valid = False
    For i = 0 To 7
        L1I_Data(i) = 0
        L1D_Data(i) = 0
    Next i
    For i = 0 To L2_LINES - 1
        L2_Valid(i) = False
        L2_Tag(i) = 0
        L2_Used(i) = 0
        For j = 0 To 7
            L2_Data(i, j) = 0
        Next j
    Next i
    UseClock = 0
    JustAccessed = False
    LastText = "Cache empty"
    LastL2Line = -1
    Cache_Paint
End Sub

Public Function Cache_Read(ByVal pa As Long, ByVal isInstr As Boolean) As Long
    Dim tag As Long, off As Long, l2 As Long
    pa = pa And &HFF
    tag = pa \ LINE_SIZE
    off = pa Mod LINE_SIZE
    BeginAccess pa, False, isInstr

    ' I/O range: never cached
    If pa >= IO_LO And pa <= IO_HI Then
        Uncached = Uncached + 1
        LastUncached = True
        LastFromRAM = True
        LastText = "UNCACHED I/O  " & Hex2(pa) & "h -> RAM"
        Cache_Read = RAM_Read(pa)
        Exit Function
    End If

    ' Level 1
    If isInstr Then
        If L1I_Valid And L1I_Tag = tag Then
            L1I_Hits = L1I_Hits + 1
            LastL1Hit = True
            LastText = "L1I HIT  " & Hex2(pa) & "h"
            Cache_Read = L1I_Data(off)
            Exit Function
        End If
        L1I_Misses = L1I_Misses + 1
    Else
        If L1D_Valid And L1D_Tag = tag Then
            L1D_Hits = L1D_Hits + 1
            LastL1Hit = True
            LastText = "L1D HIT  " & Hex2(pa) & "h"
            Cache_Read = L1D_Data(off)
            Exit Function
        End If
        L1D_Misses = L1D_Misses + 1
    End If

    ' Level 2
    l2 = FindL2(tag)
    If l2 >= 0 Then
        L2_Hits = L2_Hits + 1
        LastL2State = "HIT"
        LastText = LastL1 & " MISS -> L2 HIT (line " & l2 & ")  " & Hex2(pa) & "h"
    Else
        L2_Misses = L2_Misses + 1
        LastL2State = "MISS"
        l2 = VictimL2()
        FillL2 l2, tag
        LastFromRAM = True
        LastText = LastL1 & " MISS -> L2 MISS -> RAM: line " & Hex2(tag * LINE_SIZE) & "h-" & _
                   Hex2(tag * LINE_SIZE + 7) & "h loaded into L2 line " & l2
    End If
    UseClock = UseClock + 1
    L2_Used(l2) = UseClock
    LastL2Line = l2
    CopyToL1 isInstr, l2
    Cache_Read = L2_Data(l2, off)
End Function

Public Sub Cache_Write(ByVal pa As Long, ByVal value As Long)
    Dim tag As Long, off As Long, l2 As Long, hitTxt As String
    pa = pa And &HFF
    value = value And &HFF
    tag = pa \ LINE_SIZE
    off = pa Mod LINE_SIZE
    BeginAccess pa, True, False

    If pa >= IO_LO And pa <= IO_HI Then
        Uncached = Uncached + 1
        LastUncached = True
        LastFromRAM = True
        LastText = "UNCACHED I/O write  " & Hex2(pa) & "h -> RAM"
        RAM_Write pa, value
        Exit Sub
    End If

    ' write-through: update every level that holds the line, then RAM
    If L1D_Valid And L1D_Tag = tag Then
        L1D_Hits = L1D_Hits + 1
        LastL1Hit = True
        L1D_Data(off) = value
        hitTxt = "L1D HIT"
    Else
        L1D_Misses = L1D_Misses + 1
        hitTxt = "L1D MISS"
    End If
    If L1I_Valid And L1I_Tag = tag Then
        L1I_Data(off) = value          ' keep L1I coherent (self-modifying code)
    End If
    l2 = FindL2(tag)
    If l2 >= 0 Then
        L2_Data(l2, off) = value
        If Not LastL1Hit Then
            L2_Hits = L2_Hits + 1
            LastL2State = "HIT"
            hitTxt = hitTxt & " -> L2 HIT"
        End If
        UseClock = UseClock + 1
        L2_Used(l2) = UseClock
        LastL2Line = l2
    ElseIf Not LastL1Hit Then
        L2_Misses = L2_Misses + 1
        LastL2State = "MISS"
        hitTxt = hitTxt & " -> L2 MISS (no write-allocate)"
    End If
    RAM_Write pa, value
    LastFromRAM = True
    LastText = "WRITE " & Hex2(value) & "h to " & Hex2(pa) & "h: " & hitTxt & " -> RAM (write-through)"
End Sub

Public Sub Cache_InvalidateFrame(ByVal frame As Long)
    Dim lo As Long, hi As Long, i As Long
    lo = frame * 32 \ LINE_SIZE          ' first line number of the frame
    hi = lo + 32 \ LINE_SIZE - 1
    If L1I_Valid And L1I_Tag >= lo And L1I_Tag <= hi Then L1I_Valid = False
    If L1D_Valid And L1D_Tag >= lo And L1D_Tag <= hi Then L1D_Valid = False
    For i = 0 To L2_LINES - 1
        If L2_Valid(i) And L2_Tag(i) >= lo And L2_Tag(i) <= hi Then L2_Valid(i) = False
    Next i
End Sub

' Short text of the last access (shown by the bus diagram)
Public Function Cache_LastResult() As String
    Cache_LastResult = LastText
End Function

'==============================================================================
'  INTERNAL HELPERS
'==============================================================================
Private Sub BeginAccess(ByVal pa As Long, ByVal isWrite As Boolean, ByVal isInstr As Boolean)
    JustAccessed = True
    LastAddr = pa
    LastWrite = isWrite
    If isInstr Then LastL1 = "L1I" Else LastL1 = "L1D"
    LastL1Hit = False
    LastL2State = ""
    LastL2Line = -1
    LastFromRAM = False
    LastUncached = False
End Sub

Private Function FindL2(ByVal tag As Long) As Long
    Dim i As Long
    FindL2 = -1
    For i = 0 To L2_LINES - 1
        If L2_Valid(i) And L2_Tag(i) = tag Then
            FindL2 = i
            Exit Function
        End If
    Next i
End Function

' LRU: first an empty line, otherwise the least recently used one
Private Function VictimL2() As Long
    Dim i As Long, best As Long
    For i = 0 To L2_LINES - 1
        If Not L2_Valid(i) Then
            VictimL2 = i
            Exit Function
        End If
    Next i
    best = 0
    For i = 1 To L2_LINES - 1
        If L2_Used(i) < L2_Used(best) Then best = i
    Next i
    VictimL2 = best
End Function

Private Sub FillL2(ByVal l2 As Long, ByVal tag As Long)
    Dim j As Long
    For j = 0 To LINE_SIZE - 1
        L2_Data(l2, j) = RAM_Read(tag * LINE_SIZE + j)
    Next j
    L2_Tag(l2) = tag
    L2_Valid(l2) = True
End Sub

Private Sub CopyToL1(ByVal isInstr As Boolean, ByVal l2 As Long)
    Dim j As Long
    For j = 0 To LINE_SIZE - 1
        If isInstr Then
            L1I_Data(j) = L2_Data(l2, j)
        Else
            L1D_Data(j) = L2_Data(l2, j)
        End If
    Next j
    If isInstr Then
        L1I_Tag = L2_Tag(l2): L1I_Valid = True
    Else
        L1D_Tag = L2_Tag(l2): L1D_Valid = True
    End If
End Sub

Private Function Rate(ByVal hits As Long, ByVal misses As Long) As String
    If hits + misses = 0 Then
        Rate = "-"
    Else
        Rate = Format$(100 * hits / (hits + misses), "0.0") & " %"
    End If
End Function

'==============================================================================
'  SHEET "Cache"
'==============================================================================
Private Function Sh() As Worksheet
    Set Sh = GetSheet(SH_CACHE)
End Function

Private Sub Banner(ByVal r As Range, ByVal txt As String, ByVal color As Long)
    r.Merge
    r.Value = txt
    r.Font.Bold = True
    r.Font.Color = RGB(255, 255, 255)
    r.Interior.Color = color
    r.HorizontalAlignment = xlCenter
End Sub

Private Sub LineHeader(ByVal ws As Worksheet, ByVal row As Long, ByVal withLru As Boolean)
    Dim j As Long
    ws.Cells(row, 1).Value = "Line"
    ws.Cells(row, 2).Value = "Valid"
    ws.Cells(row, 3).Value = "Tag (addresses)"
    For j = 0 To 7
        ws.Cells(row, 4 + j).Value = "B" & j
    Next j
    If withLru Then ws.Cells(row, 12).Value = "LRU age"
    With ws.Range(ws.Cells(row, 1), ws.Cells(row, IIf(withLru, 12, 11)))
        .Font.Bold = True
        .Interior.Color = RGB(234, 240, 248)
        .HorizontalAlignment = xlCenter
    End With
End Sub

Public Sub Cache_SetupSheet()
    Dim ws As Worksheet, i As Long
    Set ws = Sh
    Application.DisplayAlerts = False           ' no "merging cells" warnings
    ws.Range("A1:L40").UnMerge
    ws.Range("A1:L40").ClearContents
    ws.Cells.Font.Name = "Calibri"
    ws.Columns("A").ColumnWidth = 9
    ws.Columns("B").ColumnWidth = 7
    ws.Columns("C").ColumnWidth = 16
    ws.Range("D:K").ColumnWidth = 5.5
    ws.Columns("L").ColumnWidth = 9
    ws.Columns("M").ColumnWidth = 3

    Banner ws.Range("A1:L1"), TEAM_NAME & "  -  CACHE HIERARCHY  L1I / L1D / L2", RGB(31, 42, 74)
    ws.Range("A1").Font.Size = 14
    ws.Range("A3:L3").Merge
    ws.Range("A3").Font.Name = "Consolas"
    ws.Range("A3").Font.Bold = True
    ws.Range("A3").Font.Size = 9
    ws.Range("A3").HorizontalAlignment = xlLeft
    ws.Range("A3").VerticalAlignment = xlCenter
    ws.Range("A3").WrapText = True
    ws.Rows(3).RowHeight = 30
    ws.Range("C7:K7,C11:K11,C15:K18").NumberFormat = "@"

    Banner ws.Range("A5:L5"), "L1I  -  instruction cache  (8 bytes = 1 line)", RGB(47, 111, 181)
    LineHeader ws, 6, False
    ws.Cells(7, 1).Value = 0
    Banner ws.Range("A9:L9"), "L1D  -  data cache  (8 bytes = 1 line)", RGB(46, 158, 98)
    LineHeader ws, 10, False
    ws.Cells(11, 1).Value = 0
    Banner ws.Range("A13:L13"), "L2  -  unified cache  (32 bytes = 4 lines, fully associative, LRU)", RGB(123, 94, 167)
    LineHeader ws, 14, True
    For i = 0 To L2_LINES - 1
        ws.Cells(15 + i, 1).Value = i
    Next i

    Banner ws.Range("A20:I20"), "HIT / MISS COUNTERS", RGB(31, 42, 74)
    For i = 21 To 25
        ws.Range("D" & i & ":F" & i).Merge
        ws.Range("G" & i & ":I" & i).Merge
    Next i
    ws.Range("A21").Value = "Level"
    ws.Range("B21").Value = "Hits"
    ws.Range("C21").Value = "Misses"
    ws.Range("D21").Value = "Accesses"
    ws.Range("G21").Value = "Hit rate"
    ws.Range("A21:I21").Font.Bold = True
    ws.Range("A21:I21").Interior.Color = RGB(234, 240, 248)
    ws.Range("A22:A25").Value = Application.Transpose(Array("L1I", "L1D", "L2", "I/O"))
    ws.Range("A22:A25").Font.Bold = True
    ws.Range("G22:G25").NumberFormat = "@"

    ws.Range("A27").Value = "Policies"
    ws.Range("A27").Font.Bold = True
    ws.Range("B27").Value = "Line = 8 bytes. L1I and L1D: 1 line each. L2: 4 lines, fully associative, LRU replacement."
    ws.Range("B28").Value = "Write-through, no write-allocate: every write goes to RAM. I/O range E0h-EFh is never cached."
    ws.Range("B29").Value = "When virtual memory replaces a frame, its lines are invalidated (Cache_InvalidateFrame)."

    With ws.Range("A6:L18")
        .HorizontalAlignment = xlCenter
        .Font.Size = 10
    End With
    ws.Range("D7:K7,D11:K11,D15:K18").Font.Name = "Consolas"
    ws.Range("A5:L18").Borders.LineStyle = xlContinuous
    ws.Range("A5:L18").Borders.Color = RGB(200, 208, 222)
    ws.Range("A21:I25").Borders.LineStyle = xlContinuous
    ws.Range("A21:I25").Borders.Color = RGB(200, 208, 222)
    ws.Range("A21:I25").HorizontalAlignment = xlCenter

    DrawFlow
    Application.DisplayAlerts = True
    Cache_Paint
End Sub

Private Sub PaintLine(ByVal ws As Worksheet, ByVal row As Long, ByVal valid As Boolean, ByVal tag As Long, _
                      ByRef data() As Long, ByVal l2 As Long, ByVal highlight As Long)
    Dim j As Long, rng As Range
    ws.Cells(row, 2).Value = IIf(valid, 1, 0)
    If valid Then
        ws.Cells(row, 3).Value = Hex2(tag * LINE_SIZE) & "h-" & Hex2(tag * LINE_SIZE + 7) & "h"
    Else
        ws.Cells(row, 3).Value = "-"
    End If
    For j = 0 To 7
        If valid Then
            If l2 >= 0 Then
                ws.Cells(row, 4 + j).Value = Hex2(L2_Data(l2, j))
            Else
                ws.Cells(row, 4 + j).Value = Hex2(data(j))
            End If
        Else
            ws.Cells(row, 4 + j).Value = "--"
        End If
    Next j
    Set rng = ws.Range(ws.Cells(row, 1), ws.Cells(row, 11))
    If highlight <> 0 Then
        rng.Interior.Color = highlight
    ElseIf valid Then
        rng.Interior.Color = RGB(255, 255, 255)
    Else
        rng.Interior.Color = RGB(245, 246, 248)
    End If
    ' byte that was accessed
    If highlight <> 0 And valid Then
        If tag = LastAddr \ LINE_SIZE Then
            ws.Cells(row, 4 + (LastAddr Mod LINE_SIZE)).Interior.Color = RGB(255, 217, 102)
        End If
    End If
End Sub

Private Sub CounterRow(ByVal ws As Worksheet, ByVal row As Long, ByVal hits As Long, ByVal misses As Long)
    ws.Cells(row, 2).Value = hits
    ws.Cells(row, 3).Value = misses
    ws.Cells(row, 4).Value = hits + misses
    ws.Cells(row, 7).Value = Rate(hits, misses)
End Sub

Public Sub Cache_Paint()
    Dim ws As Worksheet, i As Long, hl As Long, hitCol As Long, missCol As Long, dummy(0 To 7) As Long
    Set ws = Sh
    If ws.Range("A1").Value = "" Then Exit Sub        ' sheet not created yet
    hitCol = RGB(214, 240, 224)
    missCol = RGB(252, 222, 220)

    ' banner of the last access
    If JustAccessed Then
        ws.Range("A3").Value = LastText
        If LastUncached Then
            ws.Range("A3").Interior.Color = RGB(255, 242, 204)
        ElseIf LastL1Hit Or LastL2State = "HIT" Then
            ws.Range("A3").Interior.Color = hitCol
        Else
            ws.Range("A3").Interior.Color = missCol
        End If
    Else
        ws.Range("A3").Value = LastText & "   (no memory access in this pulse)"
        ws.Range("A3").Interior.Color = RGB(245, 246, 248)
    End If

    ' L1I / L1D
    hl = 0
    If JustAccessed And LastL1 = "L1I" And Not LastUncached Then hl = IIf(LastL1Hit, hitCol, RGB(214, 228, 247))
    PaintLine ws, 7, L1I_Valid, L1I_Tag, L1I_Data, -1, hl
    hl = 0
    If JustAccessed And LastL1 = "L1D" And Not LastUncached Then hl = IIf(LastL1Hit, hitCol, RGB(214, 228, 247))
    PaintLine ws, 11, L1D_Valid, L1D_Tag, L1D_Data, -1, hl

    ' L2
    For i = 0 To L2_LINES - 1
        hl = 0
        If JustAccessed And i = LastL2Line Then
            If LastL2State = "HIT" Then
                hl = hitCol
            ElseIf LastL2State = "MISS" Then
                hl = missCol
            Else
                hl = RGB(214, 228, 247)
            End If
        End If
        PaintLine ws, 15 + i, L2_Valid(i), L2_Tag(i), dummy, i, hl
        If L2_Valid(i) Then
            ws.Cells(15 + i, 12).Value = UseClock - L2_Used(i)
        Else
            ws.Cells(15 + i, 12).Value = "-"
        End If
    Next i

    ' counters
    CounterRow ws, 22, L1I_Hits, L1I_Misses
    CounterRow ws, 23, L1D_Hits, L1D_Misses
    CounterRow ws, 24, L2_Hits, L2_Misses
    ws.Range("B25").Value = "-"
    ws.Range("C25").Value = "-"
    ws.Range("D25").Value = Uncached
    ws.Range("G25").Value = "never cached"

    PaintFlow
    JustAccessed = False
End Sub

'==============================================================================
'  LIVE DIAGRAM: CPU -> L1I / L1D -> L2 -> RAM (shapes from column N)
'==============================================================================
Private Function CdShape(ByVal ws As Worksheet, ByVal nm As String, ByVal kind As Long, _
                         ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                         ByVal txt As String) As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(kind, ws.Range("N3").Left + x, ws.Range("N3").Top + y, w, h)
    shp.Name = CD & nm
    shp.Line.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
    If txt <> "" Then
        With shp.TextFrame
            .Characters.Text = txt
            .Characters.Font.Size = 9
            .Characters.Font.Bold = True
            .Characters.Font.Color = RGB(31, 42, 74)
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
    End If
    Set CdShape = shp
End Function

Private Sub CdBox(ByVal ws As Worksheet, ByVal nm As String, ByVal x As Double, ByVal y As Double, _
                  ByVal w As Double, ByVal h As Double, ByVal txt As String)
    Dim shp As Shape
    Set shp = CdShape(ws, "box_" & nm, 5, x, y, w, h, txt)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(31, 42, 74)
    shp.Line.Weight = 1.5
End Sub

Private Sub DrawFlow()
    Dim ws As Worksheet, i As Long, shp As Shape
    Set ws = Sh
    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(CD)) = CD Then ws.Shapes(i).Delete
    Next i
    Set shp = CdShape(ws, "frame", 1, 0, 0, 360, 420, "")
    shp.Fill.ForeColor.RGB = RGB(244, 247, 251)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(200, 208, 222)
    Set shp = CdShape(ws, "title", 1, 0, 0, 360, 20, "MEMORY HIERARCHY  -  last access")
    shp.Fill.ForeColor.RGB = RGB(31, 42, 74)
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)

    CdBox ws, "CPU", 120, 32, 120, 40, "CPU"
    Call CdShape(ws, "wire_L1I", 1, 148, 72, 6, 40, "")
    Call CdShape(ws, "wire_L1D", 1, 206, 72, 6, 40, "")
    CdBox ws, "L1I", 40, 112, 130, 50, "L1I  8 B"
    CdBox ws, "L1D", 190, 112, 130, 50, "L1D  8 B"
    Call CdShape(ws, "wire_L1I_L2", 1, 102, 162, 6, 40, "")
    Call CdShape(ws, "wire_L1D_L2", 1, 252, 162, 6, 40, "")
    CdBox ws, "L2", 60, 202, 240, 60, "L2  32 B  (4 lines, LRU)"
    Call CdShape(ws, "wire_L2_RAM", 1, 177, 262, 6, 40, "")
    Call CdShape(ws, "wire_IO", 1, 330, 72, 6, 230, "")
    Set shp = CdShape(ws, "lbl_IO", 1, 286, 180, 44, 26, "I/O" & vbLf & "uncached")
    shp.Fill.Visible = msoFalse
    shp.TextFrame.Characters.Font.Size = 7
    CdBox ws, "RAM", 60, 302, 240, 50, "RAM  256 B"
    Set shp = CdShape(ws, "result", 5, 20, 366, 320, 42, "")
    shp.TextFrame.Characters.Font.Size = 9
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(200, 208, 222)
    On Error GoTo 0
End Sub

Private Sub CdFill(ByVal ws As Worksheet, ByVal nm As String, ByVal color As Long)
    ws.Shapes(CD & nm).Fill.ForeColor.RGB = color
End Sub

Private Sub CdText(ByVal ws As Worksheet, ByVal nm As String, ByVal txt As String)
    With ws.Shapes(CD & nm).TextFrame
        .Characters.Text = txt
        .Characters.Font.Color = RGB(31, 42, 74)
        .Characters.Font.Bold = True
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
End Sub

Private Sub PaintFlow()
    Dim ws As Worksheet, white As Long, idle As Long, hitC As Long, missC As Long, ramC As Long
    Dim l1Hit As Boolean, l2Hit As Boolean, l2Miss As Boolean
    Set ws = Sh
    On Error Resume Next
    white = RGB(255, 255, 255)
    idle = RGB(200, 205, 215)
    hitC = RGB(46, 158, 98)
    missC = RGB(217, 83, 79)
    ramC = RGB(242, 165, 65)

    ' reset everything to idle
    CdFill ws, "box_L1I", white
    CdFill ws, "box_L1D", white
    CdFill ws, "box_L2", white
    CdFill ws, "box_RAM", white
    CdFill ws, "wire_L1I", idle
    CdFill ws, "wire_L1D", idle
    CdFill ws, "wire_L1I_L2", idle
    CdFill ws, "wire_L1D_L2", idle
    CdFill ws, "wire_L2_RAM", idle
    CdFill ws, "wire_IO", idle
    CdText ws, "box_L1I", "L1I  8 B" & vbLf & "hits " & L1I_Hits & " / misses " & L1I_Misses & vbLf & "hit rate " & Rate(L1I_Hits, L1I_Misses)
    CdText ws, "box_L1D", "L1D  8 B" & vbLf & "hits " & L1D_Hits & " / misses " & L1D_Misses & vbLf & "hit rate " & Rate(L1D_Hits, L1D_Misses)
    CdText ws, "box_L2", "L2  32 B  (4 lines, LRU)" & vbLf & "hits " & L2_Hits & " / misses " & L2_Misses & "   hit rate " & Rate(L2_Hits, L2_Misses)
    If Not JustAccessed Then
        CdText ws, "result", "No memory access in this pulse"
        CdFill ws, "result", white
        Exit Sub
    End If

    CdText ws, "result", LastText
    If LastUncached Then
        CdFill ws, "wire_IO", ramC
        CdFill ws, "box_RAM", RGB(255, 242, 204)
        CdFill ws, "result", RGB(255, 242, 204)
        Exit Sub
    End If

    l1Hit = LastL1Hit
    l2Hit = (LastL2State = "HIT")
    l2Miss = (LastL2State = "MISS")
    CdFill ws, "wire_" & LastL1, IIf(l1Hit, hitC, missC)
    CdFill ws, "box_" & LastL1, IIf(l1Hit, RGB(214, 240, 224), RGB(252, 222, 220))
    If Not l1Hit Then
        CdFill ws, "wire_" & LastL1 & "_L2", IIf(l2Hit, hitC, missC)
        CdFill ws, "box_L2", IIf(l2Hit, RGB(214, 240, 224), IIf(l2Miss, RGB(252, 222, 220), white))
    End If
    If LastFromRAM Then
        CdFill ws, "wire_L2_RAM", ramC
        CdFill ws, "box_RAM", RGB(255, 242, 204)
    End If
    If l1Hit Or l2Hit Then
        CdFill ws, "result", RGB(214, 240, 224)
    Else
        CdFill ws, "result", RGB(252, 222, 220)
    End If
    On Error GoTo 0
End Sub
