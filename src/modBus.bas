Attribute VB_Name = "modBus"
'==============================================================================
'  404 Angels - CPU Simulator
'  modBus : control / address / data buses, tri-state buffers, control word
'  OWNER  : Alicia
'
'  Rule: in every micro-step only ONE component drives each bus (Enable = 1);
'        all the others stay in High-Z (Z). Components with Load = 1 latch
'        the value that is on the bus at the clock edge.
'
'  Buses
'    Internal bus : inside the CPU  (PC, IR, MDR, AX, BX, CX, DX, TMP, MAR)
'    External bus : CPU <-> memory / cache / I/O
'                   - address bus (9 bits, driven by MAR)
'                   - data bus    (8 bits, bidirectional through MDR)
'                   - control bus (READ, WRITE, IO, CLOCK)
'
'  Panel on sheet CPU: table in columns F:N (rows 3-27) and a live
'  diagram drawn with shapes starting at column P.
'==============================================================================
Option Explicit

Public ControlWord As String     ' e.g. "En_PC Ld_MAR"
Public AddressBus As Long        ' 9 bits
Public DataBus As Long           ' 8 bits, -1 = High-Z

' Signals of the control word, in display order (24 bits)
Private Const SIGNAL_LIST As String = _
    "En_PC Ld_PC PC_Inc Ld_MAR En_MDR Ld_MDR En_IR Ld_IR " & _
    "En_AX Ld_AX En_BX Ld_BX En_CX Ld_CX En_DX Ld_DX " & _
    "En_TMP Ld_TMP ALU Mem_R Mem_W IO En_KBD Ld_LCD"

' Components on the internal bus (row 6 of the panel, columns G..N)
Private Const INTERNAL_LIST As String = "PC IR MDR AX BX CX DX TMP"
' Components on the external data bus (row 11, columns G..L)
Private Const EXTERNAL_LIST As String = "MDR MEM KBD LCD AX BX"

Private Const DG As String = "bd_"   ' prefix of every diagram shape
Private Const C0 As Long = 7     ' column G
Private Const LAST_COL As Long = 14   ' column N

' Colors
Private Function ColEnable() As Long
    ColEnable = RGB(46, 158, 98)
End Function
Private Function ColLoad() As Long
    ColLoad = RGB(47, 111, 181)
End Function
Private Function ColHighZ() As Long
    ColHighZ = RGB(205, 210, 220)
End Function
Private Function ColBusOn() As Long
    ColBusOn = RGB(242, 165, 65)
End Function
Private Function ColBusOff() As Long
    ColBusOff = RGB(200, 205, 215)
End Function
Private Function ColConflict() As Long
    ColConflict = RGB(217, 83, 79)
End Function

'==============================================================================
'  PUBLIC INTERFACE (docs/INTERFACES.md)
'==============================================================================
Public Sub Bus_Reset()
    ControlWord = ""
    AddressBus = 0
    DataBus = -1
End Sub

Public Sub Bus_ControlWord(ByVal signals As String)
    ControlWord = signals
End Sub

Public Function Bus_MemRead(ByVal vaddr As Long, ByVal isInstr As Boolean) As Long
    Dim pa As Long
    AddressBus = vaddr
    pa = VM_Translate(vaddr, False)
    DataBus = Cache_Read(pa, isInstr)
    Bus_MemRead = DataBus
End Function

Public Sub Bus_MemWrite(ByVal vaddr As Long, ByVal value As Long)
    Dim pa As Long
    AddressBus = vaddr
    DataBus = value And &HFF
    pa = VM_Translate(vaddr, True)
    Cache_Write pa, DataBus
End Sub

Public Function Bus_IoRead(ByVal port As Long) As Long
    DataBus = IO_KeyboardRead() And &HFF
    Bus_IoRead = DataBus
End Function

Public Sub Bus_IoWrite(ByVal port As Long, ByVal value As Long)
    DataBus = value And &HFF
    IO_LcdWrite DataBus
End Sub

'==============================================================================
'  SIGNAL HELPERS
'==============================================================================
Private Function Has(ByVal token As String) As Boolean
    Has = InStr(1, " " & ControlWord & " ", " " & token & " ") > 0
End Function

' Bit value (0/1) of one signal of the control word
Private Function Bit(ByVal sig As String) As Long
    Select Case sig
        Case "ALU": Bit = IIf(InStr(1, " " & ControlWord, " ALU_") > 0, 1, 0)
        Case "IO": Bit = IIf(Has("IO_R") Or Has("IO_W"), 1, 0)
        Case Else: Bit = IIf(Has(sig), 1, 0)
    End Select
End Function

Private Function AluOp() As String
    Dim p As Long, rest As String
    p = InStr(1, " " & ControlWord, " ALU_")
    If p = 0 Then Exit Function
    rest = Mid$(ControlWord, p + 4)
    If InStr(rest, " ") > 0 Then rest = Left$(rest, InStr(rest, " ") - 1)
    AluOp = rest
End Function

' Binary string of the whole control word (24 bits)
Public Function Bus_ControlBits() As String
    Dim sigs() As String, i As Long, s As String
    sigs = Split(SIGNAL_LIST, " ")
    For i = 0 To UBound(sigs)
        s = s & Bit(sigs(i))
        If i Mod 8 = 7 And i < UBound(sigs) Then s = s & " "
    Next i
    Bus_ControlBits = s
End Function

' Is the external data bus used in this micro-step?
Private Function ExternalActive() As Boolean
    ExternalActive = Has("Mem_R") Or Has("Mem_W") Or Has("IO_R") Or Has("IO_W")
End Function

' Drivers (Enable = 1) of the external data bus
Private Function ExternalDrivers() As String
    Dim s As String
    If Has("Mem_R") Then s = s & "MEM "
    If Has("IO_R") Then s = s & "KBD "
    If Has("Mem_W") And Has("En_MDR") Then s = s & "MDR "
    If Has("IO_W") Then
        If Has("En_AX") Then s = s & "AX "
        If Has("En_BX") Then s = s & "BX "
        If Has("En_CX") Then s = s & "CX "
        If Has("En_DX") Then s = s & "DX "
    End If
    ExternalDrivers = Trim$(s)
End Function

' Receivers (Load = 1) of the external data bus
Private Function ExternalLoads() As String
    Dim s As String
    If Has("Mem_R") And Has("Ld_MDR") Then s = s & "MDR "
    If Has("Mem_W") Then s = s & "MEM "
    If Has("Ld_LCD") Then s = s & "LCD "
    If Has("IO_R") Then s = s & "TMP "
    ExternalLoads = Trim$(s)
End Function

' Drivers (Enable = 1) of the internal bus
Private Function InternalDrivers() As String
    Dim comps() As String, i As Long, s As String
    If Has("Mem_W") Or Has("IO_W") Then Exit Function    ' En_ signals go to the external bus
    comps = Split(INTERNAL_LIST, " ")
    For i = 0 To UBound(comps)
        If Has("En_" & comps(i)) Then s = s & comps(i) & " "
    Next i
    InternalDrivers = Trim$(s)
End Function

' Receivers (Load = 1) of the internal bus
Private Function InternalLoads() As String
    Dim comps() As String, i As Long, s As String
    comps = Split("PC IR MAR AX BX CX DX TMP", " ")
    For i = 0 To UBound(comps)
        If Has("Ld_" & comps(i)) Then
            If Not (comps(i) = "TMP" And Has("IO_R")) Then s = s & comps(i) & " "
        End If
    Next i
    If Has("ALU_" & AluOp()) And AluOp() <> "" Then
        If InStr(s, "TMP") = 0 Then s = s & "TMP "
    End If
    InternalLoads = Trim$(s)
End Function

Private Function CountWords(ByVal s As String) As Long
    If Trim$(s) = "" Then
        CountWords = 0
    Else
        CountWords = UBound(Split(Trim$(s), " ")) + 1
    End If
End Function

Private Function InList(ByVal item As String, ByVal list As String) As Boolean
    InList = InStr(1, " " & list & " ", " " & item & " ") > 0
End Function

' Value that the internal driver puts on the bus
Private Function DriverValue(ByVal comp As String) As String
    Select Case comp
        Case "PC": DriverValue = Hex3(PC) & "h"
        Case "IR"
            If OPV >= 0 Then DriverValue = Hex3(OPV) & "h" Else DriverValue = Hex2(IR) & "h"
        Case "MDR": DriverValue = Hex2(MDR) & "h"
        Case "AX": DriverValue = Hex2(AX) & "h"
        Case "BX": DriverValue = Hex2(BX) & "h"
        Case "CX": DriverValue = Hex2(CX) & "h"
        Case "DX": DriverValue = Hex2(DX) & "h"
        Case "TMP": DriverValue = Hex2(Core_TMP()) & "h"
    End Select
End Function

'==============================================================================
'  PANEL (sheet CPU, F3:N27)
'==============================================================================
Public Sub Bus_SetupSheet()
    Dim ws As Worksheet, i As Long, comps() As String, sigs() As String, r As Long, c As Long
    Set ws = GetSheet(SH_CPU)
    ws.Columns("F").ColumnWidth = 18
    ws.Range(ws.Cells(1, C0), ws.Cells(1, LAST_COL)).ColumnWidth = 8.5

    With ws.Range("F3:N3")
        .Merge
        .Value = "SYSTEM BUSES  -  TRI-STATE BUFFERS  -  CONTROL WORD"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 42, 74)
        .HorizontalAlignment = xlCenter
    End With

    ' Internal bus
    SectionLabel ws.Range("F5"), "INTERNAL BUS"
    ws.Range("F6").Value = "Component"
    ws.Range("F7").Value = "Output (Enable)"
    ws.Range("F8").Value = "Input (Load)"
    comps = Split(INTERNAL_LIST, " ")
    For i = 0 To UBound(comps)
        ws.Cells(6, C0 + i).Value = comps(i)
    Next i

    ' External data bus
    SectionLabel ws.Range("F10"), "EXT. DATA BUS"
    ws.Range("F11").Value = "Component"
    ws.Range("F12").Value = "Output (Enable)"
    ws.Range("F13").Value = "Input (Load)"
    comps = Split(EXTERNAL_LIST, " ")
    For i = 0 To UBound(comps)
        ws.Cells(11, C0 + i).Value = comps(i)
    Next i

    ' Address and control bus
    SectionLabel ws.Range("F15"), "ADDRESS BUS"
    SectionLabel ws.Range("F16"), "CONTROL BUS"
    ws.Range("G16:J16").Value = Array("READ", "WRITE", "IO", "CLOCK")

    ' Control word
    With ws.Range("F18:N18")
        .Merge
        .Value = "CONTROL WORD (micro-instruction emitted by the control unit)"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(123, 94, 167)
        .HorizontalAlignment = xlCenter
    End With
    sigs = Split(SIGNAL_LIST, " ")
    For i = 0 To UBound(sigs)
        r = 19 + 2 * (i \ 8)
        c = C0 + (i Mod 8)
        ws.Cells(r, c).Value = sigs(i)
        ws.Cells(r, c).Font.Size = 8
        ws.Cells(r, c).Font.Bold = True
    Next i
    ws.Range("F19").Value = "Signal"
    ws.Range("F20").Value = "Bit"
    ws.Range("F21").Value = "Signal"
    ws.Range("F22").Value = "Bit"
    ws.Range("F23").Value = "Signal"
    ws.Range("F24").Value = "Bit"
    ws.Range("F25").Value = "Binary"
    ws.Range("F26").Value = "Bus rule"
    ws.Range("F27").Value = "Transfer"
    For r = 25 To 27
        ws.Range(ws.Cells(r, C0), ws.Cells(r, LAST_COL)).Merge
        ws.Cells(r, C0).Font.Name = "Consolas"
    Next r

    With ws.Range("F5:F27")
        .Font.Bold = True
        .Font.Size = 9
    End With
    With ws.Range(ws.Cells(5, C0), ws.Cells(27, LAST_COL))
        .HorizontalAlignment = xlCenter
        .Font.Size = 9
    End With
    ws.Range("F5:N16").Borders.LineStyle = xlContinuous
    ws.Range("F5:N16").Borders.Color = RGB(200, 208, 222)
    ws.Range("F18:N27").Borders.LineStyle = xlContinuous
    ws.Range("F18:N27").Borders.Color = RGB(200, 208, 222)
    DrawDiagram
    Bus_Paint
End Sub

Private Sub SectionLabel(ByVal r As Range, ByVal txt As String)
    r.Value = txt
    r.Font.Bold = True
    r.Interior.Color = RGB(31, 42, 74)
    r.Font.Color = RGB(255, 255, 255)
End Sub

' Paints one bus line (a row of cells) with its value
Private Sub PaintLine(ByVal ws As Worksheet, ByVal row As Long, ByVal active As Boolean, _
                      ByVal conflict As Boolean, ByVal txt As String)
    With ws.Range(ws.Cells(row, C0), ws.Cells(row, LAST_COL))
        .ClearContents
        If conflict Then
            .Interior.Color = ColConflict()
        ElseIf active Then
            .Interior.Color = ColBusOn()
        Else
            .Interior.Color = ColBusOff()
        End If
        .Font.Bold = True
    End With
    ws.Cells(row, C0 + 3).Value = txt
End Sub

' Paints the Enable / Load rows of a group of components
Private Sub PaintBuffers(ByVal ws As Worksheet, ByVal rowEn As Long, ByVal list As String, _
                         ByVal drivers As String, ByVal loads As String)
    Dim comps() As String, i As Long
    comps = Split(list, " ")
    For i = 0 To UBound(comps)
        With ws.Cells(rowEn, C0 + i)
            If InList(comps(i), drivers) Then
                .Value = "EN = 1"
                .Interior.Color = ColEnable()
                .Font.Color = RGB(255, 255, 255)
            Else
                .Value = "Z"
                .Interior.Color = ColHighZ()
                .Font.Color = RGB(120, 128, 140)
            End If
        End With
        With ws.Cells(rowEn + 1, C0 + i)
            If InList(comps(i), loads) Then
                .Value = "LD = 1"
                .Interior.Color = ColLoad()
                .Font.Color = RGB(255, 255, 255)
            Else
                .Value = "0"
                .Interior.Color = RGB(255, 255, 255)
                .Font.Color = RGB(160, 166, 176)
            End If
        End With
    Next i
End Sub

Public Sub Bus_Paint()
    Dim ws As Worksheet, inDrv As String, inLd As String, exDrv As String, exLd As String
    Dim sigs() As String, i As Long, r As Long, c As Long, b As Long, txt As String, rule As String
    Set ws = GetSheet(SH_CPU)
    If ws.Range("F3").Value = "" Then Exit Sub          ' panel not created yet

    inDrv = InternalDrivers()
    inLd = InternalLoads()
    exDrv = ExternalDrivers()
    exLd = ExternalLoads()

    ' --- internal bus ---
    If CountWords(inDrv) = 1 Then
        txt = inDrv & " -> " & DriverValue(inDrv)
    ElseIf CountWords(inDrv) > 1 Then
        txt = "CONFLICT: " & inDrv
    ElseIf inLd <> "" Then
        txt = "ALU -> TMP"
    Else
        txt = "High-Z (idle)"
    End If
    PaintLine ws, 5, (inDrv <> "" Or inLd <> ""), CountWords(inDrv) > 1, txt
    PaintBuffers ws, 7, INTERNAL_LIST, inDrv, Replace(inLd, "MAR", "")

    ' --- external data bus ---
    If ExternalActive() And DataBus >= 0 Then
        txt = exDrv & " -> " & Hex2(DataBus) & "h -> " & exLd
    Else
        txt = "High-Z (idle)"
    End If
    PaintLine ws, 10, ExternalActive(), CountWords(exDrv) > 1, txt
    PaintBuffers ws, 12, EXTERNAL_LIST, exDrv, exLd

    ' --- address bus (driven by MAR) ---
    If Has("Mem_R") Or Has("Mem_W") Then
        PaintLine ws, 15, True, False, "MAR -> " & Hex3(AddressBus) & "h"
    ElseIf Has("Ld_MAR") Then
        PaintLine ws, 15, False, False, "MAR <- " & Hex3(MAR) & "h (latched)"
    Else
        PaintLine ws, 15, False, False, "High-Z (idle)"
    End If

    ' --- control bus ---
    LightCell ws.Cells(16, C0), Has("Mem_R") Or Has("IO_R")
    LightCell ws.Cells(16, C0 + 1), Has("Mem_W") Or Has("IO_W")
    LightCell ws.Cells(16, C0 + 2), Has("IO_R") Or Has("IO_W")
    LightCell ws.Cells(16, C0 + 3), ControlWord <> ""

    ' --- control word bits ---
    sigs = Split(SIGNAL_LIST, " ")
    For i = 0 To UBound(sigs)
        r = 20 + 2 * (i \ 8)
        c = C0 + (i Mod 8)
        b = Bit(sigs(i))
        With ws.Cells(r, c)
            If sigs(i) = "ALU" And b = 1 Then
                .Value = AluOp()
            Else
                .Value = b
            End If
            If b = 1 Then
                .Interior.Color = IIf(Left$(sigs(i), 3) = "Ld_", ColLoad(), ColEnable())
                .Font.Color = RGB(255, 255, 255)
                .Font.Bold = True
            Else
                .Interior.Color = RGB(255, 255, 255)
                .Font.Color = RGB(160, 166, 176)
                .Font.Bold = False
            End If
        End With
    Next i
    ws.Cells(25, C0).Value = Bus_ControlBits()

    ' --- single-driver rule ---
    If CountWords(inDrv) > 1 Or CountWords(exDrv) > 1 Then
        rule = "VIOLATED: more than one driver on the same bus"
        ws.Cells(26, C0).Interior.Color = ColConflict()
    Else
        rule = "OK - internal: " & IIf(inDrv = "", "none", inDrv) & " | external: " & IIf(exDrv = "", "none", exDrv) & " | others in High-Z"
        ws.Cells(26, C0).Interior.Color = RGB(233, 246, 239)
    End If
    ws.Cells(26, C0).Value = rule
    If ControlWord = "" Or ControlWord = "-" Then
        ws.Cells(27, C0).Value = "no transfer"
    Else
        ws.Cells(27, C0).Value = ControlWord
    End If
    PaintDiagram inDrv, inLd, exDrv, exLd
End Sub

Private Sub LightCell(ByVal cell As Range, ByVal isOn As Boolean)
    If isOn Then
        cell.Interior.Color = ColBusOn()
        cell.Font.Color = RGB(31, 42, 74)
        cell.Font.Bold = True
    Else
        cell.Interior.Color = ColHighZ()
        cell.Font.Color = RGB(150, 156, 166)
        cell.Font.Bold = False
    End If
End Sub

'==============================================================================
'  GRAPHIC DIAGRAM (shapes on sheet CPU, starting at column P)
'    boxes      = components (green = drives the bus, blue = loads from it)
'    triangles  = tri-state buffers (green = Enable, gray = High-Z)
'    thick bars = buses (orange = carrying data)
'==============================================================================

Private Function DgShape(ByVal ws As Worksheet, ByVal nm As String, ByVal kind As Long, _
                         ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                         ByVal txt As String) As Shape
    Dim shp As Shape, x0 As Double, y0 As Double
    x0 = ws.Range("P3").Left
    y0 = ws.Range("P3").Top
    Set shp = ws.Shapes.AddShape(kind, x0 + x, y0 + y, w, h)
    shp.Name = DG & nm
    shp.Line.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
    If txt <> "" Then
        With shp.TextFrame
            .Characters.Text = txt
            .Characters.Font.Size = 8
            .Characters.Font.Bold = True
            .Characters.Font.Color = RGB(31, 42, 74)
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
            .MarginLeft = 1
            .MarginRight = 1
        End With
    End If
    Set DgShape = shp
End Function

Private Sub DgBox(ByVal ws As Worksheet, ByVal nm As String, ByVal x As Double, ByVal y As Double, _
                  ByVal w As Double, ByVal h As Double, ByVal txt As String)
    Dim shp As Shape
    Set shp = DgShape(ws, "box_" & nm, 5, x, y, w, h, txt)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(31, 42, 74)
    shp.Line.Weight = 1.25
End Sub

Private Sub DgBuffer(ByVal ws As Worksheet, ByVal nm As String, ByVal midX As Double, ByVal y As Double, _
                     ByVal pointDown As Boolean)
    Dim shp As Shape
    Set shp = DgShape(ws, "buf_" & nm, 7, midX - 11, y - 1, 22, 15, "")
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(110, 118, 132)
    shp.Line.Weight = 1
    If pointDown Then shp.Rotation = 180
End Sub

Private Sub DgWire(ByVal ws As Worksheet, ByVal nm As String, ByVal x As Double, ByVal y As Double, _
                   ByVal w As Double, ByVal h As Double)
    Call DgShape(ws, "wire_" & nm, 1, x, y, w, h, "")
End Sub

Private Sub DgLabel(ByVal ws As Worksheet, ByVal nm As String, ByVal x As Double, ByVal y As Double, _
                    ByVal w As Double, ByVal txt As String)
    Dim shp As Shape
    Set shp = DgShape(ws, "lbl_" & nm, 1, x, y, w, 12, txt)
    shp.Fill.Visible = msoFalse
    shp.TextFrame.HorizontalAlignment = xlHAlignLeft
    shp.TextFrame.Characters.Font.Size = 7
End Sub

Private Sub DeleteDiagram(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(DG)) = DG Then ws.Shapes(i).Delete
    Next i
End Sub

Private Function CompX(ByVal i As Long) As Double
    CompX = 14 + i * 66        ' left of component i in the top row
End Function

Private Sub DrawDiagram()
    Dim ws As Worksheet, comps() As String, i As Long, midX As Double, shp As Shape
    Set ws = GetSheet(SH_CPU)
    On Error Resume Next
    DeleteDiagram ws

    ' frame and title
    Set shp = DgShape(ws, "frame", 1, 0, 0, 620, 372, "")
    shp.Fill.ForeColor.RGB = RGB(244, 247, 251)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(200, 208, 222)
    Set shp = DgShape(ws, "title", 1, 0, 0, 620, 18, "BUS DIAGRAM  -  live view of every clock pulse")
    shp.Fill.ForeColor.RGB = RGB(31, 42, 74)
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)

    ' top row: CPU registers
    DgLabel ws, "cpu", 8, 22, 200, "CPU"
    comps = Split("PC IR MAR MDR AX BX CX DX TMP", " ")
    For i = 0 To UBound(comps)
        DgBox ws, comps(i), CompX(i), 36, 56, 34, comps(i)
        midX = CompX(i) + 28
        If comps(i) <> "MAR" Then
            DgBuffer ws, comps(i), midX, 74, True
        End If
        DgWire ws, comps(i), midX - 2, 87, 4, 13
    Next i

    ' internal bus
    Set shp = DgShape(ws, "bus_int", 1, 8, 100, 604, 12, "")
    DgLabel ws, "int", 8, 113, 200, "INTERNAL BUS (CPU)"

    ' ALU under TMP
    DgWire ws, "ALU", CompX(8) + 26, 112, 4, 14
    DgBox ws, "ALU", CompX(7) + 20, 126, 98, 26, "ALU"

    ' MAR -> address bus, MDR <-> external data bus (bus interface)
    DgWire ws, "MARaddr", CompX(2) + 26, 112, 4, 112
    DgLabel ws, "maraddr", CompX(2) + 32, 150, 60, "address"
    DgWire ws, "MDRext", CompX(3) + 26, 112, 4, 80
    DgBuffer ws, "MDRext", CompX(3) + 28, 140, True
    DgLabel ws, "mdrext", CompX(3) + 34, 150, 80, "data (bus interface)"

    ' external data bus and address bus
    Set shp = DgShape(ws, "bus_ext", 1, 8, 192, 604, 12, "")
    DgLabel ws, "ext", 8, 180, 200, "EXTERNAL DATA BUS (8 bits)"
    Set shp = DgShape(ws, "bus_addr", 1, 8, 224, 330, 10, "")
    DgLabel ws, "addr", 8, 235, 200, "ADDRESS BUS (9 bits)"

    ' bottom row: memory, keyboard, LCD
    DgWire ws, "MEM", 210, 204, 4, 46
    DgBuffer ws, "MEM", 212, 236, False
    DgWire ws, "MEMaddr", 290, 234, 4, 16
    DgBox ws, "MEM", 150, 250, 190, 40, "MEMORY" & vbLf & "cache L1/L2 + RAM (virtual)"
    DgWire ws, "KBD", 418, 204, 4, 46
    DgBuffer ws, "KBD", 420, 236, False
    DgBox ws, "KBD", 380, 250, 80, 40, "KEYBOARD" & vbLf & "queue 8 B"
    DgWire ws, "LCD", 528, 204, 4, 46
    DgBox ws, "LCD", 490, 250, 80, 40, "LCD 8x1" & vbLf & "buffer 8 B"

    ' control bus with lamps
    Set shp = DgShape(ws, "bus_ctl", 1, 8, 306, 604, 6, "")
    DgLabel ws, "ctl", 8, 294, 200, "CONTROL BUS"
    comps = Split("READ WRITE IO CLOCK", " ")
    For i = 0 To 3
        Set shp = DgShape(ws, "lamp_" & comps(i), 9, 40 + i * 90, 318, 70, 22, comps(i))
        shp.Line.Visible = msoTrue
        shp.Line.ForeColor.RGB = RGB(150, 156, 166)
    Next i

    ' legend
    Set shp = DgShape(ws, "leg1", 5, 400, 318, 14, 14, ""): shp.Fill.ForeColor.RGB = ColEnable()
    DgLabel ws, "leg1t", 416, 319, 80, "Enable (drives bus)"
    Set shp = DgShape(ws, "leg2", 5, 400, 338, 14, 14, ""): shp.Fill.ForeColor.RGB = ColLoad()
    DgLabel ws, "leg2t", 416, 339, 80, "Load (latches)"
    Set shp = DgShape(ws, "leg3", 7, 510, 318, 14, 12, ""): shp.Fill.ForeColor.RGB = ColHighZ(): shp.Line.Visible = msoTrue: shp.Line.ForeColor.RGB = RGB(110, 118, 132)
    DgLabel ws, "leg3t", 526, 319, 90, "High-Z (disconnected)"
    Set shp = DgShape(ws, "leg4", 1, 510, 342, 14, 8, ""): shp.Fill.ForeColor.RGB = ColBusOn()
    DgLabel ws, "leg4t", 526, 339, 90, "Bus carrying data"
    On Error GoTo 0
End Sub

Private Sub DgFill(ByVal ws As Worksheet, ByVal nm As String, ByVal color As Long)
    ws.Shapes(DG & nm).Fill.ForeColor.RGB = color
End Sub

Private Sub DgText(ByVal ws As Worksheet, ByVal nm As String, ByVal txt As String, ByVal fontColor As Long)
    With ws.Shapes(DG & nm).TextFrame.Characters
        .Text = txt
        .Font.Color = fontColor
    End With
End Sub

Private Sub PaintDiagram(ByVal inDrv As String, ByVal inLd As String, ByVal exDrv As String, ByVal exLd As String)
    Dim ws As Worksheet, comps() As String, i As Long, c As String, drives As Boolean, loads As Boolean
    Dim ioOut As Boolean, extOn As Boolean, addrOn As Boolean, v As String
    Set ws = GetSheet(SH_CPU)
    On Error Resume Next
    If ws.Shapes(DG & "frame") Is Nothing Then Exit Sub

    ioOut = Has("IO_W")
    extOn = ExternalActive()
    addrOn = Has("Mem_R") Or Has("Mem_W")

    ' registers: box color, buffer state, wire
    comps = Split("PC IR MAR MDR AX BX CX DX TMP", " ")
    For i = 0 To UBound(comps)
        c = comps(i)
        drives = InList(c, inDrv) Or (ioOut And InList(c, exDrv))
        loads = InList(c, inLd) Or (c = "MDR" And Has("Mem_R") And Has("Ld_MDR")) Or (c = "TMP" And Has("IO_R"))
        If drives Then
            DgFill ws, "box_" & c, RGB(214, 240, 224)
        ElseIf loads Then
            DgFill ws, "box_" & c, RGB(214, 228, 247)
        Else
            DgFill ws, "box_" & c, RGB(255, 255, 255)
        End If
        If c = "MAR" Then
            v = Hex3(MAR) & "h"
        ElseIf c = "PC" Then
            v = Hex3(PC) & "h"
        ElseIf c = "IR" Then
            v = Hex2(IR) & "h"
        ElseIf c = "TMP" Then
            v = Hex2(Core_TMP()) & "h"
        Else
            v = DriverValue(c)
        End If
        DgText ws, "box_" & c, c & vbLf & v, RGB(31, 42, 74)
        If c <> "MAR" Then
            DgFill ws, "buf_" & c, IIf(drives, ColEnable(), ColHighZ())
        End If
        If drives Then
            DgFill ws, "wire_" & c, ColEnable()
        ElseIf loads And c <> "MDR" Then
            DgFill ws, "wire_" & c, ColLoad()
        ElseIf loads And c = "MDR" And Not Has("Mem_R") Then
            DgFill ws, "wire_" & c, ColLoad()
        Else
            DgFill ws, "wire_" & c, ColBusOff()
        End If
    Next i

    ' internal bus
    If CountWords(inDrv) > 1 Then
        DgFill ws, "bus_int", ColConflict()
    ElseIf inDrv <> "" Or inLd <> "" Or ioOut Then
        DgFill ws, "bus_int", ColBusOn()
    Else
        DgFill ws, "bus_int", ColBusOff()
    End If

    ' ALU
    If AluOp() <> "" Then
        DgFill ws, "box_ALU", RGB(255, 217, 102)
        DgText ws, "box_ALU", "ALU  " & AluOp(), RGB(31, 42, 74)
        DgFill ws, "wire_ALU", ColLoad()
    Else
        DgFill ws, "box_ALU", RGB(255, 255, 255)
        DgText ws, "box_ALU", "ALU", RGB(31, 42, 74)
        DgFill ws, "wire_ALU", ColBusOff()
    End If

    ' bus interface (MDR <-> external data bus) and MAR -> address bus
    DgFill ws, "buf_MDRext", IIf(Has("Mem_W") Or ioOut, ColEnable(), ColHighZ())
    DgFill ws, "wire_MDRext", IIf(extOn And Not Has("IO_R"), ColBusOn(), ColBusOff())
    DgFill ws, "wire_MARaddr", IIf(addrOn, ColEnable(), ColBusOff())

    ' external data bus and address bus
    If CountWords(exDrv) > 1 Then
        DgFill ws, "bus_ext", ColConflict()
    Else
        DgFill ws, "bus_ext", IIf(extOn, ColBusOn(), ColBusOff())
    End If
    DgFill ws, "bus_addr", IIf(addrOn, ColBusOn(), ColBusOff())
    If extOn And DataBus >= 0 Then
        DgText ws, "lbl_ext", "EXTERNAL DATA BUS (8 bits)   " & exDrv & " -> " & Hex2(DataBus) & "h -> " & exLd, RGB(31, 42, 74)
    Else
        DgText ws, "lbl_ext", "EXTERNAL DATA BUS (8 bits)   High-Z", RGB(31, 42, 74)
    End If
    If addrOn Then
        DgText ws, "lbl_addr", "ADDRESS BUS (9 bits)   MAR -> " & Hex3(AddressBus) & "h", RGB(31, 42, 74)
    Else
        DgText ws, "lbl_addr", "ADDRESS BUS (9 bits)   High-Z", RGB(31, 42, 74)
    End If
    If inDrv <> "" And CountWords(inDrv) = 1 Then
        DgText ws, "lbl_int", "INTERNAL BUS (CPU)   " & inDrv & " -> " & DriverValue(inDrv) & IIf(inLd <> "", " -> " & inLd, ""), RGB(31, 42, 74)
    ElseIf inLd <> "" Then
        DgText ws, "lbl_int", "INTERNAL BUS (CPU)   ALU -> " & inLd, RGB(31, 42, 74)
    Else
        DgText ws, "lbl_int", "INTERNAL BUS (CPU)   High-Z", RGB(31, 42, 74)
    End If

    ' memory, keyboard, LCD
    DgFill ws, "buf_MEM", IIf(Has("Mem_R"), ColEnable(), ColHighZ())
    DgFill ws, "wire_MEM", IIf(Has("Mem_R"), ColEnable(), IIf(Has("Mem_W"), ColLoad(), ColBusOff()))
    DgFill ws, "wire_MEMaddr", IIf(addrOn, ColBusOn(), ColBusOff())
    DgFill ws, "box_MEM", IIf(Has("Mem_R"), RGB(214, 240, 224), IIf(Has("Mem_W"), RGB(214, 228, 247), RGB(255, 255, 255)))
    If Has("Mem_R") Or Has("Mem_W") Then
        DgText ws, "box_MEM", "MEMORY" & vbLf & Cache_LastResult(), RGB(31, 42, 74)
    Else
        DgText ws, "box_MEM", "MEMORY" & vbLf & "cache L1/L2 + RAM (virtual)", RGB(31, 42, 74)
    End If
    DgFill ws, "buf_KBD", IIf(Has("IO_R"), ColEnable(), ColHighZ())
    DgFill ws, "wire_KBD", IIf(Has("IO_R"), ColEnable(), ColBusOff())
    DgFill ws, "box_KBD", IIf(Has("IO_R"), RGB(214, 240, 224), RGB(255, 255, 255))
    DgFill ws, "wire_LCD", IIf(Has("Ld_LCD"), ColLoad(), ColBusOff())
    DgFill ws, "box_LCD", IIf(Has("Ld_LCD"), RGB(214, 228, 247), RGB(255, 255, 255))

    ' control bus lamps
    DgFill ws, "lamp_READ", IIf(Has("Mem_R") Or Has("IO_R"), ColBusOn(), RGB(255, 255, 255))
    DgFill ws, "lamp_WRITE", IIf(Has("Mem_W") Or Has("IO_W"), ColBusOn(), RGB(255, 255, 255))
    DgFill ws, "lamp_IO", IIf(Has("IO_R") Or Has("IO_W"), ColBusOn(), RGB(255, 255, 255))
    DgFill ws, "lamp_CLOCK", IIf(ControlWord <> "", ColBusOn(), RGB(255, 255, 255))
    DgFill ws, "bus_ctl", IIf(ControlWord <> "", ColBusOn(), ColBusOff())
    On Error GoTo 0
End Sub
