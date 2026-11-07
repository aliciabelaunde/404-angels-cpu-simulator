Attribute VB_Name = "modAssembler"
'==============================================================================
'  404 Angels - CPU Simulator
'  modAssembler : program catalog, editor, assembler (ISA v3) and loader
'  OWNER        : Paola
'
'  Syntax (one line per cell):
'     LOOP:  ADD AX, BX       ; a label ends with ':', a comment starts with ';'
'            MOV CX, 10       ; decimal number
'            LOAD AX, [1E0h]  ; hexadecimal number ends with h
'            ORG 100h         ; the next byte goes to address 100h
'     N:     DB 5, 0Ah        ; data bytes
'
'  Two passes:
'     pass 1: find the address of every line and the value of every label
'     pass 2: translate every instruction into bytes (docs/ISA.md)
'==============================================================================
Option Explicit

Private Const MAX_LINES As Long = 100
Private Const ED_ROW As Long = 11       ' first editor row
Private Const CAT_ROW As Long = 6       ' first catalog row

' --- Program image (result of the last assembly) -------------------------------
Private Image(0 To 511) As Long         ' bytes of the 512-byte virtual space
Private Used(0 To 511) As Boolean       ' True if an instruction or DB uses that byte
Private ImageLength As Long             ' last used address + 1
Private ErrorCount As Long

' --- One entry per source line -------------------------------------------------
Private LineMnem(1 To 100) As String    ' e.g. "ADD"
Private LineOps(1 To 100) As String     ' e.g. "AX,BX"
Private LineAddr(1 To 100) As Long      ' address of the line (-1 = empty line)
Private LineBytes(1 To 100) As String   ' e.g. "21"
Private LineError(1 To 100) As String   ' "" = no error

' --- Labels ----------------------------------------------------------------------
Private LabelName(1 To 100) As String
Private LabelAddr(1 To 100) As Long
Private LabelCount As Long

' --- Address being filled during pass 2 ---------------------------------------
Private EmitAddr As Long
Private EmitLine As Long

'==============================================================================
'  1. SMALL HELPERS
'==============================================================================
' Register code from docs/ISA.md: AX = 0, BX = 1, CX = 2, DX = 3 (-1 = not a register)
Private Function RegCode(ByVal t As String) As Long
    Select Case t
        Case "AX": RegCode = 0
        Case "BX": RegCode = 1
        Case "CX": RegCode = 2
        Case "DX": RegCode = 3
        Case Else: RegCode = -1
    End Select
End Function

' Position of the operation in the ALU group: MOV = 0, ADD = 1 ... XOR = 6
Private Function AluIndex(ByVal m As String) As Long
    Select Case m
        Case "MOV": AluIndex = 0
        Case "ADD": AluIndex = 1
        Case "SUB": AluIndex = 2
        Case "CMP": AluIndex = 3
        Case "AND": AluIndex = 4
        Case "OR": AluIndex = 5
        Case "XOR": AluIndex = 6
        Case Else: AluIndex = -1
    End Select
End Function

' Position of the jump: JMP = 0, JZ = 1 ... JNS = 6
Private Function JumpIndex(ByVal m As String) As Long
    Select Case m
        Case "JMP": JumpIndex = 0
        Case "JZ": JumpIndex = 1
        Case "JNZ": JumpIndex = 2
        Case "JC": JumpIndex = 3
        Case "JNC": JumpIndex = 4
        Case "JS": JumpIndex = 5
        Case "JNS": JumpIndex = 6
        Case Else: JumpIndex = -1
    End Select
End Function

' "12" -> 12   "0Ch" -> 12   "1E0h" -> 480   (ok = False if it is not a number)
Private Function ParseNumber(ByVal t As String, ByRef ok As Boolean) As Long
    Dim i As Long, c As String, v As Long, digits As String
    ok = False
    If t = "" Then Exit Function
    If Right(t, 1) = "H" Then
        digits = Left(t, Len(t) - 1)
        If digits = "" Then Exit Function
        For i = 1 To Len(digits)
            c = Mid(digits, i, 1)
            If InStr("0123456789ABCDEF", c) = 0 Then Exit Function
            v = v * 16 + InStr("0123456789ABCDEF", c) - 1
            If v > 65535 Then Exit Function
        Next i
    Else
        For i = 1 To Len(t)
            c = Mid(t, i, 1)
            If InStr("0123456789", c) = 0 Then Exit Function
            v = v * 10 + (InStr("0123456789", c) - 1)
            If v > 65535 Then Exit Function
        Next i
    End If
    ok = True
    ParseNumber = v
End Function

Private Function FindLabel(ByVal nm As String) As Long
    Dim i As Long
    FindLabel = -1
    For i = 1 To LabelCount
        If LabelName(i) = nm Then FindLabel = i
    Next i
End Function

' A label or a number
Private Function ValueOf(ByVal t As String, ByRef ok As Boolean) As Long
    Dim k As Long
    k = FindLabel(t)
    If k > 0 Then
        ok = True
        ValueOf = LabelAddr(k)
    Else
        ValueOf = ParseNumber(t, ok)
    End If
End Function

' "[100h]" or "[DATA]" -> address (ok = False if it is not a valid address)
Private Function MemOperand(ByVal t As String, ByRef ok As Boolean) As Long
    Dim v As Long
    ok = False
    If Left(t, 1) <> "[" Or Right(t, 1) <> "]" Then Exit Function
    v = ValueOf(Mid(t, 2, Len(t) - 2), ok)
    If v > 511 Then ok = False
    MemOperand = v
End Function

Private Sub AddError(ByVal i As Long, ByVal msg As String)
    If LineError(i) = "" Then
        LineError(i) = msg
        ErrorCount = ErrorCount + 1
    End If
End Sub

' Puts one byte in the image (pass 2)
Private Sub Emit(ByVal v As Long)
    If EmitAddr > 511 Then
        AddError EmitLine, "Address out of range (max 1FFh)"
        Exit Sub
    End If
    If Used(EmitAddr) Then
        AddError EmitLine, "Overlaps another line at " & Hex3(EmitAddr) & "h"
    End If
    Image(EmitAddr) = v And &HFF
    Used(EmitAddr) = True
    If LineBytes(EmitLine) <> "" Then LineBytes(EmitLine) = LineBytes(EmitLine) & " "
    LineBytes(EmitLine) = LineBytes(EmitLine) & Hex2(v)
    EmitAddr = EmitAddr + 1
    If EmitAddr > ImageLength Then ImageLength = EmitAddr
End Sub

'==============================================================================
'  2. PASS 1: addresses and labels
'==============================================================================
' Size in bytes of one instruction (0 = unknown instruction)
Private Function InstrSize(ByVal m As String, ByVal op1 As String, ByVal op2 As String) As Long
    Select Case m
        Case "NOP", "HLT", "INC", "DEC", "NOT", "IN", "OUT"
            InstrSize = 1
        Case "MOV", "ADD", "SUB", "CMP", "AND", "OR", "XOR"
            If RegCode(op2) >= 0 Then InstrSize = 1 Else InstrSize = 2
        Case "LOAD"
            If op2 = "[BX]" Then InstrSize = 1 Else InstrSize = 3
        Case "STORE"
            If op1 = "[BX]" Then InstrSize = 1 Else InstrSize = 3
        Case "JMP", "JZ", "JNZ", "JC", "JNC", "JS", "JNS"
            InstrSize = 3
        Case Else
            InstrSize = 0
    End Select
End Function

Private Sub Pass1(ByRef lines() As String, ByVal n As Long)
    Dim i As Long, s As String, p As Long, lbl As String, addr As Long
    Dim ops() As String, op1 As String, op2 As String, size As Long, v As Long, ok As Boolean
    addr = 0
    For i = 1 To n
        LineMnem(i) = "": LineOps(i) = "": LineBytes(i) = "": LineError(i) = ""
        LineAddr(i) = -1

        ' clean the line: remove the comment, tabs and extra spaces
        s = lines(i)
        p = InStr(s, ";")
        If p > 0 Then s = Left(s, p - 1)
        s = UCase(Trim(Replace(s, vbTab, " ")))

        ' label
        p = InStr(s, ":")
        If p > 0 Then
            lbl = Trim(Left(s, p - 1))
            s = Trim(Mid(s, p + 1))
            If lbl = "" Or InStr(lbl, " ") > 0 Then
                AddError i, "Invalid label"
            ElseIf FindLabel(lbl) > 0 Then
                AddError i, "Label " & lbl & " is repeated"
            ElseIf RegCode(lbl) >= 0 Then
                AddError i, "A register cannot be a label"
            Else
                LabelCount = LabelCount + 1
                LabelName(LabelCount) = lbl
                LabelAddr(LabelCount) = addr
            End If
            LineAddr(i) = addr
        End If

        If s <> "" Then
            ' mnemonic and operands ("MOV AX, 5" -> "MOV" and "AX,5")
            p = InStr(s, " ")
            If p > 0 Then
                LineMnem(i) = Left(s, p - 1)
                LineOps(i) = Replace(Mid(s, p + 1), " ", "")
            Else
                LineMnem(i) = s
            End If
            op1 = "": op2 = ""
            If LineOps(i) <> "" Then
                ops = Split(LineOps(i), ",")
                op1 = ops(0)
                If UBound(ops) >= 1 Then op2 = ops(1)
            End If

            LineAddr(i) = addr
            If LineMnem(i) = "ORG" Then
                v = ParseNumber(op1, ok)
                If ok = False Or v > 511 Then
                    AddError i, "ORG needs an address from 0 to 1FFh"
                Else
                    addr = v
                    LineAddr(i) = addr
                End If
            ElseIf LineMnem(i) = "DB" Then
                If LineOps(i) = "" Then
                    AddError i, "DB needs at least one value"
                Else
                    addr = addr + UBound(Split(LineOps(i), ",")) + 1
                End If
            Else
                size = InstrSize(LineMnem(i), op1, op2)
                If size = 0 Then
                    AddError i, "Unknown instruction " & LineMnem(i)
                Else
                    addr = addr + size
                End If
            End If
        End If
    Next i
End Sub

'==============================================================================
'  3. PASS 2: instructions -> bytes (ISA v3)
'==============================================================================
Private Sub EncodeLine(ByVal i As Long)
    Dim m As String, ops() As String, n As Long, a As String, b As String
    Dim r1 As Long, r2 As Long, k As Long, v As Long, ok As Boolean, j As Long

    m = LineMnem(i)
    If m = "" Or m = "ORG" Then Exit Sub
    EmitLine = i
    EmitAddr = LineAddr(i)
    If LineOps(i) = "" Then
        n = 0
    Else
        ops = Split(LineOps(i), ",")
        n = UBound(ops) + 1
    End If
    If n >= 1 Then a = ops(0)
    If n >= 2 Then b = ops(1)

    Select Case m
        Case "NOP", "HLT"
            If n <> 0 Then
                AddError i, m & " has no operands"
            ElseIf m = "NOP" Then
                Emit 0
            Else
                Emit 255
            End If

        Case "MOV", "ADD", "SUB", "CMP", "AND", "OR", "XOR"
            r1 = RegCode(a)
            r2 = RegCode(b)
            k = AluIndex(m)
            If n <> 2 Then
                AddError i, m & " needs 2 operands"
            ElseIf r1 < 0 Then
                AddError i, "First operand must be AX, BX, CX or DX"
            ElseIf r2 >= 0 Then
                Emit 16 * (k + 1) + r1 * 4 + r2           ' reg, reg (1 byte)
            Else
                v = ValueOf(b, ok)
                If ok = False Then
                    AddError i, "Invalid number or label: " & b
                ElseIf v > 255 Then
                    AddError i, "Value must be 0-255"
                Else
                    Emit 128 + k * 4 + r1                 ' reg, imm8 (2 bytes): 80h + ...
                    Emit v
                End If
            End If

        Case "INC", "DEC", "NOT"
            r1 = RegCode(a)
            If m = "INC" Then k = 0
            If m = "DEC" Then k = 1
            If m = "NOT" Then k = 2
            If n <> 1 Then
                AddError i, m & " needs 1 operand"
            ElseIf r1 < 0 Then
                AddError i, "Operand must be AX, BX, CX or DX"
            Else
                Emit 176 + k * 4 + r1                     ' B0h + ...
            End If

        Case "LOAD"
            r1 = RegCode(a)
            If n <> 2 Then
                AddError i, "Use: LOAD reg, [address]"
            ElseIf r1 < 0 Then
                AddError i, "First operand must be AX, BX, CX or DX"
            ElseIf b = "[BX]" Then
                Emit 168 + r1                             ' LOAD reg, [BX]: A8h + reg
            Else
                v = MemOperand(b, ok)
                If ok = False Then
                    AddError i, "Invalid address: " & b
                Else
                    Emit 160 + r1                         ' LOAD reg, [addr9]: A0h + reg
                    Emit v Mod 256
                    Emit v \ 256
                End If
            End If

        Case "STORE"
            r2 = RegCode(b)
            If n <> 2 Then
                AddError i, "Use: STORE [address], reg"
            ElseIf r2 < 0 Then
                AddError i, "Second operand must be AX, BX, CX or DX"
            ElseIf a = "[BX]" Then
                Emit 172 + r2                             ' STORE [BX], reg: ACh + reg
            Else
                v = MemOperand(a, ok)
                If ok = False Then
                    AddError i, "Invalid address: " & a
                Else
                    Emit 164 + r2                         ' STORE [addr9], reg: A4h + reg
                    Emit v Mod 256
                    Emit v \ 256
                End If
            End If

        Case "JMP", "JZ", "JNZ", "JC", "JNC", "JS", "JNS"
            If n <> 1 Then
                AddError i, m & " needs 1 operand (label or address)"
            Else
                v = ValueOf(a, ok)
                If ok = False Then
                    AddError i, "Unknown label: " & a
                ElseIf v > 511 Then
                    AddError i, "Address out of range (max 1FFh)"
                Else
                    Emit 192 + JumpIndex(m)               ' C0h + jump
                    Emit v Mod 256
                    Emit v \ 256
                End If
            End If

        Case "IN"
            r1 = RegCode(a)
            If n <> 2 Or b <> "[KBD]" Then
                AddError i, "Use: IN reg, [KBD]"
            ElseIf r1 < 0 Then
                AddError i, "First operand must be AX, BX, CX or DX"
            Else
                Emit 208 + r1                             ' D0h + reg
            End If

        Case "OUT"
            r2 = RegCode(b)
            If n <> 2 Or a <> "[LCD]" Then
                AddError i, "Use: OUT [LCD], reg"
            ElseIf r2 < 0 Then
                AddError i, "Second operand must be AX, BX, CX or DX"
            Else
                Emit 212 + r2                             ' D4h + reg
            End If

        Case "DB"
            For j = 0 To n - 1
                v = ValueOf(ops(j), ok)
                If ok = False Or v > 255 Then
                    AddError i, "DB value must be 0-255: " & ops(j)
                    Exit For
                End If
                Emit v
            Next j
    End Select
End Sub

'==============================================================================
'  4. ASSEMBLE: returns the number of errors (0 = the program is correct)
'==============================================================================
Public Function Asm_AssembleLines(ByRef lines() As String, ByVal n As Long) As Long
    Dim i As Long
    If n > MAX_LINES Then n = MAX_LINES
    For i = 0 To 511
        Image(i) = 0
        Used(i) = False
    Next i
    ImageLength = 0
    ErrorCount = 0
    LabelCount = 0
    Pass1 lines, n
    For i = 1 To n
        If LineError(i) = "" Then EncodeLine i
    Next i
    Asm_AssembleLines = ErrorCount
End Function

' Byte of the last assembled program (used by the loader and the tests)
Public Function Asm_Byte(ByVal addr As Long) As Long
    Asm_Byte = Image(addr And &H1FF)
End Function

'==============================================================================
'  5. SHEET "Programs"
'     A4:F7     catalog: the programs saved on "disk" (columns K, L, M)
'     A9:B9     result of the last assembly
'     A11:E110  editor: A line, B source code, C address, D bytes, E status
'     K5:M105   catalog programs (one line per cell, first line = name)
'
'  The programs are NOT inside the VBA code: they are written in the cells.
'==============================================================================

Private Sub AsmHeader(ByVal c As Range, ByVal txt As String)
    c.Value = txt
    c.Font.Bold = True
    c.Font.Color = RGB(255, 255, 255)
    c.Interior.Color = RGB(64, 64, 64)
    c.HorizontalAlignment = xlCenter
End Sub

Private Sub AsmButton(ByVal ws As Worksheet, ByVal nm As String, ByVal where As Range, _
                      ByVal txt As String, ByVal macro As String, ByVal color As Long)
    Dim shp As Shape
    On Error Resume Next
    ws.Shapes("btnAsm" & nm).Delete
    Set shp = ws.Shapes.AddShape(5, where.Left + 2, where.Top + 1, where.Width - 4, where.Height - 2)
    shp.Name = "btnAsm" & nm
    shp.Fill.ForeColor.RGB = color
    shp.Line.Visible = msoFalse
    shp.OnAction = macro
    shp.TextFrame.Characters.Text = txt
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.Characters.Font.Size = 9
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp.TextFrame.VerticalAlignment = xlVAlignCenter
    On Error GoTo 0
End Sub

' The name of a catalog program is its first line, without the ';'
Private Function ProgramName(ByVal ws As Worksheet, ByVal k As Long) As String
    Dim s As String
    s = Trim(CStr(ws.Cells(CAT_ROW, 10 + k).Value))
    If Left(s, 1) = ";" Then s = Trim(Mid(s, 2))
    If s = "" Then s = "(empty: paste the program in column " & Chr(Asc("J") + k) & ")"
    ProgramName = s
End Function

Public Sub Asm_SetupSheet()
    Dim ws As Worksheet, k As Long
    Set ws = GetSheet(SH_PROGRAMS)
    ' Only the layout is created. Columns K:M (the catalog) are never cleared.
    ws.Range("A1:I110").Clear
    ws.Range("A1:I110").Font.Name = "Consolas"
    ws.Range("A1:I110").Font.Size = 10

    ws.Range("A1").Value = TEAM_NAME & " - Program directory"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Catalog (disk) -> OPEN -> editor -> ASSEMBLE -> LOAD PROGRAM -> virtual memory"

    ' catalog list
    AsmHeader ws.Range("A4:F4"), "CATALOG"
    For k = 1 To 3
        ws.Cells(4 + k, 1).Value = k
        ws.Cells(4 + k, 1).HorizontalAlignment = xlCenter
        AsmButton ws, "Open" & k, ws.Cells(4 + k, 6), "OPEN " & k, "BtnOpenProgram" & k, RGB(47, 117, 181)
    Next k

    ' result line and main buttons
    ws.Range("A9").Value = "Result"
    ws.Range("A9").Font.Bold = True
    AsmButton ws, "Assemble", ws.Range("G11:H12"), "ASSEMBLE", "BtnAssemble", RGB(191, 143, 0)
    AsmButton ws, "Load", ws.Range("G14:H15"), "LOAD PROGRAM", "BtnLoadProgram", RGB(84, 130, 53)

    ' editor
    AsmHeader ws.Range("A10"), "Line"
    AsmHeader ws.Range("B10"), "Source code (edit here)"
    AsmHeader ws.Range("C10"), "Address"
    AsmHeader ws.Range("D10"), "Bytes"
    AsmHeader ws.Range("E10"), "Status"
    For k = 1 To MAX_LINES
        ws.Cells(ED_ROW + k - 1, 1).Value = k
    Next k
    ws.Range("A11:A110").HorizontalAlignment = xlCenter
    ws.Range("A11:A110").Font.Color = RGB(128, 128, 128)
    ws.Range("B11:B110").Interior.Color = RGB(255, 255, 240)
    ws.Range("B11:B110").NumberFormat = "@"
    ws.Range("C11:D110").NumberFormat = "@"
    ws.Range("C11:C110").HorizontalAlignment = xlCenter

    ' catalog columns (only the headers; the programs stay as they are)
    For k = 1 To 3
        AsmHeader ws.Cells(CAT_ROW - 1, 10 + k), "PROGRAM " & k
        ws.Range(ws.Cells(CAT_ROW, 10 + k), ws.Cells(CAT_ROW + MAX_LINES - 1, 10 + k)).NumberFormat = "@"
    Next k

    ws.Columns("A").ColumnWidth = 6
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 9
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 40
    ws.Columns("F:I").ColumnWidth = 9
    ws.Columns("J").ColumnWidth = 2
    ws.Columns("K:M").ColumnWidth = 34

    ShowCatalog ws
End Sub

Private Sub ShowCatalog(ByVal ws As Worksheet)
    Dim k As Long
    For k = 1 To 3
        ws.Cells(4 + k, 2).Value = ProgramName(ws, k)
    Next k
End Sub

' OPEN 1 / 2 / 3: copies a catalog program into the editor
Private Sub OpenProgram(ByVal k As Long)
    Dim ws As Worksheet, i As Long
    Set ws = GetSheet(SH_PROGRAMS)
    ws.Range("B11:E110").ClearContents      ' empty the editor first
    For i = 1 To MAX_LINES
        If CStr(ws.Cells(CAT_ROW + i - 1, 10 + k).Value) <> "" Then
            ws.Cells(ED_ROW + i - 1, 2).Value = ws.Cells(CAT_ROW + i - 1, 10 + k).Value
        End If
    Next i
    ws.Range("B11:E110").Interior.Color = RGB(255, 255, 240)
    ws.Range("B9").Value = "Program " & k & " opened in the editor: " & ProgramName(ws, k)
    ws.Range("B9").Interior.Color = RGB(221, 235, 247)
    ShowCatalog ws
End Sub

Public Sub BtnOpenProgram1()
    OpenProgram 1
End Sub

Public Sub BtnOpenProgram2()
    OpenProgram 2
End Sub

Public Sub BtnOpenProgram3()
    OpenProgram 3
End Sub

' Reads the editor, assembles it and shows address, bytes and errors per line
Private Function AssembleEditor() As Long
    Dim ws As Worksheet, lines() As String, i As Long, errors As Long
    Set ws = GetSheet(SH_PROGRAMS)
    ReDim lines(1 To MAX_LINES)
    For i = 1 To MAX_LINES
        lines(i) = CStr(ws.Cells(ED_ROW + i - 1, 2).Value)
    Next i

    errors = Asm_AssembleLines(lines, MAX_LINES)

    For i = 1 To MAX_LINES
        If LineAddr(i) >= 0 And Trim(lines(i)) <> "" Then
            ws.Cells(ED_ROW + i - 1, 3).Value = Hex3(LineAddr(i)) & "h"
        Else
            ws.Cells(ED_ROW + i - 1, 3).Value = ""
        End If
        ws.Cells(ED_ROW + i - 1, 4).Value = LineBytes(i)
        If LineError(i) <> "" Then
            ws.Cells(ED_ROW + i - 1, 5).Value = "ERROR: " & LineError(i)
            ws.Range(ws.Cells(ED_ROW + i - 1, 2), ws.Cells(ED_ROW + i - 1, 5)).Interior.Color = RGB(255, 199, 206)
        Else
            ws.Cells(ED_ROW + i - 1, 5).Value = ""
            ws.Range(ws.Cells(ED_ROW + i - 1, 2), ws.Cells(ED_ROW + i - 1, 5)).Interior.Color = RGB(255, 255, 240)
        End If
    Next i

    If errors = 0 Then
        ws.Range("B9").Value = "OK: " & ImageLength & " bytes (000h-" & Hex3(ImageLength - 1) & "h), no errors."
        ws.Range("B9").Interior.Color = RGB(198, 239, 206)
    Else
        ws.Range("B9").Value = errors & " error(s). Fix the red lines and press ASSEMBLE again."
        ws.Range("B9").Interior.Color = RGB(255, 199, 206)
    End If
    AssembleEditor = errors
End Function

Public Sub BtnAssemble()
    Call AssembleEditor
End Sub

'==============================================================================
'  6. LOADER: assembles the editor and puts the program in virtual memory
'==============================================================================
Public Sub BtnLoadProgram()
    Dim ws As Worksheet, img() As Long, i As Long
    Set ws = GetSheet(SH_PROGRAMS)
    If AssembleEditor() > 0 Then Exit Sub
    If ImageLength = 0 Then
        ws.Range("B9").Value = "The editor is empty: open a program first."
        Exit Sub
    End If

    ReDim img(0 To ImageLength - 1)
    For i = 0 To ImageLength - 1
        img(i) = Image(i)
    Next i

    ' same order as docs/INTERFACES.md
    VM_Reset
    VM_LoadImage img, ImageLength
    Cache_Reset
    IO_Reset
    Bus_Reset
    Core_Reset

    Core_Paint
    Bus_Paint
    Cache_Paint
    VM_Paint
    IO_Paint

    ws.Range("B9").Value = "Loaded " & ImageLength & " bytes into virtual memory (swap). Go to sheet CPU and press STEP or RUN."
    ws.Range("B9").Interior.Color = RGB(198, 239, 206)
End Sub
