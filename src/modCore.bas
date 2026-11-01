Attribute VB_Name = "modCore"
'==============================================================================
'  404 Angels - CPU Simulator
'  modCore : CPU registers, ALU, decoder, control unit, buttons
'  OWNER   : Alicia
'
'  Based on the Parcial 1 CPU (ISA v2), extended to ISA v3 (docs/ISA.md):
'    - 4 data registers AX, BX, CX, DX (8 bits)
'    - PC and MAR are 9-bit (virtual address space 000h-1FFh)
'    - LOAD/STORE [addr9] and jumps carry 2 address bytes (3-byte instructions)
'    - LOAD/STORE [BX] use effective address 100h + BX (data area)
'    - IN reg,[KBD] (D0-D3) and OUT [LCD],reg (D4-D7), flags unchanged
'    - every memory access goes through the bus: Bus_MemRead / Bus_MemWrite
'    - every micro-step emits its control word: Bus_ControlWord
'
'  One clock pulse = one micro-operation.
'    FETCH   : 3 pulses   (MAR <- PC ; MDR <- M[MAR] ; IR <- MDR, PC+1)
'    DECODE  : 1 pulse + 3 per operand byte
'    EXECUTE : 1 pulse (2 for LOAD/STORE: address + transfer)
'    STORE   : 1 pulse (write-back)
'
'  Sheet CPU layout:
'    A3:B11  registers        C3:D6  flags        C8:D11 counters, status, delay
'    A13:D17 buttons          A19:D25 control unit
'    A27:D46 micro-operation log (full history in sheet Log)
'    F3:N27  bus panel (modBus, created by Bus_SetupSheet)
'==============================================================================
Option Explicit

' --- Registers (public so other panels can display them) --------------------
Public PC As Long, IR As Long, MAR As Long, MDR As Long
Public AX As Long, BX As Long, CX As Long, DX As Long
Public ZF As Long, CF As Long, SF As Long
Public OPV As Long               ' operand value (imm8 or addr9), -1 = none
Public Halted As Boolean

' --- Control unit state --------------------------------------------------------
Private TMP As Long
Private OpLo As Long, OpHi As Long
Private Phase As String          ' phase of the NEXT micro-step
Private Micro As Long            ' micro-step inside the phase
Private LastPhase As String      ' phase of the last executed micro-step
Private InstrCount As Long
Private Pulses As Long
Private JumpTaken As Boolean
Private InstrAddr As Long        ' address where the current instruction starts
Private Paused As Boolean
Private Running As Boolean

' --- Description of the last micro-step ---------------------------------------
Private Active As String         ' e.g. "|PC|MAR|"
Private RtlText As String
Private DetailText As String
Private AluText As String
Private CtrlWord As String

' --- Decoded instruction --------------------------------------------------------
Private Type TInstr
    Valid As Boolean
    Kind As String       ' NOP MOV LOAD STORE ALU UNARY JUMP IN OUT HLT
    Mnem As String
    Op1 As String        ' AX BX CX DX [ADDR] [BX] ADDR [LCD]
    Op2 As String        ' AX BX CX DX IMM [ADDR] [BX] [KBD]
    Size As Long         ' 1, 2 or 3 bytes
End Type

' --- Sheet layout -------------------------------------------------------------
Private Const LOG_ROW1 As Long = 28
Private Const LOG_ROWN As Long = 46

'==============================================================================
'  1. UTILITIES
'==============================================================================
Private Function Sh() As Worksheet
    Set Sh = GetSheet(SH_CPU)
End Function

Private Function B8(ByVal v As Long) As Long
    B8 = ((v Mod 256) + 256) Mod 256
End Function

Private Function A9(ByVal v As Long) As Long
    A9 = ((v Mod 512) + 512) Mod 512
End Function

Private Function HX(ByVal v As Long) As String
    HX = Hex2(v) & "h"
End Function

Private Function HA(ByVal v As Long) As String
    HA = Hex3(v) & "h"
End Function

Private Function Arrow() As String
    Arrow = " " & ChrW(8592) & " "
End Function

Private Sub Act(ByVal list As String)
    Active = Active & Replace(list, ",", "|") & "|"
End Sub

Private Function IsActive(ByVal comp As String) As Boolean
    IsActive = InStr(1, Active, "|" & comp & "|") > 0
End Function

Private Sub Signals(ByVal cw As String)
    CtrlWord = cw
    Bus_ControlWord cw
End Sub

'==============================================================================
'  2. DECODER (ISA v3, see docs/ISA.md)
'==============================================================================
Private Function RegName(ByVal c As Long) As String
    RegName = Choose((c And 3) + 1, "AX", "BX", "CX", "DX")
End Function

Private Function OpName(ByVal n As Long) As String
    OpName = Choose(n + 1, "MOV", "ADD", "SUB", "CMP", "AND", "OR", "XOR")
End Function

Private Function Decode(ByVal op As Long) As TInstr
    Dim d As TInstr
    d.Valid = True
    Select Case op
        Case &H0
            d.Kind = "NOP": d.Mnem = "NOP"
        Case &H10 To &H7F
            d.Mnem = OpName((op \ 16) - 1)
            d.Op1 = RegName((op \ 4) And 3)
            d.Op2 = RegName(op And 3)
        Case &H80 To &H9B
            d.Mnem = OpName((op - &H80) \ 4)
            d.Op1 = RegName(op And 3)
            d.Op2 = "IMM"
        Case &HA0 To &HA3
            d.Kind = "LOAD": d.Mnem = "LOAD": d.Op1 = RegName(op): d.Op2 = "[ADDR]"
        Case &HA4 To &HA7
            d.Kind = "STORE": d.Mnem = "STORE": d.Op1 = "[ADDR]": d.Op2 = RegName(op)
        Case &HA8 To &HAB
            d.Kind = "LOAD": d.Mnem = "LOAD": d.Op1 = RegName(op): d.Op2 = "[BX]"
        Case &HAC To &HAF
            d.Kind = "STORE": d.Mnem = "STORE": d.Op1 = "[BX]": d.Op2 = RegName(op)
        Case &HB0 To &HBB
            d.Kind = "UNARY"
            d.Mnem = Choose((op - &HB0) \ 4 + 1, "INC", "DEC", "NOT")
            d.Op1 = RegName(op)
        Case &HC0 To &HC6
            d.Kind = "JUMP"
            d.Mnem = Choose(op - &HC0 + 1, "JMP", "JZ", "JNZ", "JC", "JNC", "JS", "JNS")
            d.Op1 = "ADDR"
        Case &HD0 To &HD3
            d.Kind = "IN": d.Mnem = "IN": d.Op1 = RegName(op): d.Op2 = "[KBD]"
        Case &HD4 To &HD7
            d.Kind = "OUT": d.Mnem = "OUT": d.Op1 = "[LCD]": d.Op2 = RegName(op)
        Case &HFF
            d.Kind = "HLT": d.Mnem = "HLT"
        Case Else
            d.Valid = False: d.Kind = "INVALID": d.Mnem = "???"
    End Select
    If d.Kind = "" Then
        If d.Mnem = "MOV" Then
            d.Kind = "MOV"
        Else
            d.Kind = "ALU"
        End If
    End If
    If d.Op1 = "ADDR" Or d.Op1 = "[ADDR]" Or d.Op2 = "[ADDR]" Then
        d.Size = 3
    ElseIf d.Op2 = "IMM" Then
        d.Size = 2
    Else
        d.Size = 1
    End If
    Decode = d
End Function

Private Function OperandText(ByVal s As String, ByVal v As Long) As String
    Select Case s
        Case "IMM"
            If v < 0 Then OperandText = "imm8" Else OperandText = HX(v)
        Case "ADDR"
            If v < 0 Then OperandText = "addr9" Else OperandText = HA(v)
        Case "[ADDR]"
            If v < 0 Then OperandText = "[addr9]" Else OperandText = "[" & HA(v) & "]"
        Case Else
            OperandText = s
    End Select
End Function

Private Function InstrText(ByRef d As TInstr, ByVal v As Long) As String
    Dim s As String
    If Not d.Valid Then
        InstrText = "(invalid opcode)"
        Exit Function
    End If
    s = d.Mnem
    If d.Op1 <> "" Then s = s & " " & OperandText(d.Op1, v)
    If d.Op2 <> "" Then s = s & ", " & OperandText(d.Op2, v)
    InstrText = s
End Function

Private Function AddrMode(ByRef d As TInstr) As String
    If d.Op1 = "[ADDR]" Or d.Op2 = "[ADDR]" Then
        AddrMode = "direct"
    ElseIf d.Op1 = "[BX]" Or d.Op2 = "[BX]" Then
        AddrMode = "register indirect (100h + BX)"
    ElseIf d.Op2 = "IMM" Then
        AddrMode = "immediate"
    ElseIf d.Op1 = "ADDR" Then
        AddrMode = "absolute (jump)"
    ElseIf d.Kind = "IN" Or d.Kind = "OUT" Then
        AddrMode = "I/O port"
    ElseIf d.Op1 <> "" Then
        AddrMode = "register"
    Else
        AddrMode = "implicit"
    End If
End Function

' Disassembles one opcode (used by other panels, e.g. Memory / Programs)
Public Function Core_Mnemonic(ByVal op As Long) As String
    Dim d As TInstr
    d = Decode(op And &HFF)
    If d.Valid Then Core_Mnemonic = InstrText(d, -1) Else Core_Mnemonic = "data"
End Function

' Temporary register of the control unit (shown by the bus panel)
Public Function Core_TMP() As Long
    Core_TMP = TMP
End Function

' Size in bytes of the instruction that starts with this opcode
Public Function Core_InstrSize(ByVal op As Long) As Long
    Dim d As TInstr
    d = Decode(op And &HFF)
    Core_InstrSize = d.Size
End Function

'==============================================================================
'  3. ALU (flags like x86)
'     ADD: CF=1 on carry (>255)   SUB/CMP: CF=1 on borrow (a<b)
'     AND/OR/XOR: CF=0   INC/DEC: CF unchanged   NOT: flags unchanged
'     ZF=1 if result is 0   SF = bit 7 of the result
'==============================================================================
Private Function ALU(ByVal op As String, ByVal a As Long, ByVal b As Long) As Long
    Dim r As Long
    Select Case op
        Case "ADD"
            r = a + b
            CF = IIf(r > 255, 1, 0)
        Case "SUB", "CMP"
            r = a - b
            CF = IIf(a < b, 1, 0)
        Case "AND"
            r = a And b: CF = 0
        Case "OR"
            r = a Or b: CF = 0
        Case "XOR"
            r = a Xor b: CF = 0
        Case "INC"
            r = a + 1
        Case "DEC"
            r = a - 1
        Case "NOT"
            ALU = 255 - a
            AluText = "NOT " & HX(a) & " = " & HX(255 - a)
            Exit Function
    End Select
    r = B8(r)
    ZF = IIf(r = 0, 1, 0)
    SF = IIf((r And &H80) <> 0, 1, 0)
    ALU = r
    If op = "INC" Or op = "DEC" Then b = 1
    AluText = HX(a) & " " & op & " " & HX(b) & " = " & HX(r)
End Function

Private Function GetReg(ByVal n As String) As Long
    Select Case n
        Case "AX": GetReg = AX
        Case "BX": GetReg = BX
        Case "CX": GetReg = CX
        Case "DX": GetReg = DX
    End Select
End Function

Private Sub SetReg(ByVal n As String, ByVal v As Long)
    Select Case n
        Case "AX": AX = B8(v)
        Case "BX": BX = B8(v)
        Case "CX": CX = B8(v)
        Case "DX": DX = B8(v)
    End Select
End Sub

Private Function FlagsText() As String
    FlagsText = "ZF=" & ZF & " CF=" & CF & " SF=" & SF
End Function

'==============================================================================
'  4. CONTROL UNIT: one clock pulse = one micro-operation
'==============================================================================
Private Function MicroCount(ByVal f As String, ByRef d As TInstr) As Long
    Select Case f
        Case "FETCH"
            MicroCount = 3
        Case "DECODE"
            If d.Valid Then
                MicroCount = 1 + 3 * (d.Size - 1)
            Else
                MicroCount = 1
            End If
        Case "EXECUTE"
            If d.Kind = "LOAD" Or d.Kind = "STORE" Then
                MicroCount = 2
            Else
                MicroCount = 1
            End If
        Case "STORE"
            MicroCount = 1
    End Select
End Function

Private Sub MicroStep()
    Dim d As TInstr
    If Halted Then Exit Sub
    Active = "|": AluText = "": CtrlWord = ""
    LastPhase = Phase
    Pulses = Pulses + 1
    d = Decode(IR)
    Select Case Phase
        Case "FETCH": UFetch
        Case "DECODE": UDecode d
        Case "EXECUTE": UExecute d
        Case "STORE": UStore d
    End Select
    If Halted Then Exit Sub
    ' Sequencer: next micro-step / next phase
    d = Decode(IR)
    Micro = Micro + 1
    If Micro >= MicroCount(Phase, d) Then
        Micro = 0
        Select Case Phase
            Case "FETCH": Phase = "DECODE"
            Case "DECODE": Phase = "EXECUTE"
            Case "EXECUTE": Phase = "STORE"
            Case "STORE": Phase = "FETCH"
        End Select
    End If
End Sub

' ---- shared bus micro-operations ---------------------------------------------
Private Sub UMar(ByVal source As String, ByVal value As Long)
    MAR = A9(value)
    Signals "En_" & Replace(source, "(op)", "") & " Ld_MAR"
    Act source & ",MAR,BUS"
    RtlText = "MAR" & Arrow() & source
    DetailText = "MAR=" & HA(MAR)
End Sub

Private Sub URead(ByVal isInstr As Boolean)
    Signals "Mem_R Ld_MDR"
    MDR = Bus_MemRead(MAR, isInstr)
    Act "MAR,MDR,BUS"
    RtlText = "MDR" & Arrow() & "M[MAR]"
    DetailText = "M[" & HA(MAR) & "]=" & HX(MDR)
End Sub

' ---- PHASE 1: FETCH ---------------------------------------------------------
Private Sub UFetch()
    Dim d As TInstr
    Select Case Micro
        Case 0
            InstrCount = InstrCount + 1
            InstrAddr = PC
            OPV = -1: OpLo = 0: OpHi = 0
            UMar "PC", PC
        Case 1
            URead True
        Case 2
            Signals "En_MDR Ld_IR PC_Inc"
            IR = MDR
            PC = A9(PC + 1)
            d = Decode(IR)
            Act "MDR,IR,PC"
            RtlText = "IR" & Arrow() & "MDR ; PC" & Arrow() & "PC+1"
            DetailText = "IR=" & HX(IR) & " (" & d.Mnem & "), PC=" & HA(PC)
    End Select
End Sub

' ---- PHASE 2: DECODE --------------------------------------------------------
Private Sub UDecode(ByRef d As TInstr)
    Dim k As Long, part As Long
    If Micro = 0 Then
        Signals "UC_Decode"
        Act "IR,UC"
        If Not d.Valid Then
            Halted = True
            RtlText = "UC: invalid opcode, CPU halted"
            DetailText = "IR=" & HX(IR) & " at " & HA(InstrAddr)
            Exit Sub
        End If
        RtlText = "UC decodes " & HX(IR) & " = " & InstrText(d, -1)
        DetailText = AddrMode(d) & ", " & d.Size & " byte(s)"
        Exit Sub
    End If
    ' operand bytes: 3 micro-steps each (MAR <- PC ; MDR <- M[MAR] ; IR(op) <- MDR, PC+1)
    k = (Micro - 1) \ 3          ' 0 = first operand byte, 1 = second
    part = (Micro - 1) Mod 3
    Select Case part
        Case 0
            UMar "PC", PC
            RtlText = RtlText & " (operand byte " & (k + 1) & ")"
        Case 1
            URead True
        Case 2
            Signals "En_MDR Ld_IR PC_Inc"
            If k = 0 Then
                OpLo = MDR
            Else
                OpHi = MDR And 1
            End If
            If d.Size = 2 Then
                OPV = OpLo
            Else
                OPV = OpLo + 256 * OpHi
            End If
            PC = A9(PC + 1)
            Act "MDR,IR,PC"
            RtlText = "IR(op" & (k + 1) & ")" & Arrow() & "MDR ; PC" & Arrow() & "PC+1"
            DetailText = "IR=" & IRText() & " -> " & InstrText(d, OPV) & ", PC=" & HA(PC)
    End Select
End Sub

Private Function SourceVal(ByRef d As TInstr, ByRef srcName As String) As Long
    If d.Op2 = "IMM" Then
        SourceVal = OPV: srcName = "IR(op)"
    Else
        SourceVal = GetReg(d.Op2): srcName = d.Op2
    End If
End Function

' ---- PHASE 3: EXECUTE -------------------------------------------------------
Private Sub UExecute(ByRef d As TInstr)
    Dim a As Long, b As Long, nb As String
    Select Case d.Kind
        Case "NOP"
            Signals "-"
            Act "UC": RtlText = "NOP: no operation": DetailText = "-"
        Case "HLT"
            Signals "UC_Halt"
            Act "UC": RtlText = "UC: HLT, clock will stop": DetailText = "-"
        Case "MOV"
            TMP = SourceVal(d, nb)
            Signals "En_" & Replace(nb, "(op)", "") & " Ld_TMP"
            Act Replace(nb, "IR(op)", "IR") & ",TMP"
            RtlText = "TMP" & Arrow() & nb
            DetailText = "TMP=" & HX(TMP)
        Case "LOAD", "STORE"
            If Micro = 0 Then
                If d.Op1 = "[BX]" Or d.Op2 = "[BX]" Then
                    UMar "BX", &H100 + BX              ' effective address 100h + BX
                    RtlText = "MAR" & Arrow() & "100h + BX"
                Else
                    UMar "IR(op)", OPV
                    Act "IR"
                End If
            ElseIf d.Kind = "LOAD" Then
                URead False
            Else
                MDR = GetReg(d.Op2)
                Signals "En_" & d.Op2 & " Ld_MDR"
                Act d.Op2 & ",MDR"
                RtlText = "MDR" & Arrow() & d.Op2
                DetailText = "MDR=" & HX(MDR)
            End If
        Case "ALU", "UNARY"
            a = GetReg(d.Op1)
            If d.Kind = "ALU" Then
                b = SourceVal(d, nb)
                Signals "En_" & Replace(nb, "(op)", "") & " ALU_" & d.Mnem & " Ld_TMP"
            Else
                nb = "1"
                Signals "ALU_" & d.Mnem & " Ld_TMP"
            End If
            TMP = ALU(d.Mnem, a, b)
            Act d.Op1 & ",ALU,TMP,FLAGS"
            If d.Kind = "ALU" Then
                Act Replace(nb, "IR(op)", "IR")
            End If
            If d.Mnem = "CMP" Then
                RtlText = "ALU: " & d.Op1 & " - " & nb & " (compare)"
            ElseIf d.Mnem = "NOT" Then
                RtlText = "TMP" & Arrow() & "NOT " & d.Op1
            Else
                RtlText = "TMP" & Arrow() & d.Op1 & " " & d.Mnem & " " & nb
            End If
            DetailText = AluText & " | " & FlagsText()
        Case "JUMP"
            Select Case d.Mnem
                Case "JMP": JumpTaken = True
                Case "JZ": JumpTaken = (ZF = 1)
                Case "JNZ": JumpTaken = (ZF = 0)
                Case "JC": JumpTaken = (CF = 1)
                Case "JNC": JumpTaken = (CF = 0)
                Case "JS": JumpTaken = (SF = 1)
                Case "JNS": JumpTaken = (SF = 0)
            End Select
            Signals "UC_Test_Flags"
            Act "UC,FLAGS"
            RtlText = "UC tests the condition of " & d.Mnem
            DetailText = FlagsText() & IIf(JumpTaken, " -> jump TAKEN", " -> jump NOT taken")
        Case "IN"
            Signals "IO_R En_KBD Ld_TMP"
            TMP = B8(Bus_IoRead(PORT_KBD))
            Act "KBD,TMP,BUS"
            RtlText = "TMP" & Arrow() & "KBD"
            DetailText = "char=" & HX(TMP) & CharText(TMP) & " (flags unchanged)"
        Case "OUT"
            Signals "IO_W En_" & d.Op2 & " Ld_LCD"
            Bus_IoWrite PORT_LCD, GetReg(d.Op2)
            Act d.Op2 & ",LCD,BUS"
            RtlText = "LCD" & Arrow() & d.Op2
            DetailText = "char=" & HX(GetReg(d.Op2)) & CharText(GetReg(d.Op2)) & " (flags unchanged)"
    End Select
End Sub

Private Function CharText(ByVal v As Long) As String
    If v >= 32 And v < 127 Then CharText = " '" & Chr$(v) & "'" Else CharText = ""
End Function

' ---- PHASE 4: STORE / WRITE-BACK --------------------------------------------
Private Sub UStore(ByRef d As TInstr)
    Select Case d.Kind
        Case "NOP", "OUT"
            Signals "-"
            RtlText = "No write-back": DetailText = "-"
        Case "HLT"
            Signals "UC_Halt"
            Halted = True
            Act "UC"
            RtlText = "HLT: clock stopped"
            DetailText = "End of program (" & InstrCount & " instructions, " & Pulses & " pulses)"
        Case "MOV", "UNARY", "IN"
            Signals "En_TMP Ld_" & d.Op1
            SetReg d.Op1, TMP
            Act "TMP," & d.Op1
            RtlText = d.Op1 & Arrow() & "TMP"
            DetailText = d.Op1 & "=" & HX(TMP)
        Case "ALU"
            If d.Mnem = "CMP" Then
                Signals "-"
                Act "FLAGS"
                RtlText = "No write-back: CMP only changes the flags"
                DetailText = FlagsText()
            Else
                Signals "En_TMP Ld_" & d.Op1
                SetReg d.Op1, TMP
                Act "TMP," & d.Op1
                RtlText = d.Op1 & Arrow() & "TMP"
                DetailText = d.Op1 & "=" & HX(TMP)
            End If
        Case "LOAD"
            Signals "En_MDR Ld_" & d.Op1
            SetReg d.Op1, MDR
            Act "MDR," & d.Op1
            RtlText = d.Op1 & Arrow() & "MDR"
            DetailText = d.Op1 & "=" & HX(MDR)
        Case "STORE"
            Signals "En_MDR Mem_W"
            Bus_MemWrite MAR, MDR
            Act "MAR,MDR,BUS"
            RtlText = "M[MAR]" & Arrow() & "MDR"
            DetailText = "M[" & HA(MAR) & "]=" & HX(MDR)
        Case "JUMP"
            If JumpTaken Then
                Signals "En_IR Ld_PC"
                PC = A9(OPV)
                Act "IR,PC"
                RtlText = "PC" & Arrow() & "IR(op)"
            Else
                Signals "-"
                Act "PC"
                RtlText = "PC unchanged"
            End If
            DetailText = "PC=" & HA(PC)
    End Select
End Sub

Private Function IRText() As String
    Dim d As TInstr, s As String
    d = Decode(IR)
    s = Hex2(IR)
    If OPV >= 0 Then
        s = s & " " & Hex2(OpLo)
        If d.Size = 3 Then s = s & " " & Hex2(OpHi)
    End If
    IRText = s
End Function

'==============================================================================
'  5. SHEET CPU: panels, highlight and log
'==============================================================================
Public Sub Core_Reset()
    PC = 0: IR = 0: MAR = 0: MDR = 0
    AX = 0: BX = 0: CX = 0: DX = 0
    ZF = 0: CF = 0: SF = 0
    TMP = 0: OPV = -1: OpLo = 0: OpHi = 0
    Phase = "FETCH": Micro = 0: LastPhase = ""
    InstrCount = 0: Pulses = 0: JumpTaken = False: InstrAddr = 0
    Halted = False: Paused = False
    Active = "|": AluText = "": CtrlWord = ""
    RtlText = "CPU reset": DetailText = "PC=000h"
    ClearLog
    Sh.Range("B24").Value = "-"
    Core_Paint
    SetStatus "READY"
End Sub

Public Sub Core_SetupSheet()
    Dim ws As Worksheet, i As Long, names As Variant
    Set ws = Sh
    ws.Cells.Font.Name = "Calibri"
    ws.Columns("A").ColumnWidth = 28
    ws.Columns("B").ColumnWidth = 13
    ws.Columns("C").ColumnWidth = 17
    ws.Columns("D").ColumnWidth = 13
    ws.Columns("E").ColumnWidth = 3

    With ws.Range("A1:D1")
        .Merge
        .Value = TEAM_NAME & "  -  CPU"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 42, 74)
        .HorizontalAlignment = xlCenter
    End With

    ' Registers A3:B11
    Header ws.Range("A3"), "REGISTER"
    Header ws.Range("B3"), "VALUE"
    names = Array("PC (Program Counter)", "IR (Instruction Reg.)", "MAR (Memory Address Reg.)", "MDR (Memory Data Reg.)", _
                  "AX (Accumulator)", "BX (Base)", "CX (Counter)", "DX (Data)")
    For i = 0 To 7
        ws.Cells(4 + i, 1).Value = names(i)
        ws.Cells(4 + i, 1).Font.Bold = True
    Next i
    ws.Range("B4:B11").NumberFormat = "@"
    ws.Range("B4:B11").HorizontalAlignment = xlCenter
    ws.Range("B4:B11").Font.Name = "Consolas"
    ws.Range("B4:B11").Font.Bold = True

    ' Flags C3:D6
    Header ws.Range("C3"), "FLAG"
    Header ws.Range("D3"), "VALUE"
    ws.Range("C4").Value = "ZF (Zero)"
    ws.Range("C5").Value = "CF (Carry)"
    ws.Range("C6").Value = "SF (Sign)"
    ws.Range("C4:C6").Font.Bold = True
    ws.Range("D4:D6").HorizontalAlignment = xlCenter
    ws.Range("D4:D6").Font.Bold = True

    ' Counters C8:D11
    ws.Range("C8").Value = "Instruction #"
    ws.Range("C9").Value = "Clock pulses"
    ws.Range("C10").Value = "Status"
    ws.Range("C11").Value = "Run delay (ms)"
    ws.Range("C8:C11").Font.Bold = True
    ws.Range("D8:D11").HorizontalAlignment = xlCenter
    ws.Range("D11").Value = 100
    ws.Range("D11").Interior.Color = RGB(255, 255, 204)
    ws.Range("A3:D11").Borders.LineStyle = xlContinuous
    ws.Range("A3:D11").Borders.Color = RGB(200, 208, 222)

    ' Control unit A19:D25
    With ws.Range("A19:D19")
        .Merge
        .Value = "CONTROL UNIT"
        .Interior.Color = RGB(31, 42, 74)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("A20:D20").Value = Array("FETCH", "DECODE", "EXECUTE", "STORE")
    ws.Range("A20:D20").Font.Bold = True
    ws.Range("A20:D20").HorizontalAlignment = xlCenter
    For i = 21 To 22
        ws.Range("A" & i & ":D" & i).Merge
        ws.Range("A" & i).HorizontalAlignment = xlCenter
        ws.Range("A" & i).Font.Name = "Consolas"
    Next i
    ws.Range("A21").Font.Bold = True
    ws.Range("A21").Font.Size = 12
    ws.Range("A22").Font.Size = 9
    ws.Range("A23").Value = "Instruction"
    ws.Range("A24").Value = "ALU"
    ws.Range("A25").Value = "Control word"
    For i = 23 To 25
        ws.Range("B" & i & ":D" & i).Merge
        ws.Range("B" & i).Font.Name = "Consolas"
        ws.Range("A" & i).Font.Bold = True
    Next i
    ws.Range("A19:D25").Borders.LineStyle = xlContinuous
    ws.Range("A19:D25").Borders.Color = RGB(200, 208, 222)

    ' Log A27:D46
    With ws.Range("A27:D27")
        .Merge
        .Value = "MICRO-OPERATION LOG (full history in sheet Log)"
        .Interior.Color = RGB(30, 38, 56)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    For i = LOG_ROW1 To LOG_ROWN
        ws.Range("A" & i & ":D" & i).Merge
    Next i
    With ws.Range("A" & LOG_ROW1 & ":D" & LOG_ROWN)
        .Font.Name = "Consolas"
        .Font.Size = 9
        .Font.Color = RGB(200, 230, 200)
        .Interior.Color = RGB(30, 38, 56)
    End With

    SetupLogSheet
    CreateButtons
    Bus_SetupSheet
End Sub

Private Sub Header(ByVal r As Range, ByVal txt As String)
    r.Value = txt
    r.Font.Bold = True
    r.Font.Color = RGB(255, 255, 255)
    r.Interior.Color = RGB(47, 111, 181)
    r.HorizontalAlignment = xlCenter
End Sub

Private Sub CreateButtons()
    Dim ws As Worksheet, shp As Shape, i As Long, x As Double, y As Double, w As Double, h As Double
    Dim labels As Variant, macros As Variant, colors As Variant
    Set ws = Sh
    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 7) = "btnCore" Then
            ws.Shapes(i).Delete
        End If
    Next i
    labels = Array("STEP (micro-op)", "STEP INSTR", "RUN", "PAUSE", "RESET", "LOAD DEMO")
    macros = Array("BtnStep", "BtnStepInstr", "BtnRun", "BtnPause", "BtnReset", "BtnLoadDemo")
    colors = Array(RGB(47, 117, 181), RGB(31, 78, 121), RGB(84, 130, 53), RGB(191, 143, 0), RGB(192, 0, 0), RGB(112, 48, 160))
    w = (ws.Range("A13:D13").Width - 18) / 2
    h = (ws.Range("A13:A17").Height - 18) / 3
    For i = 0 To 5
        x = ws.Range("A13").Left + 6 + (i Mod 2) * (w + 6)
        y = ws.Range("A13").Top + 4 + (i \ 2) * (h + 5)
        Set shp = ws.Shapes.AddShape(5, x, y, w, h)
        shp.Name = "btnCore" & macros(i)
        shp.Fill.ForeColor.RGB = colors(i)
        shp.Line.Visible = msoFalse
        shp.OnAction = macros(i)
        With shp.TextFrame
            .Characters.Text = labels(i)
            .Characters.Font.Bold = True
            .Characters.Font.Size = 10
            .Characters.Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
    Next i
    On Error GoTo 0
End Sub

Public Sub Core_Paint()
    Dim ws As Worksheet, i As Long, d As TInstr, regs As Variant, phases As Variant
    Set ws = Sh

    ws.Range("B4").Value = HA(PC)
    ws.Range("B5").Value = IRText()
    ws.Range("B6").Value = HA(MAR)
    ws.Range("B7").Value = HX(MDR)
    ws.Range("B8").Value = HX(AX)
    ws.Range("B9").Value = HX(BX)
    ws.Range("B10").Value = HX(CX)
    ws.Range("B11").Value = HX(DX)
    ws.Range("D4").Value = ZF
    ws.Range("D5").Value = CF
    ws.Range("D6").Value = SF
    ws.Range("D8").Value = InstrCount
    ws.Range("D9").Value = Pulses

    regs = Array("PC", "IR", "MAR", "MDR", "AX", "BX", "CX", "DX")
    For i = 0 To 7
        If IsActive(CStr(regs(i))) Then
            ws.Range("A" & (4 + i) & ":B" & (4 + i)).Interior.Color = RGB(255, 217, 102)
        ElseIf i Mod 2 = 0 Then
            ws.Range("A" & (4 + i) & ":B" & (4 + i)).Interior.Color = RGB(242, 246, 252)
        Else
            ws.Range("A" & (4 + i) & ":B" & (4 + i)).Interior.Color = RGB(255, 255, 255)
        End If
    Next i
    If IsActive("FLAGS") Then
        ws.Range("C4:D6").Interior.Color = RGB(255, 217, 102)
    Else
        ws.Range("C4:D6").Interior.Color = RGB(255, 242, 204)
    End If

    phases = Array("FETCH", "DECODE", "EXECUTE", "STORE")
    For i = 0 To 3
        If phases(i) = LastPhase Then
            ws.Cells(20, i + 1).Interior.Color = Choose(i + 1, RGB(47, 117, 181), RGB(123, 94, 167), RGB(217, 130, 43), RGB(46, 158, 98))
            ws.Cells(20, i + 1).Font.Color = RGB(255, 255, 255)
        Else
            ws.Cells(20, i + 1).Interior.Color = RGB(217, 217, 217)
            ws.Cells(20, i + 1).Font.Color = RGB(89, 89, 89)
        End If
    Next i
    ws.Range("A21").Value = RtlText
    ws.Range("A22").Value = DetailText
    If Phase = "FETCH" And Micro > 0 Then
        ws.Range("B23").Value = "fetching instruction at " & HA(InstrAddr) & "..."
    ElseIf InstrCount > 0 Then
        d = Decode(IR)
        ws.Range("B23").Value = HA(InstrAddr) & ": " & InstrText(d, OPV)
    Else
        ws.Range("B23").Value = "-"
    End If
    If IsActive("ALU") Then
        ws.Range("B24").Value = AluText
        ws.Range("B24").Interior.Color = RGB(255, 217, 102)
    Else
        ws.Range("B24").Interior.Color = RGB(255, 255, 255)
    End If
    ws.Range("B25").Value = CtrlWord
End Sub

Private Sub SetStatus(ByVal s As String)
    Sh.Range("D10").Value = s
End Sub

Private Sub SetupLogSheet()
    Dim wl As Worksheet
    Set wl = GetSheet(SH_LOG)
    wl.Range("A1:F1").Value = Array("Pulse", "Instruction", "Phase", "Micro-operation", "Detail", "Control word")
    wl.Range("A1:F1").Font.Bold = True
    wl.Columns("D").ColumnWidth = 40
    wl.Columns("E").ColumnWidth = 50
    wl.Columns("F").ColumnWidth = 30
End Sub

Private Sub ClearLog()
    Dim ws As Worksheet
    Set ws = Sh
    ws.Range("A" & LOG_ROW1 & ":A" & LOG_ROWN).Value = ""
    GetSheet(SH_LOG).Range("A2:F30000").ClearContents
End Sub

Private Sub WriteLog()
    Dim ws As Worksheet, wl As Worksheet, txt As String, r As Long, v As Variant, n As Long
    Set ws = Sh
    Set wl = GetSheet(SH_LOG)
    wl.Cells(Pulses + 1, 1).Value = Pulses
    wl.Cells(Pulses + 1, 2).Value = InstrCount
    wl.Cells(Pulses + 1, 3).Value = LastPhase
    wl.Cells(Pulses + 1, 4).Value = RtlText
    wl.Cells(Pulses + 1, 5).Value = DetailText
    wl.Cells(Pulses + 1, 6).Value = CtrlWord
    txt = "[" & Format$(Pulses, "0000") & "] " & LastPhase & ": " & RtlText & " | " & DetailText
    n = LOG_ROWN - LOG_ROW1 + 1
    v = ws.Range("A" & LOG_ROW1 & ":A" & LOG_ROWN).Value
    For r = 1 To n
        If CStr(v(r, 1)) = "" Then Exit For
    Next r
    If r > n Then
        For r = 1 To n - 1
            v(r, 1) = v(r + 1, 1)
        Next r
        r = n
    End If
    v(r, 1) = txt
    ws.Range("A" & LOG_ROW1 & ":A" & LOG_ROWN).Value = v
End Sub

Private Sub PaintAll()
    Core_Paint
    Bus_Paint
    Cache_Paint
    VM_Paint
    IO_Paint
End Sub

Private Sub WaitMs(ByVal ms As Long)
    Dim t0 As Single
    t0 = Timer
    Do While (Timer - t0) * 1000 < ms And (Timer - t0) >= 0
        DoEvents
        If Paused Then Exit Do
    Loop
End Sub

'==============================================================================
'  6. BUTTONS
'==============================================================================
Private Function CheckHalted() As Boolean
    If Halted Then
        MsgBox "The CPU is halted (HLT). Press RESET or LOAD DEMO.", vbInformation, TEAM_NAME
        CheckHalted = True
    End If
End Function

Public Sub BtnStep()
    ' STEP: one micro-operation (one clock pulse)
    If Running Then Exit Sub
    If CheckHalted() Then Exit Sub
    MicroStep
    WriteLog
    PaintAll
    If Halted Then
        SetStatus "HALTED"
    End If
End Sub

Public Sub BtnStepInstr()
    ' STEP INSTR: all the micro-operations of one instruction
    Dim n As Long
    If Running Then Exit Sub
    If CheckHalted() Then Exit Sub
    Do
        MicroStep
        WriteLog
        n = n + 1
    Loop Until Halted Or (Phase = "FETCH" And Micro = 0) Or n > 30
    PaintAll
    If Halted Then
        SetStatus "HALTED"
    End If
End Sub

Public Sub BtnRun()
    ' RUN: continuous execution with the delay of cell D11 (ms)
    Dim delay As Long
    If Running Then Exit Sub
    If CheckHalted() Then Exit Sub
    delay = Val(Sh.Range("D11").Value)
    If delay < 0 Then delay = 0
    Paused = False
    Running = True
    SetStatus "RUNNING..."
    Do While Not Halted And Not Paused
        MicroStep
        WriteLog
        PaintAll
        DoEvents
        If delay > 0 Then
            WaitMs delay
        End If
        If Pulses > 20000 Then Exit Do        ' protection against infinite loops
    Loop
    Running = False
    If Halted Then
        SetStatus "HALTED"
    Else
        SetStatus "PAUSED"
    End If
End Sub

Public Sub BtnPause()
    Paused = True
End Sub

Public Sub BtnReset()
    ' RESET: registers, flags and PC to 0. Memory is kept.
    If Running Then
        Paused = True
        Exit Sub
    End If
    Bus_Reset
    Core_Reset
    PaintAll
End Sub

'==============================================================================
'  7. QUICK DEMO (developer shortcut; the real catalog lives in modAssembler)
'
'  Multiplication by repeated addition: 7 x 3 = 21 = 15h
'   Addr Bytes      Label  Instruction        Comment
'   000  A1 80 00   START  LOAD BX, [080h]    BX = A
'   003  A2 81 00          LOAD CX, [081h]    CX = B (counter)
'   006  80 00             MOV AX, 0          AX = 0
'   008  8E 00             CMP CX, 0          B = 0 ?
'   00A  C1 12 00          JZ END
'   00D  21         LOOP   ADD AX, BX         AX = AX + A
'   00E  B6                DEC CX             CX = CX - 1
'   00F  C2 0D 00          JNZ LOOP
'   012  A4 82 00   END    STORE [082h], AX   RES = AX
'   015  FF                HLT
'   080  07 03 00          A, B, RES
'==============================================================================
Public Sub BtnLoadDemo()
    Dim img() As Long, code As Variant, i As Long
    If Running Then Exit Sub
    ReDim img(0 To &H82)
    code = Array(&HA1, &H80, &H0, &HA2, &H81, &H0, &H80, &H0, &H8E, &H0, &HC1, &H12, &H0, _
                 &H21, &HB6, &HC2, &HD, &H0, &HA4, &H82, &H0, &HFF)
    For i = 0 To UBound(code)
        img(i) = code(i)
    Next i
    img(&H80) = 7
    img(&H81) = 3
    img(&H82) = 0
    VM_Reset
    VM_LoadImage img, &H83
    Cache_Reset
    IO_Reset
    Bus_Reset
    Core_Reset
    PaintAll
    RtlText = "Demo loaded: 7 x 3 (result in 082h)"
    Sh.Range("A21").Value = RtlText
End Sub
