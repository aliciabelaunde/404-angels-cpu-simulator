Attribute VB_Name = "modCore"
'==============================================================================
'  404 Angels - CPU Simulator
'  modCore : CPU registers, ALU, decoder, control unit, buttons
'  OWNER   : Alicia
'
'  TODO (Alicia, feature/alicia/core-cpu):
'    - Migrate the Parcial 1 CPU (modSimulador.bas, ISA v2 with AX..DX).
'    - Every memory access goes through Bus_MemRead / Bus_MemWrite.
'    - PC and MAR become 9-bit (see docs/ISA.md).
'    - Decode IN (D0-D3) and OUT (D4-D7) -> Bus_IoRead / Bus_IoWrite.
'    - Emit the control word with Bus_ControlWord in every micro-step.
'==============================================================================
Option Explicit

' --- Registers (public so other panels can display them) --------------------
Public PC As Long, IR As Long, MAR As Long, MDR As Long
Public AX As Long, BX As Long, CX As Long, DX As Long
Public ZF As Long, CF As Long, SF As Long

Public Sub Core_Reset()
    PC = 0: IR = 0: MAR = 0: MDR = 0
    AX = 0: BX = 0: CX = 0: DX = 0
    ZF = 0: CF = 0: SF = 0
    Core_Paint
End Sub

Public Sub Core_SetupSheet()
    Dim ws As Worksheet
    Set ws = GetSheet(SH_CPU)
    ws.Range("A1").Value = TEAM_NAME & " - CPU"
    GetSheet SH_LOG
End Sub

Public Sub Core_Paint()
    ' STUB: repaint registers on sheet CPU
End Sub

Private Sub PaintAll()
    Core_Paint
    Bus_Paint
    Cache_Paint
    VM_Paint
    IO_Paint
End Sub

Public Sub BtnStep()
    ' STUB: one micro-operation
    PaintAll
End Sub

Public Sub BtnStepInstr()
    ' STUB: one full instruction
    PaintAll
End Sub

Public Sub BtnRun()
    ' STUB: run until HLT
    PaintAll
End Sub

Public Sub BtnPause()
    ' STUB
End Sub

Public Sub BtnReset()
    Bus_Reset
    Core_Reset
    PaintAll
End Sub
