Attribute VB_Name = "modAssembler"
'==============================================================================
'  404 Angels - CPU Simulator
'  modAssembler : program catalog, editor, assembler (ISA v3), loader
'  OWNER        : Paola
'==============================================================================
Option Explicit

Public Sub Asm_SetupSheet()
    GetSheet(SH_PROGRAMS).Range("A1").Value = TEAM_NAME & " - Program directory"
End Sub

Public Sub BtnAssemble()
    ' STUB: check syntax of the editor cells and show errors
End Sub

Public Sub BtnLoadProgram()
    ' STUB: assemble, then load into virtual memory and reset everything
    Dim image() As Long
    ReDim image(0 To 0)
    image(0) = &HFF                          ' HLT
    VM_Reset
    VM_LoadImage image, 1
    Cache_Reset
    IO_Reset
    Bus_Reset
    Core_Reset
End Sub
