Attribute VB_Name = "modBus"
'==============================================================================
'  404 Angels - CPU Simulator
'  modBus : control / address / data buses, tri-state buffers, control word
'  OWNER  : Alicia
'
'  Rule: in every micro-step only ONE component drives the bus (Enable);
'        all the others stay in High-Z.
'==============================================================================
Option Explicit

Public ControlWord As String     ' e.g. "En_PC Ld_MAR"
Public AddressBus As Long        ' 9 bits
Public DataBus As Long           ' 8 bits, -1 = High-Z

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
    DataBus = IO_KeyboardRead()
    Bus_IoRead = DataBus
End Function

Public Sub Bus_IoWrite(ByVal port As Long, ByVal value As Long)
    DataBus = value And &HFF
    IO_LcdWrite DataBus
End Sub

Public Sub Bus_Paint()
    ' STUB: paint buses, buffers (Enable / High-Z) and the control word on sheet CPU
End Sub
