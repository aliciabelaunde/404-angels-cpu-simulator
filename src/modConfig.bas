Attribute VB_Name = "modConfig"
'==============================================================================
'  404 Angels - CPU Simulator
'  modConfig : shared constants and helpers
'  OWNER     : BOTH (change only through a PR approved by Alicia and Paola)
'==============================================================================
Option Explicit

' --- Memory sizes ------------------------------------------------------------
Public Const PAGE_SIZE As Long = 32
Public Const NUM_PAGES As Long = 16          ' virtual pages 0..15
Public Const NUM_FRAMES As Long = 8          ' physical frames 0..7
Public Const VIRT_SIZE As Long = 512         ' bytes of virtual memory
Public Const PHYS_SIZE As Long = 256         ' bytes of physical RAM

' --- Cache sizes -------------------------------------------------------------
Public Const L1_LINE As Long = 8             ' L1I and L1D: one 8-byte line each
Public Const L2_SIZE As Long = 32            ' L2: 4 lines of 8 bytes

' --- Memory-mapped I/O (virtual addresses, see docs/ISA.md) ------------------
Public Const KBD_BASE As Long = &H1E0        ' keyboard queue, 8 bytes
Public Const LCD_BASE As Long = &H1E8        ' LCD video buffer, 8 bytes
Public Const IO_SIZE As Long = 8
Public Const IO_PAGE As Long = 15            ' pinned page
Public Const IO_FRAME As Long = 7            ' pinned frame

' --- I/O ports -----------------------------------------------------------------
Public Const PORT_KBD As Long = 0
Public Const PORT_LCD As Long = 1

' --- Sheet names (each module paints only its own sheet) ---------------------
Public Const SH_CPU As String = "CPU"            ' modCore + modBus
Public Const SH_CACHE As String = "Cache"        ' modCache
Public Const SH_MEMORY As String = "Memory"      ' modVirtualMemory
Public Const SH_IO As String = "IO"              ' modIO
Public Const SH_PROGRAMS As String = "Programs"  ' modAssembler
Public Const SH_LOG As String = "Log"            ' modCore

Public Const TEAM_NAME As String = "404 Angels"

' Returns the sheet with that name, creating it at the end if missing.
Public Function GetSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set GetSheet = ws
End Function

Public Function Hex2(ByVal v As Long) As String
    Hex2 = Right$("0" & Hex$(v And &HFF), 2)
End Function

Public Function Hex3(ByVal v As Long) As String
    Hex3 = Right$("00" & Hex$(v And &H1FF), 3)
End Function
