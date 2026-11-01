Attribute VB_Name = "modCache"
'==============================================================================
'  404 Angels - CPU Simulator
'  modCache : L1I (8B), L1D (8B), L2 (32B), hit/miss counters, hit rate
'  OWNER    : Alicia
'  Policy   : write-through (RAM is always up to date)
'==============================================================================
Option Explicit

Public L1I_Hits As Long, L1I_Misses As Long
Public L1D_Hits As Long, L1D_Misses As Long
Public L2_Hits As Long, L2_Misses As Long

Public Sub Cache_Reset()
    L1I_Hits = 0: L1I_Misses = 0
    L1D_Hits = 0: L1D_Misses = 0
    L2_Hits = 0: L2_Misses = 0
End Sub

Public Function Cache_Read(ByVal pa As Long, ByVal isInstr As Boolean) As Long
    ' STUB: no cache yet, go straight to RAM
    Cache_Read = RAM_Read(pa)
End Function

Public Sub Cache_Write(ByVal pa As Long, ByVal value As Long)
    ' STUB: write-through straight to RAM
    RAM_Write pa, value
End Sub

Public Sub Cache_InvalidateFrame(ByVal frame As Long)
    ' STUB: drop every cached line whose address is inside this frame
End Sub

Public Sub Cache_SetupSheet()
    GetSheet(SH_CACHE).Range("A1").Value = TEAM_NAME & " - Cache L1I / L1D / L2"
End Sub

Public Sub Cache_Paint()
    ' STUB
End Sub
