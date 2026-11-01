Attribute VB_Name = "modBuild"
'==============================================================================
'  404 Angels - CPU Simulator
'  modBuild : imports/exports VBA modules between src/ and the workbook
'  OWNER    : BOTH (change only through a PR approved by Alicia and Paola)
'
'  Requires: Trust Center > Macro Settings >
'            "Trust access to the VBA project object model"
'==============================================================================
Option Explicit

' This module must NOT depend on any other module, because it is the first
' one imported into an empty workbook. Other modules are called by name
' through Application.Run.

Private Const BUILD_TITLE As String = "404 Angels"
Private SrcPath As String

' Converts a OneDrive web path (https://...) into the local synced folder.
Private Function LocalPath(ByVal p As String) As String
    Dim parts() As String, i As Long, rest As String
    If LCase$(Left$(p, 4)) <> "http" Then
        LocalPath = p
        Exit Function
    End If
    parts = Split(p, "/")
    If InStr(1, p, "/Documents/", vbTextCompare) > 0 Then
        rest = Mid$(p, InStr(1, p, "/Documents/", vbTextCompare) + 11)
        LocalPath = Environ$("OneDriveCommercial") & "\" & Replace(rest, "/", "\")
    Else
        For i = 4 To UBound(parts)
            rest = rest & "\" & parts(i)
        Next i
        LocalPath = Environ$("OneDrive") & rest
    End If
End Function

' Folder with the .bas files: <workbook>\src\  or  <workbook>\..\src\
Private Function SrcFolder() As String
    Dim wbPath As String
    If SrcPath = "" Then
        wbPath = LocalPath(ThisWorkbook.Path)
        If Dir(wbPath & "\src\*.bas") <> "" Then
            SrcPath = wbPath & "\src\"
        ElseIf Dir(wbPath & "\..\src\*.bas") <> "" Then
            SrcPath = wbPath & "\..\src\"
        Else
            SrcPath = InputBox("Paste the full path of the repository's src folder:", BUILD_TITLE)
            If SrcPath <> "" And Right$(SrcPath, 1) <> "\" Then
                SrcPath = SrcPath & "\"
            End If
        End If
    End If
    SrcFolder = SrcPath
End Function

Private Sub RemoveComponent(ByVal moduleName As String)
    Dim comp As Object
    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents(moduleName)
    On Error GoTo 0
    If comp Is Nothing Then Exit Sub
    comp.Name = moduleName & "_old"          ' avoids the "modCore1" duplicate bug
    ThisWorkbook.VBProject.VBComponents.Remove comp
End Sub

' Imports every module in src/ (except this one). Run after each git pull.
Public Sub Build_ImportModules()
    Dim folder As String, f As String, moduleName As String, files As New Collection, item As Variant
    folder = SrcFolder()
    If folder = "" Then Exit Sub
    f = Dir(folder & "*.bas")
    Do While f <> ""
        files.Add f
        f = Dir()
    Loop
    For Each item In files
        moduleName = Left$(item, Len(item) - 4)
        If moduleName <> "modBuild" Then
            RemoveComponent moduleName
            ThisWorkbook.VBProject.VBComponents.Import folder & item
        End If
    Next item
    MsgBox files.Count & " module(s) imported from " & folder, vbInformation, BUILD_TITLE
End Sub

' Exports ONE module back to src/. Export only modules you own!
'   Example (Immediate window):  Build_ExportModule "modCache"
Public Sub Build_ExportModule(ByVal moduleName As String)
    ThisWorkbook.VBProject.VBComponents(moduleName).Export SrcFolder() & moduleName & ".bas"
    MsgBox moduleName & ".bas exported to src/", vbInformation, BUILD_TITLE
End Sub

' Creates every sheet and resets the whole computer.
Public Sub Build_SetupAll()
    Dim steps As Variant, stepName As Variant
    steps = Array("Core_SetupSheet", "Cache_SetupSheet", "VM_SetupSheet", "IO_SetupSheet", "Asm_SetupSheet", _
                  "VM_Reset", "Cache_Reset", "IO_Reset", "Bus_Reset", "Core_Reset")
    Application.ScreenUpdating = False
    For Each stepName In steps
        Application.Run stepName
    Next stepName
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets("CPU").Activate
End Sub
