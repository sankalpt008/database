Option Explicit

Const acModule = 5

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count = 0 Then
    WScript.Echo "Usage: cscript //nologo build_access.vbs [module3|chapter4|capstone]"
    WScript.Quit 1
End If

Dim assignment
assignment = LCase(WScript.Arguments.Item(0))

Dim projectRoot
projectRoot = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))

Dim dbPath, entrypoint, moduleFolders
Select Case assignment
    Case "module3"
        dbPath = fso.BuildPath(projectRoot, "module3_queries\module3.accdb")
        entrypoint = "Build_Module3"
        moduleFolders = Array(fso.BuildPath(projectRoot, "common\modules"))
    Case "chapter4"
        dbPath = fso.BuildPath(projectRoot, "chapter4_forms\chapter4.accdb")
        entrypoint = "Build_Chapter4"
        moduleFolders = Array(_
            fso.BuildPath(projectRoot, "common\modules"), _
            fso.BuildPath(projectRoot, "chapter4_forms\vba"))
    Case "capstone"
        dbPath = fso.BuildPath(projectRoot, "capstone1\capstone.accdb")
        entrypoint = "Build_Capstone"
        moduleFolders = Array(_
            fso.BuildPath(projectRoot, "common\modules"), _
            fso.BuildPath(projectRoot, "capstone1\vba"))
    Case Else
        WScript.Echo "Unknown assignment: " & assignment
        WScript.Quit 1
End Select

Dim accessApp
On Error Resume Next
Set accessApp = CreateObject("Access.Application")
If Err.Number <> 0 Then
    WScript.Echo "Failed to create Access instance: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

On Error Resume Next
If fso.FileExists(dbPath) Then
    accessApp.OpenCurrentDatabase dbPath
Else
    Dim dbFolder
    dbFolder = fso.GetParentFolderName(dbPath)
    If Not fso.FolderExists(dbFolder) Then
        fso.CreateFolder dbFolder
    End If
    accessApp.NewCurrentDatabase dbPath
End If
If Err.Number <> 0 Then
    WScript.Echo "Unable to open database: " & Err.Description
    accessApp.Quit
    WScript.Quit 1
End If
On Error GoTo 0

Dim i
For i = LBound(moduleFolders) To UBound(moduleFolders)
    Call ImportModulesInFolder(accessApp, moduleFolders(i))
Next

On Error Resume Next
accessApp.Run entrypoint
If Err.Number <> 0 Then
    WScript.Echo "Error running " & entrypoint & ": " & Err.Description
    accessApp.CloseCurrentDatabase
    accessApp.Quit
    WScript.Quit 1
End If
On Error GoTo 0

accessApp.CloseCurrentDatabase
accessApp.Quit
WScript.Echo "Successfully executed " & entrypoint & " for " & assignment

Sub ImportModulesInFolder(app, folderPath)
    If Not fso.FolderExists(folderPath) Then Exit Sub

    Dim folder, file
    Set folder = fso.GetFolder(folderPath)
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Path)) = "bas" Then
            Dim moduleName
            moduleName = GetModuleName(file.Path)
            If Len(moduleName) > 0 Then
                On Error Resume Next
                app.DoCmd.DeleteObject acModule, moduleName
                On Error GoTo 0
                app.LoadFromText acModule, moduleName, file.Path
            End If
        End If
    Next
End Sub

Function GetModuleName(filePath)
    Dim ts, line
    GetModuleName = ""
    On Error Resume Next
    Set ts = fso.OpenTextFile(filePath, 1)
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0

    Do Until ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If InStr(1, line, "Attribute VB_Name", vbTextCompare) = 1 Then
            Dim parts
            parts = Split(line, "=")
            If UBound(parts) >= 1 Then
                Dim rawName
                rawName = Trim(parts(1))
                rawName = Replace(rawName, Chr(34), "")
                GetModuleName = rawName
                Exit Do
            End If
        End If
    Loop
    ts.Close
End Function
