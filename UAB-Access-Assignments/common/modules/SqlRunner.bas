Attribute VB_Name = "SqlRunner"
Option Compare Database
Option Explicit

Public Sub RunSqlFile(ByVal filePath As String)
    On Error GoTo ErrHandler

    LogMessage "Running SQL file: " & filePath

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        Err.Raise vbObjectError + 513, "RunSqlFile", "SQL file not found: " & filePath
    End If

    Dim ts As Object
    Set ts = fso.OpenTextFile(filePath, 1)

    Dim buffer As String
    buffer = ""

    Do Until ts.AtEndOfStream
        Dim line As String
        line = ts.ReadLine

        Dim commentPos As Long
        commentPos = InStr(line, "--")
        If commentPos > 0 Then
            line = Left$(line, commentPos - 1)
        End If

        If Len(Trim$(line)) = 0 Then
            ' skip empty lines
        Else
            buffer = buffer & line & vbCrLf
            If InStr(line, ";") > 0 Then
                buffer = ProcessBuffer(buffer)
            End If
        End If
    Loop

    ts.Close

    If Len(Trim$(buffer)) > 0 Then
        ProcessStatement buffer
    End If

    Exit Sub
ErrHandler:
    LogError "RunSqlFile(" & filePath & ")", Err.Number, Err.Description
End Sub

Private Function ProcessBuffer(ByVal buffer As String) As String
    Dim statements As Variant
    statements = Split(buffer, ";")

    Dim i As Long
    For i = LBound(statements) To UBound(statements) - 1
        ProcessStatement statements(i)
    Next i

    ProcessBuffer = statements(UBound(statements))
End Function

Private Sub ProcessStatement(ByVal rawStatement As String)
    Dim statement As String
    statement = Trim$(rawStatement)
    If Len(statement) = 0 Then Exit Sub

    If LCase$(Left$(statement, 9)) = "savequery" Then
        HandleSaveQuery statement
    ElseIf LCase$(Left$(statement, 12)) = "create table" Then
        HandleCreateTable statement
    Else
        CurrentDb.Execute statement, dbFailOnError
        LogMessage "Executed SQL statement: " & Left$(statement, 80)
    End If
End Sub

Private Sub HandleCreateTable(ByVal statement As String)
    Dim tableName As String
    tableName = ExtractCreateName(statement)

    If Len(tableName) = 0 Then
        CurrentDb.Execute statement, dbFailOnError
        LogMessage "Executed CREATE TABLE (name unknown)."
        Exit Sub
    End If

    If TableExists(tableName) Then
        LogMessage "Skipped existing table: " & tableName
    Else
        CurrentDb.Execute statement, dbFailOnError
        LogMessage "Created table: " & tableName
    End If
End Sub

Private Sub HandleSaveQuery(ByVal statement As String)
    Dim asPos As Long
    asPos = InStr(1, statement, " AS ", vbTextCompare)
    If asPos = 0 Then
        Err.Raise vbObjectError + 514, "HandleSaveQuery", "Missing AS keyword in SAVEQUERY statement."
    End If

    Dim header As String
    header = Trim$(Mid$(statement, 10, asPos - 9))

    Dim queryName As String
    queryName = Trim$(header)

    Dim sqlText As String
    sqlText = Trim$(Mid$(statement, asPos + 4))

    If Len(queryName) = 0 Then
        Err.Raise vbObjectError + 515, "HandleSaveQuery", "Query name not provided."
    End If

    CreateOrReplaceQuery queryName, sqlText
End Sub

Private Sub CreateOrReplaceQuery(ByVal queryName As String, ByVal sqlText As String)
    On Error GoTo ErrHandler

    Dim qd As DAO.QueryDef

    On Error Resume Next
    Set qd = CurrentDb.QueryDefs(queryName)
    If Err.Number = 0 Then
        CurrentDb.QueryDefs.Delete queryName
    End If
    Err.Clear
    On Error GoTo ErrHandler

    CurrentDb.CreateQueryDef queryName, sqlText
    LogMessage "Created query: " & queryName
    Exit Sub
ErrHandler:
    LogError "CreateOrReplaceQuery(" & queryName & ")", Err.Number, Err.Description
End Sub

Private Function TableExists(ByVal tableName As String) As Boolean
    Dim tdf As DAO.TableDef
    For Each tdf In CurrentDb.TableDefs
        If StrComp(tdf.Name, tableName, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next
    TableExists = False
End Function

Private Function ExtractCreateName(ByVal statement As String) As String
    Dim tokens As Variant
    tokens = Split(statement)
    If UBound(tokens) >= 2 Then
        Dim rawName As String
        rawName = tokens(2)
        rawName = Replace(rawName, "[", "")
        rawName = Replace(rawName, "]", "")
        rawName = Replace(rawName, "(", "")
        ExtractCreateName = Trim$(rawName)
    Else
        ExtractCreateName = ""
    End If
End Function
