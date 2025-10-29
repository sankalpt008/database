Attribute VB_Name = "AccessBuild"
Option Compare Database
Option Explicit

Private Const DEVLOG_RELATIVE As String = "..\DEVLOG\development_log.md"
Private Const acCmdSaveRecord As Long = 615
Private Const acCmdDeleteRecord As Long = 223
Private Const acNewRec As Long = 5
Private Const acNext As Long = 2
Private Const acPrevious As Long = 1

Public Sub Build_Module3()
    On Error GoTo ErrHandler
    LogMessage "Starting Build_Module3"

    Dim basePath As String
    basePath = CurrentProject.Path

    RunSqlFile BuildPath(basePath, "sql\01_setup_tables.sql")
    ClearAndImport "Customers", BuildPath(basePath, "data\customers.csv")
    RunSqlFile BuildPath(basePath, "sql\02_queries.sql")

    LogMessage "Completed Build_Module3"
    Exit Sub
ErrHandler:
    LogError "Build_Module3", Err.Number, Err.Description
End Sub

Public Sub Build_Chapter4()
    On Error GoTo ErrHandler
    LogMessage "Starting Build_Chapter4"

    Dim basePath As String
    basePath = CurrentProject.Path

    RunSqlFile BuildPath(basePath, "sql\01_setup_tables.sql")
    ClearAndImport "Customers", BuildPath(basePath, "data\customers.csv")
    ClearAndImport "Orders", BuildPath(basePath, "data\orders.csv")

    CreateOneToMany "Customers", "CustomerID", "Orders", "CustomerID", True, False

    BuildCustomerForms

    LogMessage "Completed Build_Chapter4"
    Exit Sub
ErrHandler:
    LogError "Build_Chapter4", Err.Number, Err.Description
End Sub

Public Sub Build_Capstone()
    On Error GoTo ErrHandler
    LogMessage "Starting Build_Capstone"

    Dim basePath As String
    basePath = CurrentProject.Path

    RunSqlFile BuildParentPath(basePath, "sql\01_schema.sql")

    ClearAndImport "Customers", BuildParentPath(basePath, "data\customers.csv")
    ClearAndImport "Products", BuildParentPath(basePath, "data\products.csv")
    ClearAndImport "Orders", BuildParentPath(basePath, "data\orders.csv")
    ClearAndImport "OrderDetails", BuildParentPath(basePath, "data\order_details.csv")

    RunSqlFile BuildParentPath(basePath, "sql\02_seed.sql")
    RunSqlFile BuildParentPath(basePath, "sql\03_queries.sql")

    CreateOneToMany "Customers", "CustomerID", "Orders", "CustomerID", True, False
    CreateOneToMany "Orders", "OrderID", "OrderDetails", "OrderID", True, False
    CreateOneToMany "Products", "ProductID", "OrderDetails", "ProductID", False, False

    BuildCapstoneForms
    BuildCapstoneReports

    LogMessage "Completed Build_Capstone"
    Exit Sub
ErrHandler:
    LogError "Build_Capstone", Err.Number, Err.Description
End Sub

Public Sub LogMessage(ByVal message As String)
    On Error GoTo LogErrorHandler

    Dim timestamp As String
    timestamp = Format$(Now(), "yyyy-mm-dd\THH:nn:ss\Z")

    Debug.Print timestamp & " - " & message

    Dim devlogPath As String
    devlogPath = GetDevLogPath()

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim stream As Object
    Set stream = fso.OpenTextFile(devlogPath, 8, True)
    stream.WriteLine "- " & timestamp & " - " & message
    stream.Close

    Exit Sub
LogErrorHandler:
    Debug.Print "LogMessage failure: " & Err.Description
End Sub

Public Sub LogError(ByVal context As String, ByVal number As Long, ByVal description As String)
    LogMessage "ERROR in " & context & " (#" & number & "): " & description
End Sub

Public Sub ClearAndImport(ByVal tableName As String, ByVal csvPath As String)
    On Error GoTo ErrHandler

    LogMessage "Refreshing data for table " & tableName & " from " & csvPath

    On Error Resume Next
    CurrentDb.Execute "DELETE * FROM [" & tableName & "]", dbFailOnError
    On Error GoTo ErrHandler

    ImportCsv tableName, csvPath

    Exit Sub
ErrHandler:
    LogError "ClearAndImport(" & tableName & ")", Err.Number, Err.Description
End Sub

Public Function BuildPath(ByVal basePath As String, ByVal relativePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    BuildPath = fso.BuildPath(basePath, relativePath)
End Function

Public Function BuildParentPath(ByVal basePath As String, ByVal relativePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim parentFolder As String
    parentFolder = fso.GetParentFolderName(basePath)
    BuildParentPath = fso.BuildPath(parentFolder, relativePath)
End Function

Public Function FilterCustomersByCity() As Boolean
    On Error GoTo ErrHandler

    Dim frm As Form
    Set frm = Screen.ActiveForm

    Dim cityValue As Variant
    cityValue = Nz(frm.Controls("cboCityFilter").Value, "")

    If Len(cityValue & "") = 0 Then
        frm.FilterOn = False
    Else
        frm.Filter = "[City] = '" & Replace(cityValue, "'", "''") & "'"
        frm.FilterOn = True
    End If

    FilterCustomersByCity = True
    Exit Function
ErrHandler:
    LogError "FilterCustomersByCity", Err.Number, Err.Description
    FilterCustomersByCity = False
End Function

Public Function CustomerCommand(ByVal actionKey As String) As Boolean
    On Error GoTo ErrHandler

    Select Case LCase$(actionKey)
        Case "new"
            DoCmd.GoToRecord , , acNewRec
        Case "save"
            DoCmd.RunCommand acCmdSaveRecord
        Case "delete"
            DoCmd.RunCommand acCmdDeleteRecord
        Case "next"
            DoCmd.GoToRecord , , acNext
        Case "previous"
            DoCmd.GoToRecord , , acPrevious
        Case Else
            LogMessage "Unknown CustomerCommand action: " & actionKey
    End Select

    CustomerCommand = True
    Exit Function
ErrHandler:
    LogError "CustomerCommand", Err.Number, Err.Description
    CustomerCommand = False
End Function

Private Function GetDevLogPath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim currentPath As String
    currentPath = CurrentProject.Path

    Dim projectRoot As String
    projectRoot = fso.GetParentFolderName(currentPath)
    If Len(projectRoot) = 0 Then
        projectRoot = currentPath
    End If

    GetDevLogPath = fso.BuildPath(projectRoot, DEVLOG_RELATIVE)
End Function
