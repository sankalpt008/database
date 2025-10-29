Attribute VB_Name = "CsvImport"
Option Compare Database
Option Explicit

Private Const acImportDelim As Long = 0

Public Sub ImportCsv(ByVal tableName As String, ByVal csvPath As String)
    On Error GoTo ErrHandler

    DoCmd.TransferText acImportDelim, , tableName, csvPath, True
    LogMessage "DoCmd.TransferText completed for table " & tableName

    Exit Sub
ErrHandler:
    LogError "ImportCsv(" & tableName & ")", Err.Number, Err.Description
End Sub
