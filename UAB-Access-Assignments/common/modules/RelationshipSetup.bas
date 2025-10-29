Attribute VB_Name = "RelationshipSetup"
Option Compare Database
Option Explicit

Private Const dbRelationUpdateCascade As Long = &H100
Private Const dbRelationDeleteCascade As Long = &H200

Public Sub CreateOneToMany(ByVal parentTable As String, _
                           ByVal parentKey As String, _
                           ByVal childTable As String, _
                           ByVal childForeignKey As String, _
                           ByVal cascadeUpdate As Boolean, _
                           ByVal cascadeDelete As Boolean)
    On Error GoTo ErrHandler

    Dim relationName As String
    relationName = parentTable & "_" & childTable

    If RelationExists(relationName) Then
        LogMessage "Relationship already exists: " & relationName
        Exit Sub
    End If

    Dim attributes As Long
    attributes = 0
    If cascadeUpdate Then attributes = attributes Or dbRelationUpdateCascade
    If cascadeDelete Then attributes = attributes Or dbRelationDeleteCascade

    Dim db As DAO.Database
    Set db = CurrentDb

    Dim rel As DAO.Relation
    Set rel = db.CreateRelation(relationName, parentTable, childTable, attributes)

    Dim relField As DAO.Field
    Set relField = rel.CreateField(childForeignKey)
    relField.ForeignName = parentKey
    rel.Fields.Append relField

    db.Relations.Append rel
    LogMessage "Created relationship: " & relationName
    Exit Sub
ErrHandler:
    LogError "CreateOneToMany(" & relationName & ")", Err.Number, Err.Description
End Sub

Private Function RelationExists(ByVal relationName As String) As Boolean
    Dim rel As DAO.Relation
    For Each rel In CurrentDb.Relations
        If StrComp(rel.Name, relationName, vbTextCompare) = 0 Then
            RelationExists = True
            Exit Function
        End If
    Next
    RelationExists = False
End Function
