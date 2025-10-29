Attribute VB_Name = "FormsBuilder"
Option Compare Database
Option Explicit

Private Const acForm As Long = 2
Private Const acTextBox As Long = 109
Private Const acComboBox As Long = 111
Private Const acCommandButton As Long = 104
Private Const acSubform As Long = 112
Private Const acLabel As Long = 100
Private Const acDetail As Long = 0
Private Const acFormHeader As Long = 1
Private Const acViewContinuous As Long = 1
Private Const acViewSingle As Long = 0
Private Const acLayoutTabular As Long = 1
Private Const acSaveYes As Long = -1

Public Sub BuildCustomerForms()
    On Error GoTo ErrHandler

    LogMessage "Building Chapter 4 CustomerEntry form"

    DeleteIfExists "Form", "Orders_Subform"
    DeleteIfExists "Form", "CustomerEntry"

    BuildOrdersSubform
    BuildCustomerEntryForm

    LogMessage "CustomerEntry form build complete"
    Exit Sub
ErrHandler:
    LogError "BuildCustomerForms", Err.Number, Err.Description
End Sub

Private Sub BuildOrdersSubform()
    Dim frm As Form
    Set frm = CreateForm

    frm.RecordSource = "Orders"
    frm.DefaultView = acViewContinuous
    frm.AllowAdditions = False
    frm.AllowDeletions = False
    frm.NavigationButtons = False

    Dim ctlOrderDate As Control
    Set ctlOrderDate = CreateControl(frm.Name, acTextBox, acDetail, , "OrderDate", 480, 480, 2400, 360)
    ctlOrderDate.Format = "Short Date"
    ctlOrderDate.Name = "txtOrderDate"
    ctlOrderDate.ColumnOrder = 0

    Dim lblOrderDate As Control
    Set lblOrderDate = CreateControl(frm.Name, acLabel, acDetail, ctlOrderDate.Name, , 480, 120, 2400, 240)
    lblOrderDate.Caption = "Order Date"

    Dim ctlTotal As Control
    Set ctlTotal = CreateControl(frm.Name, acTextBox, acDetail, , "Total", 3000, 480, 1800, 360)
    ctlTotal.Format = "Currency"
    ctlTotal.Name = "txtTotal"
    ctlTotal.ColumnOrder = 1

    Dim lblTotal As Control
    Set lblTotal = CreateControl(frm.Name, acLabel, acDetail, ctlTotal.Name, , 3000, 120, 1800, 240)
    lblTotal.Caption = "Total"

    DoCmd.Save acForm, frm.Name
    DoCmd.Close acForm, frm.Name, acSaveYes

    DoCmd.Rename "Orders_Subform", acForm, frm.Name
End Sub

Private Sub BuildCustomerEntryForm()
    Dim frm As Form
    Set frm = CreateForm

    frm.RecordSource = "Customers"
    frm.DefaultView = acViewSingle
    frm.NavigationButtons = True
    frm.AllowAdditions = True
    frm.AllowDeletions = True
    frm.AllowEdits = True
    frm.Section(acFormHeader).Visible = True

    Dim lblHeader As Control
    Set lblHeader = CreateControl(frm.Name, acLabel, acFormHeader, , , 480, 120, 6000, 360)
    lblHeader.Caption = "Customer Entry - Use the city filter to focus records"
    lblHeader.FontSize = 12

    Dim currentTop As Long
    currentTop = 720

    AddBoundTextBox frm, "FirstName", "First Name", 480, currentTop
    currentTop = currentTop + 480
    AddBoundTextBox frm, "LastName", "Last Name", 480, currentTop
    currentTop = currentTop + 480
    AddBoundTextBox frm, "City", "City", 480, currentTop
    currentTop = currentTop + 480
    AddBoundTextBox frm, "Email", "Email", 480, currentTop, 3600

    Dim cboFilter As Control
    Set cboFilter = CreateControl(frm.Name, acComboBox, acDetail, , , 3600, 720, 2400, 360)
    cboFilter.Name = "cboCityFilter"
    cboFilter.RowSource = "SELECT DISTINCT City FROM Customers ORDER BY City"
    cboFilter.ColumnCount = 1
    cboFilter.BoundColumn = 1
    cboFilter.ColumnWidths = "2.5"
    cboFilter.AfterUpdate = "=FilterCustomersByCity()"
    cboFilter.ControlTipText = "Choose a city to filter the form"

    Dim lblFilter As Control
    Set lblFilter = CreateControl(frm.Name, acLabel, acDetail, cboFilter.Name, , 3600, 480, 2400, 240)
    lblFilter.Caption = "Filter by City"

    Dim buttonTop As Long
    buttonTop = currentTop + 480
    AddCommandButton frm, "btnNew", "New", 480, buttonTop, "New"
    AddCommandButton frm, "btnSave", "Save", 1320, buttonTop, "Save"
    AddCommandButton frm, "btnDelete", "Delete", 2160, buttonTop, "Delete"
    AddCommandButton frm, "btnPrev", "Previous", 3000, buttonTop, "Previous"
    AddCommandButton frm, "btnNext", "Next", 3840, buttonTop, "Next"

    Dim subformControl As Control
    Set subformControl = CreateControl(frm.Name, acSubform, acDetail, , , 480, buttonTop + 600, 5520, 2400)
    subformControl.Name = "sfmOrders"
    subformControl.SourceObject = "Form.Orders_Subform"
    subformControl.LinkChildFields = "CustomerID"
    subformControl.LinkMasterFields = "CustomerID"
    subformControl.ControlTipText = "Related orders"

    DoCmd.Save acForm, frm.Name
    DoCmd.Close acForm, frm.Name, acSaveYes
    DoCmd.Rename "CustomerEntry", acForm, frm.Name
End Sub

Private Sub AddBoundTextBox(ByVal frm As Form, ByVal fieldName As String, ByVal caption As String, ByVal leftPos As Long, ByVal topPos As Long, Optional ByVal width As Long = 2400)
    Dim txt As Control
    Set txt = CreateControl(frm.Name, acTextBox, acDetail, , fieldName, leftPos + 1200, topPos, width, 360)
    txt.Name = "txt" & fieldName

    Dim lbl As Control
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, txt.Name, , leftPos, topPos, 1200, 360)
    lbl.Caption = caption
End Sub

Private Sub AddCommandButton(ByVal frm As Form, ByVal name As String, ByVal caption As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal actionKey As String)
    Dim cmd As Control
    Set cmd = CreateControl(frm.Name, acCommandButton, acDetail, , , leftPos, topPos, 720, 360)
    cmd.Name = name
    cmd.Caption = caption
    cmd.OnClick = "=CustomerCommand(\"" & actionKey & "\")"
End Sub

Private Sub DeleteIfExists(ByVal objectType As String, ByVal objectName As String)
    On Error Resume Next
    DoCmd.DeleteObject GetObjectTypeConstant(objectType), objectName
    Err.Clear
    On Error GoTo 0
End Sub

Private Function GetObjectTypeConstant(ByVal objectType As String) As Long
    Select Case LCase$(objectType)
        Case "form"
            GetObjectTypeConstant = acForm
        Case Else
            GetObjectTypeConstant = acForm
    End Select
End Function
