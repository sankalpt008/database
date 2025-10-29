Attribute VB_Name = "CapstoneForms"
Option Compare Database
Option Explicit

Private Const acForm As Long = 2
Private Const acReport As Long = 3
Private Const acTextBox As Long = 109
Private Const acComboBox As Long = 111
Private Const acCommandButton As Long = 104
Private Const acSubform As Long = 112
Private Const acLabel As Long = 100
Private Const acDetail As Long = 0
Private Const acFormHeader As Long = 1
Private Const acFormFooter As Long = 2
Private Const acViewContinuous As Long = 1
Private Const acViewSingle As Long = 0
Private Const acSaveYes As Long = -1
Private Const acViewPreview As Long = 2

Public Sub BuildCapstoneForms()
    On Error GoTo ErrHandler

    LogMessage "Building Capstone forms"

    DeleteIfExists "Form", "OrderDetails_Subform"
    DeleteIfExists "Form", "OrderEntry"
    DeleteIfExists "Form", "CustomerForm"
    DeleteIfExists "Form", "MainMenu"
    DeleteIfExists "Form", "ReportsMenu"

    BuildOrderDetailsSubform
    BuildCustomerForm
    BuildOrderEntryForm
    BuildReportsMenu
    BuildMainMenu

    LogMessage "Capstone forms complete"
    Exit Sub
ErrHandler:
    LogError "BuildCapstoneForms", Err.Number, Err.Description
End Sub

Public Function LaunchCapstoneTarget(ByVal target As String) As Boolean
    On Error GoTo ErrHandler

    Select Case LCase$(target)
        Case "customerform"
            DoCmd.OpenForm "CustomerForm"
        Case "orderentry"
            DoCmd.OpenForm "OrderEntry"
        Case "reportsmenu"
            DoCmd.OpenForm "ReportsMenu"
        Case "salesbycustomerreport"
            DoCmd.OpenReport "rptSalesByCustomer", acViewPreview
        Case "topproductsreport"
            DoCmd.OpenReport "rptTopProducts", acViewPreview
        Case Else
            LogMessage "LaunchCapstoneTarget received unknown target: " & target
    End Select

    LaunchCapstoneTarget = True
    Exit Function
ErrHandler:
    LogError "LaunchCapstoneTarget", Err.Number, Err.Description
    LaunchCapstoneTarget = False
End Function

Private Sub BuildMainMenu()
    Dim frm As Form
    Set frm = CreateForm

    frm.Caption = "Capstone Main Menu"
    frm.RecordSource = ""
    frm.Section(acFormHeader).Visible = True

    Dim lblHeader As Control
    Set lblHeader = CreateControl(frm.Name, acLabel, acFormHeader, , , 480, 120, 4800, 360)
    lblHeader.Caption = "Select an area to manage"
    lblHeader.FontSize = 14

    AddMenuButton frm, "btnCustomers", "Customer Form", 480, 720, "CustomerForm"
    AddMenuButton frm, "btnOrders", "Order Entry", 480, 1200, "OrderEntry"
    AddMenuButton frm, "btnReports", "Reports", 480, 1680, "ReportsMenu"

    DoCmd.Save acForm, frm.Name
    DoCmd.Close acForm, frm.Name, acSaveYes
    DoCmd.Rename "MainMenu", acForm, frm.Name
End Sub

Private Sub BuildCustomerForm()
    Dim frm As Form
    Set frm = CreateForm

    frm.RecordSource = "Customers"
    frm.DefaultView = acViewSingle
    frm.NavigationButtons = True
    frm.AllowAdditions = True
    frm.AllowDeletions = True

    Dim currentTop As Long
    currentTop = 480
    AddBoundTextBox frm, "FirstName", "First Name", 480, currentTop
    currentTop = currentTop + 480
    AddBoundTextBox frm, "LastName", "Last Name", 480, currentTop
    currentTop = currentTop + 480
    AddBoundTextBox frm, "City", "City", 480, currentTop
    currentTop = currentTop + 480
    AddBoundTextBox frm, "Email", "Email", 480, currentTop, 3600

    DoCmd.Save acForm, frm.Name
    DoCmd.Close acForm, frm.Name, acSaveYes
    DoCmd.Rename "CustomerForm", acForm, frm.Name
End Sub

Private Sub BuildOrderDetailsSubform()
    Dim frm As Form
    Set frm = CreateForm

    frm.RecordSource = "OrderDetails"
    frm.DefaultView = acViewContinuous
    frm.AllowAdditions = True
    frm.AllowDeletions = True
    frm.Section(acFormFooter).Visible = True

    Dim cboProduct As Control
    Set cboProduct = CreateControl(frm.Name, acComboBox, acDetail, , "ProductID", 480, 480, 2400, 360)
    cboProduct.Name = "cboProduct"
    cboProduct.RowSource = "SELECT ProductID, ProductName FROM Products WHERE Active = True ORDER BY ProductName"
    cboProduct.ColumnCount = 2
    cboProduct.ColumnWidths = "2.5;0"
    cboProduct.BoundColumn = 1
    cboProduct.ControlSource = "ProductID"

    Dim lblProduct As Control
    Set lblProduct = CreateControl(frm.Name, acLabel, acDetail, cboProduct.Name, , 480, 120, 2400, 240)
    lblProduct.Caption = "Product"

    Dim txtQuantity As Control
    Set txtQuantity = CreateControl(frm.Name, acTextBox, acDetail, , "Quantity", 3000, 480, 1200, 360)
    txtQuantity.Name = "txtQuantity"

    Dim lblQuantity As Control
    Set lblQuantity = CreateControl(frm.Name, acLabel, acDetail, txtQuantity.Name, , 3000, 120, 1200, 240)
    lblQuantity.Caption = "Quantity"

    Dim txtLineTotal As Control
    Set txtLineTotal = CreateControl(frm.Name, acTextBox, acDetail, , "LineTotal", 4320, 480, 1440, 360)
    txtLineTotal.Name = "txtLineTotal"
    txtLineTotal.Format = "Currency"

    Dim lblLineTotal As Control
    Set lblLineTotal = CreateControl(frm.Name, acLabel, acDetail, txtLineTotal.Name, , 4320, 120, 1440, 240)
    lblLineTotal.Caption = "Line Total"

    Dim txtSum As Control
    Set txtSum = CreateControl(frm.Name, acTextBox, acFormFooter, , , 4320, 240, 1440, 360)
    txtSum.Name = "txtOrderSum"
    txtSum.ControlSource = "=Sum([LineTotal])"
    txtSum.Format = "Currency"

    DoCmd.Save acForm, frm.Name
    DoCmd.Close acForm, frm.Name, acSaveYes
    DoCmd.Rename "OrderDetails_Subform", acForm, frm.Name
End Sub

Private Sub BuildOrderEntryForm()
    Dim frm As Form
    Set frm = CreateForm

    frm.RecordSource = "Orders"
    frm.DefaultView = acViewSingle
    frm.AllowAdditions = True
    frm.AllowDeletions = True

    Dim cboCustomer As Control
    Set cboCustomer = CreateControl(frm.Name, acComboBox, acDetail, , "CustomerID", 480, 480, 3000, 360)
    cboCustomer.Name = "cboCustomer"
    cboCustomer.RowSource = "SELECT CustomerID, FirstName & ' ' & LastName AS FullName FROM Customers ORDER BY LastName, FirstName"
    cboCustomer.ColumnCount = 2
    cboCustomer.ColumnWidths = "3;0"
    cboCustomer.BoundColumn = 1

    Dim lblCustomer As Control
    Set lblCustomer = CreateControl(frm.Name, acLabel, acDetail, cboCustomer.Name, , 480, 120, 3000, 240)
    lblCustomer.Caption = "Customer"

    Dim txtOrderDate As Control
    Set txtOrderDate = CreateControl(frm.Name, acTextBox, acDetail, , "OrderDate", 480, 960, 1800, 360)
    txtOrderDate.Format = "Short Date"
    txtOrderDate.Name = "txtOrderDate"

    Dim lblOrderDate As Control
    Set lblOrderDate = CreateControl(frm.Name, acLabel, acDetail, txtOrderDate.Name, , 480, 720, 1800, 240)
    lblOrderDate.Caption = "Order Date"

    Dim subformControl As Control
    Set subformControl = CreateControl(frm.Name, acSubform, acDetail, , , 480, 1560, 5520, 2400)
    subformControl.Name = "sfmOrderDetails"
    subformControl.SourceObject = "Form.OrderDetails_Subform"
    subformControl.LinkMasterFields = "OrderID"
    subformControl.LinkChildFields = "OrderID"

    Dim txtTotal As Control
    Set txtTotal = CreateControl(frm.Name, acTextBox, acDetail, , , 480, 4080, 1800, 360)
    txtTotal.Name = "txtOrderTotal"
    txtTotal.ControlSource = "=Nz([sfmOrderDetails].Form!txtOrderSum,0)"
    txtTotal.Format = "Currency"
    txtTotal.Enabled = False

    Dim lblTotal As Control
    Set lblTotal = CreateControl(frm.Name, acLabel, acDetail, txtTotal.Name, , 480, 3840, 1800, 240)
    lblTotal.Caption = "Order Total"

    DoCmd.Save acForm, frm.Name
    DoCmd.Close acForm, frm.Name, acSaveYes
    DoCmd.Rename "OrderEntry", acForm, frm.Name
End Sub

Private Sub BuildReportsMenu()
    Dim frm As Form
    Set frm = CreateForm

    frm.Caption = "Reports"
    frm.RecordSource = ""

    Dim lblHeader As Control
    Set lblHeader = CreateControl(frm.Name, acLabel, acDetail, , , 480, 480, 3600, 360)
    lblHeader.Caption = "Choose a report to preview"
    lblHeader.FontSize = 12

    AddMenuButton frm, "btnSalesByCustomer", "Sales by Customer", 480, 960, "SalesByCustomerReport"
    AddMenuButton frm, "btnTopProducts", "Top Products", 480, 1440, "TopProductsReport"

    DoCmd.Save acForm, frm.Name
    DoCmd.Close acForm, frm.Name, acSaveYes
    DoCmd.Rename "ReportsMenu", acForm, frm.Name
End Sub

Private Sub AddMenuButton(ByVal frm As Form, ByVal name As String, ByVal caption As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal target As String)
    Dim cmd As Control
    Set cmd = CreateControl(frm.Name, acCommandButton, acDetail, , , leftPos, topPos, 2400, 480)
    cmd.Name = name
    cmd.Caption = caption
    cmd.OnClick = "=LaunchCapstoneTarget(\"" & target & "\")"
End Sub

Private Sub AddBoundTextBox(ByVal frm As Form, ByVal fieldName As String, ByVal caption As String, ByVal leftPos As Long, ByVal topPos As Long, Optional ByVal width As Long = 2400)
    Dim txt As Control
    Set txt = CreateControl(frm.Name, acTextBox, acDetail, , fieldName, leftPos + 1200, topPos, width, 360)
    txt.Name = "txt" & fieldName

    Dim lbl As Control
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, txt.Name, , leftPos, topPos, 1200, 360)
    lbl.Caption = caption
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
