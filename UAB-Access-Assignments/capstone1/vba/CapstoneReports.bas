Attribute VB_Name = "CapstoneReports"
Option Compare Database
Option Explicit

Private Const acReport As Long = 3
Private Const acTextBox As Long = 109
Private Const acLabel As Long = 100
Private Const acDetail As Long = 0
Private Const acReportHeader As Long = 1
Private Const acSaveYes As Long = -1

Public Sub BuildCapstoneReports()
    On Error GoTo ErrHandler

    LogMessage "Building Capstone reports"

    DeleteReportIfExists "rptSalesByCustomer"
    DeleteReportIfExists "rptTopProducts"

    BuildSalesByCustomerReport
    BuildTopProductsReport

    LogMessage "Capstone reports complete"
    Exit Sub
ErrHandler:
    LogError "BuildCapstoneReports", Err.Number, Err.Description
End Sub

Private Sub BuildSalesByCustomerReport()
    Dim rpt As Report
    Set rpt = CreateReport

    rpt.RecordSource = "q_SalesByCustomer"

    AddReportHeader rpt, "Sales by Customer"

    Dim txtCustomer As Control
    Set txtCustomer = CreateReportControl(rpt, acTextBox, acDetail, , "CustomerName", 480, 720, 3600, 360)
    txtCustomer.Name = "txtCustomerName"

    Dim txtTotal As Control
    Set txtTotal = CreateReportControl(rpt, acTextBox, acDetail, , "TotalSales", 4200, 720, 1800, 360)
    txtTotal.Format = "Currency"
    txtTotal.Name = "txtTotalSales"

    Dim lblCustomer As Control
    Set lblCustomer = CreateReportControl(rpt, acLabel, acDetail, txtCustomer.Name, , 480, 480, 3600, 240)
    lblCustomer.Caption = "Customer"

    Dim lblTotal As Control
    Set lblTotal = CreateReportControl(rpt, acLabel, acDetail, txtTotal.Name, , 4200, 480, 1800, 240)
    lblTotal.Caption = "Total Sales"

    DoCmd.Save acReport, rpt.Name
    DoCmd.Close acReport, rpt.Name, acSaveYes
    DoCmd.Rename "rptSalesByCustomer", acReport, rpt.Name
End Sub

Private Sub BuildTopProductsReport()
    Dim rpt As Report
    Set rpt = CreateReport

    rpt.RecordSource = "q_TopProducts"

    AddReportHeader rpt, "Top Products"

    Dim txtProduct As Control
    Set txtProduct = CreateReportControl(rpt, acTextBox, acDetail, , "ProductName", 480, 720, 3600, 360)
    txtProduct.Name = "txtProductName"

    Dim txtRevenue As Control
    Set txtRevenue = CreateReportControl(rpt, acTextBox, acDetail, , "TotalRevenue", 4200, 720, 1800, 360)
    txtRevenue.Format = "Currency"
    txtRevenue.Name = "txtTotalRevenue"

    Dim lblProduct As Control
    Set lblProduct = CreateReportControl(rpt, acLabel, acDetail, txtProduct.Name, , 480, 480, 3600, 240)
    lblProduct.Caption = "Product"

    Dim lblRevenue As Control
    Set lblRevenue = CreateReportControl(rpt, acLabel, acDetail, txtRevenue.Name, , 4200, 480, 1800, 240)
    lblRevenue.Caption = "Total Revenue"

    DoCmd.Save acReport, rpt.Name
    DoCmd.Close acReport, rpt.Name, acSaveYes
    DoCmd.Rename "rptTopProducts", acReport, rpt.Name
End Sub

Private Sub AddReportHeader(ByVal rpt As Report, ByVal caption As String)
    rpt.Section(acReportHeader).Visible = True
    Dim lbl As Control
    Set lbl = CreateReportControl(rpt, acLabel, acReportHeader, , , 480, 120, 3600, 360)
    lbl.Caption = caption
    lbl.FontSize = 14
End Sub

Private Function CreateReportControl(ByVal rpt As Report, ByVal controlType As Long, ByVal section As Long, Optional ByVal parent As Variant, Optional ByVal controlSource As Variant, Optional ByVal leftPos As Long = 0, Optional ByVal topPos As Long = 0, Optional ByVal width As Long = 1440, Optional ByVal height As Long = 360) As Control
    If IsMissing(parent) Then parent = Null
    If IsMissing(controlSource) Then controlSource = Null
    Set CreateReportControl = CreateControl(rpt.Name, controlType, section, parent, controlSource, leftPos, topPos, width, height)
End Function

Private Sub DeleteReportIfExists(ByVal reportName As String)
    On Error Resume Next
    DoCmd.DeleteObject acReport, reportName
    Err.Clear
    On Error GoTo 0
End Sub
