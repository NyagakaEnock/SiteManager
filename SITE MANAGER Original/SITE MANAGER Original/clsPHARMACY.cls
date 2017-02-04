VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMainMenu"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub GetMainStructure()
'On Error GoTo Err
    With SchedulingMain
        .ListView1.ListItems.Clear
        .ListView1.ColumnHeaders.Clear
        
        .ListView1.ColumnHeaders.Add , , "User Name", .ListView1.Width / 6.5
        .ListView1.ColumnHeaders.Add , , "Staff ID", .ListView1.Width / 6.5 ', lvwColumnCenter
        .ListView1.ColumnHeaders.Add , , "Full Names", .ListView1.Width / 3 ', lvwColumnCenter
        .ListView1.ColumnHeaders.Add , , "User Group", .ListView1.Width / 6.5
        .ListView1.ColumnHeaders.Add , , "Account Status", .ListView1.Width / 5
        
        .ListView1.View = lvwReport
    End With
    Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub OpenMain()
'On Error GoTo Err
    Load frmLogin
    frmLogin.Show 1
    Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub ResizeControls()
'On Error GoTo Err
With SchedulingMain
    .TreeView1.Left = 0: .TreeView1.Top = 720
    .ListView1.Left = 2760: .ListView1.Top = 720
    
    .ListView1.Width = .Width - 2880: .ListView1.Height = .Height - 1860
    .TreeView1.Height = .Height - 1860: .txtTotal.Top = .Height - 1935
End With
Exit Sub
Err:
If Err.Number = 380 Then Resume Next
    ErrorMessage
End Sub

Public Sub UnLoadMyForm()
    Call UpdateLogoutRecord
    Unload SchedulingMain
End Sub

Public Sub HelpContents()
'On Error GoTo Err
With SchedulingMain
    Dim nRet As Integer
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, .Caption
    Else
        On Error Resume Next
        nRet = WinHelp(.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub HelpIndex()
'On Error GoTo Err
With SchedulingMain
    Dim nRet As Integer

    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, .Caption
    Else
        On Error Resume Next
        nRet = WinHelp(.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub HelpAbout()
'On Error GoTo Err
    Load frmAbout
    frmAbout.Show 1
Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub CreatePharmacyTree()
'On Error GoTo Err
Dim Node1 As Node

With SchedulingMain
    Set Node1 = .TreeView1.NodeS.Add(, , "P", "Quotation Manager", 1)
    Set Node1 = .TreeView1.NodeS.Add("P", tvwChild, "P1", "Create Quotation", 1)
    Set Node1 = .TreeView1.NodeS.Add("P", tvwChild, "P2", "Approve Quotation", 1)
    Set Node1 = .TreeView1.NodeS.Add("P", tvwChild, "P3", "Authorize Quotation", 1)
    Set Node1 = .TreeView1.NodeS.Add("P", tvwChild, "P4", "Print Listing", 1)

    Set Node1 = .TreeView1.NodeS.Add(, , "C", "Job Brief Manager", 1)
    Set Node1 = .TreeView1.NodeS.Add("C", tvwChild, "C1", "Create Job Brief", 1)
    Set Node1 = .TreeView1.NodeS.Add("C", tvwChild, "C2", "Approve Job Brief", 1)
    Set Node1 = .TreeView1.NodeS.Add("C", tvwChild, "C3", "Authorize Job Brief", 1)

    Set Node1 = .TreeView1.NodeS.Add(, , "L", "Job Card Manager", 1)
    Set Node1 = .TreeView1.NodeS.Add("L", tvwChild, "L1", "Open Jobcard", 1)
    Set Node1 = .TreeView1.NodeS.Add("L", tvwChild, "L2", "Material Requisition", 1)
    Set Node1 = .TreeView1.NodeS.Add("L", tvwChild, "L3", "Material Approval", 1)
    Set Node1 = .TreeView1.NodeS.Add("L", tvwChild, "L4", "Material Authorization", 1)

    Set Node1 = .TreeView1.NodeS.Add("L", tvwChild, "L5", "Stores Approval", 1)
    Set Node1 = .TreeView1.NodeS.Add("L", tvwChild, "L6", "Close Jobcard", 1)
    Set Node1 = .TreeView1.NodeS.Add("L", tvwChild, "L7", "Project Costing", 1)
    Set Node1 = .TreeView1.NodeS.Add("L", tvwChild, "L8", "Costing Approval", 1)
    Set Node1 = .TreeView1.NodeS.Add("L", tvwChild, "L9", "Create Invoice", 1)
    Set Node1 = .TreeView1.NodeS.Add("L", tvwChild, "L10", "Approve Invoice", 1)
    
  
    Set Node1 = .TreeView1.NodeS.Add(, , "G", "General/Quick Queries", 1)
    Set Node1 = .TreeView1.NodeS.Add("G", tvwChild, "G1", "Show All Pinned Contracts And Advertisements", 1)
    Set Node1 = .TreeView1.NodeS.Add("G", tvwChild, "G2", "Show Contracts Expiring on Specific Date ", 1)
    Set Node1 = .TreeView1.NodeS.Add("G", tvwChild, "G3", "Show All CouncilFees Due On A Specific Date", 1)
    Set Node1 = .TreeView1.NodeS.Add("G", tvwChild, "G4", "Show Rent Fees Due On A Specific Date", 1)
    Set Node1 = .TreeView1.NodeS.Add("G", tvwChild, "G5", "Show List Of Empty BillBoards", 1)
    Set Node1 = .TreeView1.NodeS.Add("G", tvwChild, "G6", "Show List Of Contracts Not Pinned", 1)
    Set Node1 = .TreeView1.NodeS.Add("G", tvwChild, "G7", "Show List Of All LandLords", 1)
    Set Node1 = .TreeView1.NodeS.Add("G", tvwChild, "G8", "Show List Of All Customers", 1)
    Set Node1 = .TreeView1.NodeS.Add("G", tvwChild, "G9", "Show All Sites Not Valid And Allocated", 1)
    Set Node1 = .TreeView1.NodeS.Add("G", tvwChild, "G10", "Show All Sites Expiry On Specific Date", 1)
'    Set Node1 = .TreeView1.Nodes.Add("G", tvwChild, "G11", "Show All ", 1)
    
    Dim NodeS As Node

    Set NodeS = .TreeView1.NodeS.Add(, , "S", "System Manager", 1)
    Set NodeS = .TreeView1.NodeS.Add("S", tvwChild, "S1", "Set Rights", 1)
    Set NodeS = .TreeView1.NodeS.Add("S", tvwChild, "S2", "Give Rights", 1)
    Set NodeS = .TreeView1.NodeS.Add("S", tvwChild, "S3", "Category", 1)
    Set NodeS = .TreeView1.NodeS.Add("S", tvwChild, "S4", "Departments", 1)
    Set NodeS = .TreeView1.NodeS.Add("S", tvwChild, "S5", "Set Tasks", 1)
    Set NodeS = .TreeView1.NodeS.Add("S", tvwChild, "S6", "Show List Of Contracts Not Pinned", 1)
    Set NodeS = .TreeView1.NodeS.Add("S", tvwChild, "S7", "Show List Of All LandLords", 1)
    Set NodeS = .TreeView1.NodeS.Add("S", tvwChild, "S8", "Show List Of All Customers", 1)
    Set NodeS = .TreeView1.NodeS.Add("S", tvwChild, "S9", "Show All Sites Not Valid And Allocated", 1)
    Set NodeS = .TreeView1.NodeS.Add("S", tvwChild, "S10", "Show All Sites Expiry On Specific Date", 1)
    
    Dim NodeK As Node

    Set NodeK = .TreeView1.NodeS.Add(, , "K", "Site Manager", 1)
    Set NodeK = .TreeView1.NodeS.Add("K", tvwChild, "K1", "Site Registration", 1)
    Set NodeK = .TreeView1.NodeS.Add("K", tvwChild, "K2", "Assign Properties", 1)
    Set NodeK = .TreeView1.NodeS.Add("K", tvwChild, "K3", "Land Lord Registration", 1)
    Set NodeK = .TreeView1.NodeS.Add("K", tvwChild, "K4", "Site Allocation", 1)
    Set NodeK = .TreeView1.NodeS.Add("K", tvwChild, "K5", "Approve Sites", 1)
    Set NodeK = .TreeView1.NodeS.Add("K", tvwChild, "K6", "Authorize Sites", 1)
    Set NodeK = .TreeView1.NodeS.Add("K", tvwChild, "K7", "Available Sites", 1)
    Set NodeK = .TreeView1.NodeS.Add("K", tvwChild, "K8", "Expired Sites", 1)
    
    Dim NodeM As Node

    Set NodeM = .TreeView1.NodeS.Add(, , "M", "Site Maintenance", 1)
    Set NodeM = .TreeView1.NodeS.Add("M", tvwChild, "M1", "Set Maintenance Dates", 1)
    Set NodeM = .TreeView1.NodeS.Add("M", tvwChild, "M2", "Site Due For Maintenance", 1)
    Set NodeM = .TreeView1.NodeS.Add("M", tvwChild, "M3", "Sites Not Maintained", 1)
    
    Dim NodeJ As Node

    Set NodeJ = .TreeView1.NodeS.Add(, , "J", "Inventory Manager", 1)
    Set NodeJ = .TreeView1.NodeS.Add("J", tvwChild, "J1", "LPO Requisition", 1)
    Set NodeJ = .TreeView1.NodeS.Add("J", tvwChild, "J2", "LPO Approval", 1)
    Set NodeJ = .TreeView1.NodeS.Add("J", tvwChild, "J3", "LPO Authorization", 1)
    Set NodeJ = .TreeView1.NodeS.Add("J", tvwChild, "J4", "Receive Joods", 1)
    Set NodeJ = .TreeView1.NodeS.Add("J", tvwChild, "J5", "Set Prices", 1)
    Set NodeJ = .TreeView1.NodeS.Add("J", tvwChild, "J6", "Issue Goods", 1)
    Set NodeJ = .TreeView1.NodeS.Add("J", tvwChild, "J7", "Return Goods", 1)
    
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub ShowFullInventory()
'On Error GoTo Err
With SchedulingMain
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Product Code ", .ListView1.Width / 6.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Product Name", .ListView1.Width / 3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Current Quantity", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Units", .ListView1.Width / 8 ', lvwColumnCenter
'.ListView1.ColumnHeaders.Add , , "Total Doses", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Re-Order Level", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Order No", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Category Code", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Suplier Code", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Re-Order Status", .ListView1.Width / 5 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT GenProductsInventory.* FROM GenProductsInventory WHERE PurchaseOrderNo IS NOT NULL ORDER BY GenProductsInventory.DrugCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!drugcode))

    If Not IsNull(rsLIST!drugname) Then
        MyList.SubItems(1) = CStr(rsLIST!drugname)
    End If
    
    If Not IsNull(rsLIST!currentquantity) Then
        MyList.SubItems(2) = CStr(rsLIST!currentquantity)
    End If
        
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(3) = CStr(rsLIST!QuantityUnits)
    End If
    
'    If Not IsNull(rsLIST!totaldoses) Then
'        MyList.SubItems(4) = CStr(rsLIST!totaldoses)
'    End If
    
    If Not IsNull(rsLIST!reorderlevel) Then
        MyList.SubItems(4) = CStr(rsLIST!reorderlevel)
    End If
        
    If Not IsNull(rsLIST!PurchaseOrderNo) Then
        MyList.SubItems(5) = CStr(rsLIST!PurchaseOrderNo)
    End If

    If Not IsNull(rsLIST!categorycode) Then
        MyList.SubItems(6) = CStr(rsLIST!categorycode)
    End If
    
    If Not IsNull(rsLIST!suppliercode) Then
        MyList.SubItems(7) = CStr(rsLIST!suppliercode)
    End If
    
    If IsNull(rsLIST!reorderstatus) = True Then
        MyList.SubItems(8) = CStr("OK")
    ElseIf Not IsNull(rsLIST!reorderstatus) Then
        MyList.SubItems(8) = CStr("LOW")
    End If
    
    rsLIST.MoveNext
    
Wend

NewRecord = False

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub ShowNewInventory()
'On Error GoTo Err
With SchedulingMain
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Product Code ", .ListView1.Width / 6.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Product Name", .ListView1.Width / 3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Category Code", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Suplier Code", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Order Status", .ListView1.Width / 5 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT GenProductsInventory.* FROM GenProductsInventory WHERE PurchaseOrderNo IS NULL OR PurchaseORderNo='" & "" & "' ORDER BY GenProductsInventory.DrugCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!drugcode))

    If Not IsNull(rsLIST!drugname) Then
        MyList.SubItems(1) = CStr(rsLIST!drugname)
    End If
    
    If Not IsNull(rsLIST!categorycode) Then
        MyList.SubItems(2) = CStr(rsLIST!categorycode)
    End If
    
    If Not IsNull(rsLIST!suppliercode) Then
        MyList.SubItems(3) = CStr(rsLIST!suppliercode)
    End If
    
    If IsNull(rsLIST!reorderstatus) = True Then
        MyList.SubItems(4) = CStr("NOT ORDERED")
    ElseIf Not IsNull(rsLIST!reorderstatus) Then
        MyList.SubItems(4) = CStr("NOT ORDERED")
    End If
    
    rsLIST.MoveNext
    
Wend

NewRecord = True

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub



Public Sub FindDosageSettingsByDrugCode()
'On Error GoTo Err
With SchedulingMain

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Dosage Code", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Drug Code", .ListView1.Width / 6.2 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Drug Name", .ListView1.Width / 3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Dosage Type", .ListView1.Width / 5.7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Min. Age", .ListView1.Width / 6.3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Max. Age", .ListView1.Width / 6.3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Dose Count", .ListView1.Width / 6.2 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Package Type", .ListView1.Width / 5.2 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Quantity", .ListView1.Width / 6.2 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Units", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Times/Day", .ListView1.Width / 6.2 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Qty/Usage", .ListView1.Width / 5.9 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Units", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Total Days", .ListView1.Width / 5.5 ', lvwColumnCenter

.ListView1.View = lvwReport

INPQRY = InputBox("Please Enter a Valid Drug Code to Show Settings...", "Enter Drug Code")

If Len(INPQRY) = 0 Then
    MsgBox "No Values Entered or the Operation Has Been Cancelled!! No Work Was Done!!!", vbCritical + vbOKOnly, "No Values"
    Exit Sub
End If

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT ProductQuantitySetup.* FROM ProductQuantitySetup WHERE ProductQuantitySetup.DrugCode='" & INPQRY & "' ORDER BY ProductQuantitySetup.DrugCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing:  Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!dosagecode))

    If Not IsNull(rsLIST!drugcode) Then
        MyList.SubItems(1) = CStr(rsLIST!drugcode)
    End If
    
    If Not IsNull(rsLIST!drugname) Then
        MyList.SubItems(2) = CStr(rsLIST!drugname)
    End If
    
    If Not IsNull(rsLIST!dosagetype) Then
        MyList.SubItems(3) = CStr(rsLIST!dosagetype)
    End If
    
    If Not IsNull(rsLIST!minage) Then
        MyList.SubItems(4) = CStr(rsLIST!minage)
    End If
        
    If Not IsNull(rsLIST!maxage) Then
        MyList.SubItems(5) = CStr(rsLIST!maxage)
    End If
    
    If Not IsNull(rsLIST!dosagecount) Then
        MyList.SubItems(6) = CStr(rsLIST!dosagecount)
    End If
    
    If Not IsNull(rsLIST!packagetype) Then
        MyList.SubItems(7) = CStr(rsLIST!packagetype)
    End If
    
    If Not IsNull(rsLIST!quantity) Then
        MyList.SubItems(8) = CStr(rsLIST!quantity)
    End If
    
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(9) = CStr(rsLIST!QuantityUnits)
    End If
        
    If Not IsNull(rsLIST!dailyinterval) Then
        MyList.SubItems(10) = CStr(rsLIST!dailyinterval)
    End If
    
    If Not IsNull(rsLIST!intervalqty) Then
        MyList.SubItems(11) = CStr(rsLIST!intervalqty)
    End If
    
    If Not IsNull(rsLIST!intervalunits) Then
        MyList.SubItems(12) = CStr(rsLIST!intervalunits)
    End If
        
    If Not IsNull(rsLIST!totaldays) Then
        MyList.SubItems(13) = CStr(rsLIST!totaldays)
    End If
    
    rsLIST.MoveNext
    
Wend



Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub FindAllCostPriceSettings()
'On Error GoTo Err
With SchedulingMain

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear


.ListView1.ColumnHeaders.Add , , "Product Code", .ListView1.Width / 5.9
.ListView1.ColumnHeaders.Add , , "Product Cost", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Product Price", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Price Markup", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Product Name", .ListView1.Width / 3

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT ProductsCostPriceSetup.* FROM ProductsCostPriceSetup ORDER BY ProductsCostPriceSetup.DrugCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing:  Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!drugcode))
    
    If Not IsNull(rsLIST!DosageCost) Then
        MyList.SubItems(1) = FormatNumber(rsLIST!DosageCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!DosagePrice) Then
        MyList.SubItems(2) = FormatNumber(rsLIST!DosagePrice, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!pricemarkup) Then
        MyList.SubItems(3) = FormatNumber(rsLIST!pricemarkup, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!drugname) Then
        MyList.SubItems(4) = (rsLIST!drugname)
    End If
    
        
    rsLIST.MoveNext
    
Wend

.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight



Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub FindCostPriceSettingsByDrugCode()
'On Error GoTo Err
With SchedulingMain

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear


.ListView1.ColumnHeaders.Add , , "Product Code", .ListView1.Width / 5.9
.ListView1.ColumnHeaders.Add , , "Product Cost", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Product Price", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Price Markup", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Product Name", .ListView1.Width / 3

.ListView1.View = lvwReport

INPQRY = InputBox("Please Enter the Product Code to Show Settings...", "Enter Product Code")

If Len(INPQRY) = 0 Then
    MsgBox "No Values Entered or the Operation has been Cancelled!!", vbCritical + vbOKOnly, "No Values"
    Exit Sub
End If

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT ProductsCostPriceSetup.* FROM ProductsCostPriceSetup WHERE ProductsCostPriceSetup.DrugCode='" & Trim(INPQRY) & "' ORDER BY ProductsCostPriceSetup.DrugCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing:  Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!drugcode))
    
    If Not IsNull(rsLIST!DosageCost) Then
        MyList.SubItems(1) = FormatNumber(rsLIST!DosageCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!DosagePrice) Then
        MyList.SubItems(2) = FormatNumber(rsLIST!DosagePrice, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!pricemarkup) Then
        MyList.SubItems(3) = FormatNumber(rsLIST!pricemarkup, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!drugname) Then
        MyList.SubItems(4) = (rsLIST!drugname)
    End If
    
        
    rsLIST.MoveNext
    
Wend

.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight



Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub FindAllInjectionSettings()
'On Error GoTo Err
With SchedulingMain

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Code Number", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Inj. Type", .ListView1.Width / 5.8 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Drug Code ", .ListView1.Width / 5.8 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Drug Name", .ListView1.Width / 3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Min. Age", .ListView1.Width / 6.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Max. Age", .ListView1.Width / 6.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Quantity", .ListView1.Width / 5.8 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Units ", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Total Inj.", .ListView1.Width / 5.8 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Facility Code", .ListView1.Width / 5.2 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT PharmDrugsInjectionSetup.* FROM PharmDrugsInjectionSetup ORDER BY PharmDrugsInjectionSetup.DrugCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing:  Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!setupcode))

    If Not IsNull(rsLIST!injectiontype) Then
        MyList.SubItems(1) = CStr(rsLIST!injectiontype)
    End If
    
    If Not IsNull(rsLIST!drugcode) Then
        MyList.SubItems(2) = CStr(rsLIST!drugcode)
    End If

    If Not IsNull(rsLIST!drugname) Then
        MyList.SubItems(3) = CStr(rsLIST!drugname)
    End If
    
    If Not IsNull(rsLIST!minage) Then
        MyList.SubItems(4) = CStr(rsLIST!minage)
    End If
    
    If Not IsNull(rsLIST!maxage) Then
        MyList.SubItems(5) = CStr(rsLIST!maxage)
    End If
        
    If Not IsNull(rsLIST!quantity) Then
        MyList.SubItems(6) = CStr(rsLIST!quantity)
    End If
            
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(7) = CStr(rsLIST!QuantityUnits)
    End If
    
    If Not IsNull(rsLIST!totalinjections) Then
        MyList.SubItems(8) = CStr(rsLIST!totalinjections)
    End If
    
    If Not IsNull(rsLIST!facilitycode) Then
        MyList.SubItems(9) = CStr(rsLIST!facilitycode)
    End If
        
    rsLIST.MoveNext
    
Wend



Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub FindInjectionSettingsByDrugCode()
'On Error GoTo Err
With SchedulingMain

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Code Number", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Inj. Type", .ListView1.Width / 5.8 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Drug Code ", .ListView1.Width / 5.8 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Drug Name", .ListView1.Width / 3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Min. Age", .ListView1.Width / 6.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Max. Age", .ListView1.Width / 6.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Quantity", .ListView1.Width / 5.8 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Units ", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Total Inj.", .ListView1.Width / 5.8 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Facility Code", .ListView1.Width / 5.2 ', lvwColumnCenter

.ListView1.View = lvwReport

INPQRY = InputBox("Please Enter the Drug Code to Show Settings...", "Enter Drug Code")

If Len(INPQRY) = 0 Then
    MsgBox "No Values Entered or Operation has been Cancelled!!", vbCritical + vbOKOnly, "No Values"
    Exit Sub
End If

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT PharmDrugsInjectionSetup.* FROM PharmDrugsInjectionSetup WHERE PharmDrugsInjectionSetup.DrugCode='" & Trim(INPQRY) & "' ORDER BY PharmDrugsInjectionSetup.DrugCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing:  Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!setupcode))

    If Not IsNull(rsLIST!injectiontype) Then
        MyList.SubItems(1) = CStr(rsLIST!injectiontype)
    End If
    
    If Not IsNull(rsLIST!drugcode) Then
        MyList.SubItems(2) = CStr(rsLIST!drugcode)
    End If

    If Not IsNull(rsLIST!drugname) Then
        MyList.SubItems(3) = CStr(rsLIST!drugname)
    End If
    
    If Not IsNull(rsLIST!minage) Then
        MyList.SubItems(4) = CStr(rsLIST!minage)
    End If
    
    If Not IsNull(rsLIST!maxage) Then
        MyList.SubItems(5) = CStr(rsLIST!maxage)
    End If
        
    If Not IsNull(rsLIST!quantity) Then
        MyList.SubItems(6) = CStr(rsLIST!quantity)
    End If
            
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(7) = CStr(rsLIST!QuantityUnits)
    End If
    
    If Not IsNull(rsLIST!totalinjections) Then
        MyList.SubItems(8) = CStr(rsLIST!totalinjections)
    End If
    
    If Not IsNull(rsLIST!facilitycode) Then
        MyList.SubItems(9) = CStr(rsLIST!facilitycode)
    End If
        
    rsLIST.MoveNext
    
Wend



Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub ShowCurrentOrderRecords()
'On Error GoTo Err
With SchedulingMain
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Serial No", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Product Code ", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Product Name", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Date Ordered", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Date Expected", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Supplier Code", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Quantity Ordered", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Product Cost", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Total Cost", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Received", .ListView1.Width / 5.5 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT PharmPurchaseORders.*,PharmPurchaseOrdersData.* FROM PharmPurchaseORders,PharmPurchaseOrdersData WHERE PharmPurchaseORders.PurchaseOrderNo=PharmPurchaseOrdersData.PurchaseOrderNo AND PharmPurchaseOrdersData.ReceivedStatus IS NULL ORDER BY PharmPurchaseOrdersData.SerialNO;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SerialNo))

    If Not IsNull(rsLIST!drugcode) Then
        MyList.SubItems(1) = CStr(rsLIST!drugcode)
    End If
    
    If Not IsNull(rsLIST!drugname) Then
        MyList.SubItems(2) = CStr(rsLIST!drugname)
    End If
    
    If Not IsNull(rsLIST!OrderDate) Then
        MyList.SubItems(3) = CStr(rsLIST!OrderDate)
    End If
    
    If Not IsNull(rsLIST!DatePromised) Then
        MyList.SubItems(4) = CStr(rsLIST!DatePromised)
    End If
    
    If Not IsNull(rsLIST!suppliercode) Then
        MyList.SubItems(5) = CStr(rsLIST!suppliercode)
    End If
    
    If Not IsNull(rsLIST!quantityordered) And Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(6) = Trim(CStr(rsLIST!quantityordered)) & " " & Trim(CStr(rsLIST!QuantityUnits))
    End If
    
    If Not IsNull(rsLIST!DosageCost) Then
        MyList.SubItems(7) = FormatNumber(rsLIST!DosageCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!totalcost) Then
        MyList.SubItems(8) = FormatNumber(rsLIST!totalcost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If IsNull(rsLIST!receivedstatus) Then
        MyList.SubItems(9) = CStr("NO")
    ElseIf Not IsNull(rsLIST!receivedstatus) Then
        If rsLIST!receivedstatus = "Y" Then
            MyList.SubItems(9) = CStr("YES")
        Else
            MyList.SubItems(9) = CStr("NO")
        End If
    End If
    
    rsLIST.MoveNext
    
Wend

.ListView1.ColumnHeaders(8).Alignment = lvwColumnRight
.ListView1.ColumnHeaders(9).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub



Public Sub ShowOverDueITems()
'On Error GoTo Err
With SchedulingMain
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Serial No", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Product Code ", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Product Name", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Date Ordered", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Date Expected", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Supplier Code", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Quantity Ordered", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Product Cost", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Total Cost", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Received", .ListView1.Width / 5.5 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim ThisDate: ThisDate = Format(Date, "MMMM dd,yyyy")

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT PharmPurchaseORders.*,PharmPurchaseOrdersData.* FROM PharmPurchaseORders,PharmPurchaseOrdersData WHERE PharmPurchaseORders.PurchaseOrderNo=PharmPurchaseOrdersData.PurchaseOrderNo AND PharmPurchaseOrdersData.ReceivedStatus IS NULL AND PharmPurchaseORders.DatePromised<'" & ThisDate & "' ORDER BY PharmPurchaseOrdersData.SerialNO;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SerialNo))

    If Not IsNull(rsLIST!drugcode) Then
        MyList.SubItems(1) = CStr(rsLIST!drugcode)
    End If
    
    If Not IsNull(rsLIST!drugname) Then
        MyList.SubItems(2) = CStr(rsLIST!drugname)
    End If
    
    If Not IsNull(rsLIST!OrderDate) Then
        MyList.SubItems(3) = CStr(rsLIST!OrderDate)
    End If
    
    If Not IsNull(rsLIST!DatePromised) Then
        MyList.SubItems(4) = CStr(rsLIST!DatePromised)
    End If
    
    If Not IsNull(rsLIST!suppliercode) Then
        MyList.SubItems(5) = CStr(rsLIST!suppliercode)
    End If
    
    If Not IsNull(rsLIST!quantityordered) And Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(6) = Trim(CStr(rsLIST!quantityordered)) & " " & Trim(CStr(rsLIST!QuantityUnits))
    End If
    
    If Not IsNull(rsLIST!DosageCost) Then
        MyList.SubItems(7) = FormatNumber(rsLIST!DosageCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!totalcost) Then
        MyList.SubItems(8) = FormatNumber(rsLIST!totalcost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If IsNull(rsLIST!receivedstatus) Then
        MyList.SubItems(9) = CStr("NO")
    ElseIf Not IsNull(rsLIST!receivedstatus) Then
        If rsLIST!receivedstatus = "Y" Then
            MyList.SubItems(9) = CStr("YES")
        Else
            MyList.SubItems(9) = CStr("NO")
        End If
    End If
    
    rsLIST.MoveNext
    
Wend

.ListView1.ColumnHeaders(8).Alignment = lvwColumnRight
.ListView1.ColumnHeaders(9).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Private Function MemberAdmin() As Boolean
'On Error GoTo Err
    Set rsEditRecord = New ADODB.Recordset
    rsEditRecord.Open "SELECT * FROM AdminUSerRegister WHERE UserName='" & CurrentUserName & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsEditRecord
    If .EOF And .BOF Then
        MemberAdmin = False: GoTo OUTS
    ElseIf IsNull(!usergroup) = True Or !usergroup = "" Then
        MemberAdmin = False: GoTo OUTS
    ElseIf !usergroup = "ADMIN" Then
        MemberAdmin = True: GoTo OUTS
    Else
        MemberAdmin = False: GoTo OUTS
    End If
    End With
OUTS:
    Set rsEditRecord = Nothing
Exit Function
Err:
    ErrorMessage
End Function
Public Sub StoresRequisitionSettings()
'On Error GoTo Err
If Screen.ActiveForm.Name = "frmLogin" Then Exit Sub
With Screen.ActiveForm
    Set rsFindRecord = New ADODB.Recordset
    rsFindRecord.Open "SELECT * FROM AdminIndividualRights WHERE UserName='" & CurrentUserName & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
     
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        If MemberAdmin Then
            .mnuSettings.Enabled = True
            .mnuLEave.Enabled = True
            .mnuManage.Enabled = True
            .mnuShow.Enabled = True
            .mnuPrint.Enabled = True
            
            .mnuFileCancel.Enabled = True
            .mnuLeaveAllowances.Enabled = True
            .mnuLEaveRecalling.Enabled = True
            .mnuManagePayOff.Enabled = True
            .mnuLeaveApplication.Enabled = True
            .mnuLeaveApproval.Enabled = True
            .mnuLeaveAuthorize.Enabled = True
            .mnuLeaveSignOut.Enabled = True
            .mnuSetMonthlyDays.Enabled = True
            .mnuSettingsCoyLEave.Enabled = True
            .mnuSetPublicHolidays.Enabled = True
            .mnuSettingsBalances.Enabled = True
            .mnuSetApproveAuthorize.Enabled = True
            .mnuPrintMonthLeaveSummary.Enabled = True
            
            .Toolbar1.Buttons(1).Enabled = True
            .Toolbar1.Buttons(2).Enabled = True
            .Toolbar1.Buttons(3).Enabled = True
            .Toolbar1.Buttons(4).Enabled = True
            
            Set rsFindRecord = Nothing: Exit Sub
            
        ElseIf Not MemberAdmin Then
        
            .mnuSettings.Enabled = False
            .mnuLEave.Enabled = False
            .mnuManage.Enabled = False
            .mnuShow.Enabled = False
            .mnuPrint.Enabled = False
            
            .Toolbar1.Buttons(1).Enabled = False
            .Toolbar1.Buttons(2).Enabled = False
            .Toolbar1.Buttons(3).Enabled = False
            .Toolbar1.Buttons(4).Enabled = False
            
            Set rsFindRecord = Nothing: Exit Sub
            
        End If
    Else
        GoTo TestForRights
    End If
    
TestForRights:
    If MemberAdmin Then
        .mnuSettings.Enabled = True
        .mnuLEave.Enabled = True
        .mnuManage.Enabled = True
        .mnuShow.Enabled = True
        .mnuPrint.Enabled = True
        
        .mnuFileCancel.Enabled = True
        .mnuLeaveAllowances.Enabled = True
        .mnuLEaveRecalling.Enabled = True
        .mnuManagePayOff.Enabled = True
        .mnuLeaveApplication.Enabled = True
        .mnuLeaveApproval.Enabled = True
        .mnuLeaveAuthorize.Enabled = True
        .mnuLeaveSignOut.Enabled = True
        .mnuSetMonthlyDays.Enabled = True
        .mnuSettingsCoyLEave.Enabled = True
        .mnuSetPublicHolidays.Enabled = True
        .mnuSettingsBalances.Enabled = True
        .mnuSetApproveAuthorize.Enabled = True
        .mnuPrintMonthLeaveSummary.Enabled = True
        
        .Toolbar1.Buttons(1).Enabled = True
        .Toolbar1.Buttons(2).Enabled = True
        .Toolbar1.Buttons(3).Enabled = True
        .Toolbar1.Buttons(4).Enabled = True
        
        Set rsFindRecord = Nothing: Exit Sub
        
    ElseIf Not MemberAdmin Then
        .mnuSettings.Enabled = rsFindRecord!tsetting
        .mnuLEave.Enabled = rsFindRecord!tapplication
        .mnuManage.Enabled = rsFindRecord!tmanage
        .mnuShow.Enabled = rsFindRecord!tview
        .mnuPrint.Enabled = rsFindRecord!tprint
        
        .mnuFileCancel.Enabled = rsFindRecord!smcancel
        .mnuLeaveAllowances.Enabled = rsFindRecord!smallow
        .mnuLEaveRecalling.Enabled = rsFindRecord!smrecall
        .mnuManagePayOff.Enabled = rsFindRecord!smpayoff
        
        .mnuLeaveApplication.Enabled = rsFindRecord!sanew
        .mnuLeaveApproval.Enabled = rsFindRecord!saapprove
        .mnuLeaveAuthorize.Enabled = rsFindRecord!saauthorize
        .mnuLeaveSignOut.Enabled = rsFindRecord!sasigning
            
        .mnuSetMonthlyDays.Enabled = rsFindRecord!ssmonthly
        .mnuSettingsCoyLEave.Enabled = rsFindRecord!ssdefaults
        .mnuSetPublicHolidays.Enabled = rsFindRecord!ssholiday
        .mnuSettingsBalances.Enabled = rsFindRecord!sscontrol
        .mnuSetApproveAuthorize.Enabled = rsFindRecord!ssapproval
            
        .mnuPrintMonthLeaveSummary.Enabled = rsFindRecord!spsummary
        
        .Toolbar1.Buttons(1).Enabled = rsFindRecord!sanew
        .Toolbar1.Buttons(2).Enabled = rsFindRecord!saapprove
        .Toolbar1.Buttons(3).Enabled = rsFindRecord!saauthorize
        .Toolbar1.Buttons(4).Enabled = rsFindRecord!sasigning
        
        Set rsFindRecord = Nothing: Exit Sub
    End If
        
End With
Exit Sub
Err:
If Err.Number = 13 Then Resume Next
    ErrorMessage
End Sub

Public Sub FindAllProductsSoldToday()
'On Error GoTo Err
With SchedulingMain


.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Receipt No", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Product Code", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Product Name", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Quantity Sold", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Quantity Units", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Unit Price", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Total Price", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Date Paid", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Customer Name", .ListView1.Width / 3.5
.ListView1.ColumnHeaders.Add , , "Sold By", .ListView1.Width / 3



.ListView1.View = lvwReport

Dim ThisDate As String

ThisDate = Format(Date, "MMMM dd,yyyy")

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM PharmPointOfSale WHERE PayStatus='" & "Y" & "' AND datecreated='" & Trim(ThisDate) & "' ORDER BY drugname;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: .txtTotal.Text = 0: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ReceiptNo))
    
    If Not IsNull(rsLIST!drugcode) Then
        MyList.SubItems(1) = CStr(rsLIST!drugcode)
    End If
    
    If Not IsNull(rsLIST!drugname) Then
        MyList.SubItems(2) = CStr(rsLIST!drugname)
    End If
    
    If Not IsNull(rsLIST!quantity) Then
        MyList.SubItems(3) = CStr(rsLIST!quantity)
    End If
    
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(4) = CStr(rsLIST!QuantityUnits)
    End If
    
    If Not IsNull(rsLIST!DosagePrice) Then
        MyList.SubItems(5) = FormatNumber(rsLIST!DosagePrice, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!TotalPrice) Then
        MyList.SubItems(6) = FormatNumber(rsLIST!TotalPrice, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    
    If Not IsNull(rsLIST!DatePaid) Then
       MyList.SubItems(7) = CDate(rsLIST!DatePaid)
    End If
    
    If Not IsNull(rsLIST!PaidBy) Then
      MyList.SubItems(8) = CStr(rsLIST!PaidBy)
   End If
   
   If Not IsNull(rsLIST!AllNames) Then
     MyList.SubItems(9) = CStr(rsLIST!AllNames)
  End If
    
    rsLIST.MoveNext
    
Wend

.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
.txtTotal.Text = .ListView1.ListItems.Count & "Sales"

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub


Public Sub FindProductsSoldOnSpecifiedDate()
'On Error GoTo Err
With SchedulingMain


.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Receipt No", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Product Code", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Product Name", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Quantity Sold", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Quantity Units", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Unit Price", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Total Price", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Date Paid", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Customer Name", .ListView1.Width / 3.5
.ListView1.ColumnHeaders.Add , , "Sold By", .ListView1.Width / 3
.ListView1.View = lvwReport

INPQRY = InputBox("Please Enter the Date on Which to Show Products Sold..." & vbCrLf & vbCrLf & "Format: - dd/MM/yyyy", "Enter Date", Date)

If Len(INPQRY) = 0 Then
    MsgBox "No Values Entered or the Operation has been Cancelled!!", vbCritical + vbOKOnly, "No Values"
    Exit Sub
End If

Dim q As Date, ThisDate As String

q = CDate(INPQRY): ThisDate = Format(q, "MMMM dd,yyyy")

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM PharmPointofsale WHERE PayStatus='" & "Y" & "'AND datecreated='" & Trim(ThisDate) & "' ORDER BY drugname;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: .txtTotal.Text = 0: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ReceiptNo))
    
    If Not IsNull(rsLIST!drugcode) Then
        MyList.SubItems(1) = CStr(rsLIST!drugcode)
    End If
    
    If Not IsNull(rsLIST!drugname) Then
        MyList.SubItems(2) = CStr(rsLIST!drugname)
    End If
    
    If Not IsNull(rsLIST!quantity) Then
        MyList.SubItems(3) = CStr(rsLIST!quantity)
    End If
    
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(4) = CStr(rsLIST!QuantityUnits)
    End If
    
    If Not IsNull(rsLIST!DosagePrice) Then
        MyList.SubItems(5) = FormatNumber(rsLIST!DosagePrice, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!TotalPrice) Then
        MyList.SubItems(6) = FormatNumber(rsLIST!TotalPrice, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    
    If Not IsNull(rsLIST!DatePaid) Then
       MyList.SubItems(7) = CDate(rsLIST!DatePaid)
    End If
    
    If Not IsNull(rsLIST!PaidBy) Then
      MyList.SubItems(8) = CStr(rsLIST!PaidBy)
   End If
   
   If Not IsNull(rsLIST!AllNames) Then
     MyList.SubItems(9) = CStr(rsLIST!AllNames)
  End If
    
    rsLIST.MoveNext
    
Wend

.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
.txtTotal.Text = .ListView1.ListItems.Count & "Sales"

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub FindSalesByProductName()
'On Error GoTo Err
With SchedulingMain


.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Receipt No", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Product Code", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Product Name", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Quantity Sold", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Quantity Units", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Unit Price", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Total Price", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Date Paid", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Customer Name", .ListView1.Width / 3.5
.ListView1.ColumnHeaders.Add , , "Sold By", .ListView1.Width / 3

.ListView1.View = lvwReport
Dim ThisDate As String
ThisDate = Format(Date, "MMMM dd,yyyy")

INPQRY = InputBox("Please Enter the Product Name to Show Sales...", "Enter Product Name")

If Len(INPQRY) = 0 Then
    MsgBox "No Values Entered or the Operation has been Cancelled!!", vbCritical + vbOKOnly, "No Values"
    Exit Sub
End If

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM PharmPointOfSale WHERE DrugName LIKE '" & Trim(INPQRY) & "%' AND datecreated='" & ThisDate & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: .txtTotal.Text = 0: Exit Sub
End If


While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ReceiptNo))
    
    If Not IsNull(rsLIST!drugcode) Then
        MyList.SubItems(1) = CStr(rsLIST!drugcode)
    End If
    
    If Not IsNull(rsLIST!drugname) Then
        MyList.SubItems(2) = CStr(rsLIST!drugname)
    End If
    
    If Not IsNull(rsLIST!quantity) Then
        MyList.SubItems(3) = CStr(rsLIST!quantity)
    End If
    
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(4) = CStr(rsLIST!QuantityUnits)
    End If
    
    If Not IsNull(rsLIST!DosagePrice) Then
        MyList.SubItems(5) = FormatNumber(rsLIST!DosagePrice, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!TotalPrice) Then
        MyList.SubItems(6) = FormatNumber(rsLIST!TotalPrice, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    
    If Not IsNull(rsLIST!DatePaid) Then
       MyList.SubItems(7) = CDate(rsLIST!DatePaid)
    End If
    
    If Not IsNull(rsLIST!PaidBy) Then
      MyList.SubItems(8) = CStr(rsLIST!PaidBy)
   End If
   
   If Not IsNull(rsLIST!AllNames) Then
     MyList.SubItems(9) = CStr(rsLIST!AllNames)
  End If
    
    rsLIST.MoveNext
    
Wend

.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight

.txtTotal.Text = .ListView1.ListItems.Count & "Sales"

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub


Public Sub FindSalesByProductNameByDate()
'On Error GoTo Err
With SchedulingMain


.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Receipt No", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Product Code", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Product Name", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Quantity Sold", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Quantity Units", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Unit Price", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Total Price", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Date Paid", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Customer Name", .ListView1.Width / 3.5
.ListView1.ColumnHeaders.Add , , "Sold By", .ListView1.Width / 3

.ListView1.View = lvwReport


INPQRY = InputBox("Please Enter the Product Name to Show Sales...", "Enter Product Name")

If Len(INPQRY) = 0 Then
    MsgBox "No Values Entered or the Operation has been Cancelled!!", vbCritical + vbOKOnly, "No Values"
    Exit Sub
End If

INPQR = InputBox("Please Enter the Date on Which to Show Products Sold..." & vbCrLf & vbCrLf & "Format: - dd/MM/yyyy", "Enter Date", Date)

If Len(INPQR) = 0 Then
    MsgBox "No Values Entered or the Operation has been Cancelled!!", vbCritical + vbOKOnly, "No Values"
    Exit Sub
End If

Dim q As Date, ThisDate As String

q = CDate(INPQR): ThisDate = Format(q, "MMMM dd,yyyy")


Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM PharmPointOfSale WHERE DrugName LIKE '" & Trim(INPQRY) & "%' AND datecreated='" & ThisDate & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: .txtTotal.Text = 0: Exit Sub
End If


While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ReceiptNo))
    
    If Not IsNull(rsLIST!drugcode) Then
        MyList.SubItems(1) = CStr(rsLIST!drugcode)
    End If
    
    If Not IsNull(rsLIST!drugname) Then
        MyList.SubItems(2) = CStr(rsLIST!drugname)
    End If
    
    If Not IsNull(rsLIST!quantity) Then
        MyList.SubItems(3) = CStr(rsLIST!quantity)
    End If
    
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(4) = CStr(rsLIST!QuantityUnits)
    End If
    
    If Not IsNull(rsLIST!DosagePrice) Then
        MyList.SubItems(5) = FormatNumber(rsLIST!DosagePrice, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!TotalPrice) Then
        MyList.SubItems(6) = FormatNumber(rsLIST!TotalPrice, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    
    If Not IsNull(rsLIST!DatePaid) Then
       MyList.SubItems(7) = CDate(rsLIST!DatePaid)
    End If
    
    If Not IsNull(rsLIST!PaidBy) Then
      MyList.SubItems(8) = CStr(rsLIST!PaidBy)
   End If
   
   If Not IsNull(rsLIST!AllNames) Then
     MyList.SubItems(9) = CStr(rsLIST!AllNames)
  End If
    
    rsLIST.MoveNext
    
Wend

.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight

.txtTotal.Text = .ListView1.ListItems.Count & "Sales"

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub ShowActiveCustomers()
'On Error GoTo Err
With SchedulingMain


.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Customer Name", .ListView1.Width / 1.5
.ListView1.ColumnHeaders.Add , , "Frequencry", .ListView1.Width / 5.5


Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT DISTINCT(PaidBy) as CustomerName,COUNT(PaidBy) as Frequency FROM Pharmpointofsale GROUP BY paidby ORDER BY Frequency DESC ", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: .txtTotal.Text = 0: Exit Sub
End If


While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!CustomerName))
    
    If Not IsNull(rsLIST!Frequency) Then
        MyList.SubItems(1) = CStr(rsLIST!Frequency)
    End If
    
   
    
    rsLIST.MoveNext
    
Wend


.txtTotal.Text = .ListView1.ListItems.Count & "Customers"

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub


    
    

