VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPurchaseOrdersMake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ORDate As Variant, ReqDate As Variant, PromDate As Variant

Public Sub AddNewRecord()
If EditRecord Then Exit Sub
    NewRecord = True
    With frmPharmacyPurchaseOrders
    
        Call ClearForNewDrug
        
        .cmdSAVE.Caption = "&SAVE RECORD"
        .cmdFinish.Enabled = False
        
        If .txtOrderNO.Text = Empty Then
            .txtOrderNO.SetFocus
        Else
            .cboDrugOrdered.SetFocus
        End If
        
    End With
End Sub

Public Sub SavePurchaseData()
'On Error GoTo Err
With frmPharmacyPurchaseOrders
Dim DateRequired As Date
Dim DatePromised As Date
    DateRequired = Format(.txtDateRequired.Text, "MMMM dd,yyyy")
    DatePromised = Format(.txtDatePromised.Text, "MMMM dd,yyyy")
    
    Set rsNewRecord = New ADODB.Recordset
    
    MySQL = "INSERT INTO PharmPurchaseOrdersData(serialno,purchaseorderno,drugcode,drugname,quantityordered,quantityunits,dosagecost,totalcost,createdby,datecreated,accperiod) VALUES('" & AutoSerialNo & "','" & .txtOrderNO.Text & "','" & .txtDrugCode.Text & "','" & .cboDrugOrdered.Text & "','" & .txtTotalDoses.Text & "','" & .cboUnits.Text & "'," & CCur(.txtDosageCost.Text) & "," & CCur(.txtTotalCost.Text) & ",'" & CurrentUserName & "','" & MyCurrentDate & "','" & MyCurrentPeriod & "')"
    
    rsNewRecord.Open MySQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
    Set rsNewRecord = Nothing
    
    
    Set rsLineUpdate = New ADODB.Recordset
        
    rsLineUpdate.Open "UPDATE GenProductsInventory SET ReorderStatus='" & "Y" & "',POrderStatus='" & "Y" & "',PurchaseOrderNo= '" & .txtOrderNO & "',InPOS='" & "Y" & "',POrderDate='" & MyCurrentDate & "' WHERE DrugCode='" & Trim(.txtDrugCode.Text) & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
      
    
    Set rsLineUpdate = Nothing
    
'    If Not OrderAlreadyExists Then
'        Set rsNewRecord = New ADODB.Recordset
'
'        MySQL = "INSERT INTO PharmPurchaseOrders (PurchaseOrderNo,SupplierCode,OrderDescription,OrderDate,DateRequired,DatePromised,ShippingMethodID,FreightCharge,staffidno,createdby,datecreated,accperiod) VALUES('" & .txtOrderNO.Text & "','" & .cboSupplierCode.Text & "','" & .txtOrderDescription.Text & "','" & MyCurrentDate & "'," & DateRequired & "," & DatePromised & "," & .txtShippingMethod.Text & "," & CCur(.txtFreightCharge.Text) & ",'" & .cboStaffIdNo.Text & "','" & CurrentUserName & "','" & MyCurrentDate & "','" & MyCurrentPeriod & "')"
'
'        rsNewRecord.Open MySQL, cnCOMMON, adOpenKeyset, adLockOptimistic
'
'        Set rsNewRecord = Nothing
'    End If
    
    
    .cmdSAVE.Caption = "&NEXT ITEM"
    .cmdFinish.Enabled = True
    
    
    
    
    .cmdFinish.SetFocus
    
End With

    Exit Sub
    
Err:
    UpdateErrorMessage
End Sub

Private Function OrderAlreadyExists() As Boolean
'On Error GoTo Err
With frmPharmacyPurchaseOrders
    Set rsFindRecord = cnCOMMON.Execute("SELECT COUNT(PurchaseOrderNo) AS TCOunt FROM PharmPurchaseOrders WHERE PurchaseOrderNo='" & Trim(.txtOrderNO.Text) & "';")
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        OrderAlreadyExists = False
    ElseIf IsNull(rsFindRecord!tcount) = True Then
        OrderAlreadyExists = False
    ElseIf rsFindRecord!tcount = 0 Then
        OrderAlreadyExists = False
    ElseIf rsFindRecord!tcount >= 1 Then
        OrderAlreadyExists = True
    End If
    Set rsFindRecord = Nothing
End With

Err:
    ErrorMessage
End Function


Public Sub SaveMainPurchase()
'On Error GoTo Err
With frmPharmacyPurchaseOrders
    
    ORDate = Format(.txtOrderDate.Text, "MMMM dd,yyyy")
    ReqDate = Format(.txtDateRequired.Text, "MMMM dd,yyyy")
    PromDate = Format(.txtDatePromised.Text, "MMMM dd,yyyy")
    
    Set rsNewRecord = New ADODB.Recordset
    
    MySQL = "INSERT INTO PharmPurchaseOrders (PurchaseOrderNo,OrderDescription,companycode,SupplierCode,StaffIDNo,OrderDate,DateRequired,DatePromised,ShippingMEthodID,FreightCharge,createdby,datecreated,accperiod) VALUES('" & .txtOrderNO & "','" & .txtOrderDescription & "','" & GetMyCompanyCode & "','" & .cboSupplierCode & "','" & Trim(CurrentStaffID) & "','" & ORDate & "','" & ReqDate & "','" & PromDate & "','" & .txtShippingMethod & "'," & CCur(.txtFreightCharge.Text) & ",'" & CurrentUserName & "','" & MyCurrentDate & "','" & MyCurrentPeriod & "')"
    
    rsNewRecord.Open MySQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Set rsNewRecord = Nothing
    
    If MsgBox("Current Order Data Entry Successfully Completed!! Do you want to Print the Purchase Order Form?", vbYesNo + vbQuestion + vbDefaultButton1, "Print Order Form?") = vbYes Then
        
        If .txtDataSource.Text = "RE-ORDER PRODUCTS" Then
            Call ShowPurchaseRecords
        ElseIf .txtDataSource.Text = "NEW PRODUCTS " Then
            Call ShowNewProducts
        End If
        
        Load frmRPTPurchaseOrderForm
        frmRPTPurchaseOrderForm.Show 1, frmPharmacyPurchaseOrders
        
    Else
    
        If .txtDataSource.Text = "RE-ORDER PRODUCTS" Then
            Call ShowPurchaseRecords
        ElseIf .txtDataSource.Text = "NEW PRODUCTS" Then
            Call ShowNewProducts
        End If
        
        Exit Sub
    End If
    .cmdSAVE.Caption = "ADD &NEW"
    .cmdFinish.Enabled = False
    NewRecord = False
    
End With

Exit Sub
Err:
    UpdateErrorMessage
End Sub

Private Function GetMyCompanyCode() As String
'On Error GoTo Err

    Set rsFindRecord = cnCOMMON.Execute("SELECT CompanyCode FROM ParamCompanyMaster WHERE CompanyCode IS NOT NULL;")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetMyCompanyCode = Empty
    ElseIf IsNull(rsFindRecord!CompanyCode) Or rsFindRecord!CompanyCode = "" Then
        GetMyCompanyCode = Empty
    Else
        GetMyCompanyCode = rsFindRecord!CompanyCode
    End If
    
    Set rsFindRecord = Nothing
    
Exit Function
Err:
    ErrorMessage
End Function

Private Function CurrentStaffID() As String
'On Error GoTo Err

    Set rsFindRecord = cnCOMMON.Execute("SELECT StaffIDNO FROM AdminUserRegister WHERE UserName='" & Trim(CurrentUserName) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        CurrentStaffID = Empty
    ElseIf IsNull(rsFindRecord!StaffIdNo) Or rsFindRecord!StaffIdNo = "" Then
        CurrentStaffID = Empty
    Else
        CurrentStaffID = rsFindRecord!StaffIdNo
    End If
    
    Set rsFindRecord = Nothing
    
Exit Function
Err:
    ErrorMessage
End Function

Public Sub GetDrugDataByCode()
'On Error GoTo Err
With frmPharmacyPurchaseOrders

    If .cboDrugOrdered.Text = Empty Then Exit Sub
    
    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM ProductsCostPriceSetup A,GenProductsInventory B WHERE A.DrugCode=B.DrugCode AND B.DrugCode='" & Trim(.txtDrugCode.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .txtDrugCode.Text = rsFindRecord!Drugcode & ""
        .txtDosageCost.Text = rsFindRecord!DosageCost & ""
        .txtDosageCost.Text = FormatNumber(.txtDosageCost.Text, 2, vbUseDefault, vbUseDefault, vbTrue)
        .txtTotalDoses.SetFocus
    End If
    
    Set rsFindRecord = Nothing
    
End With
Exit Sub
Err:
    ErrorMessage
End Sub



Public Sub GetSupplierDataByCode()
'On Error GoTo Err
With frmPharmacyPurchaseOrders

    If .cboSupplierCode.Text = Empty Then Exit Sub
    
    Set rsFindRecord = cnCOMMON.Execute("SELECT ParamSuppliers.* FROM ParamSuppliers WHERE ParamSuppliers.SupplierCode='" & Trim(.cboSupplierCode.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .cboSupplierCode.Text = rsFindRecord!suppliercode & ""
        .txtShippingMethod.Text = rsFindRecord!ShippingMethodID & ""
        .txtFreightCharge.Text = rsFindRecord!FreightCharge & ""
        .cboShippingMethod.Text = GetShippingMethod
        .txtFreightCharge.Text = FormatNumber(.txtFreightCharge.Text, 2, vbUseDefault, vbUseDefault, vbTrue)
        .txtDateRequired.SetFocus
    End If
    
    Set rsFindRecord = Nothing
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub GetSupplierCode()
'On Error GoTo Err
With frmPharmacyPurchaseOrders

    If .cboSupplierCode.Text = Empty Then Exit Sub
    
    Set rsFindRecord = cnCOMMON.Execute("SELECT ParamSuppliers.* FROM ParamSuppliers WHERE ParamSuppliers.SupplierName='" & Trim(.cboSupplierCode.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .cboSupplierCode.Text = rsFindRecord!suppliercode & ""
        .txtShippingMethod.Text = rsFindRecord!ShippingMethodID & ""
        .txtFreightCharge.Text = rsFindRecord!FreightCharge & ""
        .cboShippingMethod.Text = GetShippingMethod
        .txtFreightCharge.Text = FormatNumber(.txtFreightCharge.Text, 2, vbUseDefault, vbUseDefault, vbTrue)
        .txtDateRequired.SetFocus
    End If
    
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Function GetShippingMethod() As Variant
'On Error GoTo Err
With frmPharmacyPurchaseOrders
    Set rsFindRecord = cnCOMMON.Execute("SELECT ShippingMethod FROM PharmShippingMethods WHERE ShippingMEthodID='" & Trim(.txtShippingMethod.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetShippingMethod = Empty: Set rsFindRecord = Nothing: Exit Function
    ElseIf IsNull(rsFindRecord!ShippingMethod) = True Or rsFindRecord!ShippingMethod = "" Then
        GetShippingMethod = Empty: Set rsFindRecord = Nothing: Exit Function
    Else
        GetShippingMethod = rsFindRecord!ShippingMethod
    End If
    
    Set rsFindRecord = Nothing
End With
Exit Function
Err:
    ErrorMessage
End Function

Public Sub AttachQuantityUnits()
'On Error GoTo Err
With frmPharmacyPurchaseOrders

    If .cboUnits.ListCount <> 0 Then Exit Sub
    Set rsCOMBO = cnCOMMON.Execute("SELECT descriptions FROM PharmPackagetypes WHERE descriptions IS NOT NULL;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!descriptions) And rsCOMBO!descriptions <> "" Then
            .cboUnits.AddItem rsCOMBO!descriptions
        End If
    rsCOMBO.MoveNext
    Loop
    
OUTS:
    Set rsCOMBO = Nothing
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub AttachDrugsList()
'On Error GoTo Err
With frmPharmacyPurchaseOrders

    If .cboDrugOrdered.ListCount <> 0 Then Exit Sub
    
    Set rsCOMBO = cnCOMMON.Execute("SELECT DrugName FROM GenProductsInventory WHERE DrugName IS NOT NULL AND Discontinued='" & "N" & "' AND ReOrderStatus='" & "Y" & "' AND POrderStatus IS NULL ORDER BY DrugName;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!DrugName) And rsCOMBO!DrugName <> "" Then
            .cboDrugOrdered.AddItem rsCOMBO!DrugName
        End If
    rsCOMBO.MoveNext
    Loop
    
OUTS:
    Set rsCOMBO = Nothing
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub AttachShippingMethods()
'On Error GoTo Err
With frmPharmacyPurchaseOrders

    If .cboShippingMethod.ListCount <> 0 Then Exit Sub
    
    Set rsCOMBO = cnCOMMON.Execute("SELECT ShippingMEthod FROM PharmShippingMEthods WHERE ShippingMEthod IS NOT NULL ORDER BY ShippingMEthod;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!ShippingMethod) And rsCOMBO!ShippingMethod <> "" Then
            .cboShippingMethod.AddItem rsCOMBO!ShippingMethod
        End If
    rsCOMBO.MoveNext
    Loop
    
OUTS:
    Set rsCOMBO = Nothing
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub AttachSuppliers()
'On Error GoTo Err
With frmPharmacyPurchaseOrders

If .cboSupplierCode.ListCount <> 0 Then Exit Sub
    
    Set rsCOMBO = cnCOMMON.Execute("SELECT SupplierName FROM ParamSuppliers WHERE SupplierName IS NOT NULL AND SupplierType='" & "PHAM" & "' ORDER BY SupplierName;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!SupplierName) And rsCOMBO!SupplierName <> "" Then
            .cboSupplierCode.AddItem rsCOMBO!SupplierName
        End If
    rsCOMBO.MoveNext
    Loop
    
OUTS:
    Set rsCOMBO = Nothing
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub ClearSupplierData()
With frmPharmacyPurchaseOrders
    .cboShippingMethod.Clear
    .cboSupplierCode.Clear
    .txtDatePromised.Text = Empty
    .txtDateRequired.Text = Empty
    .txtFreightCharge.Text = Empty
    .txtShippingMethod.Text = Empty
End With
End Sub

Public Sub ClearTheScreen()
With frmPharmacyPurchaseOrders
    .ListView1.ListItems.Clear
    .ListView2.ListItems.Clear
    .cboShippingMethod.Clear
    .cboSupplierCode.Clear
    .chkOrderNo.Value = 1
    .txtDataSource.Text = Empty
    .txtTotalDoses.Text = Empty
    .cboStaffIdNo.Text = Empty
    .txtDatePromised.Text = Empty
    .txtDateRequired.Text = Empty
    .txtFreightCharge.Text = Empty
    .txtOrderDate.Text = Empty
    .txtOrderDescription.Text = Empty
    .txtOrderNO.Text = Empty
    .txtShippingMethod.Text = Empty
    .cboDrugOrdered.Clear
    .cboUnits.Clear
    .txtDrugCode.Text = Empty
'    .txtQuantity.Text = Empty
    .txtDosageCost.Text = Empty
    .txtTotalCost.Text = Empty
End With
End Sub

Public Sub ClearForNewDrug()
With frmPharmacyPurchaseOrders
    .cboDrugOrdered.Clear
    .cboUnits.Clear
    .cboUnits.Locked = False
    .txtDrugCode.Text = Empty
'    .txtQuantity.Text = Empty
    .txtDosageCost.Text = Empty
    .txtTotalCost.Text = Empty
'    .txtDosageType.Text = Empty
    .txtTotalDoses.Text = Empty
    .cboStaffIdNo.Clear
End With
End Sub

Public Sub GetSupplyStructure()
'On Error GoTo Err
With frmPharmacyPurchaseOrders
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Supplier Code", .ListView2.Width / 4.5
.ListView2.ColumnHeaders.Add , , "Supplier Name", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Supply Delay", .ListView2.Width / 4.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Physical Address", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Postal Address", .ListView2.Width / 3.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Town/City", .ListView2.Width / 3.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contact Person", .ListView2.Width / 3 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Official Title", .ListView2.Width / 2 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Shipping Method", .ListView2.Width / 3.5 ', lvwColumnCenter

.ListView2.View = lvwReport
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub ShowPossibleSuppliers()
'On Error GoTo Err
With frmPharmacyPurchaseOrders
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Supplier Code", .ListView2.Width / 4.5
.ListView2.ColumnHeaders.Add , , "Supplier Name", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Supply Delay", .ListView2.Width / 4.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Physical Address", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Postal Address", .ListView2.Width / 3.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Town/City", .ListView2.Width / 3.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contact Person", .ListView2.Width / 3 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Official Title", .ListView2.Width / 2 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Shipping Method", .ListView2.Width / 3.5 ', lvwColumnCenter

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamSuppliers WHERE SupplierType IS NOT NULL ORDER BY SupplierCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!suppliercode))

    If Not IsNull(rsLIST!SupplierName) Then
        MyList.SubItems(1) = CStr(rsLIST!SupplierName)
    End If
    
    If Not IsNull(rsLIST!SupplyDuration) Then
        MyList.SubItems(2) = CStr(rsLIST!SupplyDuration)
    End If
    
    If Not IsNull(rsLIST!physicaladdress) Then
        MyList.SubItems(3) = Trim(CStr(rsLIST!physicaladdress))
    End If
    
    If Not IsNull(rsLIST!postaladdress) Then
        MyList.SubItems(4) = Trim(CStr(rsLIST!postaladdress))
    End If
    
    If Not IsNull(rsLIST!towncity) Then
        MyList.SubItems(5) = CStr(rsLIST!towncity)
    End If
    
    If Not IsNull(rsLIST!ContactPerson) Then
        MyList.SubItems(6) = Trim(CStr(rsLIST!ContactPerson))
    End If
    
    If Not IsNull(rsLIST!contacttitle) Then
        MyList.SubItems(7) = Trim(CStr(rsLIST!contacttitle))
    End If
    
    If Not IsNull(rsLIST!ShippingMethodID) Then
        MyList.SubItems(8) = CStr(rsLIST!ShippingMethodID)
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

Public Sub ShowCurrentOrderSuppliers()
'On Error GoTo Err
With frmPharmacyPurchaseOrders
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Supplier Code", .ListView2.Width / 4.5
.ListView2.ColumnHeaders.Add , , "Supplier Name", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Supply Delay", .ListView2.Width / 4.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Physical Address", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Postal Address", .ListView2.Width / 3.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Town/City", .ListView2.Width / 3.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contact Person", .ListView2.Width / 3 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Official Title", .ListView2.Width / 2 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Shipping Method", .ListView2.Width / 3.5 ', lvwColumnCenter

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT ParamSuppliers.*,PharmPurchaseORders.* FROM ParamSuppliers,PharmPurchaseORders WHERE ParamSuppliers.SupplierCode=PharmPurchaseORders.SupplierCode AND PharmPurchaseORders.PurchaseOrderNo='" & Trim(.txtOrderNO.Text) & "' ORDER BY ParamSuppliers.SupplierCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!suppliercode))

    If Not IsNull(rsLIST!SupplierName) Then
        MyList.SubItems(1) = CStr(rsLIST!SupplierName)
    End If
    
    If Not IsNull(rsLIST!SupplyDuration) Then
        MyList.SubItems(2) = CStr(rsLIST!SupplyDuration)
    End If
    
    If Not IsNull(rsLIST!physicaladdress) Then
        MyList.SubItems(3) = Trim(CStr(rsLIST!physicaladdress))
    End If
    
    If Not IsNull(rsLIST!postaladdress) Then
        MyList.SubItems(4) = Trim(CStr(rsLIST!postaladdress))
    End If
    
    If Not IsNull(rsLIST!towncity) Then
        MyList.SubItems(5) = CStr(rsLIST!towncity)
    End If
    
    If Not IsNull(rsLIST!ContactPerson) Then
        MyList.SubItems(6) = Trim(CStr(rsLIST!ContactPerson))
    End If
    
    If Not IsNull(rsLIST!contacttitle) Then
        MyList.SubItems(7) = Trim(CStr(rsLIST!contacttitle))
    End If
    
    If Not IsNull(rsLIST!ShippingMethodID) Then
        MyList.SubItems(8) = CStr(rsLIST!ShippingMethodID)
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

Public Sub GetPurchaseStructure()
'On Error GoTo Err
With frmPharmacyPurchaseOrders
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Drug Code ", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Drug Name", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Current Quantity", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Re-Order Level", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "R.O.L. Date", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Category Code", .ListView1.Width / 4

.ListView1.View = lvwReport
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub ShowPurchaseRecords()
'On Error GoTo Err
With frmPharmacyPurchaseOrders
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Product Code ", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Product Name", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Current Quantity", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Quantity Units", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Re-Order Level", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "R.O.L. Date", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Category Code", .ListView1.Width / 4

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM GenProductsInventory WHERE Discontinued='" & "N" & "' AND POrderStatus= '" & "Y" & "' AND ReorderStatus='" & "Y" & "' AND CurrentQuantity<=Reorderlevel ORDER BY DrugCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!Drugcode))

    If Not IsNull(rsLIST!DrugName) Then
        MyList.SubItems(1) = CStr(rsLIST!DrugName)
    End If
    
    If Not IsNull(rsLIST!CurrentQuantity) Then
        MyList.SubItems(2) = Trim(CStr(rsLIST!CurrentQuantity))
        End If
        
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(3) = Trim(CStr(rsLIST!QuantityUnits))
        End If
    
    If Not IsNull(rsLIST!reorderlevel) Then
        MyList.SubItems(4) = CStr(rsLIST!reorderlevel)
    End If
    
    If Not IsNull(rsLIST!PorderDate) Then
        MyList.SubItems(5) = CStr(rsLIST!PorderDate)
    End If
    
    If Not IsNull(rsLIST!categorycode) Then
        MyList.SubItems(6) = Trim(CStr(rsLIST!categorycode))
    End If
    
    rsLIST.MoveNext
    
Wend

.txtDataSource.Text = "RE-ORDER PRODUCTS"
Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub ShowNewProducts()
'On Error GoTo Err
With frmPharmacyPurchaseOrders
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Product Code ", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Product Name", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Wholesale Cost", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Package Type", .ListView1.Width / 5 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM GenProductsInventory G, ProductsCostPriceSetup C  WHERE G.Drugcode=C.drugcode AND C.costpricestatus='" & "Y" & "'    AND G.Purchaseorderno IS NULL AND G.Discontinued='" & "N" & "' ORDER BY G.DrugCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!Drugcode))

    If Not IsNull(rsLIST!DrugName) Then
        MyList.SubItems(1) = CStr(rsLIST!DrugName)
    End If
    
    If Not IsNull(rsLIST!DosageCost) Then
        MyList.SubItems(2) = Trim(CStr(rsLIST!DosageCost))
    End If
    
    If Not IsNull(rsLIST!packagetype) Then
        MyList.SubItems(3) = Trim(CStr(rsLIST!packagetype))
    End If
            
    rsLIST.MoveNext
    
Wend

.txtDataSource.Text = "NEW PRODUCTS"
Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub ShowCurrentOrderRecords()
'On Error GoTo Err
With frmPharmacyPurchaseOrders
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Serial No", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Drug Code ", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Drug Name", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Quantity Ordered", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Dosage Cost", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Total Cost", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Received", .ListView1.Width / 5.5 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM PharmPurchaseOrdersData WHERE PurchaseOrderNo='" & Trim(.txtOrderNO.Text) & "' ORDER BY PharmPurchaseOrdersData.SerialNO;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SerialNo))

    If Not IsNull(rsLIST!Drugcode) Then
        MyList.SubItems(1) = CStr(rsLIST!Drugcode)
    End If
    
    If Not IsNull(rsLIST!DrugName) Then
        MyList.SubItems(2) = CStr(rsLIST!DrugName)
    End If
    
    If Not IsNull(rsLIST!QuantityOrdered) And Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(3) = Trim(CStr(rsLIST!QuantityOrdered)) & " " & Trim(CStr(rsLIST!QuantityUnits))
    End If
    
    If Not IsNull(rsLIST!DosageCost) Then
        MyList.SubItems(4) = FormatNumber(rsLIST!DosageCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!TotalCost) Then
        MyList.SubItems(5) = FormatNumber(rsLIST!TotalCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If IsNull(rsLIST!receivedstatus) Then
        MyList.SubItems(6) = CStr("NO")
    ElseIf Not IsNull(rsLIST!receivedstatus) Then
        If rsLIST!receivedstatus = "Y" Then
            MyList.SubItems(6) = CStr("YES")
        Else
            MyList.SubItems(6) = CStr("NO")
        End If
    End If
    
    rsLIST.MoveNext
    
Wend

.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
.ListView1.ColumnHeaders(6).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Function AutoSerialNo() As Variant
'On Error GoTo Err
With frmPharmacyPurchaseOrders

Dim rsLastID As ADODB.Recordset 'used to retrieve current LastId in the Table
Dim strLastID As String 'SQL statement

Dim strTemp As String 'store current record
Dim iNumPos As Integer 'store position of the first numeral
Dim strPrefix As String 'stores Id Prefix

'Retrieve the last record in the recrdset where order is ascending

'strLastID = "SELECT max(cnCliniclDetails.cnCliniclID) as lastid from cnCliniclDetails"
strLastID = "SELECT MAX(serialno) AS LastID FROM PharmPurchaseOrdersData" ';"
Set rsLastID = New ADODB.Recordset

With rsLastID
'open the recordset
    .Open strLastID, cnCOMMON, adOpenKeyset, adLockOptimistic
    If .EOF And .BOF Then 'shows empty recordset
        AutoSerialNo = 1 'format of desired format of the string value
    ElseIf IsNull(!lastid) = True Or !lastid = "" Then
        AutoSerialNo = 1
    Else
       ' If .EOF And .BOF Then .MoveFirst
       ' .MoveLast
        strTemp = !lastid
        iNumPos = 1
        Dim sChar As String
        Dim iIDLen As Integer
        iIDLen = Len(strTemp)
        sChar = Mid(strTemp, iNumPos, 1)
        While InStr("1234567890", sChar) = 0
            iNumPos = iNumPos + 1
            sChar = Mid(strTemp, iNumPos, 1)
        Wend
        'store the ID prefix eg AP
        strPrefix = Left(strTemp, iNumPos - 1)
        'store the number portion eg and the length with leading Zeros
        strTemp = Right(strTemp, Len(strTemp) + 1 - iNumPos)
        strTemp = Format(Int(strTemp) + 1, String(iIDLen + 1 - iNumPos, "0"))
        AutoSerialNo = strPrefix & strTemp
    End If
End With
End With
    Exit Function
Err:
    ErrorMessage
End Function

Public Function AutoPurchaseOrderNo() As String
'On Error GoTo Err
With frmPharmacyPurchaseOrders

Dim rsLastID As ADODB.Recordset 'used to retrieve current LastId in the Table
Dim strLastID As String 'SQL statement

Dim strTemp As String 'store current record
Dim iNumPos As Integer 'store position of the first numeral
Dim strPrefix As String 'stores Id Prefix

'Retrieve the last record in the recrdset where order is ascending

'strLastID = "SELECT max(cnCliniclDetails.cnCliniclID) as lastid from cnCliniclDetails"
strLastID = "SELECT MAX(PurchaseOrderNo) AS LastID FROM PharmPurchaseOrders;"
Set rsLastID = New ADODB.Recordset

With rsLastID
'open the recordset
    .Open strLastID, cnCOMMON, adOpenKeyset, adLockOptimistic
    If .EOF And .BOF Then 'shows empty recordset
        AutoPurchaseOrderNo = "ORD00001" 'format of desired format of the string value
    ElseIf IsNull(!lastid) = True Or !lastid = "" Then
        AutoPurchaseOrderNo = "ORD00001"
    Else
       ' If .EOF And .BOF Then .MoveFirst
        '.MoveLast
        strTemp = !lastid
        iNumPos = 1
        Dim sChar As String
        Dim iIDLen As Integer
        iIDLen = Len(strTemp)
        sChar = Mid(strTemp, iNumPos, 1)
        While InStr("1234567890", sChar) = 0
            iNumPos = iNumPos + 1
            sChar = Mid(strTemp, iNumPos, 1)
        Wend
        'store the ID prefix eg AP
        strPrefix = Left(strTemp, iNumPos - 1)
        'store the number portion eg and the length with leading Zeros
        strTemp = Right(strTemp, Len(strTemp) + 1 - iNumPos)
        strTemp = Format(Int(strTemp) + 1, String(iIDLen + 1 - iNumPos, "0"))
        AutoPurchaseOrderNo = strPrefix & strTemp
    End If
End With
End With
    Exit Function
Err:
    ErrorMessage
End Function

Public Sub FindOrderDetails()
'On Error GoTo Err
With frmPharmacyPurchaseOrders
    Set rsFindRecord = cnCOMMON.Execute("SELECT ParamSuppliers.*,PharmPurchaseOrders.* FROM ParamSuppliers,PharmPurchaseOrders WHERE PharmPurchaseOrders.SupplierCode=ParamSuppliers.SupplierCode AND PharmPurchaseOrders.PurchaseOrderNo='" & Trim(.txtOrderNO.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
    
        Set rsFindRecord = Nothing: Exit Sub
        
    Else
    
        .txtOrderDate.Text = rsFindRecord!OrderDate & ""
        .txtOrderDescription.Text = rsFindRecord!orderdescription & ""
        .txtOrderNO.Text = rsFindRecord!PurchaseOrderNo & ""
        .txtDatePromised.Text = rsFindRecord!DatePromised & ""
        .txtDateRequired.Text = rsFindRecord!DateRequired & ""
        .txtFreightCharge.Text = FormatNumber(rsFindRecord!FreightCharge & "", 2, vbUseDefault, vbUseDefault, vbTrue)
        .txtShippingMethod.Text = rsFindRecord!ShippingMethodID & ""
        .cboSupplierCode.Text = rsFindRecord!suppliercode & ""
        .cboShippingMethod.Text = GetShippingMethod
        
        Call ShowCurrentOrderRecords
        Call ShowCurrentOrderSuppliers
        
    End If
    
End With
Exit Sub
Err:
    ErrorMessage
End Sub
Public Sub ShowReOrderItems()
'On Error GoTo Err
With frmPharmacyPurchaseOrders
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Product Code ", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "PurchaseOrderNO", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Product Name", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "CategoryCode", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "SupplierCode", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Current Quantity", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Purchase Order Date", .ListView1.Width / 4

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM GenProductsInventory WHERE Discontinued='" & "N" & "' AND PurchaseOrderNO IS NOT NULL AND CurrentQuantity <= ReOrderLevel ORDER BY DrugCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. All Products Are Above ReOrder Level")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!Drugcode))

    If Not IsNull(rsLIST!PurchaseOrderNo) Then
        MyList.SubItems(1) = CStr(rsLIST!PurchaseOrderNo)
    End If
    
     If Not IsNull(rsLIST!DrugName) Then
        MyList.SubItems(2) = CStr(rsLIST!DrugName)
    End If
    
     If Not IsNull(rsLIST!categorycode) Then
        MyList.SubItems(3) = CStr(rsLIST!categorycode)
    End If
    
     If Not IsNull(rsLIST!suppliercode) Then
        MyList.SubItems(4) = CStr(rsLIST!suppliercode)
    End If
    
        
    If Not IsNull(rsLIST!CurrentQuantity) And Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(5) = Trim(CStr(rsLIST!CurrentQuantity)) & " " & Trim(CStr(rsLIST!QuantityUnits))
    End If
    
          
    If Not IsNull(rsLIST!PorderDate) Then
        MyList.SubItems(6) = Trim(CStr(rsLIST!PorderDate))
    End If
    
    rsLIST.MoveNext
    
Wend

.txtDataSource.Text = "C"
Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Public Sub ShowCurrentOrderRecordsFromMenu()
'On Error GoTo Err
With frmPharmacyPurchaseOrders
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear
'If Not NewRecord And Not EditRecord Then: Exit Sub
    
    
.ListView1.ColumnHeaders.Add , , "Serial No", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Purchase Order No", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Product Code ", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Product Name", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Quantity Ordered", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Product Cost", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Total Cost", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Received", .ListView1.Width / 5.5 ', lvwColumnCenter


Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM PharmPurchaseOrders A, PharmPurchaseOrdersData B WHERE A.PurchaseOrderNo=B.PurchaseOrderNo AND A.Receivedstatus IS  NULL ORDER BY B.SerialNO;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SerialNo))
    
    If Not IsNull(rsLIST!PurchaseOrderNo) Then
        MyList.SubItems(1) = CStr(rsLIST!PurchaseOrderNo)
    End If
    
    
    If Not IsNull(rsLIST!Drugcode) Then
        MyList.SubItems(2) = CStr(rsLIST!Drugcode)
    End If
    
    If Not IsNull(rsLIST!DrugName) Then
        MyList.SubItems(3) = CStr(rsLIST!DrugName)
    End If
    
    If Not IsNull(rsLIST!QuantityOrdered) And Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(4) = Trim(CStr(rsLIST!QuantityOrdered)) & " " & Trim(CStr(rsLIST!QuantityUnits))
    End If
    
    If Not IsNull(rsLIST!DosageCost) Then
        MyList.SubItems(5) = FormatNumber(rsLIST!DosageCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!TotalCost) Then
        MyList.SubItems(6) = FormatNumber(rsLIST!TotalCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If IsNull(rsLIST!receivedstatus) Then
        MyList.SubItems(7) = CStr("NO")
    ElseIf Not IsNull(rsLIST!receivedstatus) Then
        If rsLIST!receivedstatus = "Y" Then
            MyList.SubItems(7) = CStr("YES")
        Else
            MyList.SubItems(7) = CStr("NO")
        End If
    End If
    
    rsLIST.MoveNext
    
Wend

.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
.ListView1.ColumnHeaders(6).Alignment = lvwColumnRight
.ListView1.View = lvwReport
If .ListView1.ListItems.Count = 0 Or .ListView1.View <> lvwReport Then Exit Sub
    
    Dim i, j, k
    j = .ListView1.ListItems.Count
    
    .ListView1.SelectedItem.Checked = False
    
    If .ListView1.ListItems.Count = 0 Or .ListView1.View <> lvwReport Then Exit Sub
    
    
Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Public Sub AttachPharmacyStaff()
'On Error GoTo Err
With frmPharmacyPurchaseOrders

    If .cboStaffIdNo.ListCount <> 0 Then Exit Sub
    
    Set rsCOMBO = cnCOMMON.Execute("SELECT AllNames FROM AdminUserRegister WHERE staffidno IS NOT NULL AND UserName='" & CurrentUserName & "' ORDER BY AllNames;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!AllNames) And rsCOMBO!AllNames <> "" Then
            .cboStaffIdNo.AddItem rsCOMBO!AllNames
        End If
    rsCOMBO.MoveNext
    Loop
    
OUTS:
    Set rsCOMBO = Nothing
End With
Exit Sub
Err:
    ErrorMessage
End Sub
Public Sub GetStaffID()
'On Error GoTo Err
With frmPharmacyPurchaseOrders

    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM AdminUserRegister WHERE allnames='" & Trim(.cboStaffIdNo.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        
        .cboStaffIdNo.Text = rsFindRecord!StaffIdNo & ""
        
'        .cmdNEW.SetFocus
        
    End If
    
End With
Exit Sub
Err:
    ErrorMessage
End Sub

