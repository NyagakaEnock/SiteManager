VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPurchaseOrdersReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private DTRec As Variant, ManDate As Variant, EXPDate As Variant, DelDate As Variant, InvDate As Variant, QBal As Double, ShipDate As Variant

Public Sub AddNewRecord()
If EditRecord Then Exit Sub
    NewRecord = True
    With frmPharmacyPurchaseOrdersReceive
    
        Call ClearForNewRecord
        
        .cmdSAVE.Caption = "&SAVE RECORD"
'        .cmdCHANGE.Enabled = False
        .txtDateReceived = MyCurrentDate
        
        If .cboOrderNO.Text = Empty Then
            .cboOrderNO.SetFocus
        Else
            .ListView2.SetFocus
        End If
        
    End With
End Sub

Public Sub SaveReceivedData()
'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive
 Dim WQuantity, NWQuantity, UpWQuantity, NUpWQuantity, ItemCost As Variant
    WQuantity = GetWholesaleAmount
    NWQuantity = .txtQuantityReceived.Text * WQuantity
    UpWQuantity = GetCurrentWholesaleAmount
    NUpWQuantity = UpWQuantity + NWQuantity
    ItemCost = .txtDosesReceived.Text
    
 
    
    DTRec = Format(.txtDateReceived.Text, "MMMM dd,yyyy")
    QBal = CDbl(.txtQuantityOrdered.Text) - CDbl(.txtQuantityReceived.Text)
    DelDate = Format(.txtDelDate.Text, "MMMM dd,yyyy")
    InvDate = Format(.txtInvoiceDate.Text, "MMMM dd,yyyy")
'    ShipDate = Format(.txtShippDate.Text, "MMMM dd,yyyy")

    Set rsNewRecord = New ADODB.Recordset
    
    MySQL = "INSERT INTO AccPeriodStock(POrderNo,ProductCode,ProductName,Cost,QuantityReceived,Accperiod,CurrentAccperiod,Datecreated,CreatedBy) VALUES('" & Trim(.cboOrderNO.Text) & "', '" & Trim(.txtDrugCode.Text) & "','" & Trim(.txtDrugName.Text) & "'," & CSng(ItemCost) & ",'" & Trim(.txtQuantityReceived.Text) & "','" & MyCurrentPeriod & "','" & CurrentAccPeriod & "','" & MyCurrentDate & "','" & CurrentUserName & "')"
    
    rsNewRecord.Open MySQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Set rsNewRecord = Nothing
    
UpdatePurchaseData:
      If QBal <> 0 Then
      Set rsEditRecord = New ADODB.Recordset
      MySQL = "UPDATE PharmPurchaseOrders SET BalanceStatus = '" & "Y" & "' WHERE PurchaseOrderNo = '" & Trim(.cboOrderNO.Text) & "';"
      rsEditRecord.Open MySQL, cnCOMMON, adOpenKeyset, adLockOptimistic
      
      MySQL = "UPDATE PharmPurchaseOrdersData SET BalanceStatus = '" & "Y" & "' WHERE PurchaseOrderNo = '" & Trim(.cboOrderNO.Text) & "' AND DrugCode = '" & Trim(.txtDrugCode.Text) & "';"
      rsEditRecord.Open MySQL, cnCOMMON, adOpenKeyset, adLockOptimistic
      
      Else:  Set rsEditRecord = New ADODB.Recordset: MySQL = "UPDATE PharmPurchaseOrders SET BalanceStatus = NULL WHERE PurchaseOrderNo = '" & Trim(.cboOrderNO.Text) & "';"
      rsEditRecord.Open MySQL, cnCOMMON, adOpenKeyset, adLockOptimistic
      
      MySQL = "UPDATE PharmPurchaseOrdersData SET BalanceStatus = NULL WHERE PurchaseOrderNo = '" & Trim(.cboOrderNO.Text) & "' AND DrugCode = '" & Trim(.txtDrugCode.Text) & "';"
      rsEditRecord.Open MySQL, cnCOMMON, adOpenKeyset, adLockOptimistic
      End If
      Set rsEditRecord = Nothing
    
      
      
      Set rsEditRecord = New ADODB.Recordset
      MySQL = "UPDATE PharmPurchaseOrdersData SET ReceivedStatus='" & "Y" & "',quantityreceived=" & CDec(.txtQuantityReceived.Text) & ",dosesreceived=" & CDbl(.txtDosesReceived.Text) & ",datereceived='" & DTRec & "',balance=" & QBal & " ,delnoteno='" & .txtDeliveryNote & "',deldate='" & DelDate & "',invoiceno='" & .txtInvoiceNumber & "',invoicedate='" & InvDate & "',Shippingdate='" & ShipDate & "' WHERE SerialNO='" & Trim(.txtSerialNo.Text) & "' AND PurchaseORderNO='" & Trim(.cboOrderNO.Text) & "';"
      rsEditRecord.Open MySQL, cnCOMMON, adOpenKeyset, adLockOptimistic
      Set rsEditRecord = Nothing
    
UpdateInventoryRecord:

    Dim CQty As Double, NQty As Double, UPQty As Double, DSQty As Double, QUnits
    NQty = CDbl(.txtQuantityReceived.Text)
    
    Set rsFindRecord = cnCOMMON.Execute("SELECT CurrentQuantity FROM GenProductsInventory WHERE DrugCode='" & Trim(.txtDrugCode.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: GoTo UpdateMainPurchaseOrder
    End If
    If IsNull(rsFindRecord!CurrentQuantity) = True Or rsFindRecord!CurrentQuantity = "" Then
        CQty = 0
    Else
        CQty = CDbl(rsFindRecord!CurrentQuantity)
'        CDoses = CDbl(rsFindRecord!totaldoses)
    End If
    
    UPQty = CDbl(CQty + NQty)
'    UpDoses = CDbl(CDoses + NDoses)
'
    Dim i, j: j = .ListView2.ListItems.Count
    For i = 1 To j
        If .ListView2.ListItems(i).Checked = True Then
            QUnits = CStr(.ListView2.ListItems(i).SubItems(4))
'            DSQty = CDbl(.ListView2.ListItems(i).SubItems(5))
        End If
    Next i
    
    Set rsFindRecord = Nothing
    
    Set rsLineUpdate = New ADODB.Recordset
    rsLineUpdate.Open "UPDATE GenProductsInventory SET TotalPieces = '" & NUpWQuantity & "', CurrentQuantity=" & UPQty & ",quantityunits='" & QUnits & "',ReOrderStatus=NULL,POrderDate=NULL,POrderStatus=NULL,PurchaseOrderNo='" & Trim(.cboOrderNO.Text) & "',ReOrderLevel=" & CDbl(GetReOrderLevel) & " WHERE DrugCode='" & Trim(.txtDrugCode.Text) & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    Set rsLineUpdate = Nothing
    
UpdateMainPurchaseOrder:

    If AllItemsReceived Then
        Set rsLineUpdate = New ADODB.Recordset
        rsLineUpdate.Open "UPDATE PharmPurchaseOrders SET ReceivedStatus='" & "Y" & "',ReceivedDate='" & DTRec & "' WHERE PurchaseORderNo='" & Trim(.cboOrderNO.Text) & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        .cboOrderNO.Clear: .txtOrderDate.Text = Empty: .txtOrderDescription.Text = Empty
        Set rsLineUpdate = Nothing
    End If
    
    .cmdSAVE.Caption = "&NEW"
'    .cmdCHANGE.Enabled = True
    
    Call ShowCurrentPendingOrdersData
    
    NewRecord = False
    
End With

    Exit Sub
    
Err:
    UpdateErrorMessage
End Sub

Private Function GetReOrderLevel() As Variant

'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive
    Set rsFindRecord = cnCOMMON.Execute("SELECT ReOrderLevel FROM GeneralProducts WHERE DrugCode = '" & Trim(.txtDrugCode.Text) & "' ;")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetReOrderLevel = 0: Set rsFindRecord = Nothing
    ElseIf IsNull(rsFindRecord!reorderlevel) = True Or rsFindRecord!reorderlevel = "" Then
        GetReOrderLevel = 0: Set rsFindRecord = Nothing
    Else
        GetReOrderLevel = CDbl(rsFindRecord!reorderlevel)
    End If
        
    Set rsFindRecord = Nothing
Exit Function
Err:
    ErrorMessage
    End With
End Function
Private Function GetWholesaleAmount() As Variant

'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive
    Set rsFindRecord = cnCOMMON.Execute("SELECT WholeSaleQuantity FROM GeneralProducts WHERE DrugCode = '" & Trim(.txtDrugCode.Text) & "' ;")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetWholesaleAmount = 0: Set rsFindRecord = Nothing
    ElseIf IsNull(rsFindRecord!WholeSaleQuantity) = True Or rsFindRecord!WholeSaleQuantity = "" Then
        GetWholesaleAmount = 0: Set rsFindRecord = Nothing
    Else
        GetWholesaleAmount = CDbl(rsFindRecord!WholeSaleQuantity)
    End If
        
    Set rsFindRecord = Nothing
Exit Function
Err:
    ErrorMessage
    End With
End Function
Private Function GetCurrentWholesaleAmount() As Variant

'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive
    Set rsFindRecord = cnCOMMON.Execute("SELECT TotalPieces FROM GenProductsInventory WHERE DrugCode = '" & Trim(.txtDrugCode.Text) & "' ;")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetCurrentWholesaleAmount = 0: Set rsFindRecord = Nothing
    ElseIf IsNull(rsFindRecord!TotalPieces) = True Or rsFindRecord!TotalPieces = "" Then
        GetCurrentWholesaleAmount = 0: Set rsFindRecord = Nothing
    Else
        GetCurrentWholesaleAmount = CDbl(rsFindRecord!TotalPieces)
    End If
        
    Set rsFindRecord = Nothing
Exit Function
Err:
    ErrorMessage
    End With
End Function
Private Function AllItemsReceived() As Boolean
'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive

    Set rsFindRecord = cnCOMMON.Execute("SELECT COUNT(SerialNo) AS TItems FROM PharmPurchaseORdersData WHERE PurchaseOrderNO='" & Trim(.cboOrderNO.Text) & "' AND ReceivedStatus IS NULL;")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        AllItemsReceived = True
    ElseIf rsFindRecord!titems = 0 Then
        AllItemsReceived = True
    ElseIf rsFindRecord!titems >= 1 Then
        AllItemsReceived = False
    End If
    
    Set rsFindRecord = Nothing
    
End With
Exit Function
Err:
    ErrorMessage
End Function

Public Sub AttachNewPurchaseOrders()
'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive

    If .cboOrderNO.ListCount <> 0 Then Exit Sub
    
    Set rsCOMBO = cnCOMMON.Execute("SELECT PurchaseOrderNo FROM PharmPurchaseORders WHERE PurchaseOrderNo IS NOT NULL AND ReceivedStatus IS NULL ORDER BY PurchaseOrderNo;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!PurchaseOrderNo) And rsCOMBO!PurchaseOrderNo <> "" Then
            .cboOrderNO.AddItem rsCOMBO!PurchaseOrderNo
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

Public Sub FindDetailsByPurchaseOrderNo()
'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive
    Set rsFindRecord = cnCOMMON.Execute("SELECT ParamSuppliers.*,PharmPurchaseOrders.* FROM ParamSuppliers,PharmPurchaseOrders WHERE PharmPurchaseOrders.SupplierCode=ParamSuppliers.SupplierCode AND PharmPurchaseOrders.PurchaseOrderNo='" & Trim(.cboOrderNO.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
    
        Set rsFindRecord = Nothing: Exit Sub
        
    Else
        .cboOrderNO.Text = rsFindRecord!PurchaseOrderNo & ""
        .txtOrderDate.Text = rsFindRecord!OrderDate & ""
        .txtOrderDescription.Text = rsFindRecord!orderdescription & ""
        
        Call ShowCurrentPurchaseOrder
        Call ShowCurrentPendingOrdersData
        
        .txtDeliveryNote.SetFocus
    End If
    
    Set rsFindRecord = Nothing
    
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub GetPurchaseStructure()
'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Order No", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Order Date", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Date Required", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Date Promised", .ListView1.Width / 4.8
.ListView1.ColumnHeaders.Add , , "Supplier Code", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Shipping Method", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Freight Charge", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Staff ID NO", .ListView1.Width / 4.5

.ListView1.View = lvwReport
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub ShowPurchaseOrdersOverDue()
'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Order No", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Order Date", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Date Required", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Date Promised", .ListView1.Width / 4.8
.ListView1.ColumnHeaders.Add , , "Supplier Code", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Shipping Method", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Freight Charge", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Staff ID NO", .ListView1.Width / 4.5

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset, TDate
TDate = Format(Date, "MMMM dd,yyyy")

Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT PharmPurchaseOrders.* FROM PharmPurchaseOrders WHERE PharmPurchaseOrders.DatePromised<'" & Trim(TDate) & "' AND PharmPurchaseOrders.ReceivedStatus IS NULL ORDER BY PharmPurchaseOrders.PurchaseORderNo;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "There are No Overdue Order Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PurchaseOrderNo))

    If Not IsNull(rsLIST!OrderDate) Then
        MyList.SubItems(1) = CStr(rsLIST!OrderDate)
    End If
    
    If Not IsNull(rsLIST!DateRequired) Then
        MyList.SubItems(2) = CStr(rsLIST!DateRequired)
    End If
    
    If Not IsNull(rsLIST!DatePromised) Then
        MyList.SubItems(3) = CStr(rsLIST!DatePromised)
    End If
    
    If Not IsNull(rsLIST!suppliercode) Then
        MyList.SubItems(4) = CStr(rsLIST!suppliercode)
    End If
    
    If Not IsNull(rsLIST!ShippingMethodID) Then
        MyList.SubItems(5) = Trim(CStr(rsLIST!ShippingMethodID))
    End If
    
    If Not IsNull(rsLIST!FreightCharge) Then
        MyList.SubItems(6) = FormatNumber(rsLIST!FreightCharge, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!StaffIdNo) Then
        MyList.SubItems(7) = Trim(CStr(rsLIST!StaffIdNo))
    End If
    
    rsLIST.MoveNext
    
Wend

.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub ShowPurchaseOrdersDueToday()
'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Order No", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Order Date", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Date Required", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Date Promised", .ListView1.Width / 4.8
.ListView1.ColumnHeaders.Add , , "Supplier Code", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Shipping Method", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Freight Charge", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Staff ID NO", .ListView1.Width / 4.5

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset, TDate
TDate = Format(Date, "MMMM dd,yyyy")

Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT PharmPurchaseOrders.* FROM PharmPurchaseOrders WHERE PharmPurchaseOrders.DatePromised='" & Trim(TDate) & "' AND PharmPurchaseOrders.ReceivedStatus IS NULL ORDER BY PharmPurchaseOrders.PurchaseORderNo;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "There are No Order Records Expected Today to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PurchaseOrderNo))

    If Not IsNull(rsLIST!OrderDate) Then
        MyList.SubItems(1) = CStr(rsLIST!OrderDate)
    End If
    
    If Not IsNull(rsLIST!DateRequired) Then
        MyList.SubItems(2) = CStr(rsLIST!DateRequired)
    End If
    
    If Not IsNull(rsLIST!DatePromised) Then
        MyList.SubItems(3) = CStr(rsLIST!DatePromised)
    End If
    
    If Not IsNull(rsLIST!suppliercode) Then
        MyList.SubItems(4) = CStr(rsLIST!suppliercode)
    End If
    
    If Not IsNull(rsLIST!ShippingMethodID) Then
        MyList.SubItems(5) = Trim(CStr(rsLIST!ShippingMethodID))
    End If
    
    If Not IsNull(rsLIST!FreightCharge) Then
        MyList.SubItems(6) = FormatNumber(rsLIST!FreightCharge, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!StaffIdNo) Then
        MyList.SubItems(7) = Trim(CStr(rsLIST!StaffIdNo))
    End If
    
    rsLIST.MoveNext
    
Wend

.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub ShowNewPurchaseOrders()
'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Order No", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Order Date", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Date Required", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Date Promised", .ListView1.Width / 4.8
.ListView1.ColumnHeaders.Add , , "Supplier Code", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Shipping Method", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Freight Charge", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Staff ID NO", .ListView1.Width / 4.5

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM PharmPurchaseOrders WHERE ReceivedStatus IS NULL  ORDER BY PurchaseOrderNo;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "There are No New Order Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PurchaseOrderNo))

    If Not IsNull(rsLIST!OrderDate) Then
        MyList.SubItems(1) = CStr(rsLIST!OrderDate)
    End If
    
    If Not IsNull(rsLIST!DateRequired) Then
        MyList.SubItems(2) = CStr(rsLIST!DateRequired)
    End If
    
    If Not IsNull(rsLIST!DatePromised) Then
        MyList.SubItems(3) = CStr(rsLIST!DatePromised)
    End If
    
    If Not IsNull(rsLIST!suppliercode) Then
        MyList.SubItems(4) = CStr(rsLIST!suppliercode)
    End If
    
    If Not IsNull(rsLIST!ShippingMethodID) Then
        MyList.SubItems(5) = Trim(CStr(rsLIST!ShippingMethodID))
    End If
    
    If Not IsNull(rsLIST!FreightCharge) Then
        MyList.SubItems(6) = FormatNumber(rsLIST!FreightCharge, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!StaffIdNo) Then
        MyList.SubItems(7) = Trim(CStr(rsLIST!StaffIdNo))
    End If
    
    rsLIST.MoveNext
    
Wend

.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub ShowCurrentPurchaseOrder()
'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Order No", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Order Date", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Date Required", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Date Promised", .ListView1.Width / 4.8
.ListView1.ColumnHeaders.Add , , "Supplier Code", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Shipping Method", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Freight Charge", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Staff ID NO", .ListView1.Width / 4.5

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT PharmPurchaseOrders.* FROM PharmPurchaseOrders WHERE PharmPurchaseOrders.PurchaseORderNo='" & Trim(.cboOrderNO.Text) & "' ORDER BY PharmPurchaseOrders.PurchaseORderNo;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PurchaseOrderNo))

    If Not IsNull(rsLIST!OrderDate) Then
        MyList.SubItems(1) = CStr(rsLIST!OrderDate)
    End If
    
    If Not IsNull(rsLIST!DateRequired) Then
        MyList.SubItems(2) = CStr(rsLIST!DateRequired)
    End If
    
    If Not IsNull(rsLIST!DatePromised) Then
        MyList.SubItems(3) = CStr(rsLIST!DatePromised)
    End If
    
    If Not IsNull(rsLIST!suppliercode) Then
        MyList.SubItems(4) = CStr(rsLIST!suppliercode)
    End If
    
    If Not IsNull(rsLIST!ShippingMethodID) Then
        MyList.SubItems(5) = Trim(CStr(rsLIST!ShippingMethodID))
    End If
    
    If Not IsNull(rsLIST!FreightCharge) Then
        MyList.SubItems(6) = FormatNumber(rsLIST!FreightCharge, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!StaffIdNo) Then
        MyList.SubItems(7) = Trim(CStr(rsLIST!StaffIdNo))
    End If
    
    rsLIST.MoveNext
    
Wend

.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub ShowCurrentPendingOrdersData()
'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Serial No", .ListView2.Width / 6
.ListView2.ColumnHeaders.Add , , "Product Code ", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Product Name", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Quantity Ordered", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Qty Units", .ListView2.Width / 6 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Product Cost", .ListView2.Width / 5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Total Cost", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Received?", .ListView2.Width / 5.4 ', lvwColumnCenter

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM PharmPurchaseOrdersData WHERE PurchaseOrderNo='" & Trim(.cboOrderNO.Text) & "' AND ReceivedStatus IS NULL ORDER BY SerialNO;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!SerialNo))

    If Not IsNull(rsLIST!Drugcode) Then
        MyList.SubItems(1) = CStr(rsLIST!Drugcode)
    End If
    
    If Not IsNull(rsLIST!DrugName) Then
        MyList.SubItems(2) = CStr(rsLIST!DrugName)
    End If
    
    If Not IsNull(rsLIST!QuantityOrdered) Then
        MyList.SubItems(3) = CStr(rsLIST!QuantityOrdered)
    End If
    
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(4) = CStr(rsLIST!QuantityUnits)
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

.ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
.ListView2.ColumnHeaders(7).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub ShowCurrentOrderRecords()
'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Serial No", .ListView2.Width / 6
.ListView2.ColumnHeaders.Add , , "Product Code ", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Product Name", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Quantity Ordered", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Qty Units", .ListView2.Width / 6 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Unit Cost", .ListView2.Width / 6 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Total Cost", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Received", .ListView2.Width / 5.5 ', lvwColumnCenter

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT PharmPurchaseOrdersData.* FROM PharmPurchaseOrdersData WHERE PharmPurchaseOrdersData.PurchaseOrderNo='" & Trim(.cboOrderNO.Text) & "' ORDER BY PharmPurchaseOrdersData.SerialNO;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!SerialNo))

    If Not IsNull(rsLIST!Drugcode) Then
        MyList.SubItems(1) = CStr(rsLIST!Drugcode)
    End If
    
    If Not IsNull(rsLIST!DrugName) Then
        MyList.SubItems(2) = CStr(rsLIST!DrugName)
    End If
    
    If Not IsNull(rsLIST!QuantityOrdered) Then
        MyList.SubItems(3) = CStr(rsLIST!QuantityOrdered)
    End If
    
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(4) = CStr(rsLIST!QuantityUnits)
    End If
    
    If Not IsNull(rsLIST!UnitCost) Then
        MyList.SubItems(5) = FormatNumber(rsLIST!UnitCost, 2, vbUseDefault, vbUseDefault, vbTrue)
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

.ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
.ListView2.ColumnHeaders(7).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Function GetMyOrderDescription() As String
'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive
    Set rsFindRecord = cnCOMMON.Execute("SELECT OrderDescription FROM PharmPurchaseOrders WHERE PurchaseOrderNO='" & Trim(CurrentOrder) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetMyOrderDescription = Empty
    ElseIf IsNull(rsFindRecord!orderdescription) = True Or rsFindRecord!orderdescription = "" Then
        GetMyOrderDescription = Empty
    Else
        GetMyOrderDescription = rsFindRecord!orderdescription
    End If
    
    Set rsFindRecord = Nothing
    
End With
Exit Function
Err:
    ErrorMessage
End Function

Public Sub ShowItemsInCurrentOrder()
'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Serial No", .ListView2.Width / 6
.ListView2.ColumnHeaders.Add , , "Product Code ", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Product Name", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Quantity Ordered", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Qty Units", .ListView2.Width / 6 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Dosage Cost", .ListView2.Width / 5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Total Cost", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Received?", .ListView2.Width / 5.4 ', lvwColumnCenter

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT PharmPurchaseOrdersData.* FROM PharmPurchaseOrdersData WHERE PharmPurchaseOrdersData.PurchaseOrderNo='" & Trim(CurrentOrder) & "' ORDER BY PharmPurchaseOrdersData.SerialNO;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!SerialNo))

    If Not IsNull(rsLIST!Drugcode) Then
        MyList.SubItems(1) = CStr(rsLIST!Drugcode)
    End If
    
    If Not IsNull(rsLIST!DrugName) Then
        MyList.SubItems(2) = CStr(rsLIST!DrugName)
    End If
    
    If Not IsNull(rsLIST!QuantityOrdered) Then
        MyList.SubItems(3) = CStr(rsLIST!QuantityOrdered)
    End If
    
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(4) = CStr(rsLIST!QuantityUnits)
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

.ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
.ListView2.ColumnHeaders(7).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub GetRecordsStructure()
'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Serial No", .ListView2.Width / 6
.ListView2.ColumnHeaders.Add , , "Order No ", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Drug Code ", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Drug Name", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Quantity Ordered", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Unit Cost", .ListView2.Width / 6 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Total Cost", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Received", .ListView2.Width / 5.5 ', lvwColumnCenter

.ListView2.View = lvwReport

End With

Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub ShowAllOrderRecords()
'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Serial No", .ListView2.Width / 6
.ListView2.ColumnHeaders.Add , , "Order No ", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Drug Code ", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Drug Name", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Quantity Ordered", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Unit Cost", .ListView2.Width / 6 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Total Cost", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Received", .ListView2.Width / 5.5 ', lvwColumnCenter

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT PharmPurchaseOrdersData.* FROM PharmPurchaseOrdersData ORDER BY PharmPurchaseOrdersData.SerialNO;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!SerialNo))
    
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
    
    If Not IsNull(rsLIST!UnitCost) Then
        MyList.SubItems(5) = FormatNumber(rsLIST!UnitCost, 2, vbUseDefault, vbUseDefault, vbTrue)
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

.ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
.ListView2.ColumnHeaders(7).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub ClearForNewRecord()
'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive
    .cboUnitsOrdered.Clear
    .cboUnitsReceived.Clear
    .txtDrugCode.Text = Empty
    .txtDrugName.Text = Empty
'    .txtExpiryDate.Text = Empty
'    .txtManufactureDate.Text = Empty
    .txtQuantityOrdered.Text = Empty
    .txtQuantityReceived.Text = Empty
    .txtSerialNo.Text = Empty
    .txtDosesOrdered.Text = Empty
    .txtDosesReceived.Text = Empty
    .txtDosesReceived.Locked = False
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub ClearTextFields()
'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive
    .cboUnitsOrdered.Clear
    .cboUnitsReceived.Clear
    .txtDateReceived.Text = Empty
    .txtDrugCode.Text = Empty
    .txtDrugName.Text = Empty
'    .txtExpiryDate.Text = Empty
'    .txtManufactureDate.Text = Empty
    .txtQuantityOrdered.Text = Empty
    .txtQuantityReceived.Text = Empty
    .txtSerialNo.Text = Empty
'    .txtShippDate.Text = Empty
    .txtDelDate.Text = Empty
    .txtDeliveryNote.Text = Empty
    .txtInvoiceDate.Text = Empty
    .txtInvoiceNumber.Text = Empty
    .txtDosesOrdered.Text = Empty
    .txtDosesReceived.Text = Empty
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub ClearTheScreen()
'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive
    .ListView1.ListItems.Clear
    .ListView2.ListItems.Clear
    .cboOrderNO.Clear
    .cboUnitsOrdered.Clear
    .cboUnitsReceived.Clear
    .txtDateReceived.Text = Empty
    .txtDrugCode.Text = Empty
    .txtDrugName.Text = Empty
    .txtOrderDate.Text = Empty
    .txtOrderDescription.Text = Empty
    .txtQuantityOrdered.Text = Empty
    .txtQuantityReceived.Text = Empty
    .txtSerialNo.Text = Empty
'    .txtShippDate.Text = Empty
    .txtDosesOrdered.Text = Empty
    .txtDosesReceived.Text = Empty
    .txtDeliveryNote.Text = Empty
    .txtDelDate.Text = Empty
    .txtInvoiceNumber.Text = Empty
    .txtInvoiceDate.Text = Empty
    
    
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub ShowOrdersWithBalances()
'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive
.ListView3.ListItems.Clear
.ListView3.ColumnHeaders.Clear

.ListView3.ColumnHeaders.Add , , "Order No", .ListView3.Width / 5 ', lvwColumnCenter
.ListView3.ColumnHeaders.Add , , "Order Date", .ListView3.Width / 5 ', lvwColumnCenter
.ListView3.ColumnHeaders.Add , , "Date Required", .ListView3.Width / 5 ', lvwColumnCenter
.ListView3.ColumnHeaders.Add , , "Date Promised", .ListView3.Width / 4.8
.ListView3.ColumnHeaders.Add , , "Supplier Code", .ListView3.Width / 5 ', lvwColumnCenter
.ListView3.ColumnHeaders.Add , , "Shipping Method", .ListView3.Width / 4 ', lvwColumnCenter
.ListView3.ColumnHeaders.Add , , "Freight Charge", .ListView3.Width / 4
.ListView3.ColumnHeaders.Add , , "Staff ID NO", .ListView3.Width / 4.5

.ListView3.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM PharmPurchaseOrders WHERE BalanceStatus IS NOT NULL  ORDER BY PurchaseOrderNo;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView3.View = lvwList
    Set MyList = .ListView3.ListItems.Add(, , "There are No Orders With Balances to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!PurchaseOrderNo))

    If Not IsNull(rsLIST!OrderDate) Then
        MyList.SubItems(1) = CStr(rsLIST!OrderDate)
    End If
    
    If Not IsNull(rsLIST!DateRequired) Then
        MyList.SubItems(2) = CStr(rsLIST!DateRequired)
    End If
    
    If Not IsNull(rsLIST!DatePromised) Then
        MyList.SubItems(3) = CStr(rsLIST!DatePromised)
    End If
    
    If Not IsNull(rsLIST!suppliercode) Then
        MyList.SubItems(4) = CStr(rsLIST!suppliercode)
    End If
    
    If Not IsNull(rsLIST!ShippingMethodID) Then
        MyList.SubItems(5) = Trim(CStr(rsLIST!ShippingMethodID))
    End If
    
    If Not IsNull(rsLIST!FreightCharge) Then
        MyList.SubItems(6) = FormatNumber(rsLIST!FreightCharge, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!StaffIdNo) Then
        MyList.SubItems(7) = Trim(CStr(rsLIST!StaffIdNo))
    End If
    
    rsLIST.MoveNext
    
Wend

.ListView3.ColumnHeaders(7).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub ShowBalanceItemsInCurrentOrder()
'On Error GoTo Err
With frmPharmacyPurchaseOrdersReceive
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Serial No", .ListView2.Width / 6
.ListView2.ColumnHeaders.Add , , "Product Code ", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Product Name", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Balance Due", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Qty Units", .ListView2.Width / 6 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Dosage Cost", .ListView2.Width / 5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Total Cost", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Received?", .ListView2.Width / 5.4 ', lvwColumnCenter

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM PharmPurchaseOrdersData  WHERE BalanceStatus IS NOT NULL AND PurchaseOrderNo='" & Trim(CurrentOrder) & "' ORDER BY PharmPurchaseOrdersData.SerialNO;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!SerialNo))

    If Not IsNull(rsLIST!Drugcode) Then
        MyList.SubItems(1) = CStr(rsLIST!Drugcode)
    End If
    
    If Not IsNull(rsLIST!DrugName) Then
        MyList.SubItems(2) = CStr(rsLIST!DrugName)
    End If
    
    If Not IsNull(rsLIST!Balance) Then
        MyList.SubItems(3) = CStr(rsLIST!Balance)
    End If
    
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(4) = CStr(rsLIST!QuantityUnits)
    End If
       
    If Not IsNull(rsLIST!DosageCost) Then
        MyList.SubItems(5) = FormatNumber(rsLIST!DosageCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!TotalCost) Then
        MyList.SubItems(6) = FormatNumber(rsLIST!TotalCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If IsNull(rsLIST!BalanceStatus) Then
        MyList.SubItems(7) = CStr("YES")
    ElseIf Not IsNull(rsLIST!BalanceStatus) Then
        If rsLIST!BalanceStatus = "Y" Then
            MyList.SubItems(7) = CStr("No")
        Else
            MyList.SubItems(7) = CStr("NO")
        End If
    End If
    
    rsLIST.MoveNext
    
Wend

.ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
.ListView2.ColumnHeaders(7).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub GetBalanceItemsInCurrentOrder()
With frmPharmacyPurchaseOrdersReceive
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Serial No", .ListView2.Width / 6
.ListView2.ColumnHeaders.Add , , "Product Code ", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Product Name", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Balance Due", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Qty Units", .ListView2.Width / 6 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Dosage Cost", .ListView2.Width / 5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Total Cost", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Received?", .ListView2.Width / 5.4 ', lvwColumnCenter

.ListView2.View = lvwReport
End With
End Sub




