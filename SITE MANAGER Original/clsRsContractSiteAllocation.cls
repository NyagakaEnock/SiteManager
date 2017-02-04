VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRsContractSiteAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private StartDate As Variant, EndDate As Variant, EXPDate As Variant, DelDate As Variant, InvDate As Variant, QBal As Double, ShipDate As Variant

Public Sub AddNewRecord()
If EditRecord Then Exit Sub
    NewRecord = True
    With frmRsContractSiteAllocation
    
        Call ClearTextFields
        
        .cmdSAVE.Caption = "&SAVE RECORD"
'        .cmdCHANGE.Enabled = False
       
        
        If .cboOrderNO.Text = Empty Then
            .cboOrderNO.SetFocus
        Else
            .ListView2.SetFocus
        End If
        
    End With
End Sub

Public Sub SaveReceivedData()
''On Error GoTo Err
With frmRsContractSiteAllocation
        
    StartDate = Format(.txtContractStartDate.Text, "MMMM dd,yyyy")
    EndDate = Format(.txtContractExpiryDate.Text, "MMMM dd,yyyy")
    
    Set rsNewRecord = New ADODB.Recordset
    
    MySQL = "INSERT INTO ContractSiteAllocationData(ContractNo,SiteNumber,BillBoardNo,SiteName,City,SitePhysicalAddress,AdvCode,AdvName,ClientCode,ClientName,StartDate,ExpDate,Createdby,DateCreated,AccPeriod) VALUES('" & .cboOrderNO.Text & "','" & Trim(.txtSiteNo.Text) & "', '" & Trim(.txtBBNo.Text) & "','" & Trim(.txtSiteName.Text) & "','" & Trim(.txtCity.Text) & "','" & Trim(.txtPhysicalAddress.Text) & "','" & Trim(.txtAdvCode.Text) & "','" & Trim(.txtAdvName.Text) & "','" & Trim(.txtClientCode.Text) & "','" & Trim(.txtClientName.Text) & "','" & Trim(StartDate) & "','" & Trim(EndDate) & "','" & CurrentUserName & "','" & MyCurrentDate & "','" & MyCurrentPeriod & "')"
    
    rsNewRecord.Open MySQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Set rsNewRecord = Nothing
    
Update:
              
      Set rsEditRecord = New ADODB.Recordset
      MySQL = "UPDATE AdvertSites SET AllocationStatus='" & "Y" & "' WHERE SiteNo='" & Trim(.txtSiteNo.Text) & "';"
      rsEditRecord.Open MySQL, cnCOMMON, adOpenKeyset, adLockOptimistic
      Set rsEditRecord = Nothing
    
    
      Set rsEditRecord = New ADODB.Recordset
      MySQL = "UPDATE AdvertContractRequisition SET AllocationStatus='" & "Y" & "' WHERE PurchaseOrderNo='" & Trim(.cboOrderNO.Text) & "';"
      rsEditRecord.Open MySQL, cnCOMMON, adOpenKeyset, adLockOptimistic
      Set rsEditRecord = Nothing
      
      Set rsEditRecord = New ADODB.Recordset
      MySQL = "UPDATE AdvertContractRequisitionData SET AllocationStatus='" & "Y" & "' WHERE PurchaseOrderNo='" & Trim(.cboOrderNO.Text) & "' AND SerialNo = '" & Trim(.txtSerialNO.Text) & "';"
      rsEditRecord.Open MySQL, cnCOMMON, adOpenKeyset, adLockOptimistic
      Set rsEditRecord = Nothing
        
      Call RemoveCurrentListItem
    .cmdSAVE.Caption = "&NEW"

    
    Call ShowCurrentPendingOrdersData
    
    NewRecord = False
    
End With

    Exit Sub
    
Err:
    UpdateErrorMessage
End Sub
Public Sub RemoveCurrentListItem()
''On Error GoTo Err
With frmRsContractSiteAllocation
Dim i, j, k
   j = .ListView2.ListItems.Count: i = 1
     If j = 0 Then Exit Sub
     
     For i = 1 To j
      If .ListView2.ListItems(i).Checked = True Then
         .ListView2.ListItems.Remove (i): Exit Sub
      End If
    Next i
End With
Exit Sub
Err:
   ErrorMessage
End Sub



Private Function AllItemsReceived() As Boolean
''On Error GoTo Err
With frmRsContractSiteAllocation

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
''On Error GoTo Err
Dim Today As Variant
Today = Format(Date, "MMMM dd,yy")
With frmRsContractSiteAllocation

    If .cboOrderNO.ListCount <> 0 Then Exit Sub
    
    Set rsCOMBO = cnCOMMON.Execute("SELECT PurchaseOrderNo FROM AdvertContractRequisition WHERE ApprovedStatus = '" & "Y" & "' AND PaidStatus = '" & "Y" & "' AND EndDate > '" & Today & "' ORDER BY PurchaseOrderNo;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo Outs
    
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!PurchaseOrderNo) And rsCOMBO!PurchaseOrderNo <> "" Then
            .cboOrderNO.AddItem rsCOMBO!PurchaseOrderNo
        End If
    rsCOMBO.MoveNext
    Loop
    
Outs:
    Set rsCOMBO = Nothing
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub FindDetailsByPurchaseOrderNo()
''On Error GoTo Err
With frmRsContractSiteAllocation
    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM AdvertContractRequisitionData  WHERE PurchaseOrderNo='" & Trim(.cboOrderNO.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
    
        Set rsFindRecord = Nothing: Exit Sub
        
    Else
        .cboOrderNO.Text = rsFindRecord!PurchaseOrderNo & ""
        .txtOrderDate.Text = rsFindRecord!DateCreated & ""
        .txtOrderDescription.Text = rsFindRecord!DateCreated & ""
        
        Call ShowCurrentPurchaseOrder
        Call ShowCurrentPendingOrdersData
        
       
    End If
    
    Set rsFindRecord = Nothing
    
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub GetPurchaseStructure()
''On Error GoTo Err
With frmRsContractSiteAllocation
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Order Date", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract Start Date", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract End Date", .ListView1.Width / 4.8
.ListView1.ColumnHeaders.Add , , "Client Code", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Client Name", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Contact Person", .ListView1.Width / 4.5

.ListView1.View = lvwReport
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub ShowPurchaseOrdersOverDue()
''On Error GoTo Err
Dim Today As Variant
Today = Format(Date, "MMMM dd,yy")
With frmRsContractSiteAllocation
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Order Date", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract Start Date", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract End Date", .ListView1.Width / 4.8
.ListView1.ColumnHeaders.Add , , "Client Code", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Client Name", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Contact Person", .ListView1.Width / 4.5

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertContractRequisition WHERE ApprovedStatus = '" & "Y" & "' AND PaidStatus = '" & "Y" & "' AND StartDate < '" & Today & "'  ORDER BY PurchaseOrderNo;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "There are No New Order Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PurchaseOrderNo))

    If Not IsNull(rsLIST!DateCreated) Then
        MyList.SubItems(1) = CStr(rsLIST!DateCreated)
    End If
    
    If Not IsNull(rsLIST!StartDate) Then
        MyList.SubItems(2) = CStr(rsLIST!StartDate)
    End If
    
    If Not IsNull(rsLIST!EndDate) Then
        MyList.SubItems(3) = CStr(rsLIST!EndDate)
    End If
    
    If Not IsNull(rsLIST!ClientCode) Then
        MyList.SubItems(4) = CStr(rsLIST!ClientCode)
    End If
    
    If Not IsNull(rsLIST!ClientName) Then
        MyList.SubItems(5) = Trim(CStr(rsLIST!ClientName))
    End If
    
    If Not IsNull(rsLIST!PhysicalAddress) Then
        MyList.SubItems(6) = Trim(CStr(rsLIST!PhysicalAddress))
    End If
    
    If Not IsNull(rsLIST!ContactPerson) Then
        MyList.SubItems(7) = Trim(CStr(rsLIST!ContactPerson))
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

Public Sub ShowPurchaseOrdersDueToday()
''On Error GoTo Err
Dim Today As Variant
Today = Format(Date, "MMMM dd,yy")
With frmRsContractSiteAllocation
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Order Date", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract Start Date", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract End Date", .ListView1.Width / 4.8
.ListView1.ColumnHeaders.Add , , "Client Code", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Client Name", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Contact Person", .ListView1.Width / 4.5

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertContractRequisition WHERE ApprovedStatus = '" & "Y" & "' AND PaidStatus = '" & "Y" & "' AND StartDate > '" & Today & "'  ORDER BY PurchaseOrderNo;", cnCOMMON, adOpenKeyset, adLockOptimistic

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
    
    If Not IsNull(rsLIST!StartDate) Then
        MyList.SubItems(2) = CStr(rsLIST!StartDate)
    End If
    
    If Not IsNull(rsLIST!EndDate) Then
        MyList.SubItems(3) = CStr(rsLIST!EndDate)
    End If
    
    If Not IsNull(rsLIST!ClientCode) Then
        MyList.SubItems(4) = CStr(rsLIST!ClientCode)
    End If
    
    If Not IsNull(rsLIST!ClientName) Then
        MyList.SubItems(5) = Trim(CStr(rsLIST!ClientName))
    End If
    
    If Not IsNull(rsLIST!PhysicalAddress) Then
        MyList.SubItems(6) = Trim(CStr(rsLIST!PhysicalAddress))
    End If
    
    If Not IsNull(rsLIST!ContactPerson) Then
        MyList.SubItems(8) = Trim(CStr(rsLIST!ContactPerson))
    End If
    
        
    rsLIST.MoveNext
    
Wend

'.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub


Public Sub ShowNewPurchaseOrders()
''On Error GoTo Err
Dim Today As Variant
Today = Format(Date, "MMMM dd,yy")
With frmRsContractSiteAllocation
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Order Date", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract Start Date", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract End Date", .ListView1.Width / 4.8
.ListView1.ColumnHeaders.Add , , "Client Code", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Client Name", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Contact Person", .ListView1.Width / 4.5

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertContractRequisition A,AdvertContractRequisitionData B WHERE A.PurchaseOrderNo = B.PurchaseOrderNo AND A.ApprovedStatus = '" & "Y" & "' AND A.PaidStatus = '" & "Y" & "' AND A.EndDate > '" & Today & "'AND A.AllocationStatus IS NULL ORDER BY A.PurchaseOrderNo;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "There are No New Order Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PurchaseOrderNo))

    If Not IsNull(rsLIST!DateCreated) Then
        MyList.SubItems(1) = CStr(rsLIST!DateCreated)
    End If
    
    If Not IsNull(rsLIST!StartDate) Then
        MyList.SubItems(2) = CStr(rsLIST!StartDate)
    End If
    
    If Not IsNull(rsLIST!EndDate) Then
        MyList.SubItems(3) = CStr(rsLIST!EndDate)
    End If
    
    If Not IsNull(rsLIST!ClientCode) Then
        MyList.SubItems(4) = CStr(rsLIST!ClientCode)
    End If
    
    If Not IsNull(rsLIST!ClientName) Then
        MyList.SubItems(5) = Trim(CStr(rsLIST!ClientName))
    End If
    
    If Not IsNull(rsLIST!PhysicalAddress) Then
        MyList.SubItems(6) = Trim(CStr(rsLIST!PhysicalAddress))
    End If
    
    If Not IsNull(rsLIST!ContactPerson) Then
        MyList.SubItems(7) = Trim(CStr(rsLIST!ContactPerson))
    End If
    
        
    rsLIST.MoveNext
    
Wend

'.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub ShowCurrentPurchaseOrder()
''On Error GoTo Err
Dim Today As Variant
Today = Format(Date, "MMMM dd,yy")
With frmRsContractSiteAllocation
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Order Date", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract Start Date", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract End Date", .ListView1.Width / 4.8
.ListView1.ColumnHeaders.Add , , "Client Code", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Client Name", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Contact Person", .ListView1.Width / 4.5

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertContractRequisition WHERE ApprovedStatus = '" & "Y" & "' AND PaidStatus = '" & "Y" & "' AND EndDate > '" & Today & "' AND PurchaseOrderNo = '" & Trim(.cboOrderNO.Text) & "'  ORDER BY PurchaseOrderNo;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "There are No New Order Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PurchaseOrderNo))

    If Not IsNull(rsLIST!DateCreated) Then
        MyList.SubItems(1) = CStr(rsLIST!DateCreated)
    End If
    
    If Not IsNull(rsLIST!StartDate) Then
        MyList.SubItems(2) = CStr(rsLIST!StartDate)
    End If
    
    If Not IsNull(rsLIST!EndDate) Then
        MyList.SubItems(3) = CStr(rsLIST!EndDate)
    End If
    
    If Not IsNull(rsLIST!ClientCode) Then
        MyList.SubItems(4) = CStr(rsLIST!ClientCode)
    End If
    
    If Not IsNull(rsLIST!ClientName) Then
        MyList.SubItems(5) = Trim(CStr(rsLIST!ClientName))
    End If
    
    If Not IsNull(rsLIST!PhysicalAddress) Then
        MyList.SubItems(6) = Trim(CStr(rsLIST!PhysicalAddress))
    End If
    
    If Not IsNull(rsLIST!ContactPerson) Then
        MyList.SubItems(7) = Trim(CStr(rsLIST!ContactPerson))
    End If
    
        
    rsLIST.MoveNext
    
Wend

'.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub ShowCurrentPendingOrdersData()
''On Error GoTo Err
Dim Today As Variant
Today = Format(Date, "MMMM dd,yy")
With frmRsContractSiteAllocation
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Serial No", .ListView2.Width / 6
.ListView2.ColumnHeaders.Add , , "Advert Code ", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Advert Name", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Advert Type", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Advert Length", .ListView2.Width / 6 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Advert Width", .ListView2.Width / 5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Duration", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Days", .ListView2.Width / 5.4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Advert Cost", .ListView2.Width / 5.4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contract Start Date", .ListView2.Width / 5.4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contract End Date", .ListView2.Width / 5.4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Approved", .ListView2.Width / 6.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Paid", .ListView2.Width / 6.5 ', lvwColumnCenter




.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertContractRequisitionData WHERE PurchaseOrderNo='" & Trim(.cboOrderNO.Text) & "' AND PaidStatus = '" & "Y" & "' AND ApprovedStatus = '" & "Y" & "' AND ContractEndDate > '" & Today & "'  ORDER BY SerialNO;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!SerialNo))

    If Not IsNull(rsLIST!Advcode) Then
        MyList.SubItems(1) = CStr(rsLIST!Advcode)
    End If
    
    If Not IsNull(rsLIST!AdvName) Then
        MyList.SubItems(2) = CStr(rsLIST!AdvName)
    End If
    
    If Not IsNull(rsLIST!AdvType) Then
        MyList.SubItems(3) = CStr(rsLIST!AdvType)
    End If
    
    If Not IsNull(rsLIST!AdvLength) Then
        MyList.SubItems(4) = CStr(rsLIST!AdvLength)
    End If
    
    If Not IsNull(rsLIST!AdvWidth) Then
        MyList.SubItems(5) = CStr(rsLIST!AdvWidth)
    End If
    
    If Not IsNull(rsLIST!Duration) Then
        MyList.SubItems(6) = CStr(rsLIST!Duration)
    End If
    
    If Not IsNull(rsLIST!Days) Then
        MyList.SubItems(7) = CStr(rsLIST!Days)
    End If
    
            
    If Not IsNull(rsLIST!AdvCost) Then
        MyList.SubItems(8) = FormatNumber(rsLIST!AdvCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!ContractStartDate) Then
        MyList.SubItems(9) = CStr(rsLIST!ContractStartDate)
    End If
    
    If Not IsNull(rsLIST!ContractEndDate) Then
        MyList.SubItems(10) = CStr(rsLIST!ContractEndDate)
    End If
    
        
    If IsNull(rsLIST!ApprovedStatus) Then
        MyList.SubItems(11) = CStr("NO")
    ElseIf Not IsNull(rsLIST!ApprovedStatus) Then
        If rsLIST!ApprovedStatus = "Y" Then
            MyList.SubItems(11) = CStr("YES")
        Else
            MyList.SubItems(11) = CStr("NO")
        End If
    End If
    
    
    If IsNull(rsLIST!PaidStatus) Then
        MyList.SubItems(12) = CStr("NO")
    ElseIf Not IsNull(rsLIST!PaidStatus) Then
        If rsLIST!PaidStatus = "Y" Then
            MyList.SubItems(12) = CStr("YES")
        Else
            MyList.SubItems(12) = CStr("NO")
        End If
    End If
    
    
    rsLIST.MoveNext
    
Wend

'.ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
'.ListView2.ColumnHeaders(7).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub ShowCurrentOrderRecords()
''On Error GoTo Err
With frmRsContractSiteAllocation
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
''On Error GoTo Err
With frmRsContractSiteAllocation
    Set rsFindRecord = cnCOMMON.Execute("SELECT OrderDescription FROM AdvertContractRequisition WHERE PurchaseOrderNO='" & Trim(CurrentOrder) & "';")
    
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
''On Error GoTo Err
Dim Today As Variant
Today = Format(Date, "MMMM dd,yy")
With frmRsContractSiteAllocation
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Serial No", .ListView2.Width / 6
.ListView2.ColumnHeaders.Add , , "Advert Code ", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Advert Name", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Advert Type", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Advert Length", .ListView2.Width / 6 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Advert Width", .ListView2.Width / 5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Duration", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Days", .ListView2.Width / 5.4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Advert Cost", .ListView2.Width / 5.4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contract Start Date", .ListView2.Width / 5.4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contract End Date", .ListView2.Width / 5.4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Approved", .ListView2.Width / 6.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Paid", .ListView2.Width / 6.5 ', lvwColumnCenter

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertContractRequisitionData A ,AdvertContractRequisition B WHERE A.PurchaseOrderNo = B.PurchaseOrderNo AND  A.PurchaseOrderNo='" & Trim(CurrentOrder) & "' AND A.PaidStatus = '" & "Y" & "' AND A.ApprovedStatus = '" & "Y" & "' AND A.ContractEndDate > '" & Today & "'AND A.AllocationStatus IS NULL  ORDER BY A.SerialNO;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing
    
    Set rsLineUpdate = New ADODB.Recordset
     rsLineUpdate.Open "UPDATE AdvertContractRequisition SET AllocationStatus = '" & "Y" & "' WHERE PurchaseOrderNo ='" & Trim(CurrentOrder) & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
    Set rsLineUpdate = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!SerialNo))

    If Not IsNull(rsLIST!Advcode) Then
        MyList.SubItems(1) = CStr(rsLIST!Advcode)
    End If
    
    If Not IsNull(rsLIST!AdvName) Then
        MyList.SubItems(2) = CStr(rsLIST!AdvName)
    End If
    
    If Not IsNull(rsLIST!AdvType) Then
        MyList.SubItems(3) = CStr(rsLIST!AdvType)
    End If
    
    If Not IsNull(rsLIST!AdvLength) Then
        MyList.SubItems(4) = CStr(rsLIST!AdvLength)
    End If
    
    If Not IsNull(rsLIST!AdvWidth) Then
        MyList.SubItems(5) = CStr(rsLIST!AdvWidth)
    End If
    
    If Not IsNull(rsLIST!Duration) Then
        MyList.SubItems(6) = CStr(rsLIST!Duration)
    End If
    
    If Not IsNull(rsLIST!Days) Then
        MyList.SubItems(7) = CStr(rsLIST!Days)
    End If
    
            
    If Not IsNull(rsLIST!AdvCost) Then
        MyList.SubItems(8) = FormatNumber(rsLIST!AdvCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!ContractStartDate) Then
        MyList.SubItems(9) = CStr(rsLIST!ContractStartDate)
    End If
    
    If Not IsNull(rsLIST!ContractEndDate) Then
        MyList.SubItems(10) = CStr(rsLIST!ContractEndDate)
    End If
    
        
    If IsNull(rsLIST!ApprovedStatus) Then
        MyList.SubItems(11) = CStr("NO")
    ElseIf Not IsNull(rsLIST!ApprovedStatus) Then
        If rsLIST!ApprovedStatus = "Y" Then
            MyList.SubItems(11) = CStr("YES")
        Else
            MyList.SubItems(11) = CStr("NO")
        End If
    End If
    
    
    If IsNull(rsLIST!PaidStatus) Then
        MyList.SubItems(12) = CStr("NO")
    ElseIf Not IsNull(rsLIST!PaidStatus) Then
        If rsLIST!PaidStatus = "Y" Then
            MyList.SubItems(12) = CStr("YES")
        Else
            MyList.SubItems(12) = CStr("NO")
        End If
    End If
    
    
    rsLIST.MoveNext
    
Wend

'.ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
'.ListView2.ColumnHeaders(7).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub GetRecordsStructure()
''On Error GoTo Err
With frmRsContractSiteAllocation
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Serial No", .ListView2.Width / 6
.ListView2.ColumnHeaders.Add , , "Advert Code ", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Advert Name", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Advert Type", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Advert Length", .ListView2.Width / 6 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Advert Width", .ListView2.Width / 5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Duration", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Days", .ListView2.Width / 5.4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Advert Cost", .ListView2.Width / 5.4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contract Start Date", .ListView2.Width / 5.4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contract End Date", .ListView2.Width / 5.4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Approved", .ListView2.Width / 6.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Paid", .ListView2.Width / 6.5 ', lvwColumnCenter

.ListView2.View = lvwReport

End With

Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub ShowAllOrderRecords()
''On Error GoTo Err
With frmRsContractSiteAllocation
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


Public Sub ClearTextFields()
For Each i In Screen.ActiveForm
    If TypeOf i Is TextBox And i.Name <> "txtTitle" Then
        i.Text = Empty
    End If
    If TypeOf i Is ComboBox Then
        i.Clear
    End If
    If TypeOf i Is Image Then
        i.Picture = LoadPicture("")
    End If
Next i
End Sub


Public Sub ShowBalanceItemsInCurrentOrder()
''On Error GoTo Err
With frmRsContractSiteAllocation
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




