VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRsContractRequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Ordate As Variant, ReqDate As Variant, PromDate As Variant

Public Sub AddNewRecord()
If EditRecord Then Exit Sub
    NewRecord = True
    With frmRsContractRequisition
    
        Call ClearForNewItem
        
        .cmdSAVE.Caption = "&SAVE RECORD"
        .cmdFinish.Enabled = False
        
        If .txtOrderNO.Text = Empty Then
            .txtOrderNO.SetFocus
        Else
            .cboAdvCode.SetFocus
        End If
        
    End With
End Sub
Public Sub ClearForNewItem()
'On Error GoTo Err
With frmRsContractRequisition
.cboAdvCode.Text = ""
.cboAdvName.Text = ""
.txtType.Text = ""
.txtLenght.Text = ""
.txtWidth.Text = ""
.txtDuration.Text = ""
.txtDays.Text = ""
.txtAdvCost.Text = ""
.txtTotalCost.Text = ""
End With
Exit Sub
Err:
ErrorMessage
End Sub
Public Sub SavePurchaseData()
''On Error GoTo Err
With frmRsContractRequisition
Dim DateRequired As Variant
Dim DatePromised As Variant
    DateRequired = Format(.txtDateRequired.Text, "MMMM dd,yyyy")
    DatePromised = Format(.txtDatePromised.Text, "MMMM dd,yyyy")
    
    Set rsNewRecord = New ADODB.Recordset
    
    MySQL = "INSERT INTO AdvertContractRequisitionData(ContractStartDate,ContractEndDate,purchaseorderno,AdvCode,Advname,AdvType,AdvLength,AdvWidth,Duration,Days,AdvCost,createdby,datecreated,accperiod) VALUES('" & DateRequired & "','" & DatePromised & "','" & Trim(.txtOrderNO.Text) & "','" & Trim(.cboAdvCode.Text) & "','" & Trim(.cboAdvName.Text) & "','" & Trim(.txtType.Text) & "','" & Trim(.txtLenght.Text) & "','" & Trim(.txtWidth.Text) & "','" & Trim(.txtDuration.Text) & "','" & Trim(.txtDays.Text) & "'," & CCur(.txtAdvCost.Text) & ",'" & CurrentUserName & "','" & MyCurrentDate & "','" & MyCurrentPeriod & "')"
    
    rsNewRecord.Open MySQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
    Set rsNewRecord = Nothing
           
    
    .cmdSAVE.Caption = "&NEXT ITEM"
    .cmdFinish.Enabled = True
    
    
    
    
    .cmdFinish.SetFocus
    
End With

    Exit Sub
    
Err:
    UpdateErrorMessage
End Sub

Private Function OrderAlreadyExists() As Boolean
''On Error GoTo Err
With frmRsContractRequisition
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




Private Function GetMyCompanyCode() As String
''On Error GoTo Err

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
''On Error GoTo Err

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




Public Sub GetSupplyStructure()

''On Error GoTo Err
With frmRsContractRequisition
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Advert Code ", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Name", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Type", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Length", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Width", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Weight", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Duration", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Days", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Cost", .ListView1.Width / 5 ', lvwColumnCenter

.ListView1.View = lvwReport
End With
Exit Sub
Err:
    ErrorMessage
End Sub



Public Sub ShowCurrentOrderSuppliers()
''On Error GoTo Err
With frmRsContractRequisition
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Client Code", .ListView2.Width / 4.5
.ListView2.ColumnHeaders.Add , , "Client Name", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contact Person", .ListView2.Width / 4.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Physical Address", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contract Start Date", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contract End Date", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contact Title", .ListView2.Width / 4.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Postal Address", .ListView2.Width / 5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "City", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Country", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Phone", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Mobile No", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "E-Mail", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Fax", .ListView2.Width / 5.5 ', lvwColumnCenter


.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertClients A,AdvertContractRequisition B WHERE A.CustomerId = B.ClientCode AND B.PurchaseOrderNo='" & Trim(.txtOrderNO.Text) & "' ORDER BY A.CustomerId;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!CustomerId))

    If Not IsNull(rsLIST!ClientName) Then
        MyList.SubItems(1) = CStr(rsLIST!ClientName)
    End If
    
    If Not IsNull(rsLIST!ContactName) Then
        MyList.SubItems(2) = CStr(rsLIST!ContactName)
    End If
    
    If Not IsNull(rsLIST!PhysicalAddress) Then
        MyList.SubItems(3) = Trim(CStr(rsLIST!PhysicalAddress))
    End If
    
    If Not IsNull(rsLIST!StartDate) Then
        MyList.SubItems(4) = Trim(CStr(rsLIST!StartDate))
    End If
    
    If Not IsNull(rsLIST!EndDate) Then
        MyList.SubItems(5) = CStr(rsLIST!EndDate)
    End If
    
    If Not IsNull(rsLIST!ContactTitle) Then
        MyList.SubItems(6) = Trim(CStr(rsLIST!ContactTitle))
    End If
    
    If Not IsNull(rsLIST!Address) Then
        MyList.SubItems(7) = Trim(CStr(rsLIST!Address))
    End If
    
    If Not IsNull(rsLIST!City) Then
        MyList.SubItems(8) = CStr(rsLIST!City)
    End If
    
    If Not IsNull(rsLIST!Country) Then
        MyList.SubItems(9) = CStr(rsLIST!Country)
    End If
    
    If Not IsNull(rsLIST!Phone) Then
        MyList.SubItems(10) = CStr(rsLIST!Phone)
    End If
    
    If Not IsNull(rsLIST!MobilePhone) Then
        MyList.SubItems(11) = CStr(rsLIST!MobilePhone)
    End If
    
    If Not IsNull(rsLIST!Email) Then
        MyList.SubItems(12) = CStr(rsLIST!Email)
    End If
    
    If Not IsNull(rsLIST!Fax) Then
        MyList.SubItems(13) = CStr(rsLIST!Fax)
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
''On Error GoTo Err
''On Error GoTo Err
With frmRsContractRequisition
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Client Code", .ListView2.Width / 4.5
.ListView2.ColumnHeaders.Add , , "Client Name", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contact Person", .ListView2.Width / 4.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Physical Address", .ListView2.Width / 2.5 ', lvwColumnCenter

.ListView2.View = lvwReport

End With
Exit Sub
Err:
    ErrorMessage
End Sub



Public Sub ShowCurrentOrderRecords()
''On Error GoTo Err
With frmRsContractRequisition
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Serial No", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Contract No ", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Code", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Name", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Type", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Lenght", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Width", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Duration", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "No Of Days", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Cost", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Paid", .ListView1.Width / 5.5 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertContractRequisitionData WHERE PurchaseOrderNo='" & Trim(.txtOrderNO.Text) & "' ORDER BY SerialNO;", cnCOMMON, adOpenKeyset, adLockOptimistic

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
    
    If Not IsNull(rsLIST!Advcode) Then
        MyList.SubItems(2) = CStr(rsLIST!Advcode)
    End If
    
     If Not IsNull(rsLIST!AdvName) Then
        MyList.SubItems(3) = CStr(rsLIST!AdvName)
    End If
    
     If Not IsNull(rsLIST!AdvType) Then
        MyList.SubItems(4) = CStr(rsLIST!AdvType)
    End If
    
     If Not IsNull(rsLIST!AdvLength) Then
        MyList.SubItems(5) = CStr(rsLIST!AdvLength)
    End If
    
     If Not IsNull(rsLIST!AdvWidth) Then
        MyList.SubItems(6) = CStr(rsLIST!AdvWidth)
    End If
    
     If Not IsNull(rsLIST!Duration) Then
        MyList.SubItems(7) = CStr(rsLIST!Duration)
    End If
    
     If Not IsNull(rsLIST!Days) Then
        MyList.SubItems(8) = CStr(rsLIST!Days)
    End If
           
    If Not IsNull(rsLIST!AdvCost) Then
        MyList.SubItems(9) = FormatNumber(rsLIST!AdvCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    
    If IsNull(rsLIST!PaidStatus) Then
        MyList.SubItems(10) = CStr("NO")
    ElseIf Not IsNull(rsLIST!PaidStatus) Then
        If rsLIST!PaidStatus = "Y" Then
            MyList.SubItems(10) = CStr("YES")
        Else
            MyList.SubItems(10) = CStr("NO")
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
''On Error GoTo Err
With frmRsContractRequisition

Dim rsLastID As ADODB.Recordset 'used to retrieve current LastId in the Table
Dim strLastID As String 'SQL statement

Dim strTemp As String 'store current record
Dim iNumPos As Integer 'store position of the first numeral
Dim strPrefix As String 'stores Id Prefix

'Retrieve the last record in the recrdset where order is ascending

'strLastID = "SELECT max(cnCliniclDetails.cnCliniclID) as lastid from cnCliniclDetails"
strLastID = "SELECT MAX(serialno) AS LastID FROM AdvertContractRequisitionData" ';"
Set rsLastID = New ADODB.Recordset

With rsLastID
'open the recordset
    .Open strLastID, cnCOMMON, adOpenKeyset, adLockOptimistic
    If .EOF And .BOF Then 'shows empty recordset
        AutoSerialNo = 1     'format of desired format of the string value
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
''On Error GoTo Err
With frmRsContractRequisition

Dim rsLastID As ADODB.Recordset 'used to retrieve current LastId in the Table
Dim strLastID As String 'SQL statement

Dim strTemp As String 'store current record
Dim iNumPos As Integer 'store position of the first numeral
Dim strPrefix As String 'stores Id Prefix

'Retrieve the last record in the recrdset where order is ascending

'strLastID = "SELECT max(cnCliniclDetails.cnCliniclID) as lastid from cnCliniclDetails"
strLastID = "SELECT MAX(PurchaseOrderNo) AS LastID FROM AdvertContractRequisition;"
Set rsLastID = New ADODB.Recordset

With rsLastID
'open the recordset
    .Open strLastID, cnCOMMON, adOpenKeyset, adLockOptimistic
    If .EOF And .BOF Then 'shows empty recordset
        AutoPurchaseOrderNo = "CNTRCT00001" 'format of desired format of the string value
    ElseIf IsNull(!lastid) = True Or !lastid = "" Then
        AutoPurchaseOrderNo = "CNTRCT00001"
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
''On Error GoTo Err
With frmRsContractRequisition
    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM AdvertClients A,AdvertContractRequisition B WHERE B.ClientCode=A.CustomerId AND B.PurchaseOrderNo='" & Trim(.txtOrderNO.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
    
        Set rsFindRecord = Nothing: Exit Sub
        
    Else
    
        .txtOrderDate.Text = rsFindRecord!DateCreated & ""
        .txtOrderDescription.Text = rsFindRecord!orderdescription & ""
        .txtOrderNO.Text = rsFindRecord!PurchaseOrderNo & ""
        .txtDatePromised.Text = rsFindRecord!EndDate & ""
        .txtDateRequired.Text = rsFindRecord!StartDate & ""
        .txtContactPerson.Text = rsFindRecord!ContactName & ""
        .txtPhysicalAddress.Text = rsFindRecord!PhysicalAddress & ""
        .cboClientId.Text = rsFindRecord!ClientCode & ""
        .cboAdvName.Text = rsFindRecord!ClientName & ""
        
        Call ShowCurrentOrderRecords
        Call ShowCurrentOrderSuppliers
        
    End If
    
End With
Exit Sub
Err:
    ErrorMessage
End Sub
Public Sub ShowReOrderItems()
''On Error GoTo Err
With frmRsContractRequisition
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
    
     If Not IsNull(rsLIST!Categorycode) Then
        MyList.SubItems(3) = CStr(rsLIST!Categorycode)
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
''On Error GoTo Err
With frmRsContractRequisition
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

