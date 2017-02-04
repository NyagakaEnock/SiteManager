VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub GetDeliveryMethodID()
'On Error GoTo err
    With frmODASPSuppliers
        
        Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM ODASPDeliveryMethods WHERE DeliveryMethod='" & Trim(.cboDeliveryMethod.Text) & "';")
        
        If rsFindRecord.EOF And rsFindRecord.BOF Then
            Set rsFindRecord = Nothing: Exit Sub
        Else
            .cboDeliveryMethod.Text = rsFindRecord!DeliveryMethodID & ""
            .txtFreightCharge.SetFocus
        End If
        
        Set rsFindRecord = Nothing
        
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub GetCurrencyCode()
'On Error GoTo err
    With frmODASPSuppliers
        
        Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM ParamCurrencies WHERE CurrencyName='" & Trim(.cboCurrencyCode.Text) & "';")
        
        If rsFindRecord.EOF And rsFindRecord.BOF Then
            Set rsFindRecord = Nothing: Exit Sub
        Else
            .cboCurrencyCode.Text = rsFindRecord!CurrencyCode & ""
            .txtSupplyduration.SetFocus
        End If
        
        Set rsFindRecord = Nothing
        
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub GetSupplierTypeCode()
'On Error GoTo err
    With frmODASPSuppliers
        
        Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM ParamSupplierTypes WHERE Descriptions='" & Trim(.txtSupplierType.Text) & "';")
        
        If rsFindRecord.EOF And rsFindRecord.BOF Then
            Set rsFindRecord = Nothing: Exit Sub
        Else
            .txtSupplierType.Text = rsFindRecord!SupplierType & ""
            .cboDeliveryMethod.SetFocus
        End If
        
        Set rsFindRecord = Nothing
        
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Sub AddNew()
'On Error GoTo err
Dim strQRE As Variant
Dim rsFIND As ADODB.Recordset

Set rsFIND = New ADODB.Recordset

Select Case frmODASPSuppliers.cmdAddNew.Caption
Case "&NEW"
    ClearControls
    EnableCB
    EnableControls
    Save = True
        
    If Save Then
        frmODASPSuppliers.cmdAddNew.Caption = "SAVE &CHANGES"
        frmODASPSuppliers.txtSupplierCode.SetFocus
    End If
    
Case "SAVE &CHANGES"

    Dim rsFinder As ADODB.Recordset
    Set rsFinder = New ADODB.Recordset
    
    If Validate = True Then
        With frmODASPSuppliers
            rsFinder.Open "INSERT INTO ODASPSuppliers(SupplierCode,SupplierType,SupplyDuration,SupplierName,Physicaladdress,PostalAddress,TownCity,TelephoneNo,FaxNo,MobileNo,EmailAddress,ContactPerson,ContactTitle,DeliveryMethod, createdby,datecreated) VALUES('" & .txtSupplierCode.Text & "','" & .txtSupplierType.Text & "','" & .txtSupplyduration.Text & "','" & .txtSupplierName.Text & "','" & .txtPhysicalAddress.Text & "','" & .txtPostalAddress.Text & "','" & .cboTownCity.Text & "','" & .txtTelephoneNo.Text & "','" & .txtfaxNo.Text & "','" & .txtMobileNo.Text & "','" & .txtEmail.Text & "','" & .txtContactPerson.Text & "','" & .cboContactTitle.Text & "','" & .cboDeliveryMethod & "','" & CurrentUserName & "','" & MyCurrentDate & "')", cnCOMMON, adOpenKeyset, adLockOptimistic
        End With
    Else
        Exit Sub
    End If
    
    Save = False
                
    frmODASPSuppliers.cmdAddNew.Caption = "&NEW"
    ShowALLSuppliers

Case Else
        
    Exit Sub

End Select

Exit Sub
err:
    ErrorMessage
End Sub

Public Sub AttachOfficialTitle()
'On Error GoTo err
With frmODASPSuppliers
    If .cboContactTitle.ListCount <> 0 Then Exit Sub
    
    Set rsCOMBO = cnCOMMON.Execute("SELECT Designation FROM ParamDesignations WHERE Designation IS NOT NULL ORDER BY Designation;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!Designation) And rsCOMBO!Designation <> "" Then
            .cboContactTitle.AddItem rsCOMBO!Designation
        End If
    rsCOMBO.MoveNext
    Loop
    
OUTS:
    Set rsCOMBO = Nothing
End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub AttachCurrencies()
'On Error GoTo err
With frmODASPSuppliers

    If .cboCurrencyCode.ListCount <> 0 Then Exit Sub
    
    Set rsCOMBO = cnCOMMON.Execute("SELECT CurrencyName FROM ParamCurrencies WHERE CurrencyName IS NOT NULL ORDER BY CurrencyName;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!CurrencyName) And rsCOMBO!CurrencyName <> "" Then
            .cboCurrencyCode.AddItem rsCOMBO!CurrencyName
        End If
    rsCOMBO.MoveNext
    Loop
    
OUTS:
    Set rsCOMBO = Nothing
    
End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub AttachDeliveryMethods()
'On Error GoTo err
With frmODASPSuppliers
    If .cboDeliveryMethod.ListCount <> 0 Then Exit Sub
    
    Set rsCOMBO = cnCOMMON.Execute("SELECT DeliveryMethod FROM ODASPDeliveryMethods WHERE DeliveryMethod IS NOT NULL ORDER BY DeliveryMethod;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!DeliveryMethod) And rsCOMBO!DeliveryMethod <> "" Then
            .cboDeliveryMethod.AddItem rsCOMBO!DeliveryMethod
        End If
    rsCOMBO.MoveNext
    Loop
    
OUTS:
    Set rsCOMBO = Nothing
End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub AttachTownCity()
'On Error GoTo err
With frmODASPSuppliers
    If .cboTownCity.ListCount <> 0 Then Exit Sub
    
    Set rsCOMBO = cnCOMMON.Execute("SELECT Town FROM ODASPTown WHERE Town IS NOT NULL ORDER BY Town;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!Town) And rsCOMBO!Town <> "" Then
            .cboTownCity.AddItem rsCOMBO!Town
        End If
    rsCOMBO.MoveNext
    Loop
    
OUTS:
    Set rsCOMBO = Nothing
End With
Exit Sub
err:
    ErrorMessage
End Sub

Sub EnableControls()
With frmODASPSuppliers
  .txtSupplierCode.Locked = False
  .txtSupplierType.Locked = True
  .txtSupplyduration.Locked = False
  .txtSupplierName.Locked = False
  .txtPhysicalAddress.Locked = False
  .txtPostalAddress.Locked = False
  .cboTownCity.Locked = False
  .txtTelephoneNo.Locked = False
  .txtfaxNo.Locked = False
  .txtMobileNo.Locked = False
  .txtEmail.Locked = False
  .txtContactPerson.Locked = False
  .cboContactTitle.Locked = False
  
End With
End Sub

Sub disableControls()
With frmODASPSuppliers
'  .txtSupplierCode.Locked = True
  .txtSupplierType.Locked = True
  .txtSupplyduration.Locked = True
  .txtSupplierName.Locked = True
  .txtPhysicalAddress.Locked = True
  .txtPostalAddress.Locked = True
  .cboTownCity.Locked = True
  .txtTelephoneNo.Locked = True
  .txtfaxNo.Locked = True
  .txtMobileNo.Locked = True
   .txtEmail.Locked = True
  .txtContactPerson.Locked = True
  .cboContactTitle.Locked = True
 
 End With
End Sub

Sub ClearControls()
With frmODASPSuppliers
    .txtSupplierCode.Text = Empty
    .txtSupplyduration.Text = Empty
    .txtSupplierName.Text = Empty
    .txtPhysicalAddress.Text = Empty
    .txtPostalAddress.Text = Empty
    .txtTelephoneNo.Text = Empty
    .txtfaxNo.Text = Empty
    .txtMobileNo.Text = Empty
    .txtEmail.Text = Empty
    .txtContactPerson.Text = Empty
    .cboTownCity.Clear
    .cboContactTitle.Clear
    .cboDeliveryMethod.Clear
End With
End Sub

Sub EnableCB()
  With frmODASPSuppliers
            .cmdAddNew.Enabled = True
            .cmdEditRecord.Enabled = True
            .cmdRefresh.Enabled = True
End With
End Sub

Sub DisableCB()
With frmODASPSuppliers
    .cmdAddNew.Enabled = False
    .cmdEditRecord.Enabled = False
    .cmdRefresh.Enabled = False
End With
End Sub

Sub EditRecord()
'On Error GoTo err
    Dim rsFIND As ADODB.Recordset
    Set rsFIND = New ADODB.Recordset
    
    beditRECORD = True
    
    Select Case frmODASPSuppliers.cmdEditRecord.Caption
    Case "E&DIT"
        Call FindRecord
        
        If Edit Then
            frmODASPSuppliers.cmdEditRecord.Caption = "SAVE &CHANGES"
        End If
        
    Case "SAVE &CHANGES"
        Dim rsFinder As ADODB.Recordset
        Set rsFinder = New ADODB.Recordset
        
        With frmODASPSuppliers
            rsFinder.Open "Update ODASPSuppliers SET SupplierCode='" & .txtSupplierCode.Text & "', SupplierType='" & .txtSupplierType.Text & "',SupplyDuration='" & .txtSupplyduration.Text & "',SupplierName='" & .txtSupplierName.Text & "',PhysicalAddress='" & .txtPhysicalAddress.Text & "',PostalAddress='" & .txtPostalAddress.Text & "',TownCity='" & .cboTownCity.Text & "',TelephoneNo='" & .txtTelephoneNo.Text & "',FaxNo='" & .txtfaxNo.Text & "',MobileNo='" & .txtMobileNo.Text & "',EmailAddress='" & .txtEmail.Text & "',ContactPerson='" & .txtContactPerson.Text & "', ContactTitle='" & .cboContactTitle.Text & "',DeliveryMethod='" & .cboDeliveryMethod.Text & "' WHERE SupplierCode='" & .txtSupplierCode.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        End With
                       
        Edit = False
                    
        Found = False
        
        frmODASPSuppliers.cmdEditRecord.Caption = "E&DIT"
    
    Case Else
            Exit Sub
    End Select

    Exit Sub
err:
    ErrorMessage
End Sub

 Sub Refresh()

'On Error GoTo err
If MsgBox("This will cancel any ongoing process,want to continue", vbYesNo) = vbYes Then
EnableControls
ClearControls
frmODASPSuppliers.cmdAddNew.Caption = "&NEW"
frmODASPSuppliers.cmdEditRecord.Caption = "E&DIT"
Save = False
Edit = False
Else
Exit Sub
End If
Exit Sub
err:
ErrorMessage
End Sub

Sub SearchRecord()
'On Error GoTo err

    Dim rsFIND As ADODB.Recordset
    Set rsFIND = New ADODB.Recordset

    EnableCB
    rsFIND.Open "SELECT * FROM ODASPSuppliers WHERE SupplierCode ='" & frmODASPSuppliers.txtSupplierCode & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsFIND
    If .EOF And .BOF Then
        Found = False
        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
    Else
        Found = True
        With frmODASPSuppliers
        
            .txtSupplierCode.Text = rsFIND!supplierCode
            .txtSupplierType.Text = rsFIND!SupplierType
            .txtSupplyduration.Text = rsFIND!SupplyDuration
            .txtSupplierName.Text = rsFIND!SupplierName
            .txtPhysicalAddress.Text = rsFIND!PhysicalAddress
            .txtPostalAddress.Text = rsFIND!postaladdress
            .cboTownCity.Text = rsFIND!towncity
            .txtTelephoneNo.Text = rsFIND!TelephoneNo
            .txtfaxNo.Text = rsFIND!FaxNo
            .txtMobileNo.Text = rsFIND!MobileNo
            .txtEmail.Text = rsFIND!EmailAddress
            .txtContactPerson.Text = rsFIND!ContactPerson
            .cboContactTitle.Text = rsFIND!ContactTitle
            .cboCurrencyCode.Text = rsFIND!CurrencyCode
            .cboDeliveryMethod.Text = rsFIND!DeliveryMethodID
            .txtFreightCharge.Text = rsFIND!FreightCharge
            
        End With
    End If
    End With
        
    If Found = True Then
         Edit = True
    End If
    
Exit Sub
err:
ErrorMessage
End Sub

Function Validate() As Boolean
Dim strvalid As String

With frmODASPSuppliers

If .txtSupplierCode.Text = "" Then
    MsgBox "Supplier Code is required"
    .txtSupplierCode.SetFocus
    
ElseIf .txtSupplierType.Text = "" Then
    MsgBox "The Supplier type is required "
    .txtSupplierType.SetFocus
    
ElseIf .txtSupplierName.Text = "" Then
    MsgBox " The supplier Name is required"
    .txtSupplierName.SetFocus
    Exit Function

ElseIf .txtTelephoneNo.Text = "" Then
    MsgBox "Telephone number required "
    .txtTelephoneNo.SetFocus
    Exit Function
    
ElseIf .cboDeliveryMethod.Text = "" Then
    MsgBox "Shipping Method required "
    .cboDeliveryMethod.SetFocus
    Exit Function
    
    
Else
    Validate = True
End If
End With
End Function

Public Sub GetCurrentSuppliers()
With frmSuppliersList
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Supplier Code", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Supplier Type", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Supply Duration", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Supplier Name", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Postal Address", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Town / City", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Telephone No", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Fax No", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Mobile No", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "E-Mail Address", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Contact Person", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Shipping Method Id", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Freight Charge", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Currency Code", .ListView1.Width / 6.5

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset
rsLIST.Open "SELECT * FROM ODASPSuppliers ", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
.ListView1.View = lvwList
Set MyList = .ListView1.ListItems.Add(, , "Sorry there are no records to display")
Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub

End If



While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!supplierCode))

    If Not IsNull(rsLIST!SupplierType) Then
     MyList.SubItems(1) = CStr(rsLIST!SupplierType)
     End If
     
     If Not IsNull(rsLIST!SupplyDuration) Then
     MyList.SubItems(2) = CStr(rsLIST!SupplyDuration)
     End If
     
     
    If Not IsNull(rsLIST!SupplierName) Then
     MyList.SubItems(3) = CStr(rsLIST!SupplierName)
     End If
     
     If Not IsNull(rsLIST!PhysicalAddress) Then
     MyList.SubItems(4) = CStr(rsLIST!PhysicalAddress)
     End If
     
     If Not IsNull(rsLIST!postaladdress) Then
     MyList.SubItems(5) = CStr(rsLIST!postaladdress)
     End If
     
     If Not IsNull(rsLIST!towncity) Then
     MyList.SubItems(6) = CStr(rsLIST!towncity)
     End If
     
     If Not IsNull(rsLIST!TelephoneNo) Then
     MyList.SubItems(7) = CStr(rsLIST!TelephoneNo)
     End If
     
     If Not IsNull(rsLIST!FaxNo) Then
     MyList.SubItems(8) = CStr(rsLIST!FaxNo)
     End If
     
     If Not IsNull(rsLIST!MobileNo) Then
     MyList.SubItems(9) = CStr(rsLIST!MobileNo)
     End If
     
     If Not IsNull(rsLIST!EmailAddress) Then
     MyList.SubItems(10) = CStr(rsLIST!EmailAddress)
     End If
     
     If Not IsNull(rsLIST!ContactPerson) Then
     MyList.SubItems(11) = CStr(rsLIST!ContactPerson)
     End If
     
     If Not IsNull(rsLIST!DeliveryMethodID) Then
     MyList.SubItems(12) = CStr(rsLIST!DeliveryMethodID)
     End If
     
     If Not IsNull(rsLIST!FreightCharge) Then
     MyList.SubItems(13) = CStr(rsLIST!FreightCharge)
     End If
     
     If Not IsNull(rsLIST!CurrencyCode) Then
     MyList.SubItems(14) = CStr(rsLIST!SupplyDuration)
     End If
     rsLIST.MoveNext
     
     Wend
     
     Set MyList = Nothing: Set rsLIST = Nothing
     
    
     End With
     Exit Sub
                  
End Sub

Public Sub FindRecord()
'On Error GoTo err

With frmODASPSuppliers
    
    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM ODASPSuppliers WHERE Suppliercode='" & Trim(.txtSupplierCode.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
    MsgBox "Sorry there is no record to edit Please enter a Supplier Code", vbCritical + vbOKOnly, "Error Editing"
                Set rsFindRecord = Nothing: Exit Sub
    Else
        .txtSupplierCode.Text = rsFindRecord!supplierCode & ""
        .txtSupplierType.Text = rsFindRecord!SupplierType
            .txtSupplyduration.Text = rsFindRecord!SupplyDuration
            .txtSupplierName.Text = rsFindRecord!SupplierName
            .txtPhysicalAddress.Text = rsFindRecord!PhysicalAddress
            .txtPostalAddress.Text = rsFindRecord!postaladdress
            .cboTownCity.Text = rsFindRecord!towncity
            .txtTelephoneNo.Text = rsFindRecord!TelephoneNo
            .txtfaxNo.Text = rsFindRecord!FaxNo
            .txtMobileNo.Text = rsFindRecord!MobileNo
            .txtEmail.Text = rsFindRecord!EmailAddress
            .txtContactPerson.Text = rsFindRecord!ContactPerson
            .cboContactTitle.Text = rsFindRecord!ContactTitle
            .cboDeliveryMethod.Text = rsFindRecord!DeliveryMethod
            
    End If
    
    Set rsFindRecord = Nothing
      
       
    Edit = True
    
End With

Exit Sub
err:
    ErrorMessage
End Sub

Sub AddNewCustomer()
'On Error GoTo err
Dim strQRE As Variant
Dim rsFIND As ADODB.Recordset

Set rsFIND = New ADODB.Recordset

Select Case frmcustomers.cmdAddNew.Caption
Case "&NEW"
    ClearControls
    EnableCB
    EnableControls
    Save = True
        
    If Save Then
        frmcustomers.cmdAddNew.Caption = "SAVE &CHANGES"
        frmcustomers.txtSupplierCode.SetFocus
    End If
    
Case "SAVE &CHANGES"

    Dim rsFinder As ADODB.Recordset
    Set rsFinder = New ADODB.Recordset
    
    If ValidateCustomers = True Then
        With frmcustomers
            rsFinder.Open ("INSERT INTO customers(Customerid,DoorNo,companyname,contactname,contacttitle,address,city,postalcode,country,phone,fax,FloorNo,email,mobilephone,createdby,datecreated,accperiod) VALUES('" & .txtSupplierCode.Text & "','" & .txtDoorNo.Text & "','" & .txtSupplierName.Text & "','" & .txtContactPerson.Text & "','" & .cboContactTitle.Text & "','" & .txtPostalAddress.Text & "','" & .cboTownCity.Text & "','" & .txtPostalCode.Text & "','" & .txtCountry.Text & "','" & .txtTelephoneNo.Text & "','" & .txtfaxNo.Text & "','" & .txtPhysicalAddress.Text & "','" & .txtEmail.Text & "','" & .txtMobileNo.Text) & "','" & CurrentUserName & "','" & MyCurrentDate & "','" & MyCurrentPeriod & "')", cnCOMMON, adOpenKeyset, adLockOptimistic
        End With
    Else
        Exit Sub
    End If
    
    Save = False
                
    frmcustomers.cmdAddNew.Caption = "&NEW"
    
Case Else
        
    Exit Sub

End Select

Exit Sub
err:
    ErrorMessage
End Sub
Sub RefreshCustomers()

'On Error GoTo err
If MsgBox("This will cancel any ongoing process,want to continue", vbYesNo) = vbYes Then
EnableControlsCustomers
ClearControlsCustomers
frmcustomers.cmdAddNew.Caption = "&NEW"
frmcustomers.cmdEditRecord.Caption = "E&DIT"
Save = False
Edit = False
Else
Exit Sub
End If
Exit Sub
err:
ErrorMessage
End Sub
Sub EnableControlsCustomers()
With frmcustomers
  .txtSupplierCode.Locked = False
'  .txtSupplierType.Locked = False
'  .txtSupplyduration.Locked = False
  .txtSupplierName.Locked = False
  .txtPhysicalAddress.Locked = False
  .txtPostalAddress.Locked = False
  .cboTownCity.Locked = False
  .txtTelephoneNo.Locked = False
  .txtfaxNo.Locked = False
  .txtMobileNo.Locked = False
  .txtEmail.Locked = False
  .txtContactPerson.Locked = False
  .cboContactTitle.Locked = False
  .txtPostalCode.Locked = False
  .txtCountry.Locked = False
  .txtDoorNo.Locked = False
  
End With
End Sub
Sub ClearControlsCustomers()
With frmcustomers
    .txtSupplierCode.Text = Empty
'    .txtSupplyduration.Text = Empty
    .txtSupplierName.Text = Empty
    .txtPhysicalAddress.Text = Empty
    .txtPostalAddress.Text = Empty
    .txtTelephoneNo.Text = Empty
    .txtfaxNo.Text = Empty
    .txtMobileNo.Text = Empty
    .txtEmail.Text = Empty
    .txtContactPerson.Text = Empty
'    .txtFreightCharge.Text = Empty
'    .txtSupplierType.Clear
    .cboTownCity.Clear
    .cboContactTitle.Clear
'    .cboCurrencyCode.Clear
'    .CBODeliveryMethod.Clear
    .txtCountry.Text = Empty
    .txtPostalCode.Text = Empty
    .txtDoorNo.Text = Empty
End With
End Sub
Public Sub GetCurrentCustomers()
With frmCustomersList
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Customer Id", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Company Name", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Contact Person", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Contact Title", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Postal Address", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Town / City", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Country", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Postal Code", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Telephone", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Fax No", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Floor No", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Door No", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "E-Mail", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Mobile No", .ListView1.Width / 6.5

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset
rsLIST.Open "SELECT * FROM customers ", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
.ListView1.View = lvwList
Set MyList = .ListView1.ListItems.Add(, , "Sorry there are no records to display")
Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub

End If



While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!customerid))

    If Not IsNull(rsLIST!CompanyName) Then
     MyList.SubItems(1) = CStr(rsLIST!CompanyName)
     End If
     
     If Not IsNull(rsLIST!Contactname) Then
     MyList.SubItems(2) = CStr(rsLIST!Contactname)
     End If
     
     
    If Not IsNull(rsLIST!ContactTitle) Then
     MyList.SubItems(3) = CStr(rsLIST!ContactTitle)
     End If
     
     If Not IsNull(rsLIST!Address) Then
     MyList.SubItems(4) = CStr(rsLIST!Address)
     End If
     
     If Not IsNull(rsLIST!city) Then
     MyList.SubItems(5) = CStr(rsLIST!city)
     End If
     
     If Not IsNull(rsLIST!country) Then
     MyList.SubItems(6) = CStr(rsLIST!country)
     End If
     
     If Not IsNull(rsLIST!postalcode) Then
     MyList.SubItems(7) = CStr(rsLIST!postalcode)
     End If
     
     If Not IsNull(rsLIST!phone) Then
     MyList.SubItems(8) = CStr(rsLIST!phone)
     End If
     
     If Not IsNull(rsLIST!Fax) Then
     MyList.SubItems(9) = CStr(rsLIST!Fax)
     End If
     
     If Not IsNull(rsLIST!floorno) Then
     MyList.SubItems(10) = (rsLIST!floorno)
     End If
     
     If Not IsNull(rsLIST!doorno) Then
     MyList.SubItems(11) = CStr(rsLIST!doorno)
     End If
     
     If Not IsNull(rsLIST!Email) Then
     MyList.SubItems(12) = CStr(rsLIST!Email)
     End If
     
     If Not IsNull(rsLIST!Mobilephone) Then
     MyList.SubItems(12) = CStr(rsLIST!Mobilephone)
     End If
     
     rsLIST.MoveNext
     
     Wend
     
     Set MyList = Nothing: Set rsLIST = Nothing
     
    
     End With
     Exit Sub
                  
End Sub
Function ValidateCustomers() As Boolean
Dim strvalid As String

With frmcustomers

If .txtSupplierCode.Text = "" Then
    MsgBox "Customer Code is required"
    .txtSupplierCode.SetFocus
    
ElseIf .txtCountry.Text = "" Then
    MsgBox "Country is required "
    .txtCountry.SetFocus
    
ElseIf .txtSupplierName.Text = "" Then
    MsgBox " The Customer Name is required"
    .txtSupplierName.SetFocus
    Exit Function

ElseIf .txtTelephoneNo.Text = "" Then
    MsgBox "Telephone number required "
    .txtTelephoneNo.SetFocus
    Exit Function
    
ElseIf .txtPostalCode.Text = "" Then
    MsgBox "Postal Code required "
    .txtPostalCode.SetFocus
    Exit Function
    
ElseIf .txtDoorNo.Text = "" Then
    MsgBox "Door No Reqiured "
    .txtDoorNo.SetFocus
    Exit Function

ElseIf .txtPhysicalAddress.Text = "" Then
    MsgBox "Floor Number Required "
    .txtPhysicalAddress.SetFocus
    Exit Function
    
Else
    ValidateCustomers = True
End If
End With
End Function
Sub EditRecordCustomers()
'On Error GoTo err
    Dim rsFIND As ADODB.Recordset
    Set rsFIND = New ADODB.Recordset

    Select Case frmcustomers.cmdEditRecord.Caption
    Case "E&DIT"
        Call FindRecordCustomers
'        If Found = False Then
'            MsgBox "Please Do a fresh searching first", vbOKOnly
'            Exit Sub
'        Else
'            Edit = True
'        End If
        
        If Edit Then
            frmcustomers.cmdEditRecord.Caption = "SAVE &CHANGES"
        End If
        
    Case "SAVE &CHANGES"
        Dim rsFinder As ADODB.Recordset
        Set rsFinder = New ADODB.Recordset
        
        With frmcustomers
            rsFinder.Open "Update customers SET customerid='" & .txtSupplierCode.Text & "',DoorNo= '" & .txtDoorNo.Text & "', postalcode='" & .txtPostalCode.Text & "',country='" & .txtCountry.Text & "',companyname='" & .txtSupplierName.Text & "',FloorNo='" & .txtPhysicalAddress.Text & "',Address='" & .txtPostalAddress.Text & "',City='" & .cboTownCity.Text & "',phone='" & .txtTelephoneNo.Text & "',Fax='" & .txtfaxNo.Text & "',Mobilephone='" & .txtMobileNo.Text & "',Email='" & .txtEmail.Text & "',Contactname='" & .txtContactPerson.Text & "', ContactTitle='" & .cboContactTitle.Text & "',datemodified='" & MyCurrentDate & "',modifiedby='" & CurrentUserName & "' WHERE customerid='" & .txtSupplierCode.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        End With
                       
        Edit = False
                    
        Found = False
        
        frmcustomers.cmdEditRecord.Caption = "E&DIT"
    
    Case Else
            Exit Sub
    End Select

    Exit Sub
err:
    ErrorMessage
End Sub
Public Sub FindRecordCustomers()
'On Error GoTo err

With frmcustomers
    
    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM customers WHERE customerid='" & Trim(.txtSupplierCode.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
    MsgBox "Sorry there is no record to edit Please enter a Customer Id", vbCritical + vbOKOnly, "Error Editing"
                Set rsFindRecord = Nothing: Exit Sub
    Else
        .txtSupplierCode.Text = rsFindRecord!customerid & ""
            .txtSupplierName.Text = rsFindRecord!CompanyName
            .txtPhysicalAddress.Text = rsFindRecord!floorno
            .txtPostalAddress.Text = rsFindRecord!Address
            .cboTownCity.Text = rsFindRecord!city
            .txtTelephoneNo.Text = rsFindRecord!phone
            .txtfaxNo.Text = rsFindRecord!Fax
            .txtMobileNo.Text = rsFindRecord!Mobilephone
            .txtEmail.Text = rsFindRecord!Email
            .txtContactPerson.Text = rsFindRecord!Contactname
            .cboContactTitle.Text = rsFindRecord!ContactTitle
'            .cboCurrencyCode.Text = rsFindRecord!CurrencyCode
            .txtCountry.Text = rsFindRecord!country
            .txtPostalCode.Text = rsFindRecord!postalcode
            .txtDoorNo.Text = rsFindRecord!doorno
            
    End If
    
    Set rsFindRecord = Nothing
      
       
    Edit = True
    
End With

Exit Sub
err:
    ErrorMessage
End Sub
