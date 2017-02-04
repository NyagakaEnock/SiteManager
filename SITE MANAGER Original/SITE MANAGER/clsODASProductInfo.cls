VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASProductInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Dim strQRE As Variant

Public Sub GetSupplierCode()
'On Error GoTo err

With frmODASPProducts
    
    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM ODASPSuppliers WHERE suppliername='" & Trim(.cboSupplierCode.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .cboSupplierCode.Text = rsFindRecord!supplierCode & ""
        .txtSupplierName.Text = rsFindRecord!SupplierName & ""
        .txtProductCode.SetFocus
    End If
    
    Set rsFindRecord = Nothing
    
End With

Exit Sub
err:
    ErrorMessage
End Sub

Public Sub GetCategoryCode()
'On Error GoTo err

With frmODASPProducts
    
    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM ODASPProductCategory WHERE ProductCategoryDescription='" & Trim(.txtCategoryCode.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .txtCategoryCode.Text = rsFindRecord!categorycode & ""
        .cboSupplierCode.SetFocus
    End If
    
    Set rsFindRecord = Nothing
    
End With

Exit Sub
err:
    ErrorMessage
End Sub

Public Sub GetPackageType()
'On Error GoTo err

With frmODASPProducts
    
    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM PharmPackageTypes WHERE Descriptions='" & Trim(.cboPackageType.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .cboPackageType.Text = rsFindRecord!packagetype & ""
        .txtProductName.SetFocus
    End If
    
    Set rsFindRecord = Nothing
    
End With

Exit Sub
err:
    ErrorMessage
End Sub

Public Sub AttachPackageTypes()
'On Error GoTo err
With frmODASPProducts

    If .cboPackageType.ListCount <> 0 Then Exit Sub
    
    Set rsCOMBO = cnCOMMON.Execute("SELECT Descriptions FROM PharmPackageTypes WHERE Descriptions IS NOT NULL ORDER BY PackageType;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!descriptions) And rsCOMBO!descriptions <> "" Then
            .cboPackageType.AddItem rsCOMBO!descriptions
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

Sub AddNewCat()
'On Error GoTo err
Dim strQRE As Variant
Dim rsFIND As ADODB.Recordset

Set rsFIND = New ADODB.Recordset

Select Case frmODASPProducts.cmdAddNew.Caption
Case "&NEW"
    ClearControlsCat
    EnableCBCat
    EnableControlsCat
    Save = True
         
    If Save Then
        frmODASPProducts.cmdAddNew.Caption = "SAVE &CHANGES"
        frmODASPProducts.txtCategoryCode.SetFocus
    End If

Case "SAVE &CHANGES"
    Dim rsFinder As ADODB.Recordset
    Set rsFinder = New ADODB.Recordset
    
    If ValidateCat = True Then
        With frmODASPProducts
            rsFinder.Open "INSERT INTO ODASPProductCategory(CategoryCode,CategoryName) VALUES('" & .txtCategoryCode.Text & "','" & .txtCategoryName.Text & "')", cnCOMMON, adOpenKeyset, adLockOptimistic
        End With
    Else
        Exit Sub
    End If
    
    Save = False
    
    showProductCategory
    frmODASPProducts.cmdAddNew.Caption = "&NEW"

Case Else
        
    Exit Sub

End Select

Exit Sub
err:
ErrorMessage

End Sub
Sub EnableControlsCat()
With frmODASPProducts
 .txtCategoryCode.Locked = False
 .txtCategoryName.Locked = False
        
End With

End Sub
Sub disableControlsCat()
With frmODASPProducts
    .txtCategoryCode.Locked = True
    .txtCategoryName.Locked = True
End With
End Sub
Sub ClearControlsCat()
 With frmODASPProducts
    .txtCategoryCode.Text = ""
    .txtCategoryName.Text = ""
    End With
End Sub
Sub EnableCBCat()
  With frmODASPProducts
            .cmdAddNew.Enabled = True
            .cmdEditRecord.Enabled = True
            .cmdRefresh.Enabled = True
End With
End Sub

Sub DisableCBCat()
With frmODASPProducts
            .cmdAddNew.Enabled = False
            .cmdEditRecord.Enabled = False
            .cmdRefresh.Enabled = False
End With
End Sub

Sub EditRecordCat()
'On Error GoTo err

Dim rsFIND As ADODB.Recordset
Set rsFIND = New ADODB.Recordset

        Select Case frmODASPProducts.cmdEditRecord.Caption
                Case "E&DIT"
                    Call FindRecord
'                      If Found = False Then
'                      MsgBox "please Refresh and search a fresh", vbOKOnly
'                       Exit Sub
'                        Else
'                            Edit = True
'                        End If
        
                        If Edit Then
                                frmODASPProducts.cmdEditRecord.Caption = "SAVE &CHANGES"
                        End If
    
                Case "SAVE &CHANGES"
                        Dim rsFinder As ADODB.Recordset
                        Set rsFinder = New ADODB.Recordset
    With frmODASPProducts
  rsFinder.Open "UPDATE ODASPProductCategory SET CategoryCode='" & .txtCategoryCode.Text & "', ProductCategoryDescription='" & .txtCategoryName.Text & "' WHERE CategoryCode='" & .txtCategoryCode.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

    End With
                     
                            Edit = False
                   
                         Found = False
                    frmODASPProducts.cmdEditRecord.Caption = "E&DIT"
            Case Else
        
            Exit Sub

        End Select

Exit Sub
err:
ErrorMessage
End Sub

 Sub RefreshCat()

'On Error GoTo err
If MsgBox("This will cancel any ongoing process,want to continue", vbYesNo) = vbYes Then
EnableControlsCat
ClearControlsCat
frmODASPProducts.cmdAddNew.Caption = "&NEW"
frmODASPProducts.cmdEditRecord.Caption = "E&DIT"
Save = False
Edit = False
Else
Exit Sub
End If
Exit Sub
err:
ErrorMessage
End Sub

Sub SearchRecordCat()

'On Error GoTo err

Dim rsFIND As ADODB.Recordset
Set rsFIND = New ADODB.Recordset

                      EnableCBCat
                      EnableCBCat
                        
                       rsFIND.Open "SELECT * FROM ODASPProductCategory WHERE CategoryCode ='" & frmODASPProducts.txtCategoryCode & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

                        With rsFIND
                                If .EOF And .BOF Then
                                Found = False
                                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                                        
                                Else
                                Found = True
                                frmODASPProducts.txtCategoryCode = rsFIND!categorycode
                                frmODASPProducts.txtCategoryName = rsFIND!categoryname
                                       
                                End If
                        End With
                       If Found = True Then
                            Edit = True
                        End If
Exit Sub
err:
ErrorMessage
End Sub

Function ValidateCat() As Boolean
With frmODASPProducts
If .txtCategoryCode.Text = "" Then
    MsgBox "The Category code is required"
    .txtCategoryCode.SetFocus
    
    ElseIf .txtCategoryName.Text = "" Then
    MsgBox "The Category name is required "
    .txtCategoryName.SetFocus
     
     Else: ValidateCat = True
End If
End With
End Function




Sub AddNewDetails()
'On Error GoTo err
Dim strQRE As Variant
Dim rsFIND As ADODB.Recordset

Set rsFIND = New ADODB.Recordset

Select Case frmODASPProducts.cmdAddNew.Caption
Case "&NEW"
    ClearControlsDetails
    EnableCBDetails
    EnableControlsDetails
    Save = True

    If Save Then
        frmODASPProducts.cmdAddNew.Caption = "SAVE &CHANGES"
    End If

Case "SAVE &CHANGES"

    Dim rsFinder As ADODB.Recordset
    Set rsFinder = New ADODB.Recordset
    
    If ValidateDetails = True Then
        With frmODASPProducts
            rsFinder.Open "INSERT INTO ODASPProduct(ProductCode,ProductName,CategoryCode,Preparedby,DatePrepared) VALUES('" & .txtProductCode.Text & "','" & StrConv(.txtProductName.Text, vbProperCase) & "','" & .txtCategoryCode.Text & "','" & CurrentUserName & "','" & MyCurrentDate & "')", cnCOMMON, adOpenKeyset, adLockOptimistic
        End With
    Else
        Exit Sub
    End If
    
    Save = False
    showALLProducts
    frmODASPProducts.cmdAddNew.Caption = "&NEW"
    
Case Else
        
    Exit Sub

End Select
    Exit Sub
err:
    ErrorMessage
End Sub

Sub EnableControlsDetails()
'On Error GoTo err
    With frmODASPProducts
      .txtCategoryCode.Locked = False
      .txtProductCode.Locked = False
      .txtProductName.Locked = False
      .txtCategoryName.Locked = False
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Sub disableControlsDetails()
    With frmODASPProducts
      .txtCategoryCode.Locked = True
      .txtProductCode.Locked = True
      .txtProductName.Locked = True
      .cboSupplierCode.Locked = True
    End With
End Sub

Sub ClearControlsDetails()
    With frmODASPProducts
      .txtProductCode.Text = ""
      .txtProductName.Text = ""
    End With
End Sub

Sub EnableCBDetails()
    With frmODASPProducts
        .cmdAddNew.Enabled = True
        .cmdEditRecord.Enabled = True
        .cmdRefresh.Enabled = True
    End With
End Sub

Sub DisableCBDetails()
With frmODASPProducts
    .cmdAddNew.Enabled = False
    .cmdEditRecord.Enabled = False
    .cmdRefresh.Enabled = False
End With
End Sub

Sub EditRecordDetails()
''On Error GoTo Err
    Dim rsFIND As ADODB.Recordset
    Set rsFIND = New ADODB.Recordset

    Select Case frmODASPProducts.cmdEditRecord.Caption
    Case "E&DIT"
       Call FindDetailsRecord
                      
'        If Found = False Then
'            MsgBox "Please Do a fresh searching first", vbOKOnly
'            Exit Sub
'        Else
'            Edit = True
'        End If

'        If Edit Then
'            frmODASPProducts.cmdEditRecord.Caption = "SAVE &CHANGES"
'        End If
'
    Case "SAVE &CHANGES"
    
        Dim rsFinder As ADODB.Recordset
        Set rsFinder = New ADODB.Recordset
        
        With frmODASPProducts
            rsFinder.Open "Update ODASPProduct SET ProductCode ='" & .txtProductCode.Text & "',ProductName='" & StrConv(.txtProductName.Text, vbProperCase) & "',CategoryCode='" & .txtCategoryCode.Text & "' WHERE Productcode ='" & Trim(.txtProductCode.Text) & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        End With
                       
        Edit = False
                    
        Found = False
        
        frmODASPProducts.cmdEditRecord.Caption = "E&DIT"
        
    Case Else
        
        Exit Sub

    End Select

Exit Sub
err:
    ErrorMessage
End Sub

Sub RefreshDetails()
'On Error GoTo err
If MsgBox("This action will cancel any ongoing process,do you want to continue", vbYesNo) = vbYes Then
    
    EnableControlsDetails
    ClearControlsDetails
    
    frmODASPProducts.cmdAddNew.Caption = "&NEW"
    frmODASPProducts.cmdEditRecord.Caption = "E&DIT"
    
    Save = False
    Edit = False
    
Else
    Exit Sub
End If

Exit Sub
err:
    ErrorMessage
End Sub

Sub SearchRecordDetails()
'On Error GoTo err

Dim rsFIND As ADODB.Recordset
Set rsFIND = New ADODB.Recordset

                      EnableCBDetails
                      EnableCBDetails
                        
'                        rsFind.Open "SELECT * FROM ParamDrugDetails WHERE DrugId ='" & frmODASPProducts.txtDrugId.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

                        With rsFIND
                                If .EOF And .BOF Then
                                Found = False
                                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                                        
                                Else
                                Found = True
                        With frmODASPProducts
                            .txtProductCode.Text = rsFIND!ProductCode
                            .txtProductName.Text = rsFIND!ProductName
                            .txtCategoryCode.Text = rsFIND!categorycode & ""
                            .cboSupplierCode.Text = rsFIND!supplierCode
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

Function ValidateDetails() As Boolean
Dim strvalid As String

With frmODASPProducts

    
    If .txtCategoryCode.Text = "" Then
    MsgBox "The Product code is required "
    .txtCategoryCode.SetFocus
    
    ElseIf .txtCategoryName.Text = "" Then
    MsgBox " Name of the Category required"
    .txtCategoryName.SetFocus
    
Else: ValidateDetails = True
End If
End With
End Function


Public Sub FindRecord()
'On Error GoTo err

With frmODASPProducts
    
    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM ODASPProductCategory WHERE CategoryCode='" & Trim(.txtCategoryCode.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
    MsgBox "Sorry there is no record to edit Please enter a Category Code", vbCritical + vbOKOnly, "Error Editing"
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .txtCategoryCode.Text = rsFindRecord!categorycode & ""
        .txtCategoryName.Text = rsFindRecord!categoryname
    End If
    
    Set rsFindRecord = Nothing
    
    Edit = True
    
End With

Exit Sub
err:
    ErrorMessage
End Sub
Private Function RecordExists() As Boolean
'On Error GoTo err
With frmODASPProducts
    If .txtCategoryCode.Text = Empty Then
        strMessage = "There is no Current Record to Edit. Search and Display a Record then Retry Editing!!!"
'        .txtDeliveryMethodID.SetFocus
    Else
        RecordExists = True
    End If
    If Not RecordExists Then
        MsgBox strMessage, vbCritical + vbOKOnly, "Error Editing"
    End If
End With
Exit Function
err:
    ErrorMessage
End Function
Public Sub GetDetailsAllProducts()
With frmODASPProducts
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear
.Frame2.Caption = "List Of All Products Per Supplier"

.ListView1.ColumnHeaders.Add , , "Supplier Code", .ListView1.Width / 7.5
.ListView1.ColumnHeaders.Add , , "Supplier Name", .ListView1.Width / 3.5
.ListView1.ColumnHeaders.Add , , "Product Name", .ListView1.Width / 3.5
.ListView1.ColumnHeaders.Add , , "Product Code", .ListView1.Width / 7.5
.ListView1.ColumnHeaders.Add , , "Package Type", .ListView1.Width / 7.5


.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset
rsLIST.Open "SELECT * FROM ODASPProduct A,ODASPSuppliers B WHERE B.SupplierCode=A.SupplierCode ORDER BY B.SupplierName  ", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
.ListView1.View = lvwList
Set MyList = .ListView1.ListItems.Add(, , "Sorry there are no records to display")
Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub

End If



While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!supplierCode))

    If Not IsNull(rsLIST!SupplierName) Then
     MyList.SubItems(1) = CStr(rsLIST!SupplierName)
     End If
     
     If Not IsNull(rsLIST!ProductName) Then
     MyList.SubItems(2) = CStr(rsLIST!ProductName)
     End If
     
     If Not IsNull(rsLIST!ProductCode) Then
     MyList.SubItems(3) = CStr(rsLIST!ProductCode)
     End If
     
     If Not IsNull(rsLIST!packagetype) Then
     MyList.SubItems(4) = CStr(rsLIST!packagetype)
     End If
     
     
               
     rsLIST.MoveNext
     
     Wend
     
     Set MyList = Nothing: Set rsLIST = Nothing
     
    
     End With
     Exit Sub
                  
End Sub
     
Public Sub FindDetailsRecord()
''On Error GoTo Err

With frmODASPProducts
    
    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM ODASPProduct WHERE ProductCode ='" & Trim(.txtProductCode.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
    MsgBox "Sorry there is no record to edit Please enter a Product Code ", vbCritical + vbOKOnly, "Error Editing"
       
        Set rsFindRecord = Nothing: Exit Sub
    .txtProductCode.SetFocus
    Else
        .txtProductCode.Text = rsFindRecord!ProductCode & ""
        .txtProductName.Text = rsFindRecord!ProductName
        .txtCategoryCode.Text = rsFindRecord!categorycode
    End If
    
    Set rsFindRecord = Nothing
    
    Edit = True
    .cmdEditRecord.Caption = "SAVE &CHANGES"
    
End With

Exit Sub
err:
    ErrorMessage
End Sub

