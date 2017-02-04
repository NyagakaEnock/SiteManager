VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowDrugsDosages()
'On Error GoTo err
With frmODASPProductPriceSetup
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Product Code", .ListView1.Width / 6#
.ListView1.ColumnHeaders.Add , , "Product Name", .ListView1.Width / 3.2
.ListView1.ColumnHeaders.Add , , "Package Type", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Quantity", .ListView1.Width / 6#
.ListView1.ColumnHeaders.Add , , "Qty Units", .ListView1.Width / 7

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM GenProductsInventory WHERE CostPriceStatus IS NULL ORDER BY DrugCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: .txtTotal.Text = 0: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!drugcode))


    If Not IsNull(rsLIST!drugname) Then
        MyList.SubItems(1) = CStr(rsLIST!drugname)
    End If
     
    If Not IsNull(rsLIST!packagetype) Then
        MyList.SubItems(2) = CStr(rsLIST!packagetype)
    End If
    
    If Not IsNull(rsLIST!currentquantity) Then
        MyList.SubItems(3) = CStr(rsLIST!currentquantity)
    End If
    
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(4) = CStr(rsLIST!QuantityUnits)
    End If
    
    rsLIST.MoveNext
    
Wend

.txtTotal.Text = .ListView1.ListItems.Count

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub AttachDrugCategories()
'On Error GoTo err
With frmODASPProductPriceSetup

    If .cboCategoryCode.ListCount <> 0 Then Exit Sub
    
    Set rsCOMBO = cnCOMMON.Execute("SELECT CategoryName FROM ParamProductCategories WHERE CategoryName IS NOT NULL ORDER BY CategoryName;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!categoryname) And rsCOMBO!categoryname <> "" Then
            .cboCategoryCode.AddItem rsCOMBO!categoryname
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

Public Function GetPriceMarkup() As Double
'On Error GoTo err
With frmODASPProductPriceSetup

    Set rsFindRecord = cnCOMMON.Execute("SELECT PriceMarkup FROM PharmPriceMarkup;")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetPriceMarkup = 1.33
    ElseIf IsNull(rsFindRecord!pricemarkup) = True Or rsFindRecord!pricemarkup = "" Then
        GetPriceMarkup = 1.33
    Else
        GetPriceMarkup = CDbl(rsFindRecord!pricemarkup)
    End If
    
    Set rsFindRecord = Nothing
    
End With
Exit Function
err:
    ErrorMessage
End Function

Public Sub ClearForNewRecord()
With frmODASPProductPriceSetup
    .txtDosageCost.Text = Empty
    .txtDosagePrice.Text = Empty
    .txtProductCode.Text = Empty
    .txtDrugName.Text = Empty
End With
End Sub

Public Sub ClearTheScreen()
With frmODASPProductPriceSetup
    .ListView1.ListItems.Clear
    .cboCategoryCode.Clear
    .chkSearchEdit.Value = 0
    .txtCategoryName.Text = Empty
    .txtDosageCost.Text = Empty
    .txtDosagePrice.Text = Empty
    .txtProductCode.Text = Empty
    .txtDrugName.Text = Empty
    .txtPriceMarkup.Text = Empty
    .txtTotal.Text = Empty
End With
End Sub

Public Sub GetCategoryByName()
'On Error GoTo err
With frmODASPProductPriceSetup

    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM ParamProductCategories WHERE CategoryName='" & Trim(.cboCategoryCode.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .txtCategoryName.Text = rsFindRecord!categoryname & ""
        .cboCategoryCode.Text = rsFindRecord!categorycode & ""
        
        .ListView1.SetFocus
        
        If .cboCategoryCode.Text = "AAA" Then
            Call ShowDrugsDosages
        Else
            Call ShowDrugsDosagesByCategory
        End If
        
    End If
    
End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub GetMainStructure()
'On Error GoTo err
With frmODASPProductPriceSetup

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear


.ListView1.ColumnHeaders.Add , , "Product Code", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Product Name", .ListView1.Width / 3.2
.ListView1.ColumnHeaders.Add , , "Package Type", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Quantity", .ListView1.Width / 6#  ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Qty Units", .ListView1.Width / 7

.ListView1.View = lvwReport

End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub CheckForRecord()
'On Error GoTo err
    If RecordExists Then
        AllowEdit = True
    Else
        AllowEdit = False
    End If
Exit Sub
err:
    ErrorMessage
End Sub

Private Function RecordExists() As Boolean
'On Error GoTo err
With frmODASPProductPriceSetup
    If .txtPriceMarkup.Text = Empty Then
        strMessage = "There is no Current Record to Edit. Search and Display a Record then Retry Editing!!!"

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

Public Sub EditCurrentRecord()
'On Error GoTo err
With frmODASPProductPriceSetup
    Set rsEditRecord = New ADODB.Recordset
    
    MySQL = "UPDATE ProductsCostPriceSetup SET dosagecost=" & CDbl(.txtDosageCost.Text) & ",dosageprice=" & CDbl(.txtDosagePrice.Text) & " WHERE DrugCode='" & Trim(.txtProductCode.Text) & "';"
    
    rsEditRecord.Open MySQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Set rsEditRecord = Nothing

    .cmdCHANGE.Caption = "&CHANGE"
        
    EditRecord = False: AllowEdit = False
    
End With
    Exit Sub
err:
    UpdateErrorMessage
End Sub


Public Sub ShowDrugsDosagesByCategory()
'On Error GoTo err
With frmODASPProductPriceSetup
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear


.ListView1.ColumnHeaders.Add , , "Product Code", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Product Name", .ListView1.Width / 3.2
.ListView1.ColumnHeaders.Add , , "Package Type", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Quantity", .ListView1.Width / 6#  ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Qty Units", .ListView1.Width / 7

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM GenProductsInventory WHERE CategoryCode='" & Trim(.cboCategoryCode.Text) & "' AND CostPriceStatus IS NULL ORDER BY DrugCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: .txtTotal.Text = 0: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!drugcode))

    If Not IsNull(rsLIST!drugname) Then
        MyList.SubItems(1) = CStr(rsLIST!drugname)
    End If
    
    If Not IsNull(rsLIST!packagetype) Then
        MyList.SubItems(2) = CStr(rsLIST!packagetype)
    End If
    
    If Not IsNull(rsLIST!currentquantity) Then
        MyList.SubItems(3) = CStr(rsLIST!currentquantity)
    End If
    
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(4) = CStr(rsLIST!QuantityUnits)
    End If
    
    
    
    rsLIST.MoveNext
    
Wend

.txtTotal.Text = .ListView1.ListItems.Count

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub FindAllSettings()
'On Error GoTo err
With frmODASPProductPriceSetup

If .chkSearchEdit.Value = 0 Then Exit Sub

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Product Code", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Product Name", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Product Cost", .ListView1.Width / 5.9
.ListView1.ColumnHeaders.Add , , "Product Price", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Price Markup", .ListView1.Width / 5.5 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ProductsCostPriceSetup WHERE costpricestatus IS NOT NULL ORDER BY DrugCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: .txtTotal.Text = 0: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!drugcode))
    
    If Not IsNull(rsLIST!drugname) Then
        MyList.SubItems(1) = CStr(rsLIST!drugname)
    End If
    
    
    If Not IsNull(rsLIST!DosageCost) Then
        MyList.SubItems(2) = FormatNumber(rsLIST!DosageCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!dosageprice) Then
        MyList.SubItems(3) = FormatNumber(rsLIST!dosageprice, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!pricemarkup) Then
        MyList.SubItems(4) = FormatNumber(rsLIST!pricemarkup, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    
    rsLIST.MoveNext
    
Wend

.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight

.txtTotal.Text = .ListView1.ListItems.Count

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub


Public Sub FindSettingsBySpecifiedDate()
'On Error GoTo err
With frmODASPProductPriceSetup

If .chkSearchEdit.Value = 0 Then Exit Sub

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear


.ListView1.ColumnHeaders.Add , , "Product Code", .ListView1.Width / 5.9
.ListView1.ColumnHeaders.Add , , "Product Cost", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Product Price", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Price Markup", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Product Name", .ListView1.Width / 3

.ListView1.View = lvwReport

INPQRY = InputBox("Please Enter the Date on Which to Show Settings..." & vbCrLf & vbCrLf & "Format: - dd/MM/yyyy", "Enter Date", Date)

If Len(INPQRY) = 0 Then
    MsgBox "No Values Entered or the Operation has been Cancelled!!", vbCritical + vbOKOnly, "No Values"
    Exit Sub
End If

Dim q As Date, ThisDate As String

q = CDate(INPQRY): ThisDate = Format(q, "MMMM dd,yyyy")

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ProductsCostPriceSetup WHERE datecreated='" & Trim(ThisDate) & "' ORDER BY ProductsCostPriceSetup.DrugCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: .txtTotal.Text = 0: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!drugcode))
    
    If Not IsNull(rsLIST!DosageCost) Then
        MyList.SubItems(1) = FormatNumber(rsLIST!DosageCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!dosageprice) Then
        MyList.SubItems(2) = FormatNumber(rsLIST!dosageprice, 2, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!pricemarkup) Then
        MyList.SubItems(3) = FormatNumber(rsLIST!pricemarkup, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!drugname) Then
        MyList.SubItems(4) = CStr(rsLIST!drugname)
    End If
    
       
    rsLIST.MoveNext
    
Wend

.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight

.txtTotal.Text = .ListView1.ListItems.Count

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub FindSettingsForDateToday()
'On Error GoTo err
With frmODASPProductPriceSetup

If .chkSearchEdit.Value = 0 Then Exit Sub

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear


.ListView1.ColumnHeaders.Add , , "Product Code", .ListView1.Width / 5.9
.ListView1.ColumnHeaders.Add , , "Product Cost", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Product Price", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Price Markup", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Drug Name", .ListView1.Width / 3

.ListView1.View = lvwReport

Dim ThisDate As String

ThisDate = Format(Date, "MMMM dd,yyyy")

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ProductsCostPriceSetup WHERE datecreated='" & Trim(ThisDate) & "' ORDER BY ProductsCostPriceSetup.DrugCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: .txtTotal.Text = 0: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!drugcode))
    
    If Not IsNull(rsLIST!DosageCost) Then
        MyList.SubItems(1) = FormatNumber(rsLIST!DosageCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!dosageprice) Then
        MyList.SubItems(2) = FormatNumber(rsLIST!dosageprice, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!pricemarkup) Then
        MyList.SubItems(3) = FormatNumber(rsLIST!pricemarkup, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!drugname) Then
        MyList.SubItems(4) = CStr(rsLIST!drugname)
    End If
    
        
    rsLIST.MoveNext
    
Wend

.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight

.txtTotal.Text = .ListView1.ListItems.Count

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub AddNewCostSetup()
If EditRecord Then Exit Sub
    NewRecord = True
    With frmODASPProductPriceSetup
    
        Call ClearForNewRecord
        
        .cmdNEW.Caption = "&SAVE RECORD"
        .cmdCHANGE.Enabled = False
        
    End With
End Sub

Public Sub SaveCostPriceSetup()
'On Error GoTo err
With frmODASPProductPriceSetup
    
    Set rsNewRecord = New ADODB.Recordset
    
    MySQL = "INSERT INTO ProductsCostPriceSetup (drugcode,drugname,dosagecost,pricemarkup,dosageprice,createdby,datecreated,accperiod) VALUES('" & .txtProductCode & "','" & .txtDrugName & "'," & CDbl(.txtDosageCost.Text) & "," & CDbl(.txtPriceMarkup.Text) & "," & CDbl(.txtDosagePrice.Text) & ",'" & CurrentUserName & "','" & MyCurrentDate & "','" & MyCurrentPeriod & "')"
    
    rsNewRecord.Open MySQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Set rsNewRecord = Nothing
    
    
    Set rsLineUpdate = New ADODB.Recordset
    
    rsLineUpdate.Open "UPDATE GenProductsInventory SET CostPriceStatus='" & "Y" & "' WHERE drugCode='" & Trim(.txtProductCode.Text) & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Set rsLineUpdate = Nothing
    
    Set rsLineUpdate = New ADODB.Recordset
    
    rsLineUpdate.Open "UPDATE ProductsCostPriceSetup SET CostPriceStatus='" & "Y" & "' WHERE drugCode='" & Trim(.txtProductCode.Text) & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Set rsLineUpdate = Nothing
    
    .cmdNEW.Caption = "&NEW"
    .cmdCHANGE.Enabled = True
    
    If .cboCategoryCode.Text = Empty Then
        Call ShowDrugsDosages
    ElseIf .cboCategoryCode.Text <> Empty Then
        Call ShowDrugsDosagesByCategory
    End If
    
    NewRecord = False
    
End With

    Exit Sub
    
err:
    UpdateErrorMessage
End Sub

Public Sub EditCostPriceSettings()
'On Error GoTo err
With frmODASPProductPriceSetup
    Set rsEditRecord = New ADODB.Recordset
    
'    MySQL = "UPDATE ProductQuantitySetup SET drugcode='" & Trim(.txtProductCode.Text) & "',drugname='" & Trim(.txtDrugName.Text) & "',dosagetype='" & Trim(.cboDosageType.Text) & "',minage=" & CDbl(.txtMinAge.Text) & ",maxage=" & CDbl(.txtMaxAge.Text) & ",packagetype='" & Trim(.txtPackageType.Text) & "',quantity=" & CDbl(.txtTotalQuantity.Text) & ",quantityunits='" & Trim(.cboQuanityUnits.Text) & "',dailyinterval=" & CDbl(.txtTimesPerDay.Text) & ",intervalqty=" & CDbl(.txtQtyPerUse.Text) & ",intervalunits='" & Trim(.cboUnitsOfQty.Text) & "',totaldays=" & CDbl(.txtTotalDays.Text) & " WHERE ShippingMethodID='" & Trim(.txtShippingMEthodID.Text) & "';"
    
    rsEditRecord.Open MySQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Set rsEditRecord = Nothing

    .cmdCHANGE.Caption = "&CHANGE"
        
    EditRecord = False: AllowEdit = False
    
End With
    Exit Sub
err:
    UpdateErrorMessage
End Sub

Public Sub GetDosageRecords()
With frmODASPProductPriceSetup

    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM ProductsCostPriceSetup WHERE drugcode='" & Trim(.txtProductCode.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
    MsgBox "Sorry there is no record to edit ", vbCritical + vbOKOnly, "Error Editing"
                Set rsFindRecord = Nothing: Exit Sub
    Else
             .txtDosageCost.Text = rsFindRecord!DosageCost & ""
             .txtDosagePrice.Text = rsFindRecord!dosageprice
            
    End If
    
    Set rsFindRecord = Nothing
      
       
    Edit = True
    
End With

Exit Sub
err:
    ErrorMessage
End Sub

