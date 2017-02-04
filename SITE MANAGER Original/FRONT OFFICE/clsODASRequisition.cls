VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASRequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub LoadItemCOST()
On Error GoTo err
        With frmODASMOpenJobCard
        
                Set rsCONTROL = New ADODB.Recordset
                strSQL = "SELECT * FROM ODASPSupplierProduct WHERE ProductCode = '" & .txtProductCode.Text & "' and AccountNo = '" & .txtAccountNo.Text & "'; "
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
                If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
                    .txtProductCode.Text = rsCONTROL!ProductCode
                    .txtUnitPrice.Text = FormatNumber(rsCONTROL!UnitPrice)
                    .txtVATRate.Text = rsCONTROL!VATRate
                    .txtItemQuantity.Text = 1
                    .txtExchangeRate.Text = 1
                    .txtUnitCode.Text = rsCONTROL!UnitCode

        
        End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadCostCenter()
On Error GoTo err
        With frmODASMOpenJobCard
        
                Set rsCONTROL = New ADODB.Recordset
                strSQL = "SELECT * FROM ODASPProduct WHERE ProductCode = '" & .txtProductCode.Text & "' ; "
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
                If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
                .txtCostCenter.Text = rsCONTROL!CostCentre
 
        End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub LoadDEFAULT()
On Error GoTo err
        
        
        With frmODASMOpenJobCard
        
            Set rsCONTROL = New ADODB.Recordset
            strSQL = "SELECT * FROM ODASMJobCard WHERE JobCardNo = '" & .txtJobCardNo.Text & "' and DepartmentCode = '" & .txtDepartmentCode.Text & "'; "
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

            .txtUnitPrice.Text = 0
            .txtVATAmount.Text = 0
            .txtItemQuantity.Text = 1
            .UpDownQuantity.Value = 1
            .txtSupervisedBy = rsCONTROL!openedBy & ""
            .txtDoneBy.Text = CurrentUserName
            .txtDateOfCommencement.Text = Date
            .txtRequisitionDate.Text = Date
            .txtExchangeRate.Text = 1
            .txtTotalPriceExcl.Text = 0
            .txtTotalCost.Text = 0
            .txtTotalUnitPriceExcl.Text = 0
            
            
            
        End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadJOBBRIEF()
On Error GoTo err
    With Screen.ActiveForm
        
            Set rsCONTROL = New ADODB.Recordset
            strSQL = "SELECT * FROM ODASMJobBrief, ODASPAccount WHERE ODASMJobBrief.AccountNo = ODASPAccount.AccountNo and ODASMJobBrief.JobBriefNo = '" & .txtJobCardNo.Text & "' ; "
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
            
            .txtCustomerName = rsCONTROL!CompanyName
            .txtDeadlineDate = rsCONTROL!deadlineDate
            .txtLPONo.Text = rsCONTROL!LPONo & ""
            .txtDescriptionOfOrder = rsCONTROL!descriptionOfOrder
            .txtJobBriefDate = rsCONTROL!JobBriefDate

    End With
    
strSQL = Empty
rsCONTROL.Close
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub obtainBASECURRENCY()
On Error GoTo err
       With frmODASMOpenJobCard
            Set rsCONTROL = New ADODB.Recordset
            strSQL = "SELECT * FROM ODASPCurrency WHERE BaseCurrency = 'Y'; "
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
            
            .txtExchangeRate.Text = 1
            .cboCurrencyCode.Text = rsCONTROL!CurrencyCode
        End With
rsCONTROL.Close
strSQL = Empty
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub updateREQUISITIONTOTALS()
On Error GoTo err
        With frmODASMOpenJobCard

                Set rsSAVE = New ADODB.Recordset
                strSQL = "SELECT * FROM ODASMRequisition WHERE RequisitionNo = '" & .txtRequisitionNo.Text & "' "
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
                    
                If rsSAVE.EOF And rsSAVE.BOF Then Exit Sub
                 
                '/ Sum The Quantities on the Requisition
                
                Set rsCONTROL = New ADODB.Recordset
                strCONTROL = "SELECT sum(Quantity) as Totals FROM ODASMRequisitionItems WHERE RequisitionNo = '" & .txtRequisitionNo.Text & "' "
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                            
                If rsCONTROL.BOF Or rsCONTROL.EOF Then Exit Sub
                
                rsSAVE!Quantity = CDbl(rsCONTROL!TOTALS)
                            
                '/ Procedure to sum the total price on the Requisition
                
                Set rsCONTROL = New ADODB.Recordset
                strCONTROL = "SELECT sum(PriceExclusive) as Totals FROM ODASMRequisitionItems WHERE RequisitionNo = '" & .txtRequisitionNo.Text & "' "
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                            
                If rsCONTROL.BOF Or rsCONTROL.EOF Then Exit Sub

                
                rsSAVE!PriceExclusive = CDbl(rsCONTROL!TOTALS)
                            
                '/Procedure to sum the Total Price Inclusive of VAT
                
                Set rsCONTROL = New ADODB.Recordset
                strCONTROL = "SELECT sum(PriceInclusive) as Totals FROM ODASMRequisitionItems WHERE RequisitionNo = '" & .txtRequisitionNo.Text & "' "
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                            
                If rsCONTROL.BOF Or rsCONTROL.EOF Then Exit Sub

                rsSAVE!PriceInclusive = CDbl(rsCONTROL!TOTALS)
                
                '/Procedure to sum up VAT on All the Requisitions Raised
                
                Set rsCONTROL = New ADODB.Recordset
                strCONTROL = "SELECT sum(VATAmount) as Totals FROM ODASMRequisitionItems WHERE RequisitionNo = '" & .txtRequisitionNo.Text & "' "
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                            
                If rsCONTROL.BOF Or rsCONTROL.EOF Then Exit Sub

                rsSAVE!VATAmount = CDbl(rsCONTROL!TOTALS)
                            
                rsSAVE.Update
                rsSAVE.Requery
                        
        End With

Exit Sub

err:
    ErrorMessage
End Sub


Public Sub saveREQUISITION()
On Error GoTo err
        With frmODASMOpenJobCard
'                Set rsFindRecord = New ADODB.Recordset
'                rsFindRecord.Open "SELECT * FROM ;"
                Set rsSAVE = New ADODB.Recordset
'                strSQL = "SELECT * FROM ODASMRequisition WHERE AccountNo ='" & .txtAccountNo.Text & "' and JobCardNo = '" & .txtJobCardNo.Text & "';"
                strSQL = "SELECT * FROM ODASMRequisition WHERE RequisitionNo = '" & .txtRequisitionNo.Text & "';"
'                If rsSAVE.EOF And rsSAVE.BOF Then
'                rsREQ.generateRequisitionNO
'                Set rsSAVE = Nothing
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
                    
                            If rsSAVE.EOF And rsSAVE.BOF Then
                                    rsSAVE.AddNew
                                    rsSAVE!RequisitionNo = .txtRequisitionNo.Text
                                    rsSAVE!items = 0
                                    rsSAVE!Prepared = "Y"
                                    rsSAVE!Preparedby = CurrentUserName
                                    rsSAVE!dateprepared = Date
                                    rsSAVE!Approved = "N"
                                    rsSAVE!Authorized = "N"
                                    rsSAVE!Issued = "N"
                                    rsSAVE!Used = "N"
                                    rsSAVE!Invoiced = "N"
                                    rsSAVE!LPOIssued = "N"
                                    rsSAVE!Accountno = .txtAccountNo.Text

                             End If
                             
                            rsSAVE!items = CDbl(rsSAVE!items) + 1
                            rsSAVE!DepartmentCode = .txtDepartmentCode.Text
                            rsSAVE!JobCardNo = .txtJobCardNo.Text
                            rsSAVE!RequisitionDate = .txtRequisitionDate.Text
                            rsSAVE!Quantity = CDbl(.txtItemQuantity.Text)
                            rsSAVE!Status = "REQ-PLACED"
                            rsSAVE!PriceExclusive = CDbl(.txtTotalUnitPriceExcl)
                            rsSAVE!PriceInclusive = CDbl(.txtTotalUnitPriceIncl.Text)
                            rsSAVE!VATAmount = CDbl(.txtVATAmount.Text)
                            .txtItemNo.Text = Trim(.txtRequisitionNo.Text) + "-" + Trim(rsSAVE!items)
                            rsSAVE.Update
                            rsSAVE.Requery
                        

        End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub saveREQUISITIONITEMS()
On Error GoTo err
        With frmODASMOpenJobCard

                Set rsSAVE = New ADODB.Recordset
                'strSQL = "SELECT * FROM ODASMRequisitionItems WHERE RequisitionNo = '" & .txtRequisitionNo.Text & "' and ProductCode = '" & .txtProductCode.Text & "'"

                strSQL = "SELECT * FROM ODASMRequisitionItems WHERE ItemNo = '" & .txtItemNo.Text & "' "
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
                    
                            If rsSAVE.EOF And rsSAVE.BOF Then
                                    rsSAVE.AddNew
                                    rsSAVE!RequisitionNo = .txtRequisitionNo.Text
                                    rsSAVE!ItemNo = .txtItemNo.Text
                                    rsSAVE!Prepared = "Y"
                                    rsSAVE!Preparedby = CurrentUserName
                                    rsSAVE!dateprepared = Date
                                    rsSAVE!Approved = "N"
                                    rsSAVE!Authorized = "N"
                                    rsSAVE!Issued = "N"
                                    rsSAVE!Closed = "N"
                                    rsSAVE!returned = "N"
                                    rsSAVE!QuantityReturned = 0
                                    rsSAVE!QuantityIssued = 0
                                    rsSAVE!Used = "N"
                             End If
                            
                            rsSAVE!ProductCode = .txtProductCode.Text
                            rsSAVE!DepartmentCode = .txtDepartmentCode.Text
                            rsSAVE!DepartmentCode = .txtDepartmentCode.Text
                            rsSAVE!JobCardNo = .txtJobCardNo.Text
                            rsSAVE!CostCenter = .txtCostCenter.Text
                            rsSAVE!ItemSize = .txtItemSize.Text
                            rsSAVE!Quantity = CDbl(.txtItemQuantity.Text)
                            rsSAVE!Balance = CDbl(rsSAVE!Quantity)
                            rsSAVE!UnitPrice = CDbl(.txtUnitPrice.Text)
                            rsSAVE!Status = "REQ-PLACED"
                            rsSAVE!PriceExclusive = CDbl(.txtTotalUnitPriceExcl)
                            rsSAVE!PriceInclusive = CDbl(.txtTotalUnitPriceIncl.Text)
                            rsSAVE!VATRate = CDbl(.txtVATRate.Text)
                            rsSAVE!VATAmount = CDbl(.txtVATAmount.Text)
                            rsSAVE!ExchangeRate = CDbl(.txtExchangeRate.Text)
                            rsSAVE!CurrencyCode = .cboCurrencyCode.Text
                            
                            rsSAVE.Update
                            rsSAVE.Requery
                        

        End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub updateTripExpences()
With frmODASMOpenJobCard
    Set rsSAVE = New ADODB.Recordset
    rsSAVE.Open "SELECT * FROM  ODASMJobCard WHERE JobCardNo = '" & .txtJobCardNo.Text & "' and DepartmentCode = '" & .txtDepartmentCode.Text & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsSAVE!RequireTrip = "Y" Then
            Set rsFindRecord = New ADODB.Recordset
            rsFindRecord.Open "SELECT * FROM ODASMTripExpenses WHERE JobCardNo = '" & .txtJobCardNo.Text & "' and DepartmentCode = '" & .txtDepartmentCode.Text & "' and AccountingCode = '" & .txtProductCode.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
            If rsFindRecord.EOF And rsFindRecord.BOF Then
                rsFindRecord.AddNew
                rsFindRecord!JobCardNo = .txtJobCardNo.Text
                rsFindRecord!DepartmentCode = .txtDepartmentCode.Text
                rsFindRecord!AccountingCode = .txtProductCode.Text
            End If
                rsFindRecord!UnitCost = .txtUnitPrice.Text
                rsFindRecord!Quantity = .UpDownQuantity.Value
                rsFindRecord!TotalCost = .txtTotalPriceExcl.Text
                rsFindRecord!TotalCostInc = .txtTotalCost.Text
                rsFindRecord!VATAmount = .txtVATAmount.Text
           
                rsFindRecord.Update
                rsFindRecord.Requery
        End If
End With
Set rsFindRecord = Nothing
End Sub

Public Sub saveJOBCARD()
On Error GoTo err

        With frmODASMOpenJobCard

                Set rsSAVE = New ADODB.Recordset
                strSQL = "SELECT * FROM ODASMJobCard WHERE JobCardNo = '" & .txtJobCardNo.Text & "' and DepartmentCode = '" & .txtDepartmentCode & "'"
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        
                If rsSAVE.EOF And rsSAVE.BOF Then
                        rsSAVE.AddNew
                        rsSAVE!JobCardNo = .txtJobCardNo.Text
                        rsSAVE!DepartmentCode = .txtDepartmentCode.Text
                        rsSAVE!Preparedby = CurrentUserName
                        rsSAVE!dateprepared = Date
                        rsSAVE!TotalCost = 0
                        rsSAVE!TripCreated = "N"
                        rsSAVE!CurrentPeriod = CurrentPeriod
                 End If
                If bRequisitionAPPROVAL = True Then
                        rsSAVE!Status = "REQ-PREPARATION"
                ElseIf bRequisitionAPPROVAL = True Then
                        rsSAVE!Status = "REQ-APPROVAL"
                ElseIf bRequisitionAUTHORIZATION = True Then
                        rsSAVE!Status = "REQ-AUTHORIZATION"
                End If
                
                rsSAVE!Remarks = .txtRemarks.Text
                rsSAVE!PriceExclusive = CDbl(.txtTotalPriceExcl.Text)
                rsSAVE!VATAmount = CDbl(.txtTotalVATAmount.Text)
                rsSAVE!TotalCost = CDbl(.txtTotalCost.Text)
                rsSAVE!DateOfCommencement = .txtDateOfCommencement.Text
                If .optYes.Value = True Then
                    rsSAVE!RequireTrip = "Y"
                    Else
                    rsSAVE!RequireTrip = "N"
                End If
                bsaveRECORD = False
                rsSAVE.Update
                rsSAVE.Requery
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub validateRECORD()
On Error GoTo err
        With frmODASMOpenJobCard
            
            If .txtJobCardNo.Text <= "" Then
                    MsgBox "The job Card Number Cannot be Blank"
                    .txtJobCardNo.SetFocus
            
            ElseIf .txtAccountNo.Text = Empty Then
                    MsgBox "The Account Number is Required for the Supplier"
                    .txtAccountNo.SetFocus
                    
            ElseIf .cboCurrencyCode.Text = Empty Then
                    MsgBox "The Currency Code cannot be Left Blank"
                    .cboCurrencyCode.SetFocus
            ElseIf .optNo.Value = False And .optYes.Value = False Then
                    MsgBox "Please specify if you require trip or not!"
                    .optYes.SetFocus
            ElseIf .txtDepartmentCode.Text = Empty Then
                    MsgBox "The Department Code cannot be Left Blank"
                    .txtDepartmentCode.SetFocus
            
            ElseIf CDbl(.txtExchangeRate.Text) = Empty Then
                    MsgBox "The Exchange Rate MUST > 0"
                    .txtExchangeRate.SetFocus

            ElseIf .txtProductCode.Text = Empty Then
                    MsgBox "The Item Code Cannot be Left Blank"
                    .txtProductCode.SetFocus
            
            ElseIf .txtItemSize.Text <= Empty Then
                    MsgBox "The Item Size Cannot be Left Blank"
                    .txtItemSize.SetFocus
                    
            ElseIf CDbl(.txtItemQuantity.Text) = 0 Then
                    MsgBox "The Quantity MUST be Greater Than Zero ........"
                    .txtItemQuantity.SetFocus
            
            ElseIf CDbl(.txtUnitPrice.Text) = 0 Then
                    MsgBox "The Cost of the Item is Required ......."
                    .txtUnitPrice.SetFocus
            
            ElseIf CDbl(.txtTotalUnitPriceIncl.Text) = 0 Then
                    MsgBox "The total item cost MUST be Greater Than Zero"
                    txtTotalUnitPriceIncl.SetFocus
            
            ElseIf CDbl(.txtTotalUnitPriceExcl.Text) = 0 Then
                    MsgBox "The total Price Exclusive of VAT MUST be Greater Than Zero"
                    .txtTotalUnitPriceExcl.SetFocus

            Else
                    bsaveRECORD = True
            End If
        
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub generateRequisitionNO()
On Error GoTo err
        Set rsCONTROL = New Recordset
        
        strSQL = "Select * from ODASPLAstNumbers Where AutoRequisitionNo = 'Y';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .EOF Or .BOF = True Then Exit Sub
                
                frmODASMOpenJobCard.txtRequisitionNo.Text = !RequisitionNo
                
                Select Case Len(frmODASMOpenJobCard.txtRequisitionNo.Text)
                        Case 1: frmODASMOpenJobCard.txtRequisitionNo.Text = Trim(!RequisitionPrefix) + "0000" + Trim(frmODASMOpenJobCard.txtRequisitionNo.Text)
                        Case 2: frmODASMOpenJobCard.txtRequisitionNo.Text = Trim(!RequisitionPrefix) + "000" + Trim(frmODASMOpenJobCard.txtRequisitionNo.Text)
                        Case 3: frmODASMOpenJobCard.txtRequisitionNo.Text = Trim(!RequisitionPrefix) + "00" + Trim(frmODASMOpenJobCard.txtRequisitionNo.Text)
                        Case 4: frmODASMOpenJobCard.txtRequisitionNo.Text = Trim(!RequisitionPrefix) + "0" + Trim(frmODASMOpenJobCard.txtRequisitionNo.Text)
                        Case 5: frmODASMOpenJobCard.txtRequisitionNo.Text = Trim(!RequisitionPrefix) + Trim(frmODASMOpenJobCard.txtRequisitionNo.Text)
                End Select
                                
                !RequisitionNo = Str(Val(!RequisitionNo) + 1)
                .Update
                .Requery
        
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub enableRECORD()
On Error GoTo err
    With frmODASMOpenJobCard
        .txtExchangeRate.Locked = False
        .txtUnitPrice.Locked = False
        .txtItemQuantity.Locked = False
        .txtItemSize.Locked = False
        .txtVATAmount.Locked = False
        .UpDownQuantity.Enabled = True
        .cboCurrencyCode.Locked = False
        .Frame7.Enabled = True
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub clearRECORD()
On Error GoTo err
    
    With frmODASMOpenJobCard
        .txtExchangeRate.Text = 1
        .txtUnitPrice.Text = 0
        .txtItemQuantity.Text = 1
        .UpDownQuantity.Value = 1
        .txtItemSize.Text = Empty
        .txtRequisitionDate.Text = Date
        .txtRequisitionNo.Text = Empty
        .txtProductCode.Text = Empty
        .txtTotalCost.Text = 0
        .txtTotalUnitPriceIncl.Text = 0
        .txtVATAmount.Text = 0
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub


Public Sub clearRECORDPartially()
On Error GoTo err
    
    With frmODASMOpenJobCard
        .txtExchangeRate.Text = 1
        .txtUnitPrice.Text = 0
        .txtItemQuantity.Text = 1
        .UpDownQuantity.Value = 1
        .txtItemSize.Text = Empty
        .txtProductCode.Text = Empty
        .txtTotalCost.Text = 0
        .txtTotalUnitPriceIncl.Text = 0
        .txtTotalVATAmount.Text = 0
        .txtVATAmount.Text = 0
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub


Public Sub updateRECORD()
        If bsaveRECORD = True Then
                saveREQUISITION
                saveREQUISITIONITEMS
                updateREQUISITIONTOTALS
                saveJOBCARD
                updateTripExpences
                disableALLRECORD
                calculateTOTALS
        End If
        
        'showALLREQUISITIONSRAISED
End Sub

Public Sub calculateTOTALS()
On Error GoTo err
        With frmODASMOpenJobCard
                
                '/ Calculate To Items Requisitioned for This Job Card
                
                Set rsCONTROL = New ADODB.Recordset
                strSQL = "SELECT SUM(ODASMRequisitionItems.PriceExclusive) as TOTAL FROM ODASMRequisitionItems  WHERE ODASMRequisitionItems.JobCardNo = '" & .txtJobCardNo.Text & "' and ODASMRequisitionItems.DepartmentCode = '" & .txtDepartmentCode.Text & "' and ODASMRequisitionItems.Prepared = 'Y' and (ODASMRequisitionItems.Approved = 'N' or ODASMRequisitionItems.Approved is Null ) "
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
                
                If IsNull(rsCONTROL!total) = True Then
                            .txtRequisitionPrepared.Text = 0
                Else: .txtRequisitionPrepared.Text = FormatNumber(rsCONTROL!total)
                End If
                
                strSQL = Empty
                
                '/ * CALCULATE iTEMS rEQUISITION AND APPROVED
                
                Set rsCONTROL = New ADODB.Recordset
                strSQL = "SELECT SUM(PriceExclusive) as TOTAL FROM ODASMRequisitionItems WHERE JobCardNo = '" & .txtJobCardNo.Text & "' and DepartmentCode = '" & .txtDepartmentCode.Text & "' and Approved = 'Y' and (Authorized = 'N' or Authorized is Null ) "
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
                
                If IsNull(rsCONTROL!total) = True Then
                            .txtRequisitionApproved.Text = 0
                Else: .txtRequisitionApproved.Text = FormatNumber(rsCONTROL!total)
                End If
                
                strSQL = Empty

                '/ * CALCULATE iTEMS rEQUISITIONED AND AUTHORIZED
                
                Set rsCONTROL = New ADODB.Recordset
                strSQL = "SELECT SUM(PriceExclusive) as TOTAL FROM ODASMRequisitionItems WHERE JobCardNo = '" & .txtJobCardNo.Text & "' and DepartmentCode = '" & .txtDepartmentCode.Text & "' and Authorized = 'Y' "
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
                
                If IsNull(rsCONTROL!total) = True Then
                            .txtTotalCost.Text = 0
                Else: .txtTotalCost.Text = FormatNumber(rsCONTROL!total)
                End If
                
                strSQL = Empty

                '/ * CALCULATE VAT FOR ALL iTEMS rEQUISITIONED AND AUTHORIZED
                
                Set rsCONTROL = New ADODB.Recordset
                strSQL = "SELECT SUM(VATAmount) as TOTAL FROM ODASMRequisitionItems WHERE JobCardNo = '" & .txtJobCardNo.Text & "' and DepartmentCode = '" & .txtDepartmentCode.Text & "' and Authorized = 'Y' "
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
                
                If IsNull(rsCONTROL!total) = True Then
                            .txtTotalVATAmount.Text = 0
                Else: .txtTotalVATAmount.Text = FormatNumber(rsCONTROL!total)
                End If
                strSQL = Empty
                
                '/ * CALCULATE Price Exclusive of VAT FOR ALL iTEMS rEQUISITIONED AND AUTHORIZED
                
                Set rsCONTROL = New ADODB.Recordset
                strSQL = "SELECT SUM(PriceExclusive) as TOTAL FROM ODASMRequisitionItems WHERE JobCardNo = '" & .txtJobCardNo.Text & "' and DepartmentCode = '" & .txtDepartmentCode.Text & "' and Authorized = 'Y' "
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
                
                If IsNull(rsCONTROL!total) = True Then
                            .txtTotalPriceExcl.Text = 0
                Else: .txtTotalUnitPriceExcl.Text = FormatNumber(rsCONTROL!total)
                End If
                
                strSQL = Empty

               
               
               '/ Calculate Qty of Items Requisitioned for This Job Card
                
                Set rsCONTROL = New ADODB.Recordset
                strSQL = "SELECT SUM(Quantity) as TOTAL FROM ODASMRequisitionItems WHERE JobCardNo = '" & .txtJobCardNo.Text & "' and DepartmentCode = '" & .txtDepartmentCode.Text & "' and Prepared = 'Y' and (Approved = 'N' or Approved is Null ) "
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
                
                If IsNull(rsCONTROL!total) = True Then
                            .txtQuantityPrepared.Text = 0
                Else: .txtQuantityPrepared.Text = FormatNumber(rsCONTROL!total)
                End If
                
                strSQL = Empty

                
                
                '/ * CALCULATE Qty iTEMS rEQUISITION AND APPROVED
                
                Set rsCONTROL = New ADODB.Recordset
                strSQL = "SELECT SUM(Quantity) as TOTAL FROM ODASMRequisitionItems WHERE JobCardNo = '" & .txtJobCardNo.Text & "' and DepartmentCode = '" & .txtDepartmentCode.Text & "' and Approved = 'Y' and (Authorized = 'N' or Authorized is Null ) "
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
                
                If IsNull(rsCONTROL!total) = True Then
                            .txtQuantityApproved.Text = 0
                Else: .txtQuantityApproved.Text = FormatNumber(rsCONTROL!total)
                End If
                
                
        End With
Exit Sub

err:
    ErrorMessage
End Sub
