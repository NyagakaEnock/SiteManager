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
                strSQL = "SELECT * FROM ODASPInventory WHERE ItemCode = '" & .txtItemCode.Text & "'; "
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
                If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
                
                .txtItemCode.Text = rsCONTROL!ItemCode
                 .txtItemCost.Text = rsCONTROL!ItemCost
                 .txtVATAmount.Text = rsCONTROL!VATAmount
        
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

            .txtItemCost.Text = 0
            .txtTotalVATAmount.Text = 0
            .txtVATAmount.Text = 0
            .txtItemQuantity.Text = 1
            .UpDownQuantity.Value = 1
            .txtSupervisedBy = rsCONTROL!OpenedBy
            .txtDoneBy.Text = CurrentUserName
            .txtDateOfCommencement.Text = Date
            .txtRequisitionDate.Text = Date
            .txtExchangeRate.Text = 0
    
            
            
        End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub obtainBASECURRENCY()
On Error GoTo err
            
            Set rsCONTROL = New ADODB.Recordset
            strSQL = "SELECT * FROM ODASPCurrency WHERE BaseCurrency = 'Y'; "
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
            
            Screen.ActiveForm.txtExchangeRate.Text = 1
            Screen.ActiveForm.cboCurrencyCode.Text = rsCONTROL!CurrencyCode
            Screen.ActiveForm.txtCurrencySymbol.Text = rsCONTROL!CurrencySymbol

rsCONTROL.Close
strSQL = Empty
Exit Sub

err:
    ErrorMessage
End Sub


Public Sub SaveRECORD()
On Error GoTo err
        With frmODASMOpenJobCard

                Set rsSAVE = New ADODB.Recordset
                strSQL = "SELECT * FROM ODASMRequisition WHERE JobCardNo = '" & .txtJobCardNo.Text & "' and RequisitionNo = '" & .txtRequisitionNo.Text & "' and DepartmentCode = '" & .txtDepartmentCode.Text & "' and ItemCode = '" & .txtItemCode.Text & "'"
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
                    
                            If rsSAVE.EOF And rsSAVE.BOF Then
                                    rsSAVE.AddNew
                                    rsSAVE!JobCardNo = .txtJobCardNo.Text
                                    rsSAVE!RequisitionNo = .txtRequisitionNo.Text
                                    rsSAVE!ItemCode = .txtItemCode.Text
                                    rsSAVE!DepartmentCode = .txtDepartmentCode.Text
                                    rsSAVE!Prepared = "Y"
                                    rsSAVE!PreparedBy = CurrentUserName
                                    rsSAVE!DatePrepared = Date
                                    rsSAVE!Approved = "N"
                                    rsSAVE!Authorized = "N"
                             End If
                            
                            rsSAVE!CostingCode = .cboCostingCode.Text
                            rsSAVE!RequisitionDate = .txtRequisitionDate.Text
                            rsSAVE!ItemSize = .txtItemSize.Text
                            rsSAVE!itemQuantity = CDbl(.txtItemQuantity.Text)
                            rsSAVE!ItemCost = CCur(.txtItemCost.Text)
                            rsSAVE!Status = "REQ-PLACED"
                            rsSAVE!TotalItemCost = CCur(.txtTotalItemCost.Text)
                            rsSAVE!VATAmount = CCur(.txtTotalVATAmount.Text)
                            rsSAVE!ExchangeRate = CCur(.txtExchangeRate.Text)
                            rsSAVE!CurrencyCode = .cboCurrencyCode.Text
                            
                            rsSAVE.Update
                            rsSAVE.Requery
                        

        End With

Exit Sub

err:
    ErrorMessage
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
                        rsSAVE!PreparedBy = CurrentUserName
                        rsSAVE!DatePrepared = Date
                        rsSAVE!TotalCost = 0
                        rsSAVE!CurrentPeriod = CurrentPeriod
                 End If
                
                rsSAVE!Remarks = .txtRemarks.Text
                rsSAVE!DateOfCommencement = .txtDateOfCommencement.Text
                bSaveRECORD = False
                rsSAVE.Update
                rsSAVE.Requery
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub ValidateRECORD()
On Error GoTo err
        With frmODASMOpenJobCard
            
            If .txtJobCardNo.Text <= "" Then
                    MsgBox "The job Card Number Cannot be Blank"
                    .txtJobCardNo.SetFocus
            
            ElseIf .cboCurrencyCode.Text = Empty Then
                    MsgBox "The Currency Code cannot be Left Blank"
                    .cboCurrencyCode.SetFocus
            
            ElseIf .txtDepartmentCode.Text = Empty Then
                    MsgBox "The Department Code cannot be Left Blank"
                    .txtDepartmentCode.SetFocus
            
            ElseIf CDbl(.txtExchangeRate.Text) = Empty Then
                    MsgBox "The Exchange Rate MUST > 0"
                    .txtExchangeRate.SetFocus

            ElseIf .txtItemCode.Text = Empty Then
                    MsgBox "The Item Code Cannot be Left Blank"
                    .txtItemCode.SetFocus
            
            ElseIf .txtItemSize.Text <= Empty Then
                    MsgBox "The Item Size Cannot be Left Blank"
                    .txtItemSize.SetFocus
                    
            ElseIf CDbl(.txtItemQuantity.Text) = 0 Then
                    MsgBox "The Quantity MUST be Greater Than Zero ........"
                    .txtItemQuantity.SetFocus
            
            ElseIf CCur(.txtItemCost.Text) = 0 Then
                    MsgBox "The Cost of the Item is Required ......."
                    .txtItemCost.SetFocus
            
            ElseIf CCur(.txtTotalItemCost.Text) = 0 Then
                    MsgBox "The total item cost MUST be Greater Than Zero"
                    .txtTotalItemCost.SetFocus
            
            Else
                    bSaveRECORD = True
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
        .txtItemCost.Locked = False
        .txtItemQuantity.Locked = False
        .txtItemSize.Locked = False
        .txtVATAmount.Locked = False
        .UpDownQuantity.Enabled = True
        .cboCostingCode.Locked = False
        .cboCurrencyCode.Locked = False
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub clearRECORD()
On Error GoTo err
    
    With frmODASMOpenJobCard
        .txtExchangeRate.Text = 1
        .txtItemCost.Text = 0
        .txtItemQuantity.Text = 1
        .UpDownQuantity.Value = 1
        .txtItemSize.Text = Empty
        .txtRequisitionDate.Text = Date
        .txtRequisitionNo.Text = Empty
        .txtItemCode.Text = Empty
        .txtTotalCost.Text = 0
        .txtTotalItemCost.Text = 0
        .txtTotalVATAmount.Text = 0
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
        .txtItemCost.Text = 0
        .txtItemQuantity.Text = 1
        .UpDownQuantity.Value = 1
        .txtItemSize.Text = Empty
        .txtItemCode.Text = Empty
        .txtTotalCost.Text = 0
        .txtTotalItemCost.Text = 0
        .txtTotalVATAmount.Text = 0
        .txtVATAmount.Text = 0
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub


Public Sub updateRECORD()
        ValidateRECORD
        If bSaveRECORD = True Then
                SaveRECORD
                saveJOBCARD
                disableALLRECORD
                calculateTOTALS
        End If
        showALLREQUISITIONSRAISED
End Sub

Public Sub calculateTOTALS()
On Error GoTo err
        With frmODASMOpenJobCard
                
                '/ Calculate To Items Requisitioned for This Job Card
                
                Set rsCONTROL = New ADODB.Recordset
                strSQL = "SELECT SUM(TotalItemCost) as TOTAL FROM ODASMRequisition WHERE JobCardNo = '" & .txtJobCardNo.Text & "' and DepartmentCode = '" & .txtDepartmentCode.Text & "' and Prepared = 'Y' and (Approved = 'N' or Approved is Null ) "
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
                
                If IsNull(rsCONTROL!Total) = True Then
                            .txtRequisitionPrepared.Text = 0
                Else: .txtRequisitionPrepared.Text = FormatCurrency(rsCONTROL!Total)
                End If
                
                strSQL = Empty
                
                '/ * CALCULATE iTEMS rEQUISITION AND APPROVED
                
                Set rsCONTROL = New ADODB.Recordset
                strSQL = "SELECT SUM(TotalItemCost) as TOTAL FROM ODASMRequisition WHERE JobCardNo = '" & .txtJobCardNo.Text & "' and DepartmentCode = '" & .txtDepartmentCode.Text & "' and Approved = 'Y' and (Authorized = 'N' or Authorized is Null ) "
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
                
                If IsNull(rsCONTROL!Total) = True Then
                            .txtRequisitionApproved.Text = 0
                Else: .txtRequisitionApproved.Text = FormatCurrency(rsCONTROL!Total)
                End If
                
                strSQL = Empty

                '/ * CALCULATE iTEMS rEQUISITIONED AND AUTHORIZED
                
                Set rsCONTROL = New ADODB.Recordset
                strSQL = "SELECT SUM(TotalItemCost) as TOTAL FROM ODASMRequisition WHERE JobCardNo = '" & .txtJobCardNo.Text & "' and DepartmentCode = '" & .txtDepartmentCode.Text & "' and Authorized = 'Y' "
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
                
                If IsNull(rsCONTROL!Total) = True Then
                            .txtTotalCost.Text = 0
                Else: .txtTotalCost.Text = FormatCurrency(rsCONTROL!Total)
                End If

                
                strSQL = Empty

               '/ Calculate Qty of Items Requisitioned for This Job Card
                
                Set rsCONTROL = New ADODB.Recordset
                strSQL = "SELECT SUM(ItemQuantity) as TOTAL FROM ODASMRequisition WHERE JobCardNo = '" & .txtJobCardNo.Text & "' and DepartmentCode = '" & .txtDepartmentCode.Text & "' and Prepared = 'Y' and (Approved = 'N' or Approved is Null ) "
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
                
                If IsNull(rsCONTROL!Total) = True Then
                            .txtQuantityPrepared.Text = 0
                Else: .txtQuantityPrepared.Text = FormatNumber(rsCONTROL!Total)
                End If
                
                strSQL = Empty

                '/ * CALCULATE Qty iTEMS rEQUISITION AND APPROVED
                
                Set rsCONTROL = New ADODB.Recordset
                strSQL = "SELECT SUM(ItemQuantity) as TOTAL FROM ODASMRequisition WHERE JobCardNo = '" & .txtJobCardNo.Text & "' and DepartmentCode = '" & .txtDepartmentCode.Text & "' and Approved = 'Y' and (Authorized = 'N' or Authorized is Null ) "
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
                
                If IsNull(rsCONTROL!Total) = True Then
                            .txtQuantityApproved.Text = 0
                Else: .txtQuantityApproved.Text = FormatNumber(rsCONTROL!Total)
                End If
                

                
                
                
        End With
Exit Sub

err:
    ErrorMessage
End Sub
