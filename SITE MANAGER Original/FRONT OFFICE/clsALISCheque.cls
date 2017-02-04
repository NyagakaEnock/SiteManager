VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsALISCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim MYCOUNT As Double

Private Sub clearREQUISITION()
On Error GoTo err
        With Screen.ActiveForm
                .txtBankNo.Text = ""
                .txtChequeNo.Text = ""
                .txtChequeDate.Text = ""
                .txtDocumentNo.Text = ""
                .txtPayeeDetails.Text = ""
                .txtVoucherDate.Text = ""
                .txtAmountPaid.Text = ""
                .txtChequeAmount.Text = ""
                .txtPaymentFlag.Text = ""
                .txtAmountDue.Text = ""
                .txtChequeAmount.Text = ""
                .txtNoOfEntries.Text = ""
        End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub pclearREQUISITION()
On Error GoTo err
        With frmODASMCheck
                If bmakePAYMENT = True Then
                        .txtRemark.Text = Empty
                End If
        End With

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub enableREQUISITION()
    bVAL = False
    testREQUISITION
End Sub


Private Sub disableREQUISITION()
    bVAL = True
    testREQUISITION
End Sub
Public Sub loadVOUCHER()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
                
        strSQL = "Select * from ODASMVoucher V, ODASPAccount A  where V.VoucherNo = '" & Screen.ActiveForm.txtVoucherNo.Text & "' and A.AccountNo = V.AccountNo ;"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsCONTROL
                
                If .BOF Or .EOF Then Exit Sub
                
                If bmakePAYMENT = True Then
                        Screen.ActiveForm.txtChequeNo.Text = !ChequeNo & ""
                        Screen.ActiveForm.txtPayeeDetails.Text = !CompanyName & ""
                        Screen.ActiveForm.txtVoucherDate.Text = !VoucherDate & ""
                        Screen.ActiveForm.txtAccountNo.Text = !Accountno
                        Screen.ActiveForm.txtPaymentFlag.Text = "Y"
                        Screen.ActiveForm.txtAmountDue.Text = FormatNumber(!Amount)
                        Screen.ActiveForm.txtChequeDate.Text = Date
                        Screen.ActiveForm.txtCostCenter.Text = !CostCenter & ""
                        Screen.ActiveForm.DTPickerChequeDate.Value = Date
                        Screen.ActiveForm.txtStatus.Text = "CHK-PREPARED"
                End If
        End With


Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadCHEQUE()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
                
        strSQL = "Select * from ODASMVoucher V, ODASPAccount A, ALISMCheque C  where C.ChequeNo = V.ChequeNo and V.VoucherNo = '" & Screen.ActiveForm.txtVoucherNo.Text & "' and A.AccountNo = V.AccountNo ;"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsCONTROL
                
                If .BOF Or .EOF Then Exit Sub
                
                Screen.ActiveForm.txtChequeNo.Text = !ChequeNo & ""
                Screen.ActiveForm.txtDocumentNo.Text = !JobBriefNo & ""
                Screen.ActiveForm.txtPayeeDetails.Text = !CompanyName & ""
                Screen.ActiveForm.txtVoucherDate.Text = !VoucherDate & ""
                Screen.ActiveForm.txtAccountNo.Text = !Accountno
                Screen.ActiveForm.txtPaymentFlag.Text = !ChequePrepared & ""
                Screen.ActiveForm.txtAmountDue.Text = FormatNumber(!Amount)
                Screen.ActiveForm.txtChequeDate.Text = !ChequeDate
                Screen.ActiveForm.DTPickerChequeDate.Value = Date
                Screen.ActiveForm.txtStatus.Text = !Status
                Screen.ActiveForm.txtChequeAmount.Text = !ChequeAmount
                Screen.ActiveForm.txtBankNo.Text = !BankNo & ""
        End With


Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadSCHEDULEDRECORD()
On Error GoTo err
    
    If bscheduledCHEQUES = True Then
            strSQL = "Select * from ALISMBankAccountCfg, ALISPBankAccount Where ALISMBankAccountCfg.Reference = '" & frmALISMCheckSchedule.txtReference.Text & "' and ALISMBankAccountCfg.BankNo = ALISPBankAccount.BankNo "
    ElseIf bschedulePAYMENT = True Then
            strSQL = "Select * from ALISPBankAccount Where BankNo = '" & frmALISMCheckSchedule.txtAccountNo.Text & "'"
    End If
    
    Set rsCONTROL = New ADODB.Recordset
    
    With rsCONTROL
            .Open strSQL, cnCOMMON, adOpenDynamic, adLockOptimistic
            
            If .BOF Or .EOF Then Exit Sub
            
            If bscheduledCHEQUES = True Then
                    
                    frmALISMCheckSchedule.txtDetails.Text = !Details
                    frmALISMCheckSchedule.txtAccountNo.Text = !Accountno
                    frmALISMCheckSchedule.txtBankNo.Text = !BankNo
                    frmALISMCheckSchedule.txtDateBanked.Text = !DateBanked
                    frmALISMCheckSchedule.txtNoOfEntries.Text = !NoOfEntries
                    frmALISMCheckSchedule.txtTotalAmount.Text = !TotalAmount
                    frmALISMCheckSchedule.cboTemplateCode.Text = !templateCode
                    
            ElseIf bschedulePAYMENT = True Then
                    frmALISMCheckSchedule.txtDetails.Text = !Details
                    frmALISMCheckSchedule.txtAccountNo.Text = !Accountno
                    frmALISMCheckSchedule.txtBankNo.Text = !BankNo
                    frmALISMCheckSchedule.txtDateBanked.Text = Date
                    frmALISMCheckSchedule.DTPickerDateBanked.Value = Date
                    frmALISMCheckSchedule.txtNoOfEntries.Text = 0
                    frmALISMCheckSchedule.txtTotalAmount.Text = 0

            End If
    End With

strSQL = ""
rsCONTROL.Close

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadAPPROVEDCHECKS()
On Error GoTo err
        With frmODASMCheckIssuance
        
                Dim rsLOAD As ADODB.Recordset
                Set rsLOAD = New ADODB.Recordset
                
                strSQL = "Select *,V.VoucherNo, V.CostCenter  from ALISMCheque C, ALISPBank B, ODASMVoucher V where V.ChequeNo = C.ChequeNo and C.BankNo = B.BankNo and C.chequeNo = '" & frmODASMCheckIssuance.txtChequeNo & "'"
                rsLOAD.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                If rsLOAD.BOF Or rsLOAD.EOF Then Exit Sub
                
                .txtChequeNo.Text = rsLOAD!ChequeNo & ""
                .txtPayeeDetails.Text = rsLOAD!PayeeDetails & ""
                
                If IsNull(rsLOAD!ChequeAmount) = True Then
                        .txtChequeAmount.Text = 0
                Else: .txtChequeAmount.Text = FormatNumber(rsLOAD!ChequeAmount)
                End If
                
                .txtDateCollected.Text = Date
                .txtIssuedBy.Text = CurrentUserName
                .DTPickerDateCollected.Value = Date
                .txtStatus.Text = "CHK-ISSUED"
                .txtCostCenter.Text = rsLOAD!CostCenter & ""
                .txtCurrentPeriod.Text = CurrentPeriod
                .txtIssuedBy.Text = CurrentUserName
                .txtCollectedBy.Text = rsLOAD!PayeeDetails & ""
                .txtBankNo.Text = rsLOAD!BankNo & ""
                .txtVoucherNo.Text = rsLOAD!vOUCHERnO & ""
                .txtAccountNo.Text = rsLOAD!Accountno & ""
                .txtBankName.Text = rsLOAD!CompanyName & ""
        End With


Exit Sub

err:
    ErrorMessage
End Sub

Private Sub testREQUISITION()
On Error GoTo err
        If bmakePAYMENT = True Then
                With Screen.ActiveForm
                        .txtChequeNo.Locked = bVAL
                        .txtChequeDate.Locked = bVAL
                        .txtDocumentNo.Locked = bVAL
                        .txtPayeeDetails.Locked = bVAL
                        .txtVoucherDate.Locked = True
                        .txtAmountPaid.Locked = bVAL
                        .txtChequeAmount.Locked = True
                        .txtPaymentFlag.Locked = True
                        .txtAmountDue.Locked = True
                        .txtChequeAmount.Locked = True
                End With
        ElseIf breversePAYMENT = True Then
                        With Screen.ActiveForm
                        .txtChequeNo.Locked = bVAL
                        .txtChequeDate.Locked = bVAL
                        .txtDocumentNo.Locked = bVAL
                        .txtPayeeDetails.Locked = bVAL
                        .txtVoucherDate.Locked = True
                        .txtAmountPaid.Locked = bVAL
                        .txtChequeAmount.Locked = True
                        .txtPaymentFlag.Locked = True
                        .txtAmountDue.Locked = True
                        .txtChequeAmount.Locked = True
                End With
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub checkSTATUS()
On Error GoTo err
        With Screen.ActiveForm
                If CDbl(.txtAmountPaid.Text) > CDbl(.txtAmountDue.Text) Then
                        MsgBox "The Amount Paid cannot be more than The Amount Due", vbOKOnly
                        .txtAmountPaid.SetFocus
                ElseIf CDbl(.txtAmountDue) <= 0 Then
                        MsgBox "The Amount Due cannot be Zero", vbOKOnly
                        .txtAmountDue.SetFocus
                ElseIf CDbl(.txtAmountPaid) <= 0 Then
                        MsgBox "The Amount Paid cannot be Zero", vbOKOnly
                        .txtAmountPaid.SetFocus
                Else
                        If CDbl(.txtAmountPaid.Text) < CDbl(.txtAmountDue.Text) Then
                                .txtPaymentFlag.Text = "P"
                                .txtChequeAmount.Text = CDbl(.txtAmountPaid.Text)
                        ElseIf CDbl(.txtAmountPaid.Text) = CDbl(.txtAmountDue.Text) Then
                                 .txtPaymentFlag.Text = "Y"
                        End If
                End If

        End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub CostCenterLostFocus()
On Error GoTo err
        
        Dim rsCOSTCENTER As ADODB.Recordset
        Set rsCOSTCENTER = New ADODB.Recordset
        
        rsCOSTCENTER.Open "SELECT * FROM ODASMVoucher, ODASPCostCentre WHERE ODASMVoucher.VoucherNo = '" & Screen.ActiveForm.cboRequisition.Text & "' and ODASMVoucher.CostCenter = ODASPCostCentre.CostCenter ", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCOSTCENTER
                If .EOF And .BOF Then Exit Sub
                
                If rsCOSTCENTER!Materials = "Y" Then
                        strSQL = "select * from ODASMVoucher, ALISMClaimTotal WHERE ODASMVoucher.DocumentNo = ALISMClaimTotal.ClaimNo and ODASMVoucher.AccountNo = ALISMClaimTotal.ClaimNo "
                ElseIf rsCOSTCENTER!Machinery = "Y" Then
                ElseIf rsCOSTCENTER!Licenses = "Y" Then
                ElseIf rsCOSTCENTER!AdminCost = "Y" Then
                        'PaymentCodeLostFocusLoan
                ElseIf rsCOSTCENTER!ManPower = "Y" Then
                        'PaymentCodeLostFocusMedical
                ElseIf rsCOSTCENTER!Reinsurer = "Y" Then
                
                End If
        End With


rsCOSTCENTER.Close

Exit Sub

err:
        ErrorMessage
End Sub
 
Public Sub displayREQUISITION()
On Error GoTo err
        clearREQUISITION
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub ClearListview()
On Error GoTo err

 Dim j, i As Integer
       
                j = Screen.ActiveForm.ListView1.ListItems.Count
            
        For i = 1 To j
                Screen.ActiveForm.ListView1.ListItems(i).Checked = False
        Next i
Exit Sub
err:
ErrorMessage
    
End Sub

Private Sub GetCheckedBoxes()
On Error GoTo err

     Dim j, i As Integer, strRequirementcode As String

     j = Screen.ActiveForm.ListView1.ListItems.Count
     
     For i = 1 To j
             If Screen.ActiveForm.ListView1.ListItems(i).Checked = True Then
                 strRequirementcode = Screen.ActiveForm.ListView1.ListItems(i).Text
                 
                 UpdateALLRECORDS
             End If
             strRequirementcode = ""
     Next i
            
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub addRECORD()
        bcheckREQUISITION = False
        bcheckREQUISITION = True
        bcheckREQENTRY = True
        
        If bcheckREQUISITION = True Then
                bcheckREQUISITION = False
                pclearREQUISITION
                calculateTOTALPAID
                GetVouchers
                
                If frmODASMCheck.txtChequeNo.Text <= Empty Then
                Else
                        countENTRIES
                End If
        End If
End Sub
Public Sub clearCHEQUE()
On Error GoTo err
    
    With frmODASMCheckIssuance
        .cboIdType.Text = ""
        .txtIdentityNo.Text = ""
    End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub contructDATA()
On Error GoTo err

        strREQ = "Select * from ALISmCheque "
        Set rsREQ = New ADODB.Recordset
        rsREQ.Open strREQ, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsREQ
            
            If .EOF And .BOF Then Exit Sub
            
            Do While Not .EOF
                
                    If IsNull(!Prepared) = True Then
                            !Prepared = "Y"
                    End If
                    
                    If IsNull(!Approved) = True Then
                            !Approved = "N"
                    End If
                    
                    If IsNull(!Authorized) = True Then
                            !Authorized = "N"
                    End If
                    
                    If IsNull(!Scheduled) = True Then
                            !Scheduled = "N"
                    End If

                    If IsNull(!NoOfEntries) = True Then
                            !NoOfEntries = 0
                    End If
                    
                    !PaymentFlag = "N"
                    .Update
                    .MoveNext
            
            Loop
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub enableISSUANCE()
On Error GoTo err
    With frmODASMCheckIssuance
            .txtCollectedBy.Locked = False
            .cboIdType.Locked = False
            .txtIdentityNo.Locked = False
            .txtComment.Locked = False
            .txtDateCollected.Locked = False
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub enableRECORD()
On Error GoTo err
    With frmODASMCheck
            .txtChequeNo.Locked = False
            .txtRemark.Locked = False
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub clearRECORD()
On Error GoTo err
    With frmODASMCheck
        .txtChequeEntryNo.Text = Empty
        .txtVoucherNo.Text = Empty
        .txtVoucherDate.Text = Empty
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub UpdateALLRECORDS()
           bsaveRECORD = False
            validateRECORD

        If bsaveRECORD = True Then
                countENTRIES
                SaveChequeENTRIES
                calculateTOTALPAID
                saveCHEQUE
                updateREQUISITION
                GetVouchers
                updateCostCenterPrepared
                disableALLRECORD
                showALLCHKENTRIES
                NewRecord = False
                beditRECORD = False
        End If
End Sub

Public Sub updateINSTALLMENTS()
On Error GoTo err
                
                Set rsCOSTCENTER = New ADODB.Recordset
                
                rsCOSTCENTER.Open "SELECT * FROM ODASPCostCentre WHERE CostCentre = '" & Screen.ActiveForm.txtCostCenter.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic

                If rsCOSTCENTER!Materials = "Y" Then
                ElseIf rsCOSTCENTER!Machinery = "Y" Then
                ElseIf rsCOSTCENTER!Rent = "Y" Then
                ElseIf rsCOSTCENTER!Rate = "Y" Then
                ElseIf rsCOSTCENTER!AdminCost = "Y" Then
                ElseIf rsCOSTCENTER!ManPower = "Y" Then
                
                End If

Exit Sub
err:
    ErrorMessage
End Sub
Public Sub updateCostCenter()
On Error GoTo err
        Set rsCOSTCENTER = New ADODB.Recordset
        rsCOSTCENTER.Open "SELECT * FROM ODASPCostCentre WHERE CostCentre = '" & Screen.ActiveForm.txtCostCenter.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsCOSTCENTER!Materials = "Y" Then
                saveINVOICEIssued
        ElseIf rsCOSTCENTER!Machinery = "Y" Then
                saveINVOICEIssued
        ElseIf rsCOSTCENTER!Rent = "Y" Then
                saveINSTALLMENTISSUED
        ElseIf rsCOSTCENTER!Rate = "Y" Then
                saveRATESCHEDULEIssued
        ElseIf rsCOSTCENTER!AdminCost = "Y" Then
        
        ElseIf rsCOSTCENTER!ManPower = "Y" Then
        
        End If
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub saveRATESCHEDULEPrepared()
On Error GoTo err

    With frmODASMCheck
    
                Set rsSAVE = New Recordset
                strSQL = "SELECT * from ODASMCouncilRateDue where ReferenceNo = '" & .txtDocumentNo.Text & " ' ; "
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
                
                rsSAVE!Status = "CHK-PREPARED"
                rsSAVE!StatusDate = Date
                rsSAVE.Update
                rsSAVE.Requery
        rsSAVE.Close
        strSQL = ""
    
   End With
Exit Sub

err:
    UpdateErrorMessage
End Sub


Private Sub saveRATESCHEDULEIssued()
On Error GoTo err

    With frmODASMCheckIssuance
    
        Set rsSAVE = New Recordset
        strSQL = "SELECT * from ODASMCouncilRateDue where ReferenceNo = '" & .txtDocumentNo.Text & " ' ; "
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
        
        rsSAVE!Status = "CHK-ISSUED"
        rsSAVE!StatusDate = Date
        rsSAVE.Update
        rsSAVE.Requery
        
        rsSAVE.Close
        strSQL = ""
    
   End With
Exit Sub

err:
    UpdateErrorMessage
End Sub


Public Sub updateCostCenterPrepared()
On Error GoTo err
        Set rsCOSTCENTER = New ADODB.Recordset
        rsCOSTCENTER.Open "SELECT * FROM ODASPCostCentre WHERE CostCentre = '" & Screen.ActiveForm.txtCostCenter.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsCOSTCENTER!Materials = "Y" Then
                saveINVOICEPrepared
        ElseIf rsCOSTCENTER!Machinery = "Y" Then
                saveINVOICEPrepared
        ElseIf rsCOSTCENTER!Rent = "Y" Then
                saveINSTALLMENT
        ElseIf rsCOSTCENTER!Rate = "Y" Then
                saveRATESCHEDULEPrepared
        ElseIf rsCOSTCENTER!AdminCost = "Y" Then
        
        ElseIf rsCOSTCENTER!ManPower = "Y" Then
        
        End If
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub saveINVOICEPrepared()
On Error GoTo err

    With frmODASMCheck
                
                Set rsCONTROL = New ADODB.Recordset
                strCONTROL = "SELECT * from ODASMInvoice where VoucherNo = '" & .txtVoucherNo.Text & " ' ; "
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
                
                Do While Not rsCONTROL.EOF
                    
                        Set rsSAVE = New Recordset
                        strSQL = "SELECT * from ODASMInvoice where InvoiceNo = '" & rsCONTROL!InvoiceNo & " ' and LPONo = '" & rsCONTROL!LPONo & "'; "
                        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                        If rsSAVE.EOF Or rsSAVE.BOF Then
                        Else
                            rsSAVE!Status = "CHK-PREPARED"
                            rsSAVE!StatusDate = Date
                            
                            rsSAVE.Update
                            rsSAVE.Close
                            strSQL = ""
                        End If
                        
                        rsCONTROL.MoveNext
                Loop
   End With
err:
    UpdateErrorMessage
End Sub
Private Sub saveINVOICEIssued()
On Error GoTo err

    With frmODASMCheckIssuance
                
                Set rsCONTROL = New ADODB.Recordset
                strCONTROL = "SELECT * from ODASMInvoice where VoucherNo = '" & .txtVoucherNo.Text & " ' ; "
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
                
                Do While Not rsCONTROL.EOF
                    
                        Set rsSAVE = New Recordset
                        strSQL = "SELECT * from ODASMInvoice where InvoiceNo = '" & rsCONTROL!InvoiceNo & " ' and LPONo = '" & rsCONTROL!LPONo & "'; "
                        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                        If rsSAVE.EOF Or rsSAVE.BOF Then
                        Else
                            rsSAVE!Status = "CHK-ISSUED"
                            rsSAVE!StatusDate = Date
                            
                            rsSAVE.Update
                            rsSAVE.Close
                            strSQL = ""
                        End If
                        
                        rsCONTROL.MoveNext
                Loop
   End With
err:
    UpdateErrorMessage
End Sub

Public Sub UpdateChecksIssued()
        bsaveRECORD = False
        validateCHECKSISSUED

        If bsaveRECORD = True Then
                If NewRecord = True Then
                        GenerateIssuanceNo
                End If
                saveRecord
                saveCHECKSISSUED
                updateCostCenter
                updateREQUISITIONIssued
                UpDatetheInvoice
                disableALLRECORD
                GetChequesISSUEDTHISPERIOD
                NewRecord = False
                beditRECORD = False

        End If
End Sub
Public Sub updateInvoiceTotals()
On Error GoTo err
    With Screen.ActiveForm
            Set rsFindRecord = New ADODB.Recordset
            rsFindRecord.Open "Select sum(PaymentDue) as PDue From ODASMInstallment Where  VoucherNo = '" & frmODASMCheckIssuance.txtVoucherNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsFindRecord.EOF Or rsFindRecord.BOF Then
                .txtInstallmentTotal = 0
            ElseIf IsNull(rsFindRecord!PDue) Then
                .txtInstallmentTotal = 0
            Else
                .txtInstallmentTotal = FormatNumber(rsFindRecord!PDue)
                
            End If
    End With
Exit Sub
err:
ErrorMessage
End Sub
Private Sub saveINSTALLMENT()
On Error GoTo err

    With frmODASMCheck
    
        Set rsCONTROL = New Recordset
        strCONTROL = "SELECT * from ODASMInstallment where InvoiceNo = '" & .txtDocumentNo.Text & " ' ; "
        rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
    
        Do While Not rsCONTROL.EOF
                
                Set rsSAVE = New Recordset
                strSQL = "SELECT * from ODASMInstallment where InstallmentNo like '" & rsCONTROL!InstallmentNo & " ' ; "
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub

                rsSAVE!Status = "CHK-PREPARED"
                rsSAVE!StatusDate = Date
                rsSAVE!ChequeDate = CDate(.txtChequeDate.Text)
                
                rsSAVE!ChequeNo = Trim(.txtChequeNo.Text)
                    
                rsSAVE.Update
                                
                rsCONTROL.MoveNext
        Loop
    End With
    
            
        rsCONTROL.Close
        strSQL = ""
    
Exit Sub
err:
    UpdateErrorMessage
End Sub
Private Sub saveINSTALLMENTISSUED()
On Error GoTo err

    With frmODASMCheckIssuance
    
        Set rsCONTROL = New Recordset
        strCONTROL = "SELECT * from ODASMInstallment I where I.VoucherNo = '" & .txtVoucherNo & " ' ; "
        rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
    
                
        'Sum the Total Amount Paid from the Installment selected to be
        ' paid by the requisitions
        
        Dim rsINSTALL As ADODB.Recordset, strINSTALL As String, numAmountDue, numBALANCE As Double
        numAmountDue = 0
        
        Set rsINSTALL = New Recordset
        strINSTALL = "SELECT sum(TotalRent) as Totals from ODASMInstallment I where I.VoucherNo = '" & .txtVoucherNo & " ' ; "
        rsINSTALL.Open strINSTALL, cnCOMMON, adOpenKeyset, adLockOptimistic

        If rsINSTALL.EOF Or rsINSTALL.BOF Then
                numAmountDue = 0
        ElseIf IsNull(rsINSTALL!TOTALS) = True Then
                numAmountDue = 0
        Else: numAmountDue = CDbl(rsINSTALL!TOTALS)
        End If
                
        numBALANCE = numAmountDue
        
        Do While Not rsCONTROL.EOF

                Set rsSAVE = New Recordset
                strSQL = "SELECT * from ODASMInstallment where Installment = '" & rsCONTROL!Installment & " ' and  ContractNo = '" & rsCONTROL!ContractNo & " ' and ContractYear = '" & rsCONTROL!ContractYear & " ' and  PaymentMode = '" & rsCONTROL!PaymentMode & " '; "
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                Else
                        rsSAVE!Status = "CHK-ISSUED"
                        rsSAVE!StatusDate = Date
                        rsSAVE!CurrentPeriod = .txtCurrentPeriod.Text
                        
                        If CDbl(numAmountDue) >= CDbl(.txtChequeAmount.Text) Or PaymentDue = CDbl(.txtChequeAmount.Text) Then
                                rsSAVE!PaymentFlag = "Y"
                                rsSAVE!PaymentDue = 0
                                rsSAVE!Balance = 0
                                rsSAVE!AmountPaid = CDbl(rsCONTROL!TotalRent)
                        Else:
                                If CDbl(rsCONTROL!TotalRent) <= numBALANCE Then
                                        rsSAVE!PaymentDue = 0
                                        rsSAVE!PaymentFlag = "Y"
                                        rsSAVE!Balance = 0
                                        rsSAVE!AmountPaid = CDbl(rsCONTROL!TotalRent)
                                        numBALANCE = numBALANCE - CDbl(rsCONTROL!TotalRent)
                                Else
                                        rsSAVE!PaymentDue = rsCONTROL!TotalRent - numBALANCE
                                        rsSAVE!PaymentFlag = "P"
                                        rsSAVE!Balance = rsSAVE!PaymentDue
                                        rsSAVE!AmountPaid = numBALANCE
                                        numBALANCE = 0
                                End If
                        End If
                        
                        rsSAVE!PaymentDate = Date
                                               
                        rsSAVE.Update
                End If
                rsCONTROL.MoveNext
        Loop
    End With
    
    Exit Sub
            
        rsCONTROL.Close
        strSQL = ""
    
   
err:
    UpdateErrorMessage
End Sub

Private Sub updatePAYMENTFLAG()
On Error GoTo err
            
            Set rsSAVE = New ADODB.Recordset

            strSQL = "Select * from ODASMVoucher ;"
            rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

            With rsSAVE
                    Dim MYCOUNT As Double
                    MYCOUNT = .RecordCount
                    If .EOF Or .BOF Then Exit Sub
                    
                    Do While Not .EOF
                            
                            Select Case !Status
                            Case "REQ-PREP":
                                            !AmountPaid = 0
                                            !PaymentFlag = "N"
                                            !NoOfCheques = 0
                            Case "Payment Requisition - Authorization":
                                            !AmountPaid = !Amount
                                            !PaymentFlag = "Y"
                                            !NoOfCheques = 1
                            Case "Payment Requisition - Approval": !AmountPaid = !Amount
                                            !AmountPaid = !Amount
                                            !PaymentFlag = "Y"
                                            !NoOfCheques = 1
                            Case "CHQ -Prepared": !AmountPaid = !Amount
                                            !AmountPaid = !Amount
                                            !PaymentFlag = "Y"
                                            !NoOfCheques = 1
                            
                            Case "CHQ-PREPARED": !AmountPaid = !Amount
                                            !AmountPaid = !Amount
                                            !PaymentFlag = "Y"
                                            !NoOfCheques = 1
                            End Select
                    
                            .Update
                            .MoveNext
                    Loop
                    
            End With
Exit Sub

err:
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
            UpdateErrorMessage
    End If
End Sub
Private Sub updateCHECKDETAILS()
On Error GoTo err
            
            Set rsSAVE = New ADODB.Recordset

            strSQL = "Select * from ALISMChequeEntry ;"
            rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

            With rsSAVE
                    Dim MYCOUNT As Double
                    MYCOUNT = .RecordCount
                    
                    If .EOF Or .BOF Then Exit Sub
                    
                    Do While Not .EOF
                            Select Case !Status
                            Case "REQ-PREP": !AmountPaid = 0
                                            !PaymentFlag = "N"
                                            !NoOfCheques = 0
                            Case "Payment Requisition - Authorization":
                                            !AmountPaid = !Amount
                                            !PaymentFlag = "Y"
                                            !NoOfCheques = 1
                            Case "CHQ-PREPARED":
                                            !AmountDue = !Amount
                                            !PaymentFlag = "Y"
                                             End Select
                    
                        .Update
                        .MoveNext
                    Loop

            End With
Exit Sub

err:
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
            UpdateErrorMessage
    End If
End Sub
Public Sub updateCHECKPAYMENTFLAG()
On Error GoTo err
            
            Set rsSAVE = New ADODB.Recordset

            strSQL = "Select * from ALISMCheque ;"
            rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

            With rsSAVE
                    Dim MYCOUNT As Double
                    MYCOUNT = .RecordCount
                    
                    If .EOF Or .BOF Then Exit Sub
                    
                    Do While Not .EOF
                          !PaymentFlag = "Y"
                          !NoOfEntries = 1
                        .Update
                        .MoveNext
                    Loop

            End With
Exit Sub

err:
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
            UpdateErrorMessage
    End If
End Sub

Public Sub ChangeDATE()
On Error GoTo err
       Screen.ActiveForm.txtChequeDate.Text = Date
       Screen.ActiveForm.txtChequeDate.Text = Screen.ActiveForm.DTPickerChequeDate.Value
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub GetCheckedSCHEDULE()
   
    Dim j, i As Integer
    j = Screen.ActiveForm.ListView1.ListItems.Count

    For i = 1 To j
            If frmALISMCheckSchedule.ListView1.ListItems(i).Checked = True Then
                strCheckNo = Screen.ActiveForm.ListView1.ListItems(i).Text
                
                saveSCHEDULERECORD
                saveSCHEDULE
            
            End If
            
            strCheckNo = ""
    Next i
End Sub

Private Sub ValidateSchedule()
On Error GoTo err

        With frmALISMCheckSchedule
                If .txtBankNo.Text = "" Then
                        MsgBox "BankNo is Required"
                        .txtBankNo.SetFocus
                        
                ElseIf bscheduledCHEQUES = True And .txtReference.Text <= "" Then
                        MsgBox "The Reference No is REQUIRED", vbOKOnly
                        .txtReference.SetFocus
                
                ElseIf .txtDetails.Text = "" Then
                        MsgBox "The Names is  required"
                        .txtDetails.SetFocus
                        
                ElseIf .txtAccountNo.Text = "" Then
                        MsgBox "The AccountNo is  required"
                        .txtAccountNo.SetFocus
                        
                ElseIf .cboTemplateCode.Text = "" Then
                        MsgBox "You MUST select at Least one template", vbOKOnly
                        .cboTemplateCode.SetFocus
                
                ElseIf .txtDateBanked.Text <= "" Then
                        MsgBox "The Date banked MUST NOT be left Blank", vbOKOnly
                        .txtDateBanked.SetFocus
                        
                ElseIf DateDiff("D", .txtDateBanked, Date) < 0 Then
                        MsgBox "The Banking Date cannot be in the future", vbOKOnly
                        .txtDateBanked.SetFocus
                Else
                        bsaveRECORD = True
                End If
        End With
        
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub GetScheduledCheques()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView2.ListItems.Clear
                .ListView2.ColumnHeaders.Clear
                
                .ListView2.ColumnHeaders.Add , , "Cheque No", .ListView2.Width / 6
                .ListView2.ColumnHeaders.Add , , "Payee Details", .ListView2.Width / 4
                .ListView2.ColumnHeaders.Add , , "Bank No", .ListView2.Width / 4
                .ListView2.ColumnHeaders.Add , , "Cheque Date", .ListView2.Width / 6
                .ListView2.ColumnHeaders.Add , , "Amount Paid", .ListView2.Width / 6
                .ListView2.ColumnHeaders.Add , , "Scheduled", .ListView2.Width / 8

                .ListView2.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!ChequeNo))
                        
                        If Not IsNull(rsLIST!BankNo) Then
                            MyList.SubItems(2) = CStr(rsLIST!BankNo)
                        End If
                                                
                        If Not IsNull(rsLIST!ChequeDate) Then
                                MyList.SubItems(3) = CStr(rsLIST!ChequeDate)
                        End If

                        If Not IsNull(rsLIST!Amount) Then
                                MyList.SubItems(4) = CStr(rsLIST!Amount)
                        End If
                        
                        If Not IsNull(rsLIST!Sheduled) Then
                                MyList.SubItems(5) = CStr(rsLIST!Scheduled)
                        End If
                        
                        If Not IsNull(rsLIST!PayeeDetails) Then
                                MyList.SubItems(1) = CStr(rsLIST!PayeeDetails)
                        End If


                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Private Sub txtBankNo_GotFocus()
On Error GoTo err
        strSQL = "SELECT * FROM ALISPbankAccount;"
        bankNoGotFocus
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub txtBankNo_LostFocus()
    strSQL = "Select * from ALISPBankAccount  "
    BankNoLostFocus
End Sub

Public Sub checkREQUISITION()
On Error GoTo err
        
          Dim rsREQNO As ADODB.Recordset, strREQNO As String
          Set rsREQNO = New ADODB.Recordset
          
          strREQNO = "SELECT * FROM ODASMVoucher where ChequePrepared = 'N' ;"
          rsREQNO.Open strREQNO, cnCOMMON, adOpenKeyset, adLockOptimistic

          With rsREQNO
                  If .EOF Or .BOF Then
                        MsgBox "There are no Requisitions to be Processed, All Payment have been Made", vbOKOnly
                        Exit Sub
                  End If
                    
                  bcheckREQUISITION = True
                  bcheckREQENTRY = True
          End With
            
          rsREQNO.Close
          strREQNO = ""
         
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub GenerateChequeNo()
On Error GoTo err

        Set rsCONTROL = New Recordset
        
        strSQL = "SELECT * FROM ODASPLastNumbers WHERE AutoChequeNo = 'Y'"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
        With rsCONTROL
                If .BOF Or .EOF = True Then Exit Sub
                        Screen.ActiveForm.txtChequeNo.Text = !ChequeNo & ""
                
                Select Case Len(Trim(frmODASMCheck.txtChequeNo))
                        Case 1: frmODASMCheck.txtChequeNo.Text = Trim(!ChequePrefix) + "00000" + Trim(frmODASMCheck.txtChequeNo)
                        Case 2: frmODASMCheck.txtChequeNo.Text = Trim(!ChequePrefix) + "0000" + Trim(frmODASMCheck.txtChequeNo)
                        Case 3: frmODASMCheck.txtChequeNo.Text = Trim(!ChequePrefix) + "000" + Trim(frmODASMCheck.txtChequeNo)
                        Case 4: frmODASMCheck.txtChequeNo.Text = Trim(!ChequePrefix) + "00" + Trim(frmODASMCheck.txtChequeNo)
                        Case 5: frmODASMCheck.txtChequeNo.Text = Trim(!ChequePrefix) + "0" + Trim(frmODASMCheck.txtChequeNo)
                        Case 6: frmODASMCheck.txtChequeNo.Text = Trim(!ChequePrefix) + Trim(frmODASMCheck.txtChequeNo)
                End Select
                
                !ChequeNo = !ChequeNo + 1
                .Update

        End With
        
rsCONTROL.Close
strTRANS = ""

Exit Sub
err:
    ErrorMessage
End Sub
Public Sub GenerateIssuanceNo()
On Error GoTo err

        Set rsCONTROL = New Recordset
        
        strSQL = "SELECT * FROM ODASPLastNumbers WHERE AutoIssuanceNo = 'Y'"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
        With rsCONTROL
                If .BOF Or .EOF = True Then Exit Sub
                        Screen.ActiveForm.txtIssuanceNo.Text = !IssuanceNo & ""
                
                Select Case Len(Trim(frmODASMCheckIssuance.txtIssuanceNo))
                        Case 1: frmODASMCheckIssuance.txtIssuanceNo.Text = Trim(!IssuancePrefix) + "00000" + Trim(frmODASMCheckIssuance.txtIssuanceNo)
                        Case 2: frmODASMCheckIssuance.txtIssuanceNo.Text = Trim(!IssuancePrefix) + "0000" + Trim(frmODASMCheckIssuance.txtIssuanceNo)
                        Case 3: frmODASMCheckIssuance.txtIssuanceNo.Text = Trim(!IssuancePrefix) + "000" + Trim(frmODASMCheckIssuance.txtIssuanceNo)
                        Case 4: frmODASMCheckIssuance.txtIssuanceNo.Text = Trim(!IssuancePrefix) + "00" + Trim(frmODASMCheckIssuance.txtIssuanceNo)
                        Case 5: frmODASMCheckIssuance.txtIssuanceNo.Text = Trim(!IssuancePrefix) + "0" + Trim(frmODASMCheckIssuance.txtIssuanceNo)
                        Case 6: frmODASMCheckIssuance.txtIssuanceNo.Text = Trim(!IssuancePrefix) + Trim(frmODASMCheckIssuance.txtIssuanceNo)
                End Select
                
                !IssuanceNo = !IssuanceNo + 1
                .Update
        End With
        
rsCONTROL.Close
strTRANS = ""

Exit Sub
err:
    ErrorMessage
End Sub

Public Sub validateRECORD()
On Error GoTo err
    With frmODASMCheck
            bsaveRECORD = False
            
            If .txtAmountDue.Text = "" Then
                    MsgBox "The Amount Due cannot be Zero", vbOKOnly
                    .txtAmountDue.SetFocus
            
            ElseIf .txtChequeNo <= Empty Then
                    MsgBox "The cheque number must be entered", vbOKOnly
                    .txtChequeNo.SetFocus
            ElseIf .txtVoucherDate <= Empty Then
                    MsgBox "The Requisition date cannot be Left Blank", vbOKOnly
                    .txtVoucherDate.SetFocus
            
            ElseIf .txtChequeDate.Text <= Empty Then
                    MsgBox "The Cheque Date cannot be left Blank", vbOKOnly
                    .txtChequeDate.SetFocus
            
            ElseIf DateDiff("D", .txtChequeDate.Text, Date) < 0 Then
                    MsgBox "The Cheque Date cannot be in the FUTURE", vbOKOnly
                    .txtChequeDate.SetFocus

            ElseIf .txtBankNo.Text <= Empty Then
                    MsgBox "The bank No cannot be left Blank", vbOKOnly
                    .txtBankNo.SetFocus
            
            ElseIf .txtVoucherNo <= Empty Then
                    MsgBox "The Requisition date cannot be Left Blank", vbOKOnly
                    .txtVoucherNo.SetFocus

            ElseIf .txtPayeeDetails <= Empty Then
                    MsgBox "you must cobfirm the Payee Details prior to Approval", vbOKOnly
                    .txtPayeeDetails.SetFocus
            Else:
                    bsaveRECORD = True
            End If
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub validateCHECKSISSUED()
On Error GoTo err
     With frmODASMCheckIssuance
        
        If CDbl(.txtChequeAmount.Text) < 0 Then
                MsgBox "The Amount cannot be Zero", vbOKOnly
                .txtChequeAmount.SetFocus
        
        ElseIf .txtDateCollected <= Empty Then
                MsgBox "The colelction date cannot be Left Blank", vbOKOnly
                .txtDateCollected.SetFocus
        
        ElseIf DateDiff("D", .txtDateCollected.Text, Date) < 0 Then
                MsgBox "The Collection Date cannot be in the FUTURE", vbOKOnly
                .txtDateCollected.SetFocus

        ElseIf .txtBankNo.Text <= Empty Then
                MsgBox "The bank No cannot be left Blank", vbOKOnly
                .txtBankNo.SetFocus
        ElseIf .txtIdentityNo.Text <= Empty Then
                MsgBox "The Identity No cannot be left Blank", vbOKOnly
                .txtIdentityNo.SetFocus
        ElseIf .cboIdType.Text <= Empty Then
                MsgBox "The Identity Type cannot be left Blank", vbOKOnly
                .cboIdType.SetFocus
        ElseIf .txtChequeNo <= Empty Then
                MsgBox "The Cheque No cannot be Left Blank", vbOKOnly
                .txtChequeNo.SetFocus
                
        ElseIf .txtPayeeDetails <= Empty Then
                MsgBox "you must cobfirm the Payee Details prior to Approval", vbOKOnly
                .txtPayeeDetails.SetFocus
        Else:
                    bsaveRECORD = True
        End If
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub SaveChequeENTRIES()
On Error GoTo err
    Set rsSAVE = New Recordset
    
    strSQL = "SELECT * from ALISMChequeEntry where chequeEntryNo = '" & Screen.ActiveForm.txtChequeEntryNo.Text & "';"
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsSAVE
            If .EOF Or .BOF Then
                    .AddNew
                    !ChequeNo = Screen.ActiveForm.txtChequeNo.Text
                    !ChequeEntryNo = Screen.ActiveForm.txtChequeEntryNo.Text
                    !vOUCHERnO = Screen.ActiveForm.txtVoucherNo
                    !dateprepared = Date
                    !Preparedby = CurrentUserName
            End If
            
            !AmountDue = Screen.ActiveForm.txtAmountDue.Text
            !Amount = Screen.ActiveForm.txtAmountDue.Text
            !Status = Screen.ActiveForm.txtStatus.Text
            !PaymentFlag = Screen.ActiveForm.txtPaymentFlag.Text
            !BankNo = Screen.ActiveForm.txtBankNo.Text
            !remark = Screen.ActiveForm.txtRemark.Text
            bsaveRECORD = False
            .Update
            .Requery
        End With

strSAVE = ""
rsSAVE.Close

Exit Sub

err:
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
        Else
            UpdateErrorMessage
        End If
End Sub

Private Sub reverseChequeENTRIES()
On Error GoTo err
    Set rsSAVE = New Recordset
    
    strSAVE = "SELECT * from ALISMChequeEntry where chequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "' and (paymentflag = 'Y' or paymentflag = 'P');"

    rsSAVE.Open strSAVE, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsSAVE
            
            MYCOUNT = .RecordCount
            
            If .EOF Or .BOF Then Exit Sub
            
            Do While Not .EOF
                !Status = Screen.ActiveForm.txtStatus.Text
                !PaymentFlag = "V"
                .Update
                .MoveNext
            Loop
        End With

strSAVE = ""
rsSAVE.Close

Exit Sub

err:
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
        Else
            UpdateErrorMessage
        End If
End Sub

Private Sub voidCHEQUEENTRY()
On Error GoTo err

    Dim strCONTROL As String
    Set rsSAVE = New Recordset
    
    strCONTROL = "SELECT * from ALISMChequeEntry where chequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "' and paymentflag = 'V';"
    strSQL = "SELECT * from ALISMChequeEntry where chequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "' and paymentflag = 'C';"

    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic

    
    With rsCONTROL
            If .EOF Or .BOF Then Exit Sub
            MYCOUNT = .RecordCount
            
            Do While Not .EOF
                    rsSAVE.AddNew
                    rsSAVE!AmountDue = rsCONTROL!AmountDue
                    rsSAVE!Amount = (rsCONTROL!Amount) * "Y"
                    rsSAVE!vOUCHERnO = rsCONTROL!vOUCHERnO
                    rsSAVE!Status = Screen.ActiveForm.txtStatus.Text
                    rsSAVE!ChequeNo = rsCONTROL!ChequeNo
                    rsSAVE!CurrentPeriod = CurrentPeriod
                    rsSAVE!PaymentFlag = "C"
            
                    rsSAVE!Preparedby = CurrentUserName
                    rsSAVE!dateprepared = Date
                    rsSAVE.Update
                    .MoveNext
            Loop

    End With

strSQL = ""
rsSAVE.Close
rsCONTROL.Close

Exit Sub

err:
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
        Else
            UpdateErrorMessage
        End If
End Sub
Private Sub saveNEGATIVEENTRY()
On Error GoTo err
    Set rsSAVE = New Recordset
    
    If bmakePAYMENT = True Then
            strSAVE = "SELECT * from ALISMChequeEntry where chequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "' and VoucherNo = '" & Screen.ActiveForm.txtVoucherNo.Text & "';"
    ElseIf breversePAYMENT = True Then
            strSAVE = "SELECT * from ALISMChequeEntry where chequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "' and VoucherNo = '" & Screen.ActiveForm.txtVoucherNo.Text & "' and (paymentflag = 'Y' or paymentflag = 'P');"
    End If
    
    rsSAVE.Open strSAVE, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsSAVE
            .AddNew
            !AmountDue = Screen.ActiveForm.txtAmountDue.Text
            !Amount = CDbl(Screen.ActiveForm.txtAmountPaid.Text) * "Y"
            !vOUCHERnO = Screen.ActiveForm.txtVoucherNo
            !Status = Screen.ActiveForm.txtStatus.Text
            !ChequeNo = Screen.ActiveForm.txtChequeNo.Text
            !PaymentFlag = "C"

            !Preparedby = CurrentUserName
            !dateprepared = Date
            bsaveRECORD = False
            .Update
        End With

strSAVE = ""
rsSAVE.Close

Exit Sub

err:
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
        Else
            UpdateErrorMessage
        End If
End Sub

Public Sub saveCHEQUE()
On Error GoTo err
    
    Set rsSAVE = New ADODB.Recordset
    strSQL = "SELECT * from ALISMCheque where chequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "';"
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsSAVE
            If .EOF Or .BOF Then
                    .AddNew
                    !ChequeNo = Screen.ActiveForm.txtChequeNo.Text
                    !ChequeDate = Screen.ActiveForm.txtChequeDate.Text
                    !Preparedby = CurrentUserName
                    !dateprepared = Date
                    !Prepared = "Y"
                    !Issued = "N"
                    !Authorized = "N"
                    !Approved = "N"
                    !Scheduled = "N"
                    !CurrentPeriod = CurrentPeriod
                    !NoOfEntries = 0
                    
                    If rsDEFAULT!AutoApproval = "Y" Then
                            !ApprovedBy = CurrentUserName
                            !DateApproved = Date
                            !Status = "CHK-APPROVED"
                            !Approved = "Y"
                    Else:
                            !Approved = "N"
                            !Status = Screen.ActiveForm.txtStatus.Text
                    End If

            End If
            
            !Accountno = Screen.ActiveForm.txtAccountNo.Text
            !ChequeAmount = CDbl(Screen.ActiveForm.txtChequeAmount.Text)
            !PaymentFlag = frmODASMCheck.txtPaymentFlag.Text
            !NoOfEntries = Screen.ActiveForm.txtNoOfEntries.Text
            !BankNo = Screen.ActiveForm.txtBankNo.Text
            !PayeeDetails = Screen.ActiveForm.txtPayeeDetails
            
            .Update
            .Requery
        End With

strSAVE = ""

Exit Sub

err:
If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
    rsSAVE.CancelUpdate
    rsSAVE.Requery
End If
    UpdateErrorMessage
End Sub

Public Sub saveCHECKSISSUED()
On Error GoTo err
    
    Set rsSAVE = New ADODB.Recordset
    
    strSQL = "SELECT * from ALISMCheque where chequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "' and approved = 'Y';"
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsSAVE
            If .EOF Or .BOF Then Exit Sub
                    !DateIssued = frmODASMCheckIssuance.txtDateCollected.Text
                    !IssuedBy = CurrentUserName
                    !Banked = "N"
                    !Issued = "Y"
                    .Update
                    .Requery
        End With

rsSAVE.Close
strSAVE = ""

Exit Sub

err:
If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
    rsSAVE.CancelUpdate
    rsSAVE.Requery
End If
    UpdateErrorMessage
End Sub

Public Sub saveRecord()
On Error GoTo err
    
    Set rsSAVE = New ADODB.Recordset
    
    strSQL = "SELECT * from ALISMChecksIssued where IssuanceNo = '" & Screen.ActiveForm.txtIssuanceNo.Text & "';"
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsSAVE
            If .EOF Or .BOF Then
                    .AddNew
                    !IssuanceNo = frmODASMCheckIssuance.txtIssuanceNo.Text
                    !ChequeNo = frmODASMCheckIssuance.txtChequeNo.Text
                    !dateprepared = Date
                    !Preparedby = CurrebtUserName
            End If
            
        !CurrentPeriod = CurrentPeriod
        !DateCollected = frmODASMCheckIssuance.txtDateCollected.Text
        !CollectedBy = frmODASMCheckIssuance.txtCollectedBy.Text
        !Amount = frmODASMCheckIssuance.txtChequeAmount.Text
        !IdentityNo = frmODASMCheckIssuance.txtIdentityNo.Text
        !IDType = frmODASMCheckIssuance.cboIdType.Text
        !IssuedBy = CurrentUserName
        !Collected = "Y"
        !Comment = frmODASMCheckIssuance.txtComment.Text
        .Update
        .Requery
        End With

strSAVE = ""

Exit Sub

err:
If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
    rsSAVE.CancelUpdate
    rsSAVE.Requery
End If
    UpdateErrorMessage
End Sub

Public Sub voidCHEQUE()
On Error GoTo err
    
    Set rsSAVE = New ADODB.Recordset
    
    strSQL = "SELECT * from ALISMCheque where chequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "' and paymentFlag = 'V';"
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsSAVE
            .AddNew
                !Amount = CDbl(Screen.ActiveForm.txtAmountPaid.Text) * "Y"
                !NoOfEntries = Screen.ActiveForm.txtNoOfEntries.Text
                !ChequeNo = Screen.ActiveForm.txtChequeNo.Text
                !Status = Screen.ActiveForm.txtStatus.Text
                !BankNo = Screen.ActiveForm.txtBankNo.Text
                !ChequeDate = Screen.ActiveForm.txtChequeDate.Text
                !PayeeDetails = Screen.ActiveForm.txtPayeeDetails
                !CurrentPeriod = CurrentPeriod
                !Prepared = "Y"
                !PaymentFlag = "C"
                !Preparedby = CurrentUserName
                !dateprepared = Date
            .Update
        End With

rsSAVE.Close
strSAVE = ""

Exit Sub

err:
If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
    rsSAVE.CancelUpdate
    rsSAVE.Requery
End If
    UpdateErrorMessage
End Sub

Private Sub updateREQUISITION()
On Error GoTo err

        Set rsSAVE = New ADODB.Recordset
        
        rsSAVE.Open "SELECT * FROM ODASMVoucher WHERE VoucherNo = '" & Screen.ActiveForm.txtVoucherNo.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
  
        With rsSAVE
            If .BOF Or .EOF = True Then Exit Sub
            !AmountPaid = Screen.ActiveForm.txtAmountDue.Text
            !Status = Screen.ActiveForm.txtStatus.Text
            !ChequeNo = Screen.ActiveForm.txtChequeNo.Text
            !ChequePrepared = "Y"
            .Update
            .Requery
        End With

rsSAVE.Close

Exit Sub

err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
        rsSAVE.CancelUpdate
        rsSAVE.Requery
    Else
        UpdateErrorMessage
    End If
    
End Sub

Private Sub updateREQUISITIONIssued()
On Error GoTo err
        
        Set rsSAVE = New ADODB.Recordset
        
        rsSAVE.Open "SELECT * FROM ODASMVoucher WHERE VoucherNo = '" & Screen.ActiveForm.txtVoucherNo.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
  
        With rsSAVE
            If .BOF Or .EOF = True Then Exit Sub
            !Status = Screen.ActiveForm.txtStatus.Text
            !ChequeIssued = "Y"
            !ChequeIssuedBy = CurrentUserName
            !DateChequeIssued = Date
            .Update
            .Requery
        End With

rsSAVE.Close

Exit Sub

err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
        rsSAVE.CancelUpdate
        rsSAVE.Requery
    Else
        UpdateErrorMessage
    End If
    
End Sub

Private Sub updateREQUISITIONApproved()
On Error GoTo err
        
        Set rsSAVE = New ADODB.Recordset
        
        rsSAVE.Open "SELECT * FROM ODASMVoucher WHERE VoucherNo = '" & Screen.ActiveForm.txtVoucherNo.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
  
        With rsSAVE
            If .BOF Or .EOF = True Then Exit Sub
            !Status = Screen.ActiveForm.txtStatus.Text
            !ChequeApproved = "Y"
            !ChequeApprovedBy = CurrentUserName
            !DateChequeApproved = Date
            .Update
            .Requery
        End With

rsSAVE.Close

Exit Sub

err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
        rsSAVE.CancelUpdate
        rsSAVE.Requery
    Else
        UpdateErrorMessage
    End If
    
End Sub
Private Sub updateREQUISITIONAuthorized()
On Error GoTo err
        
        Set rsSAVE = New ADODB.Recordset
        
        rsSAVE.Open "SELECT * FROM ODASMVoucher WHERE VoucherNo = '" & Screen.ActiveForm.txtVoucherNo.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
  
        With rsSAVE
            If .BOF Or .EOF = True Then Exit Sub
            !Status = Screen.ActiveForm.txtStatus.Text
            !ChequeAuthorized = "Y"
            !ChequeAuthorizedBy = CurrentUserName
            !DateAuthorizedApproved = Date
            .Update
            .Requery
        End With

rsSAVE.Close

Exit Sub

err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
        rsSAVE.CancelUpdate
        rsSAVE.Requery
    Else
        UpdateErrorMessage
    End If
    
End Sub

Public Sub UpDatetheInvoice()
On Error GoTo err

        Set rsSAVE = New ADODB.Recordset
        
        rsSAVE.Open "SELECT V.*,C.*,VI.*,(I.AmountPaid)as AP,(V.AmountPaid)as VAP, I.* FROM ODASMVoucher V,ALISMChequeEntry C,ODASMVoucherItem VI, ODASMInvoice I WHERE C.ChequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "' and C.VoucherNo = V.VoucherNo and V.VoucherNo = VI.VoucherNo and VI.DocumentNo = I.InvoiceNo", cnCOMMON, adOpenKeyset, adLockOptimistic
  
        With rsSAVE
            If .BOF Or .EOF = True Then Exit Sub
            !AP = !VAP
            !InvoiceBalance = !PriceInclusive - !VAP
            !ChequeNo = Screen.ActiveForm.txtChequeNo.Text
            If CDbl(!PriceInclusive - !VAP) <= 0 Then
                !Paid = "Y"
                !Status = "FULLY PAID"
            Else
                !Paid = "N"
                !Status = "PARTIALLY PAID"
            End If
            .Update
            .Requery
        End With

rsSAVE.Close

Exit Sub

err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
        rsSAVE.CancelUpdate
        rsSAVE.Requery
    Else
        UpdateErrorMessage
    End If
End Sub
    
Public Sub countENTRIES()
On Error GoTo err
    
    Set rsSAVE = New ADODB.Recordset
    strSQL = "SELECT * from ALISMChequeEntry where ChequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "' and PaymentFlag <> 'C';"
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsSAVE
            If .BOF Or .EOF Then
                    Screen.ActiveForm.txtNoOfEntries.Text = 1
            Else
                    Screen.ActiveForm.txtNoOfEntries.Text = .RecordCount + 1
            End If
            
            frmODASMCheck.txtChequeEntryNo.Text = Trim(frmODASMCheck.txtChequeNo.Text) + "-" + Trim(Screen.ActiveForm.txtNoOfEntries.Text)
    End With

rsSAVE.Close
strSAVE = ""

Exit Sub

err:
If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
    rsSAVE.CancelUpdate
    rsSAVE.Requery
End If
    UpdateErrorMessage
End Sub

Public Sub calculateTOTALPAID()
On Error GoTo err
    
    Set rsSAVE = New ADODB.Recordset
    strSQL = "SELECT sum(Amount)as TotalPaid from ALISMChequeEntry where ChequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "';"
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsSAVE
                
                If .EOF Or .BOF Then
                        Screen.ActiveForm.txtChequeAmount.Text = FormatNumber(Screen.ActiveForm.txtAmountDue.Text)
                ElseIf IsNull(!TotalPaid) = True Then
                        frmODASMCheck.txtChequeAmount.Text = frmODASMCheck.txtAmountDue.Text
                ElseIf (Screen.ActiveForm.txtAmountDue.Text) = Empty Then
                       Screen.ActiveForm.txtChequeAmount.Text = FormatNumber(!TotalPaid)
                Else: Screen.ActiveForm.txtChequeAmount.Text = FormatNumber(!TotalPaid)
                End If
        End With

rsSAVE.Close
strSAVE = ""

Exit Sub

err:
If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
    rsSAVE.CancelUpdate
    rsSAVE.Requery
End If
    UpdateErrorMessage
End Sub
Public Sub saveREFERENCE()
On Error GoTo err

    Set rsSAVE = New ADODB.Recordset
    
    strSQL = ""
    strSQL = "SELECT * FROM ALISMBankAccountCfg WHERE Reference =  '" & frmALISMCheckSchedule.txtReference & " ' ;"
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                      
    With rsSAVE
            If .BOF Or .EOF Then
                    .AddNew
                    !Reference = frmALISMCheckSchedule.txtReference.Text
             End If
             
             !BankNo = frmALISMCheckSchedule.txtBankNo.Text
             !dateprepared = Date
             !Preparedby = CurrentUserName
             !Accountno = frmALISMCheckSchedule.txtAccountNo.Text
             !ReceivedPayed = "Paid"
             !CurrentPeriod = CurrentPeriod
             !templateCode = frmALISMCheckSchedule.cboTemplateCode.Text
             !DateBanked = frmALISMCheckSchedule.txtDateBanked.Text
             !TotalAmount = CDbl(frmALISMCheckSchedule.txtTotalAmount.Text)
             !NoOfEntries = CDbl(frmALISMCheckSchedule.txtNoOfEntries.Text)
            .Update
            .Requery
    End With

rsSAVE.Close
strSQL = ""
            
Exit Sub

err:
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
                    rsSAVE.CancelUpdate
                    rsSAVE.Requery
        Else
                    UpdateErrorMessage
        End If
  
End Sub
Private Sub ClearControls()
On Error GoTo err
        With frmALISMCheckSchedule
            .txtReference.Text = ""
            .txtTotalAmount.Text = ""
            .txtNoOfEntries.Text = ""
            .cboTemplateCode.Text = ""
            .ListView1.ListItems.Clear
            .ListView2.ListItems.Clear
        End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub EnableControls()
On Error GoTo err
        With frmALISMCheckSchedule
            .txtBankNo.Locked = True
            .txtAccountNo.Locked = True
            .txtDetails.Locked = True
            .txtReference.Locked = True
            .txtTotalAmount.Locked = True
            .txtNoOfEntries.Locked = True
            .cboTemplateCode.Locked = False

        End With
Exit Sub

err:
        ErrorMessage
End Sub

Private Sub disableControls()
On Error GoTo err
   With frmALISMCheckSchedule
            .txtBankNo.Locked = True
            .txtDetails.Locked = True
            .txtAccountNo.Locked = True
            .txtReference.Locked = True
            .txtTotalAmount.Locked = True
            .txtNoOfEntries.Locked = True
            .cboTemplateCode.Locked = True

    End With
Exit Sub

err:
    ErrorMessage
End Sub


Public Sub addSCHEDULE()
        bAddNew = True
        If bschedulePAYMENT = True Then
                    ClearControls
                    EnableControls
                    disableButtons
                    strSQL = "Select ALISMCheque.ChequeNo, ALISMCheque.PayeeDetails, ALISMCheque.ChequeDate, ALISMCheque.Amount, ALISMCheque.PaymentFlag from ALISMCheque Where ALISMCheque.issued = 'Y' and ALISMCheque.Scheduled = 'N'"
                    GetCheques
                    strSQL = "Select ALISMCheque.ChequeNo, ALISMCheque.PayeeDetails, BankNo, ALISMCheque.ChequeDate, ALISMCheque.Amount, ALISMCheque.PaymentFlag from ALISMCheque Where ALISMCheque.issued = 'Y' and ALISMCheque.Scheduled = 'Y' and Reference = '" & frmALISMCheckSchedule.txtReference.Text & "'"
                    GetScheduledCheques
                    strSQL = ""
        ElseIf bscheduledCHEQUES = True Then
                    EnableControls
                    disableButtons
                    strSQL = "Select ALISMCheque.ChequeNo, ALISMCheque.PayeeDetails, ALISMCheque.ChequeDate, ALISMCheque.Amount, ALISMCheque.PaymentFlag from ALISMCheque Where ALISMCheque.issued = 'Y' and ALISMCheque.Scheduled = 'N'"
                    GetCheques
                    strSQL = "Select ALISMCheque.ChequeNo, ALISMCheque.PayeeDetails, BankNo, ALISMCheque.ChequeDate, ALISMCheque.Amount, ALISMCheque.PaymentFlag from ALISMCheque Where ALISMCheque.issued = 'Y' and ALISMCheque.Scheduled = 'Y' and Reference = '" & frmALISMCheckSchedule.txtReference.Text & "'"
                    GetScheduledCheques
                    strSQL = ""
        End If
End Sub

Public Sub cancelSCHEDULE()
        enableButtons
        ClearControls
        disableControls
End Sub

Public Sub updateSCHEDULE()
        bsaveRECORD = False
        ValidateSchedule
        
        If bsaveRECORD = True Then
            If bschedulePAYMENT = True Then
                    generateReference
            End If
            
            GetCheckedSCHEDULE
            countSCHEDULEENTRIES
            calculateTOTAL
            saveREFERENCE
            bsaveRECORD = False
        End If
        
        enableButtons
        disableALLRECORD
        
        strSQL = "Select ALISMCheque.ChequeNo, ALISMCheque.PayeeDetails, ALISMCheque.ChequeDate, ALISMCheque.Amount, ALISMCheque.PaymentFlag from ALISMCheque Where ALISMCheque.issued = 'Y' and ALISMCheque.Scheduled = 'N'"
        GetCheques
        strSQL = "Select ALISMCheque.ChequeNo, ALISMCheque.PayeeDetails, BankNo, ALISMCheque.ChequeDate, ALISMCheque.Amount, ALISMCheque.PaymentFlag from ALISMCheque Where ALISMCheque.issued = 'Y' and ALISMCheque.Scheduled = 'Y' and Reference = '" & frmALISMCheckSchedule.txtReference.Text & "'"
        GetScheduledCheques
        strSQL = ""


End Sub

Private Sub saveSCHEDULERECORD()
On Error GoTo err
        
        Dim rsSAVE As ADODB.Recordset
        Set rsSAVE = New ADODB.Recordset
        
        strSQL = ""
        strSQL = "Select * from ALISMChequeSchedule where Reference = '" & frmALISMCheckSchedule.txtReference & "' and chequeNo = '" & strCheckNo & "'  ;"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                        
        With rsSAVE
                                
                    If .BOF Or .EOF = True Then
                        .AddNew
                        !Reference = frmALISMCheckSchedule.txtReference.Text
                        !ChequeNo = strCheckNo
                    End If
                    
                        !BankNo = frmALISMCheckSchedule.txtBankNo.Text
                        !CurrentPeriod = CurrentPeriod
                        !Status = "CHK-SCHEDULED"
                        !Scheduled = "Y"
                        !dateprepared = Date
                        !Preparedby = CurrentUserName
                        .Update
        End With
 
Exit Sub
err:
        If err.Number = 3021 Then Resume Next

        UpdateErrorMessage
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
        End If
  
End Sub
Private Sub countSCHEDULEENTRIES()
On Error GoTo err
        
        Set rsSAVE = New ADODB.Recordset
        
        strSQL = "Select * from ALISMChequeSchedule where Reference = '" & frmALISMCheckSchedule.txtReference & "' ;"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                        
        With rsSAVE
                    If .EOF Or .BOF Then Exit Sub
                    
                    frmALISMCheckSchedule.txtNoOfEntries.Text = .RecordCount
        End With
 
Exit Sub
err:
        If err.Number = 3021 Then Resume Next

        UpdateErrorMessage
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
        End If
  
End Sub
Private Sub calculateTOTAL()
On Error GoTo err
        Dim TOTALS As Double
        Set rsSAVE = New ADODB.Recordset
        
        strSQL = "Select sum(Amount)as Totals from ALISMCheque where Reference = '" & frmALISMCheckSchedule.txtReference & "' ;"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                        
        With rsSAVE
                    If .EOF Or .BOF Then Exit Sub
                    
                    If IsNull(!TOTALS) = True Then !TOTALS = 0
                    frmALISMCheckSchedule.txtTotalAmount.Text = !TOTALS
        End With
 
Exit Sub
err:
        If err.Number = 3021 Then Resume Next

        UpdateErrorMessage
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
        End If
  
End Sub


Private Sub saveSCHEDULE()
On Error GoTo err
        
        Dim rsSAVE As ADODB.Recordset
        Set rsSAVE = New ADODB.Recordset
        
        strSQL = ""
        
        strSQL = "Select * from ALISMCheque where Scheduled = 'N' and chequeNo = '" & strCheckNo & "'  ;"
        
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                        
        With rsSAVE
                                
                  If .BOF Or .EOF = True Then Exit Sub
                              
                        !Reference = frmALISMCheckSchedule.txtReference.Text
                        !Status = "CHK-SCHEDULED"
                        !Scheduled = "Y"
                        !CurrentPeriod = CurrentPeriod
                        !DateScheduled = Date
                        !ScheduledBy = CurrentUserName
                        .Update
        End With
 
Exit Sub
err:
        If err.Number = 3021 Then Resume Next

        UpdateErrorMessage
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
        End If
  
End Sub

Public Sub generateReference()
On Error GoTo err
        Dim rsLAST As ADODB.Recordset, strLAST As String
        
        Set rsLAST = New Recordset
      
        strLAST = "SELECT * FROM ALISPBankAccount Where BankNo = '" & frmALISMCheckSchedule.txtBankNo.Text & "' ;"
        rsLAST.Open strLAST, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsLAST
                If .BOF Or .EOF Then Exit Sub
                
                frmALISMCheckSchedule.txtReference = !Reference
                
                Select Case Len(Trim(frmALISMCheckSchedule.txtReference))
                    Case 1: frmALISMCheckSchedule.txtReference = Trim(!ReferencePrefix) + "/00000" + Trim(frmALISMCheckSchedule.txtReference)
                    Case 2: frmALISMCheckSchedule.txtReference = Trim(!ReferencePrefix) + "/0000" + Trim(frmALISMCheckSchedule.txtReference)
                    Case 3: frmALISMCheckSchedule.txtReference = Trim(!ReferencePrefix) + "/000" + Trim(frmALISMCheckSchedule.txtReference)
                    Case 4: frmALISMCheckSchedule.txtReference = Trim(!ReferencePrefix) + "/00" + Trim(frmALISMCheckSchedule.txtReference)
                    Case 5: frmALISMCheckSchedule.txtReference = Trim(!ReferencePrefix) + "/0" + Trim(frmALISMCheckSchedule.txtReference)
                    Case 6: frmALISMCheckSchedule.txtReference = Trim(!ReferencePrefix) + "/" + Trim(frmALISMCheckSchedule.txtReference)
                End Select
                
                !Reference = !Reference + 1
            
                .Update
        End With

Exit Sub

err:
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
                    rsLAST.CancelUpdate
                    rsLAST.Requery
        Else
                    UpdateErrorMessage
        End If

End Sub

