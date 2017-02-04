VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsALISReverseCheque"
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
        .txtTotalPaid.Text = ""
        .txtNoOfEntries.Text = ""
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub pclearREQUISITION()
On Error GoTo err
    With Screen.ActiveForm
        If bmakePAYMENT = True Then
            .txtBankNo.Text = ""
            .txtChequeNo.Text = ""
            .txtTotalPaid.Text = ""
            .txtNoOfEntries.Text = ""
            
        ElseIf breversePAYMENT = True Then
            .txtComment.Text = ""
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

Public Sub loadRECORD()
On Error GoTo err

        Dim rsLOAD As ADODB.Recordset
        Set rsLOAD = New ADODB.Recordset
                
        If bmakePAYMENT = True Then
            strSQL = "Select * from ODASMVoucher where VoucherNo = '" & Screen.ActiveForm.txtVoucherNo.Text & "' ;"
        ElseIf breversePAYMENT = True Then
            strSQL = "Select * from ALISMChequeEntry,ALISMCheque where ALISMCheque.ChequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "' and ALISMCheque.chequeNo = ALISMCheque.chequeNo and ALISMCheque.PaymentFlag = 'Y' ;"
        ElseIf bApproveCheque = True Then
            strSQL = "Select * from ALISMChequeEntry,ALISMCheque where ALISMCheque.ChequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "' and ALISMCheque.chequeNo = ALISMCheque.chequeNo and ALISMCheque.Prepared = 'Y' and ALISmCheque.approved = 'N' ;"
        ElseIf bAuthorizeCheque = True Then
            strSQL = "Select * from ALISMChequeEntry,ALISMCheque where ALISMCheque.ChequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "' and ALISMCheque.chequeNo = ALISMCheque.chequeNo and ALISMCheque.Approved = 'Y' and ALISmCheque.Authorized = 'N' ;"
        End If
        
        rsLOAD.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsLOAD
                
                If .BOF Or .EOF Then Exit Sub
                
                If bmakePAYMENT = True Then
                        Screen.ActiveForm.txtChequeNo.Text = !ChequeNo & ""
                        Screen.ActiveForm.txtVoucherDate.Text = !VoucherDate & ""
                        Screen.ActiveForm.txtAmountPaid.Text = !Amount - !AmountPaid & ""
                        Screen.ActiveForm.txtChequeAmount.Text = CDbl(!Amount) - CDbl(!AmountPaid) & ""
                        Screen.ActiveForm.txtPaymentFlag.Text = "Y"
                        Screen.ActiveForm.txtAmountDue.Text = CDbl(!Amount) - CDbl(!AmountPaid)
                        Screen.ActiveForm.txtChequeDate.Text = Date
                        Screen.ActiveForm.DTPickerChequeDate.Value = Date
                        Screen.ActiveForm.txtStatus.Text = "CHK-PREPARED"
                        Screen.ActiveForm.txtTotalPaid.Text = !AmountPaid
                
                ElseIf breversePAYMENT = True Or bApproveCheque = True Or bAuthorizeCheque = True Then
                        
                        Screen.ActiveForm.txtBankNo.Text = !BankNo
                        Screen.ActiveForm.txtChequeNo.Text = !ChequeNo
                        Screen.ActiveForm.txtNoOfEntries.Text = !NoOfEntries & ""
                        Screen.ActiveForm.txtPayeeDetails.Text = !PayeeDetails & ""
                        Screen.ActiveForm.txtVoucherNo.Text = !vOUCHERnO & ""
                        Screen.ActiveForm.txtAmountPaid.Text = !Amount & ""
                        Screen.ActiveForm.txtChequeAmount.Text = CDbl(!Amount) & ""
                        Screen.ActiveForm.txtPaymentFlag.Text = !PaymentFlag
                        Screen.ActiveForm.txtAmountDue.Text = CDbl(!Amount)
                        Screen.ActiveForm.txtChequeDate.Text = !ChequeDate
                        Screen.ActiveForm.DTPickerChequeDate.Value = !ChequeDate
                        Screen.ActiveForm.txtStatus.Text = !Status
                        
                End If
        End With

Exit Sub

err:
    ErrorMessage
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
                        Screen.ActiveForm.txtDocumentNo.Text = !JobBriefNo & ""
                        Screen.ActiveForm.txtPayeeDetails.Text = !CompanyName & ""
                        Screen.ActiveForm.txtVoucherDate.Text = !VoucherDate & ""
                        Screen.ActiveForm.txtAmountPaid.Text = FormatNumber(!Amount) & ""

                        Screen.ActiveForm.txtChequeAmount.Text = FormatNumber(!Amount) & ""
                        Screen.ActiveForm.txtPaymentFlag.Text = "Y"
                        Screen.ActiveForm.txtAmountDue.Text = FormatNumber(!Amount)
                        Screen.ActiveForm.txtChequeDate.Text = Date
                        Screen.ActiveForm.DTPickerChequeDate.Value = Date
                        Screen.ActiveForm.txtStatus.Text = "CHK-PREPARED"
                        Screen.ActiveForm.txtTotalPaid.Text = FormatNumber(!Amount)
                End If
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
            strSQL = "Select * from ALISPBankAccount Where BankNo = '" & frmALISMCheckSchedule.txtBankNo.Text & "'"
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

        Dim rsLOAD As ADODB.Recordset
        Set rsLOAD = New ADODB.Recordset
                
        rsLOAD.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsLOAD
                
                If .BOF Or .EOF Then Exit Sub
                
                frmODASMCheckIssuance.txtChequeNo.Text = !ChequeNo
                frmODASMCheckIssuance.txtPayeeDetails.Text = !PayeeDetails
                frmODASMCheckIssuance.txtChequeAmount.Text = !Amount & ""
                frmODASMCheckIssuance.txtDateCollected.Text = Date
                frmODASMCheckIssuance.DTPickerDateCollected.Value = Date
                frmODASMCheckIssuance.txtStatus.Text = "CHK-ISSUED"
                frmODASMCheckIssuance.txtCurrentPeriod.Text = CurrentPeriod
                frmODASMCheckIssuance.txtIssuedBy.Text = CurrentUserName
                frmODASMCheckIssuance.txtCollectedBy.Text = !PayeeDetails & ""
                frmODASMCheckIssuance.txtBankNo.Text = !BankNo & ""
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
                        .txtTotalPaid.Locked = True
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
                        .txtTotalPaid.Locked = True
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
        loadRECORD
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
                        'screen.activeform.txtRequirementCode.Text = strRequirementcode
                        
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
                enableALLRECORD
                disableButtons
                
                pclearREQUISITION
                calculateTOTALPAID
                calculateAMOUNTDUE
                
                If bmakePAYMENT = True Then
                        GetPreviousPayment
                ElseIf breversePAYMENT = True Then
                        GetOtherPayment
                End If
                
                countENTRIES
        End If
End Sub

Public Sub Cancelrecord()
On Error GoTo err
        enableButtons
        clearALLRECORD
        disableALLRECORD
Exit Sub

err:
    ErrorMessage
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
Public Sub enableRECORD()
On Error GoTo err
    With frmODASMCheck
            
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
                GenerateChequeNo
                
                SaveChequeENTRIES
                
                countENTRIES
                calculateTOTALPAID
                calculateAMOUNTDUE
                saveCHEQUE
                
                updateREQUISITION
                GetPreviousPayment
                disableALLRECORD
                showALLCHKENTRIES

        End If
End Sub

Public Sub UpdateChecksIssued()
        bsaveRECORD = False
            validateCHECKSISSUED

        If bsaveRECORD = True Then
                saveRecord
                If bsaveRECORD = False Then
                saveCHECKSISSUED
                enableButtons
                disableALLRECORD
                End If
        End If
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

Public Sub GetApprovedChecks()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Cheque No", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Payee Details", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Date", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Scheduled", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ChequeNo))
                        
                        If Not IsNull(rsLIST!ChequeDate) Then
                            MyList.SubItems(2) = CStr(rsLIST!ChequeDate)
                        End If

                        If Not IsNull(rsLIST!Amount) Then
                                MyList.SubItems(3) = CStr(rsLIST!Amount)
                        End If
                        
                        If Not IsNull(rsLIST!PaymentFlag) Then
                                MyList.SubItems(4) = CStr(rsLIST!PaymentFlag)
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
Public Sub GetSchedule()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Bank No", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "AccountNo", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Details", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select BankNo, AccountNo, Details from ALISPBankAccount "
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!BankNo))
                        
                        If Not IsNull(rsLIST!Accountno) Then
                            MyList.SubItems(1) = CStr(rsLIST!Accountno)
                        End If

                        If Not IsNull(rsLIST!Details) Then
                                MyList.SubItems(2) = CStr(rsLIST!Details)
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
        
          Set rsCONTROL = New ADODB.Recordset
          
          strSQL = "SELECT * FROM ALISPLastNumber where AutoChequeNo = 'Y'; "
          rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

          With rsCONTROL
                    If .EOF Or .BOF Then
                            If Screen.ActiveForm.txtChequeNo.Text <= "" Then
                                    MsgBox "The Cheque No cannot be left Blank", vbOKOnly
                                    Screen.ActiveForm.txtChequeNo.SetFocus
                            End If
                            Exit Sub
                    End If
                    
                    Screen.ActiveForm.txtChequeNo.Text = !ChequeNo
                    !ChequeNo = CDbl(!ChequeNo) + 1
                    .Update
                    .Requery
          End With

rsCONTROL.Close
strSQL = ""
         
Exit Sub

err:
    ErrorMessage

End Sub
Public Sub validateRECORD()
On Error GoTo err
            
            bCONTINUE = False
            
            If breversePAYMENT = True Then
                    If Screen.ActiveForm.cboReversalType <= "" Then
                            MsgBox "The Reversal Type cannot be left Blank", vbOKOnly
                            Screen.ActiveForm.cboReversalType.SetFocus
                    ElseIf Screen.ActiveForm.txtComment.Text <= "" Then
                            MsgBox "The Comment is Manadatory and MUST be greater than 50 Characters", vbOKOnly
                            Screen.ActiveForm.txtComment.Text
                    Else
                            bCONTINUE = True
                    End If
            
            
            End If
            
            If bmakePAYMENT = True Then
                    If CDbl(Screen.ActiveForm.txtAmountDue) < 0 Then
                            MsgBox "The Amount Due cannot be Zero", vbOKOnly
                            Screen.ActiveForm.txtAmountDue.SetFocus
                    
                    ElseIf CDbl(Screen.ActiveForm.txtAmountPaid) <= 0 Then
                            MsgBox "The Amount Paid cannot be left Blank", vbOKOnly
                            Screen.ActiveForm.txtAmountPaid.SetFocus
                    
                    ElseIf Screen.ActiveForm.txtDocumentNo <= "" Then
                            MsgBox "The Document No is Used to Determine the Payees details", vbOKOnly
                            Screen.ActiveForm.txtDocumentNo.SetFocus
                    
                    ElseIf Screen.ActiveForm.txtVoucherDate <= "" Then
                            MsgBox "The Requisition date cannot be Left Blank", vbOKOnly
                            Screen.ActiveForm.txtVoucherDate.SetFocus
                    Else
                            bCONTINUE = True
                    End If
            End If
            
            If bCONTINUE = False Then Exit Sub
            
            If Screen.ActiveForm.txtChequeDate.Text <= "" Then
                    MsgBox "The Cheque Date cannot be left Blank", vbOKOnly
                    Screen.ActiveForm.txtChequeDate.SetFocus
            
            ElseIf DateDiff("D", Screen.ActiveForm.txtChequeDate.Text, Date) < 0 Then
                    MsgBox "The Cheque Date cannot be in the FUTURE", vbOKOnly
                    Screen.ActiveForm.txtChequeDate.SetFocus

            ElseIf Screen.ActiveForm.txtBankNo.Text <= "" Then
                    MsgBox "The bank No cannot be left Blank", vbOKOnly
                    Screen.ActiveForm.txtBankNo.SetFocus
            
            ElseIf Screen.ActiveForm.txtVoucherNo <= "" Then
                    MsgBox "The Requisition date cannot be Left Blank", vbOKOnly
                    Screen.ActiveForm.txtVoucherNo.SetFocus

            ElseIf Screen.ActiveForm.txtPayeeDetails <= "" Then
                    MsgBox "you must cobfirm the Payee Details prior to Approval", vbOKOnly
                    Screen.ActiveForm.txtPayeeDetails.SetFocus
            
            Else:
                    bsaveRECORD = True
            End If

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub validateCHECKSISSUED()
On Error GoTo err
            
        If Trim(frmODASMCheckIssuance.txtComment.Text) <= "   " Then
                MsgBox "The Comment is Manadatory and MUST be greater than 50 Characters", vbOKOnly
                frmODASMCheckIssuance.txtComment.SetFocus
        
        ElseIf CDbl(frmODASMCheckIssuance.txtChequeAmount.Text) < 0 Then
                MsgBox "The Amount cannot be Zero", vbOKOnly
                frmODASMCheckIssuance.txtChequeAmount.SetFocus
        
        ElseIf frmODASMCheckIssuance.txtDateCollected <= "" Then
                MsgBox "The colelction date cannot be Left Blank", vbOKOnly
                frmODASMCheckIssuance.txtDateCollected.SetFocus
        
        ElseIf DateDiff("D", frmODASMCheckIssuance.txtDateCollected.Text, Date) < 0 Then
                MsgBox "The Collection Date cannot be in the FUTURE", vbOKOnly
                frmODASMCheckIssuance.txtDateCollected.SetFocus

        ElseIf frmODASMCheckIssuance.txtBankNo.Text <= "" Then
                MsgBox "The bank No cannot be left Blank", vbOKOnly
                frmODASMCheckIssuance.txtBankNo.SetFocus
            
        ElseIf frmODASMCheckIssuance.txtChequeNo <= "" Then
                MsgBox "The Cheque No cannot be Left Blank", vbOKOnly
                frmODASMCheckIssuance.txtChequeNo.SetFocus
                
        ElseIf frmODASMCheckIssuance.txtPayeeDetails <= "" Then
                MsgBox "you must cobfirm the Payee Details prior to Approval", vbOKOnly
                frmODASMCheckIssuance.txtPayeeDetails.SetFocus
        Else:
                    bsaveRECORD = True
        End If

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub SaveChequeENTRIES()
On Error GoTo err
    Set rsSAVE = New Recordset
    
    If bmakePAYMENT = True Then
            strSAVE = "SELECT * from ALISMChequeEntry where chequeEntryNo = '" & Screen.ActiveForm.txtChequeEntryNo.Text & "';"
    ElseIf breversePAYMENT = True Then
            strSAVE = "SELECT * from ALISMChequeEntry where chequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "' and (paymentflag = 'Y' or paymentflag = 'P');"
    End If
    
    rsSAVE.Open strSAVE, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsSAVE
            If .EOF Or .BOF Then
                    .AddNew
                    !ChequeNo = Screen.ActiveForm.txtChequeNo.Text
                    !ChequeEntryNo = Screen.ActiveForm.txtChequeEntryNo.Text
                    !vOUCHERnO = Screen.ActiveForm.txtVoucherNo
                    !dateprepared = Date
                    !Preparedby = CurrentUserName
            End If
            
            If bmakePAYMENT = True Then
                    !AmountDue = Screen.ActiveForm.txtAmountDue.Text
                    !Amount = Screen.ActiveForm.txtAmountPaid.Text
                    !Status = Screen.ActiveForm.txtStatus.Text
                    !PaymentFlag = Screen.ActiveForm.txtPaymentFlag.Text
                    !BankNo = Screen.ActiveForm.txtBankNo.Text
                    
            ElseIf breversePAYMENT = True Then
                    !Status = Screen.ActiveForm.txtStatus.Text
                    !PaymentFlag = "V"
            End If
            
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

Public Sub saveCHEQUE()
On Error GoTo err
    
    Set rsSAVE = New ADODB.Recordset
    
    If bmakePAYMENT = True Then
            strSQL = "SELECT * from ALISMCheque where chequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "';"
    ElseIf breversePAYMENT = True Then
            strSQL = "SELECT * from ALISMCheque where chequeNo = '" & frmALISMCheckReverse.txtChequeNo.Text & "' and PaymentFlag = 'Y';"
    End If
    
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsSAVE
            If .EOF Or .BOF Then
                    .AddNew
                    !ChequeNo = Screen.ActiveForm.txtChequeNo.Text
            End If
            
                If bmakePAYMENT = True Then
                        !Amount = Screen.ActiveForm.txtAmountPaid.Text
                        !Status = Screen.ActiveForm.txtStatus.Text
                        !PaymentFlag = "Y"
                        !NoOfEntries = Screen.ActiveForm.txtNoOfEntries.Text
                        !BankNo = Screen.ActiveForm.txtBankNo.Text
                        !ChequeDate = Screen.ActiveForm.txtChequeDate.Text
                        !PayeeDetails = Screen.ActiveForm.txtPayeeDetails
                        !CurrentPeriod = CurrentPeriod
                        !Prepared = "Y"
                        !Issued = "N"
                        !Authorized = "N"
                        !Approved = "N"
                        !Scheduled = "N"
                        !Preparedby = CurrentUserName
                        !dateprepared = Date
                ElseIf breversePAYMENT = True Then
                        !Status = Screen.ActiveForm.cboReversalType.Text
                        !PaymentFlag = "V"
                End If
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

Public Sub saveCHECKSISSUED()
On Error GoTo err
    
    Set rsSAVE = New ADODB.Recordset
    
    strSQL = "SELECT * from ALISMCheque where chequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "' and approved = 'Y';"
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsSAVE
            If .EOF Or .BOF Then Exit Sub
                    !DateIssued = frmODASMCheckIssuance.txtDateCollected.Text
                    !IssuedBy = CurrentUserName
                    !Scheduled = "N"
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
    
    strSQL = "SELECT * from ALISMChecksIssued where chequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "';"
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsSAVE
            If .EOF Or .BOF Then
                    .AddNew
                    !ChequeNo = frmODASMCheckIssuance.txtChequeNo.Text
            End If
                    !DateCollected = frmODASMCheckIssuance.txtDateCollected.Text
                    !CollectedBy = frmODASMCheckIssuance.txtCollectedBy.Text
                    !Amount = frmODASMCheckIssuance.txtChequeAmount.Text
                    !IdentityNo = frmODASMCheckIssuance.txtIdentityNo.Text
                    !IDType = frmODASMCheckIssuance.cboIdType.Text
                    !IssuedBy = CurrentUserName
                    !Collected = "Y"
                    !dateprepared = Date
                    !Preparedby = CurrebtUserName
                    bsaveRECORD = False
                    .Update
                    frmODASMCheckIssuance.txtIssuanceNo.Text = !IssuanceNo
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


Private Sub updateREQUISITION()
On Error GoTo err

        Set rsSAVE = New ADODB.Recordset
        
        rsSAVE.Open "SELECT * FROM ODASMVoucher WHERE VoucherNo = '" & Screen.ActiveForm.txtVoucherNo.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
  
        With rsSAVE
            If .BOF Or .EOF = True Then Exit Sub
            !AmountPaid = Screen.ActiveForm.txtTotalPaid.Text
            !Status = Screen.ActiveForm.txtStatus.Text
            !PaymentFlag = Screen.ActiveForm.txtPaymentFlag.Text
            !ChequeNo = Screen.ActiveForm.txtChequeNo.Text
            !ChequeDate = Screen.ActiveForm.txtChequeDate.Text
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
Public Sub countENTRIES()
On Error GoTo err
    
    Set rsSAVE = New ADODB.Recordset
    
    If bmakePAYMENT = True Then
            strSQL = "SELECT * from ALISMChequeEntry where VoucherNo = '" & Screen.ActiveForm.txtVoucherNo.Text & "' and PaymentFlag <> 'C';"
    Else
            strSQL = "SELECT * from ALISMChequeEntry where VoucherNo = '" & Screen.ActiveForm.txtVoucherNo.Text & "' and PaymentFlag <> 'C';"
    End If
    
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsSAVE
                If .BOF Or .EOF Then Exit Sub
                Screen.ActiveForm.txtNoOfEntries.Text = .RecordCount
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
    
    If bmakePAYMENT = True Then
            strSQL = "SELECT sum(Amount)as TotalPaid from ALISMChequeEntry where VoucherNo = '" & Screen.ActiveForm.txtVoucherNo.Text & "';"
    ElseIf breversePAYMENT = True Then
             strSQL = "SELECT sum(Amount)as TotalPaid from ALISMChequeEntry where VoucherNo = '" & Screen.ActiveForm.txtVoucherNo.Text & "' and PaymentFlag = 'Y';"
    End If
    
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsSAVE
                
                If .EOF Or .BOF Then Exit Sub
                
                
                If IsNull(!TotalPaid) = True Then
                        Screen.ActiveForm.txtTotalPaid.Text = 0
                Else: Screen.ActiveForm.txtTotalPaid.Text = !TotalPaid
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
Public Sub calculateAMOUNTDUE()
On Error GoTo err
    
    Set rsSAVE = New ADODB.Recordset
    
    If bmakePAYMENT = True Then
            strSQL = "SELECT * from ODASMVoucher where VoucherNo = '" & Screen.ActiveForm.txtVoucherNo.Text & "';"
    ElseIf breversePAYMENT = True Then
            strSQL = "SELECT * from ODASMVoucher where VoucherNo = '" & Screen.ActiveForm.txtVoucherNo.Text & "';"
    End If
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsSAVE
                If .EOF Or .BOF Then Exit Sub
                Screen.ActiveForm.txtAmountDue.Text = !Amount - CDbl(Screen.ActiveForm.txtTotalPaid.Text)
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

