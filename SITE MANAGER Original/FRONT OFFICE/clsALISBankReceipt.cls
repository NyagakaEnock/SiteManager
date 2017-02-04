VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsALISBankReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim MYCOUNT As Double
Dim rsREQUISITION As New clsALISCheque

Private Sub clearREQUISITION()
On Error GoTo err
    With Screen.ActiveForm
            .cboBankNo.Text = ""
            .txtChequeNo.Text = ""
            .txtChequeDate.Text = ""
            .txtDocumentNo.Text = ""
            .txtPayeeDetails.Text = ""
            .txtRequisitionDate.Text = ""
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
                    .cboBankNo.Text = ""
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
                strSQL = "Select * from ODASMVoucher where RequisitionNo = '" & Screen.ActiveForm.cboRequisitionNo.Text & "' ;"
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
                        Screen.ActiveForm.txtDocumentNo.Text = !DocumentNo & ""
                        Screen.ActiveForm.txtPayeeDetails.Text = !PayeeDetails & ""
                        Screen.ActiveForm.txtRequisitionDate.Text = !RequisitionDate & ""
                        Screen.ActiveForm.txtAmountPaid.Text = !Amount - !AmountPaid & ""
                        Screen.ActiveForm.txtChequeAmount.Text = CDbl(!Amount) - CDbl(!AmountPaid) & ""
                        Screen.ActiveForm.txtPaymentFlag.Text = "Y"
                        Screen.ActiveForm.txtAmountDue.Text = CDbl(!Amount) - CDbl(!AmountPaid)
                        Screen.ActiveForm.txtChequeDate.Text = Date
                        Screen.ActiveForm.DTPickerChequeDate.Value = Date
                        Screen.ActiveForm.txtStatus.Text = "CHK-PREPARED"
                        Screen.ActiveForm.txtTotalPaid.Text = !AmountPaid
                
                ElseIf breversePAYMENT = True Or bApproveCheque = True Or bAuthorizeCheque = True Then
                        
                        Screen.ActiveForm.cboBankNo.Text = !BankNo
                        Screen.ActiveForm.txtChequeNo.Text = !ChequeNo
                        Screen.ActiveForm.txtNoOfEntries.Text = !NoOfEntries & ""
                        Screen.ActiveForm.txtPayeeDetails.Text = !PayeeDetails & ""
                        Screen.ActiveForm.cboRequisitionNo.Text = !RequisitionNo & ""
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

Public Sub loadBANK()
On Error GoTo err
    
    strSQL = "Select * from ALISPBankAccount Where ALISPBankAccount.BankNo = '" & frmALISMBankReceipt.txtBankNo.Text & "' "
    Set rsCONTROL = New ADODB.Recordset
    
    With rsCONTROL
            .Open strSQL, cnCOMMON, adOpenDynamic, adLockOptimistic
            
            If .BOF Or .EOF Then Exit Sub
                frmALISMBankReceipt.txtDetails.Text = !Details
                frmALISMBankReceipt.txtAccountNo.Text = !Accountno
                frmALISMBankReceipt.txtBankNo.Text = !BankNo
                frmALISMBankReceipt.txtDateBanked.Text = Date
                frmALISMBankReceipt.txtNoOfEntries.Text = 0
                frmALISMBankReceipt.txtTotalAmount.Text = 0
                frmALISMBankReceipt.cboTemplateCode.Text = ""
    End With

strSQL = ""
rsCONTROL.Close

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadSCHEDULEDRECORD()
On Error GoTo err
    
    
    Set rsCONTROL = New ADODB.Recordset
    
    With rsCONTROL
            strSQL = "Select * from ALISMBankAccountCfg, ALISPBankAccount Where ALISMBankAccountCfg.Reference = '" & frmALISMBankReceipt.txtReference.Text & "' and ALISMBankAccountCfg.BankNo = ALISPBankAccount.BankNo "
            .Open strSQL, cnCOMMON, adOpenDynamic, adLockOptimistic
            
            If .BOF Or .EOF Then
                    loadBANK
                    Exit Sub
            End If
            
                    frmALISMBankReceipt.txtDetails.Text = !Details
                    frmALISMBankReceipt.txtAccountNo.Text = !Accountno
                    frmALISMBankReceipt.txtBankNo.Text = !BankNo
                    frmALISMBankReceipt.txtDateBanked.Text = !DateBanked
                    frmALISMBankReceipt.txtNoOfEntries.Text = !NoOfEntries
                    frmALISMBankReceipt.txtTotalAmount.Text = !TotalAmount
                    frmALISMBankReceipt.cboTemplateCode.Text = !templateCode
        
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
                        .txtRequisitionDate.Locked = True
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
                        .txtRequisitionDate.Locked = True
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
        
        rsCOSTCENTER.Open "SELECT * FROM ODASMVoucher, ODASPCostCentre WHERE ODASMVoucher.RequisitionNo = '" & Screen.ActiveForm.cboRequisition.Text & "' and ODASMVoucher.CostCenter = ODASPCostCentre.CostCenter ", cnCOMMON, adOpenKeyset, adLockOptimistic
        
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
Public Sub UpdateALLRECORDS()
        bsaveRECORD = False
            validateRECORD

        If bsaveRECORD = True Then
                If bmakePAYMENT = True Then
                        GenerateChequeNo
                        SaveChequeENTRIES
                End If
                
                If breversePAYMENT = True Then
                       reverseChequeENTRIES
                       voidCHEQUEENTRY
                End If
                
                countENTRIES
                calculateTOTALPAID
                calculateAMOUNTDUE
                saveCHEQUE
                
                If breversePAYMENT = True Then
                    voidCHEQUE
                End If
                
                updateREQUISITION
                GetPreviousPayment
                enableButtons
                disableALLRECORD
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
                    If frmALISMBankReceipt.ListView1.ListItems(i).Checked = True Then
                        strCheckNo = Screen.ActiveForm.ListView1.ListItems(i).Text
                        
                        saveSCHEDULERECORD
                        saveRECEIPT
                    
                    End If
                    
                    strCheckNo = ""
            Next i
            
            

End Sub


Public Sub GetBankedReceipts()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView2.ListItems.Clear
                .ListView2.ColumnHeaders.Clear
                .ListView2.ColumnHeaders.Add , , "Receipt No", .ListView2.Width / 7
                .ListView2.ColumnHeaders.Add , , "Payer Details", .ListView2.Width / 4
                .ListView2.ColumnHeaders.Add , , "Method", .ListView1.Width / 8
                .ListView2.ColumnHeaders.Add , , "Receipt Date", .ListView2.Width / 5
                .ListView2.ColumnHeaders.Add , , "Amount", .ListView2.Width / 5
                .ListView2.ColumnHeaders.Add , , "Status", .ListView2.Width / 8
                .ListView2.ColumnHeaders.Add , , "Banked", .ListView2.Width / 9
                .ListView2.ColumnHeaders.Add , , "Date Banked", .ListView2.Width / 7
                .ListView2.ColumnHeaders.Add , , "Reference", .ListView2.Width / 7
                .ListView2.ColumnHeaders.Add , , "AccountNo", .ListView2.Width / 7


                .ListView2.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select ALISMReceiptNew.txtReceiptNo, ALISMReceiptNew.Payer, ALISMReceiptNew.PaymentMethod, ALISMReceiptNew.ReceiptDate, ALISMReceiptNew.ReceiptAmount, ALISMReceiptNew.PaymentStatus, Banked, BankingDate, Reference, BankAccount from ALISMReceiptNew Where ALISMReceiptNew.Banked = 'Y'"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!ReceiptNo))
                        
                        If Not IsNull(rsLIST!Payer) Then
                                MyList.SubItems(1) = CStr(rsLIST!Payer)
                        End If

                        If Not IsNull(rsLIST!PaymentMethod) Then
                            MyList.SubItems(2) = CStr(rsLIST!PaymentMethod)
                        End If

                        If Not IsNull(rsLIST!ReceiptDate) Then
                            MyList.SubItems(3) = CStr(rsLIST!ReceiptDate)
                        End If

                        If Not IsNull(rsLIST!ReceiptAmount) Then
                                MyList.SubItems(4) = CStr(rsLIST!ReceiptAmount)
                        End If
                        
                        If Not IsNull(rsLIST!PaymentStatus) Then
                                MyList.SubItems(5) = CStr(rsLIST!PaymentStatus)
                        End If
                        
                        If Not IsNull(rsLIST!Banked) Then
                                MyList.SubItems(6) = CStr(rsLIST!Banked)
                        End If
                        
                        If Not IsNull(rsLIST!BankingDate) Then
                                MyList.SubItems(7) = CStr(rsLIST!BankingDate)
                        End If
                        
                        If Not IsNull(rsLIST!Reference) Then
                                MyList.SubItems(8) = CStr(rsLIST!Reference)
                        End If
                        
                        If Not IsNull(rsLIST!BankAccount) Then
                                MyList.SubItems(9) = CStr(rsLIST!BankAccount)
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
                
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
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


Public Sub getRECEIPT()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Receipt No", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Payer Details", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Method", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Receipt Date", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Receipt Amount", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Payment Flag", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Banked", .ListView1.Width / 9
                .ListView1.ColumnHeaders.Add , , "Date Banked", .ListView1.Width / 8


                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select ALISMReceiptNew.txtReceiptNo, ALISMReceiptNew.Payer, ALISMReceiptNew.PaymentMethod, ALISMReceiptNew.ReceiptDate, ALISMReceiptNew.ReceiptAmount, ALISMReceiptNew.PaymentStatus, Banked, BankingDate from ALISMReceiptNew Where (ALISMReceiptNew.Banked = 'N' or ALISMReceiptNew.Banked is null) and paymentStatus = 'PAID'"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ReceiptNo))
                        
                        If Not IsNull(rsLIST!Payer) Then
                                MyList.SubItems(1) = CStr(rsLIST!Payer)
                        End If
                                                
                        If Not IsNull(rsLIST!PaymentMethod) Then
                                MyList.SubItems(2) = CStr(rsLIST!PaymentMethod)
                        End If
                        
                        If Not IsNull(rsLIST!ReceiptDate) Then
                                MyList.SubItems(3) = CStr(rsLIST!ReceiptDate)
                        End If

                        If Not IsNull(rsLIST!ReceiptAmount) Then
                                MyList.SubItems(4) = CStr(rsLIST!ReceiptAmount)
                        End If
                        
                                
                        If Not IsNull(rsLIST!PaymentStatus) Then
                                MyList.SubItems(5) = CStr(rsLIST!PaymentStatus)
                        End If
                        
                        If Not IsNull(rsLIST!Banked) Then
                                MyList.SubItems(6) = CStr(rsLIST!Banked)
                        End If
                        
                        If Not IsNull(rsLIST!BankingDate) Then
                                MyList.SubItems(7) = CStr(rsLIST!BankingDate)
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

        With frmALISMBankReceipt
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
                
                df = rsLIST.RecordCount
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


Public Sub GetPreviousPayment()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView3.ListItems.Clear
                .ListView3.ColumnHeaders.Clear
                
                If bmakePAYMENT = True Then
                        .ListView3.ColumnHeaders.Add , , "Req No", .ListView3.Width / 6
                        .ListView3.ColumnHeaders.Add , , "Payee Details", .ListView3.Width / 4
                ElseIf breversePAYMENT = True Then
                        .ListView3.ColumnHeaders.Add , , "Cheque No", .ListView3.Width / 6
                        .ListView3.ColumnHeaders.Add , , "RequsitionNo", .ListView3.Width / 4
                End If
                
                .ListView3.ColumnHeaders.Add , , "Amount", .ListView3.Width / 6
                .ListView3.ColumnHeaders.Add , , "Amount Due", .ListView3.Width / 6
                .ListView3.ColumnHeaders.Add , , "Payment Flag", .ListView3.Width / 8


                .ListView3.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                If bmakePAYMENT = True Then
                        strSQL = "Select ALISMChequeEntry.RequisitionNo, ALISMCheque.PayeeDetails, ALISMChequeEntry.Amount, ALISMChequeEntry.AmountDue, ALISMChequeEntry.PaymentFlag from ALISMChequeEntry, ALISMCheque Where ALISMCheque.ChequeNo = ALISMChequeEntry.ChequeNo and ALISMChequeEntry.RequisitionNo = '" & Screen.ActiveForm.cboRequisitionNo.Text & "' and ALISMChequeEntry.PaymentFlag <> 'C';"
                
                ElseIf breversePAYMENT = True Then
                         strSQL = "Select ALISMChequeEntry.ChequeNo, ALISMChequeEntry.RequisitionNo, ALISMChequeEntry.Amount, ALISMChequeEntry.AmountDue, ALISMChequeEntry.PaymentFlag from ALISMChequeEntry Where ALISMChequeEntry.ChequeNo = '" & frmALISMCheckReverse.txtChequeNo & "' ;"
                End If
                
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        If bmakePAYMENT = True Then
                                Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!RequisitionNo))
                        ElseIf breversePAYMENT = True Then
                                Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!ChequeNo))
                        End If
                        
                        If Not IsNull(rsLIST!Amount) Then
                                MyList.SubItems(2) = CStr(rsLIST!Amount)
                        End If
                        
                        If Not IsNull(rsLIST!AmountDue) Then
                                MyList.SubItems(3) = CStr(rsLIST!AmountDue)
                        End If
                        
                        If Not IsNull(rsLIST!PaymentFlag) Then
                                MyList.SubItems(4) = CStr(rsLIST!PaymentFlag)
                        End If
                        
                        If bmakePAYMENT = True Then
                                If Not IsNull(rsLIST!PayeeDetails) Then
                                        MyList.SubItems(1) = CStr(rsLIST!PayeeDetails)
                                End If
                        ElseIf breversePAYMENT = True Then
                                If Not IsNull(rsLIST!RequisitionNo) Then
                                        MyList.SubItems(1) = CStr(rsLIST!RequisitionNo)
                                End If
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
Public Sub GetOtherPayment()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView3.ListItems.Clear
                .ListView3.ColumnHeaders.Clear
                
                .ListView3.ColumnHeaders.Add , , "Cheque No", .ListView3.Width / 6
                .ListView3.ColumnHeaders.Add , , "Payee Details", .ListView3.Width / 4
                .ListView3.ColumnHeaders.Add , , "Amount", .ListView3.Width / 6
                .ListView3.ColumnHeaders.Add , , "Amount Due", .ListView3.Width / 6
                .ListView3.ColumnHeaders.Add , , "Payment Flag", .ListView3.Width / 8
                .ListView3.ColumnHeaders.Add , , "Requisition No", .ListView3.Width / 8


                .ListView3.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select ALISMChequeEntry.ChequeNo, ALISMCheque.PayeeDetails, ALISMChequeEntry.Amount, ALISMChequeEntry.AmountDue, ALISMChequeEntry.PaymentFlag, ALISMChequeEntry.RequisitionNo from ALISMChequeEntry, ALISMCheque Where ALISMCheque.ChequeNo = ALISMChequeEntry.ChequeNo and ALISMChequeEntry.chequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "' and ALISMChequeEntry.PaymentFlag = 'Y';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!ChequeNo))
                        
                        If Not IsNull(rsLIST!Amount) Then
                                MyList.SubItems(2) = CStr(rsLIST!Amount)
                        End If
                        
                        If Not IsNull(rsLIST!AmountDue) Then
                                MyList.SubItems(3) = CStr(rsLIST!AmountDue)
                        End If
                        
                        If Not IsNull(rsLIST!PaymentFlag) Then
                                MyList.SubItems(4) = CStr(rsLIST!PaymentFlag)
                        End If
                        
                        If Not IsNull(rsLIST!PayeeDetails) Then
                                MyList.SubItems(1) = CStr(rsLIST!PayeeDetails)
                        End If
                        
                        If Not IsNull(rsLIST!RequisitionNo) Then
                                MyList.SubItems(5) = CStr(rsLIST!RequisitionNo)
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

Private Sub cbobankNo_GotFocus()
On Error GoTo err
        strSQL = "SELECT * FROM ALISPbankAccount;"
        bankNoGotFocus
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cbobankNo_LostFocus()
    strSQL = "Select * from ALISPBankAccount  "
    BankNoLostFocus
    'loadBANK
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
                    
                    ElseIf Screen.ActiveForm.txtRequisitionDate <= "" Then
                            MsgBox "The Requisition date cannot be Left Blank", vbOKOnly
                            Screen.ActiveForm.txtRequisitionDate.SetFocus
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

            ElseIf Screen.ActiveForm.cboBankNo.Text <= "" Then
                    MsgBox "The bank No cannot be left Blank", vbOKOnly
                    Screen.ActiveForm.cboBankNo.SetFocus
            
            ElseIf Screen.ActiveForm.cboRequisitionNo <= "" Then
                    MsgBox "The Requisition date cannot be Left Blank", vbOKOnly
                    Screen.ActiveForm.cboRequisitionNo.SetFocus

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
            strSAVE = "SELECT * from ALISMChequeEntry where chequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "' and requisitionNo = '" & Screen.ActiveForm.cboRequisitionNo.Text & "';"
    ElseIf breversePAYMENT = True Then
            strSAVE = "SELECT * from ALISMChequeEntry where chequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "' and (paymentflag = 'Y' or paymentflag = 'P');"
    End If
    
    rsSAVE.Open strSAVE, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsSAVE
            If .EOF Or .BOF Then
                    .AddNew
                    !ChequeNo = Screen.ActiveForm.txtChequeNo.Text
                    !RequisitionNo = Screen.ActiveForm.cboRequisitionNo
            End If
            
            If bmakePAYMENT = True Then
                    !AmountDue = Screen.ActiveForm.txtAmountDue.Text
                    !Amount = Screen.ActiveForm.txtAmountPaid.Text
                    !Status = Screen.ActiveForm.txtStatus.Text
                    !PaymentFlag = Screen.ActiveForm.txtPaymentFlag.Text
                    !Preparedby = CurrentUserName
                    !BankNo = Screen.ActiveForm.cboBankNo.Text
                    !dateprepared = Date
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
                    rsSAVE!RequisitionNo = rsCONTROL!RequisitionNo
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
            strSAVE = "SELECT * from ALISMChequeEntry where chequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "' and requisitionNo = '" & Screen.ActiveForm.cboRequisitionNo.Text & "';"
    ElseIf breversePAYMENT = True Then
            strSAVE = "SELECT * from ALISMChequeEntry where chequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "' and requisitionNo = '" & Screen.ActiveForm.cboRequisitionNo.Text & "' and (paymentflag = 'Y' or paymentflag = 'P');"
    End If
    
    rsSAVE.Open strSAVE, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsSAVE
            .AddNew
            !AmountDue = Screen.ActiveForm.txtAmountDue.Text
            !Amount = CDbl(Screen.ActiveForm.txtAmountPaid.Text) * "Y"
            !RequisitionNo = Screen.ActiveForm.cboRequisitionNo
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
                        !BankNo = Screen.ActiveForm.cboBankNo.Text
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
                !BankNo = Screen.ActiveForm.cboBankNo.Text
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
        
        rsSAVE.Open "SELECT * FROM ODASMVoucher WHERE RequisitionNo = '" & Screen.ActiveForm.cboRequisitionNo.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
  
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
            strSQL = "SELECT * from ALISMChequeEntry where RequisitionNo = '" & Screen.ActiveForm.cboRequisitionNo.Text & "' and PaymentFlag <> 'C';"
    Else
            strSQL = "SELECT * from ALISMChequeEntry where RequisitionNo = '" & Screen.ActiveForm.cboRequisitionNo.Text & "' and PaymentFlag <> 'C';"
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
            strSQL = "SELECT sum(Amount)as TotalPaid from ALISMChequeEntry where RequisitionNo = '" & Screen.ActiveForm.cboRequisitionNo.Text & "';"
    ElseIf breversePAYMENT = True Then
             strSQL = "SELECT sum(Amount)as TotalPaid from ALISMChequeEntry where RequisitionNo = '" & Screen.ActiveForm.cboRequisitionNo.Text & "' and PaymentFlag = 'Y';"
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
            strSQL = "SELECT * from ODASMVoucher where RequisitionNo = '" & Screen.ActiveForm.cboRequisitionNo.Text & "';"
    ElseIf breversePAYMENT = True Then
            strSQL = "SELECT * from ODASMVoucher where RequisitionNo = '" & Screen.ActiveForm.cboRequisitionNo.Text & "';"
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
    strSQL = "SELECT * FROM ALISMBankAccountCfg WHERE Reference =  '" & frmALISMBankReceipt.txtReference & " ' ;"
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                      
    With rsSAVE
            If .BOF Or .EOF Then
                    .AddNew
                    !Reference = frmALISMBankReceipt.txtReference.Text
             End If
             
             !BankNo = frmALISMBankReceipt.txtBankNo.Text
             !dateprepared = Date
             !Preparedby = CurrentUserName
             !Accountno = frmALISMBankReceipt.txtAccountNo.Text
             !ReceivedPayed = "Paid"
             !CurrentPeriod = CurrentPeriod
             !templateCode = frmALISMBankReceipt.cboTemplateCode.Text
             !DateBanked = frmALISMBankReceipt.txtDateBanked.Text
             !TotalAmount = CDbl(frmALISMBankReceipt.txtTotalAmount.Text)
             !NoOfEntries = CDbl(frmALISMBankReceipt.txtNoOfEntries.Text)
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
        With frmALISMBankReceipt
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
        With frmALISMBankReceipt
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
   With frmALISMBankReceipt
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
                    EnableControls
                    disableButtons
                    strSQL = ""
End Sub

Public Sub cancelSCHEDULE()
        enableButtons
        ClearControls
        disableControls
End Sub
Public Sub generateReference()
On Error GoTo err
        Dim rsLAST As ADODB.Recordset, strLAST As String
        
        Set rsLAST = New Recordset
      
        strLAST = "SELECT * FROM ALISPBankAccount Where BankNo = '" & Screen.ActiveForm.txtBankNo.Text & "' ;"
        rsLAST.Open strLAST, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsLAST
                If .BOF Or .EOF Then Exit Sub
                
                frmALISMBankReceipt.txtReference = !Reference
                
                Select Case Len(Trim(frmALISMBankReceipt.txtReference))
                    Case 1: frmALISMBankReceipt.txtReference = Trim(!ReferencePrefix) + "/00000" + Trim(frmALISMBankReceipt.txtReference)
                    Case 2: frmALISMBankReceipt.txtReference = Trim(!ReferencePrefix) + "/0000" + Trim(frmALISMBankReceipt.txtReference)
                    Case 3: frmALISMBankReceipt.txtReference = Trim(!ReferencePrefix) + "/000" + Trim(frmALISMBankReceipt.txtReference)
                    Case 4: frmALISMBankReceipt.txtReference = Trim(!ReferencePrefix) + "/00" + Trim(frmALISMBankReceipt.txtReference)
                    Case 5: frmALISMBankReceipt.txtReference = Trim(!ReferencePrefix) + "/0" + Trim(frmALISMBankReceipt.txtReference)
                    Case 6: frmALISMBankReceipt.txtReference = Trim(!ReferencePrefix) + "/" + Trim(frmALISMBankReceipt.txtReference)
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

Public Sub updateSCHEDULE()
        bsaveRECORD = False
        ValidateSchedule
        
            If bsaveRECORD = True Then
                    generateReference
            GetCheckedSCHEDULE
            countSCHEDULEENTRIES
            calculateTOTAL
            saveREFERENCE
            GetBankedReceipts
            getRECEIPT

            bsaveRECORD = False
        End If
        
        enableSButtons
        disableALLRECORD
        getRECEIPT
        GetBankedReceipts
        strSQL = ""


End Sub

Private Sub saveSCHEDULERECORD()
On Error GoTo err
        
        Dim rsSAVE As ADODB.Recordset
        Set rsSAVE = New ADODB.Recordset
        
        strSQL = ""
        strSQL = "Select * from ALISMChequeSchedule where Reference = '" & frmALISMBankReceipt.txtReference & "' and chequeNo = '" & strCheckNo & "'  ;"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                        
        With rsSAVE
                                
                    If .BOF Or .EOF = True Then
                        .AddNew
                        !Reference = frmALISMBankReceipt.txtReference.Text
                        !ChequeNo = strCheckNo
                    End If
                    
                        !BankNo = frmALISMBankReceipt.txtBankNo.Text
                        !Amount = 0
                        !CurrentPeriod = CurrentPeriod
                        !Status = "RCPT-BANKED"
                        !Banked = "Y"
                        !Cleared = "N"
                        !Scheduled = "N"
                        !DateBanked = Date
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
        
        strSQL = "Select * from ALISMChequeSchedule where Reference = '" & frmALISMBankReceipt.txtReference & "' ;"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                        
        With rsSAVE
                    If .EOF Or .BOF Then Exit Sub
                    
                    frmALISMBankReceipt.txtNoOfEntries.Text = .RecordCount
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
        
        strSQL = "Select sum(ReceiptAmount)as Totals from ALISMReceiptNew where Reference = '" & frmALISMBankReceipt.txtReference & "' ;"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                        
        With rsSAVE
                    If .EOF Or .BOF Then Exit Sub
                    
                    If IsNull(!TOTALS) = True Then !TOTALS = 0
                    frmALISMBankReceipt.txtTotalAmount.Text = !TOTALS
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


Private Sub saveRECEIPT()
On Error GoTo err
        
        Set rsSAVE = New ADODB.Recordset
        
        strSQL = ""
        
        strSQL = "Select * from ALISMReceiptNew where (Banked = 'N'or Banked is Null) and ReceiptNo = '" & strCheckNo & "'  ;"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                        
        With rsSAVE
                                
                  If .BOF Or .EOF = True Then Exit Sub
                              
                        !Reference = frmALISMBankReceipt.txtReference.Text
                        !BankAccount = frmALISMBankReceipt.txtAccountNo.Text
                        !CoyBankNo = frmALISMBankReceipt.txtBankNo.Text
                        !Banked = "Y"
                        !BankingDate = Date
                        !BankedBy = CurrentUserName
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

