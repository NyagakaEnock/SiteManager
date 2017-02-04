VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPaymentRequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public rsCOSTCENTER As ADODB.Recordset, strCostCenterlf As String
Public rsClaimApproval As clsALISApproval

Public Sub selectCostCenterGotFocus()
On Error GoTo err

          Set rsCONTROL = New ADODB.Recordset
          
          strSQL = "SELECT * FROM ODASPCostCentre ;"
          rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

          frmODASMVoucher.txtCostCenter.Clear

          With rsCONTROL
                  If .EOF Or .BOF Then Exit Sub
                  Do Until .EOF
                      frmODASMVoucher.txtCostCenter.AddItem !Description
                      .MoveNext
                  Loop
          End With
            
rsCONTROL.Close
strSQL = ""
         
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub obtainMEDICALFEE()
On Error GoTo err

            Dim rsclaimno As ADODB.Recordset, strCLAIMNO As String
            Set rsclaimno = New ADODB.Recordset
            
            strCLAIMNO = " select sum(invoiceAmount) as Totals from ALISMMedicalReport where proposalno = '" & Trim(frmODASMVoucher.cboDocumentNo) & "' and DoctorCode = '" & frmODASMVoucher.cboPaymentCode & "' and resultsreceived = 'Y' and RequisitionPrepared = 'N'; "
            rsclaimno.Open strCLAIMNO, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            With rsclaimno
            
                    If .BOF Or .EOF Then Exit Sub
                    frmODASMVoucher.txtAmountPaid.Text = rsclaimno!Totals
'                    frmODASMVoucher.txtTotalPayments.Text = rsclaimno!TOTALS
'                    frmODASMVoucher.txtTotalDeductions.Text = 0
                    frmODASMVoucher.txtStatus = "REQ-PREP"
                    frmODASMVoucher.txtReference = Trim(frmODASMVoucher.txtReference) + "- " + Trim(frmODASMVoucher.cboDocumentNo.Text)
                            End With
rsclaimno.Close
strCLAIMNO = ""

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub loadDOCTORSDETAILS()
On Error GoTo err

            'Obtain Doctors Details
            
            Dim rsDOCTOR As ADODB.Recordset, strDOCTOR As String
            Set rsDOCTOR = New ADODB.Recordset
            
            rsDOCTOR.Open "SELECT * FROM ALISPDOCTOR WHERE DoctorNo = '" & frmODASMVoucher.cboPaymentCode & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
            
            With rsDOCTOR
                    
                    If .BOF Or .EOF Then Exit Sub
                        frmODASMVoucher.txtPayeeDetails = Trim(!othernames) + " " + Trim(!CompanyName)
                                              
            End With
                                            '/ End with Policy

rsDOCTOR.Close
 
Exit Sub

err:
    ErrorMessage
            
End Sub


Public Sub GenerateVoucherNo()
On Error GoTo err

        Set rsCONTROL = New Recordset
        
        strSQL = "SELECT * FROM ODASPLastNumbers WHERE AutoVoucherNo = 'Y'"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
        With rsCONTROL
                If .BOF Or .EOF = True Then Exit Sub
                        Screen.ActiveForm.txtVoucherNo.Text = !vOUCHERnO & ""
                
                Select Case Len(Trim(frmODASMVoucher.txtVoucherNo))
                        Case 1: frmODASMVoucher.txtVoucherNo.Text = Trim(!VoucherPrefix) + "00000" + Trim(frmODASMVoucher.txtVoucherNo)
                        Case 2: frmODASMVoucher.txtVoucherNo.Text = Trim(!VoucherPrefix) + "0000" + Trim(frmODASMVoucher.txtVoucherNo)
                        Case 3: frmODASMVoucher.txtVoucherNo.Text = Trim(!VoucherPrefix) + "000" + Trim(frmODASMVoucher.txtVoucherNo)
                        Case 4: frmODASMVoucher.txtVoucherNo.Text = Trim(!VoucherPrefix) + "00" + Trim(frmODASMVoucher.txtVoucherNo)
                        Case 5: frmODASMVoucher.txtVoucherNo.Text = Trim(!VoucherPrefix) + "0" + Trim(frmODASMVoucher.txtVoucherNo)
                        Case 6: frmODASMVoucher.txtVoucherNo.Text = Trim(!VoucherPrefix) + Trim(frmODASMVoucher.txtVoucherNo)
                End Select
                
                !vOUCHERnO = !vOUCHERnO + 1
                .Update
                .Requery

        End With
        
rsCONTROL.Close
strTRANS = ""

Exit Sub
err:
    ErrorMessage
End Sub
 
Public Sub LockedVoucherNo()
On Error GoTo err

        Dim rsLAST As ADODB.Recordset, strLAST As String
        
        Set rsLAST = New Recordset
      
        strLAST = "SELECT * FROM ALISPLastNumber;"
        rsLAST.Open strLAST, cnCOMMON, adOpenKeyset, adLockOptimistic

        If rsLAST!AutoVoucherNo = "Yes" Then
                txtVoucherNo.Locked = True
        End If
        
        
Exit Sub

err:
        UpdateErrorMessage
End Sub


Public Sub EnableCONTROL()
On Error GoTo err

        With frmODASMVoucher
            .txtInvoiceAmount.Locked = False
            .txtRequisitionDate.Locked = False
            .txtVoucherNo.Locked = False
            .txtStatus.Locked = False
            .cboPaymentCode.Locked = False
            .cboDocumentNo.Locked = False
            .txtReference.Locked = False
            .txtPayeeDetails.Locked = False
            .txtCostCenter.Locked = False
            .txtPaymentDescription.Locked = False
'            .txtTotalDeductions.Locked = True
'            .txtTotalPayments.Locked = True
        End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub disableCONTROL()
On Error GoTo err

        With frmODASMVoucher
            .txtInvoiceAmount.Locked = True
            .txtRequisitionDate.Locked = True
            .txtVoucherNo.Locked = True
            .txtStatus.Locked = True
            .cboPaymentCode.Locked = True
            .cboDocumentNo.Locked = True
            .txtReference.Locked = True
            .txtPayeeDetails.Locked = True
            
           .txtCostCenter.Locked = True
            .txtPaymentDescription.Locked = True
'            .txtTotalDeductions.Locked = True
'            .txtTotalPayments.Locked = True
        End With

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub loadRECORD()
On Error GoTo err

        With RsCode
            frmODASMVoucher.txtInvoiceAmount.Text = !Amount
            frmODASMVoucher.txtRequisitionDate.Text = !RequisitionDate
            frmODASMVoucher.txtVoucherNo.Text = !vOUCHERnO
            frmODASMVoucher.txtStatus.Text = !Status
            frmODASMVoucher.cboPaymentCode.Text = !PaymentCode
            frmODASMVoucher.cboDocumentNo.Text = !DocumentNo
            frmODASMVoucher.txtReference.Text = !Reference
            frmODASMVoucher.txtPayeeDetails.Text = !PayeeDetails
            frmODASMVoucher.txtCostCenter = !CostCenter

        End With
  
        loadPaymentDescription
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub GenerateVoucherItem()
On Error GoTo err
        With frmODASMVoucher
                Set rsSAVE = New ADODB.Recordset
                
                strSQL = "SELECT * from ODASMVoucher where VoucherNo = '" & .txtVoucherNo.Text & "';"
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        .txtItems.Text = 1
                ElseIf IsNull(rsSAVE!Items) = True Then
                        .txtItems.Text = 1
                Else
                        .txtItems.Text = CDbl(rsSAVE!Items) + 1
                End If
                
                .txtVoucherItemNo.Text = Trim(.txtVoucherNo) + "-" + Trim(.txtItems.Text)
           
                
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub calculateVOUCHERSUM()
On Error GoTo err
        With frmODASMVoucher
                Set rsSAVE = New ADODB.Recordset
                
                strSQL = "SELECT sum(Amount) As VoucherValue from ODASMVoucherItem where VoucherNo = '" & frmODASMVoucher.txtVoucherNo.Text & "';"
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                  
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        .txtTotalVoucherAmount.Text = 0 + CDbl(.txtVoucherAmount.Text)
                ElseIf IsNull(rsSAVE!voucherValue) = True Then
                        If .txtVoucherAmount.Text = Empty Then .txtVoucherAmount.Text = 0
                        .txtTotalVoucherAmount.Text = 0 + CDbl(.txtVoucherAmount.Text)
                Else
                        .txtTotalVoucherAmount.Text = CDbl(.txtVoucherAmount.Text) + CDbl(rsSAVE!voucherValue) + 1
                End If
                            
                .txtInvoiceBalance.Text = FormatNumber(CDbl(.txtInvoiceAmount) - CDbl(.txtTotalVoucherAmount.Text))
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub calculateReqTotals()
On Error GoTo err
        With frmODASMVoucher
                Set rsSAVE = New ADODB.Recordset
                
                strSQL = "SELECT sum(AllowanceTotals) As VoucherValue from ODASMRequisitionItems where RequisitionNo = '" & frmODASMVoucher.txtReqNo.Text & "' and issued='Y' and Request='Y';"
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                  
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        .txtTotalVoucherAmount.Text = 0
                ElseIf IsNull(rsSAVE!voucherValue) = True Then
                        .txtTotalVoucherAmount.Text = 0
                Else
                        .txtTotalVoucherAmount.Text = CDbl(rsSAVE!voucherValue)
                End If
                            
                  End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub calculateINVOICEBALANCE()
On Error GoTo err

        With frmODASMVoucher
                Set rsSAVE = New ADODB.Recordset
                
                strSQL = "SELECT sum(Amount) As InvoiceBALANCE from ODASMVoucherItem where DocumentNo = '" & frmODASMVoucher.cboDocumentNo.Text & "';"
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        .txtTotalVoucherAmount.Text = 0 + CDbl(.txtVoucherAmount.Text)
                ElseIf IsNull(rsSAVE!InvoiceBalance) = True Then
                        .txtTotalVoucherAmount.Text = 0 + CDbl(.txtVoucherAmount.Text)
                Else
                        .txtTotalVoucherAmount.Text = CDbl(.txtVoucherAmount.Text) + CDbl(rsSAVE!InvoiceBalance) + 1
                End If
        End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub saveRecord()
On Error GoTo err
            
            Dim rsSAVE As ADODB.Recordset
            Set rsSAVE = New ADODB.Recordset
            strSQL = ""
            strSQL = "SELECT * from ODASMVoucher where VoucherNo = '" & frmODASMVoucher.txtVoucherNo.Text & "';"
            rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            With rsSAVE
                If .EOF Or .BOF Then
                        .AddNew
                        !vOUCHERnO = frmODASMVoucher.txtVoucherNo
                        !PaymentCode = frmODASMVoucher.cboPaymentCode.Text
                        !CostCenter = frmODASMVoucher.txtCostCenter
                        !VoucherDate = frmODASMVoucher.txtRequisitionDate
                        !AccountNo = frmODASMVoucher.txtAccountNo.Text
                        !dateprepared = Date
                        !Preparedby = CurrentUserName
                        !Prepared = "Y"
                        !ChequePrepared = "N"
                        !Authorized = "N"
                        !Items = 0
                        !CurrentPeriod = CurrentPeriod
                    LoadDEFAULT
                        If rsDEFAULT!VoucherApproval = "Y" Then
                                !ApprovedBy = CurrentUserName
                                !DateApproved = Date
                                !Status = "VCH-APPROVED"
                                !Approved = "Y"
                        Else:
                                !Approved = "N"
                                !Status = frmODASMVoucher.txtStatus
                        End If
                End If
                !Amount = CDbl(frmODASMVoucher.txtVoucherAmount.Text)
                !Items = frmODASMVoucher.txtItems.Text
                !Reference = frmODASMVoucher.txtReference.Text
                !remark = frmODASMVoucher.txtPayeeDetails
                 If frmODASMVoucher.cboPaymentCode.Text = "OTP" Then
                 Else:
                 End If
                .Update
                .Requery
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
Private Sub saveVOUCHERITEMS()
On Error GoTo err
            
            Dim rsSAVE As ADODB.Recordset
            Set rsSAVE = New ADODB.Recordset
            strSQL = ""
            strSQL = "SELECT * from ODASMVOUCHERITEM where VoucherItemNo = '" & frmODASMVoucher.txtVoucherItemNo.Text & "';"
            rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

            With rsSAVE
                If .EOF Or .BOF Then
                        .AddNew
                        !vOUCHERnO = frmODASMVoucher.txtVoucherNo
                        !VoucherItemNo = frmODASMVoucher.txtVoucherItemNo
                        !Preparedby = CurrentUserName
                        !dateprepared = Date
                        !LPONo = frmODASMVoucher.txtLPONo.Text
                        !DocumentNo = frmODASMVoucher.cboDocumentNo
                       
                End If
                
                frmODASMVoucher.txtInvoiceBalance = CDbl(frmODASMVoucher.txtInvoiceAmount) - CDbl(PartialPaid)
                If CDbl(frmODASMVoucher.txtInvoiceBalance) = 0 Then
                    !PaymentFlag = "Y"
                ElseIf CDbl(frmODASMVoucher.txtInvoiceBalance) = CDbl(frmODASMVoucher.txtInvoiceAmount) Then
                    !PaymentFlag = "N"
                Else:
                    !PaymentFlag = "P"
                End If
                
                !InvoiceAmount = CDbl(frmODASMVoucher.txtInvoiceAmount)
                !Balance = CDbl(frmODASMVoucher.txtInvoiceBalance.Text)
                !Remarks = frmODASMVoucher.txtRemark.Text
                !Status = frmODASMVoucher.txtStatus
                If frmODASMVoucher.cboPaymentCode.Text = "OTP" Then
                   !ItemName = frmODASMVoucher.txtJobDetails.Text + "-" + "Allowances"
                   !AmountPaid = CDbl(frmODASMVoucher.txtTotalVoucherAmount.Text)
                 Else:
                   !ItemName = frmODASMVoucher.txtReference.Text
                   If PartialPaid = 0 Then
                   !AmountPaid = CDbl(frmODASMVoucher.txtAmountPaid.Text)
                   Else
                    !AmountPaid = PartialPaid
                   End If
                   '!ExpiryPeriod = frmODASMVoucher.txtExpiryDate.Text
                   '!InstallmentNo = frmODASMVoucher.cboDocumentNo.Text
                 End If
                .Update
                .Requery
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


Private Sub PaymentCodeGotFocusClaim()
On Error GoTo err
          Dim rsPaymentCodeGF As ADODB.Recordset, strPaymentCodeGF As String
          Set rsPaymentCodeGF = New ADODB.Recordset
          
          strPaymentCodeGF = "SELECT * FROM ODASPPaymentCode ;"
          rsPaymentCodeGF.Open strPaymentCodeGF, cnCOMMON, adOpenKeyset, adLockOptimistic

          frmODASMVoucher.cboPaymentCode.Clear

          With rsPaymentCodeGF
                  If .EOF Or .BOF Then Exit Sub
                  Do Until .EOF
                      frmODASMVoucher.cboPaymentCode.AddItem !PaymentCodeDescription
                      .MoveNext
                  Loop
          End With
            
rsPaymentCodeGF.Close
strPaymentCodeGF = ""

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub PaymentCodeGotFocusLoan()
On Error GoTo err

        Dim rsLOANTYPE As ADODB.Recordset, strLOANTYPE As String
        Set rsLOANTYPE = New ADODB.Recordset
          
        strLOANTYPE = "SELECT * FROM ALISPLoanType ;"
        rsLOANTYPE.Open strLOANTYPE, cnCOMMON, adOpenKeyset, adLockOptimistic

        frmODASMVoucher.cboPaymentCode.Clear

        With rsLOANTYPE
                  If .EOF Or .BOF Then Exit Sub
                  Do Until .EOF
                      frmODASMVoucher.cboPaymentCode.AddItem !LoanDescription
                      .MoveNext
                  Loop
        End With
            
rsLOANTYPE.Close
strLOANTYPE = ""

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub PaymentCodeGotFocusMedical()
On Error GoTo err

    Dim rsDOCTORCODEGF As ADODB.Recordset, strDOCTORCODEGF As String
    Set rsDOCTORCODEGF = New ADODB.Recordset
    
    strDOCTORCODEGF = "SELECT distinct(ALISPdoctor.CompanyName) FROM ALISMMedicalReport, ALISPDoctor where ALISMMedicalReport.doctorCode = ALISPDoctor.DoctorNo and ResultsReceived = 'Y' and requisitionPrepared = 'N';"
    rsDOCTORCODEGF.Open strDOCTORCODEGF, cnCOMMON, adOpenKeyset, adLockOptimistic

    frmODASMVoucher.cboPaymentCode.Clear

    With rsDOCTORCODEGF
            If .EOF Or .BOF Then Exit Sub
            Do Until .EOF
                frmODASMVoucher.cboPaymentCode.AddItem !CompanyName
                .MoveNext
            Loop
    End With
      
rsDOCTORCODEGF.Close
strDOCTORCODEGF = ""

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub PaymentCodeGotFocus()
        Set rsCOSTCENTER = New ADODB.Recordset
        
        rsCOSTCENTER.Open "SELECT * FROM ODASPCostCentre WHERE CostCenter = '" & Screen.ActiveForm.txtCostCenter.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
    
        If rsCOSTCENTER!Materials = "Y" Then
                PaymentCodeGotFocusClaim
        ElseIf rsCOSTCENTER!Machinery = "Y" Then
        
        ElseIf rsCOSTCENTER!Licenses = "Y" Then
    
        ElseIf rsCOSTCENTER!AdminCost = "Y" Then
                PaymentCodeGotFocusLoan
        ElseIf rsCOSTCENTER!ManPower = "Y" Then
                PaymentCodeGotFocusMedical
        ElseIf rsCOSTCENTER!Reinsurer = "Y" Then
        
        End If
End Sub
Private Sub PaymentCodeLostFocusClaim()
On Error GoTo err
    
    Dim rsPaymentCodeLF As ADODB.Recordset, strPaymentCodelf As String
    Set rsPaymentCodeLF = New Recordset
    
    rsPaymentCodeLF.Open "SELECT * FROM ODASPPaymentCode WHERE PaymentCodeDescription = '" & frmODASMVoucher.cboPaymentCode.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsPaymentCodeLF
            If .EOF And .BOF Then Exit Sub
            frmODASMVoucher.cboPaymentCode.Text = !PaymentCode
            frmODASMVoucher.txtReference = Trim(!PaymentCodeDescription)
            

            If !PayLifeAssured = 1 Then
                    bPayee = True
            Else: bPayee = False
            End If
    End With
    
    frmODASMVoucher.txtRequisitionDate = Date
    frmODASMVoucher.cboDocumentNo.SetFocus
    
rsPaymentCodeLF.Close
strPaymentCodelf = ""

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub PaymentCodeLostFocusLoan()
On Error GoTo err

        Dim rsLOAN As ADODB.Recordset, strLOAN As String
        Set rsLOAN = New Recordset
        
        rsLOAN.Open "SELECT * FROM ALISPLoanType WHERE LoanDescription = '" & frmODASMVoucher.cboPaymentCode.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsLOAN
                If .EOF And .BOF Then Exit Sub
                frmODASMVoucher.cboPaymentCode.Text = !loantype
                frmODASMVoucher.txtReference = Trim(!LoanDescription)
        End With
        
        frmODASMVoucher.txtRequisitionDate = Date
        frmODASMVoucher.cboDocumentNo.SetFocus
        
rsLOAN.Close
strLOAN = ""

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub PaymentCodeLostFocusMedical()
On Error GoTo err
        bMedicalRequisition = True
        Dim rsMEDICAL As ADODB.Recordset, strMEDICAL As String
        Set rsMEDICAL = New Recordset
        
        rsMEDICAL.Open "SELECT * FROM ALISPDoctor WHERE ALISPDoctor.CompanyName = '" & frmODASMVoucher.cboPaymentCode.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsMEDICAL
                If .EOF And .BOF Then Exit Sub
                frmODASMVoucher.cboPaymentCode.Text = !DoctorNo
                frmODASMVoucher.txtReference = Trim(!DoctorNo)
        End With
        
        frmODASMVoucher.txtRequisitionDate = Date
        frmODASMVoucher.cboDocumentNo.SetFocus
        
rsMEDICAL.Close
strMEDICAL = ""

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub PaymentCodeLostFocus()
        If frmODASMVoucher.cboPaymentCode.Text <= "" Then Exit Sub
        
        Set rsCOSTCENTER = New ADODB.Recordset
        
        rsCOSTCENTER.Open "SELECT * FROM ODASPCostCentre WHERE CostCenter = '" & Screen.ActiveForm.txtCostCenter.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic

        If rsCOSTCENTER!Materials = "Y" Then
                PaymentCodeLostFocusClaim
        ElseIf rsCOSTCENTER!Machinery = "Y" Then
        ElseIf rsCOSTCENTER!Licenses = "Y" Then
        ElseIf rsCOSTCENTER!AdminCost = "Y" Then
                PaymentCodeLostFocusLoan
        ElseIf rsCOSTCENTER!ManPower = "Y" Then
                PaymentCodeLostFocusMedical
        ElseIf rsCOSTCENTER!Reinsurer = "Y" Then
        
        End If

End Sub
Private Sub documentNoGotFocusCLAIM()
On Error GoTo err
        
'        Dim rsCLAIMNoGF As ADODB.Recordset, strCLAIMNoGF As String
'        Set rsCLAIMNoGF = New ADODB.Recordset
'
'        GlobalClaimNo = frmODASMVoucher.cboDocumentNo
'
'        strCLAIMNoGF = "SELECT distinct(AlisMClaim.ClaimNo) FROM ALISMClaim, ALISmClaimTotal where ALISMClaim.claimNo = ALISMClaimTotal.claimNo and AlISMclaim.PaymentCode = '" & frmODASMVoucher.cboPaymentCode & "' and (ALISMClaimTotal.RequisitionPrepared <> 'Y' or ALISMClaimTotal.RequisitionPrepared is NULL ) and ALISMClaimTotal.Authorized = 'Y' ;"
'        rsCLAIMNoGF.Open strCLAIMNoGF, cnCOMMON, adOpenKeyset, adLockOptimistic
'
'        frmODASMVoucher.cboDocumentNo.Clear
'
'        With rsCLAIMNoGF
'                If .EOF Or .BOF Then Exit Sub
'                Do Until .EOF
'                    frmODASMVoucher.cboDocumentNo.AddItem !claimno
'                    .MoveNext
'                Loop
'        End With
                
'rsCLAIMNoGF.Close
'strCLAIMNoGF = ""
         

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub selectCLAIM()
On Error GoTo err
                
         

Exit Sub

err:
    ErrorMessage
End Sub


Public Sub LoadRequisition()
On Error GoTo err
    
    Set rsCONTROL = New ADODB.Recordset
    strSQL = "Select * from ODASMVoucher  Where ODASMVoucher.VoucherNo = '" & frmODASMVoucher.txtVoucherNo.Text & "' "
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With frmODASMVoucher
            If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
            
            .cboPaymentCode.Text = rsCONTROL!PaymentCode
            '.cboDocumentNo.Text = rsCONTROL!DocumentNo
            .txtCostCenter.Text = rsCONTROL!CostCenter
            .txtInvoiceAmount.Text = rsCONTROL!Amount
            .txtReference.Text = rsCONTROL!Reference & ""
            .txtRequisitionDate.Text = rsCONTROL!VoucherDate
            .txtStatus.Text = rsCONTROL!Status & ""
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub GetSelectedClaims()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Claim No", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Proceeds", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Deductions", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Net Amount", .ListView1.Width / 5


                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                DF = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!claimno))
                        
                        If Not IsNull(rsLIST!proceeds) Then
                            MyList.SubItems(1) = CStr(rsLIST!proceeds)
                        End If

                        If Not IsNull(rsLIST!deductions) Then
                                MyList.SubItems(2) = CStr(rsLIST!deductions)
                        End If
                        
                        If Not IsNull(rsLIST!netpayable) Then
                                MyList.SubItems(3) = CStr(rsLIST!netpayable)
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

Public Sub GetProceeds()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Proceed/Deduction", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Particulars", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                DF = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ClaimTypeDescription))
                        
                        If Not IsNull(rsLIST!ClaimAmount) Then
                            MyList.SubItems(1) = CStr(rsLIST!ClaimAmount)
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


Public Sub DocumentNoGotFocus()
On Error GoTo err
        Set rsCOSTCENTER = New ADODB.Recordset
        
        rsCOSTCENTER.Open "SELECT * FROM ODASPCostCentre WHERE CostCenter = '" & frmODASMVoucher.txtCostCenter.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsCOSTCENTER.EOF Or rsCOSTCENTER.BOF Then Exit Sub
       
       ' Set rsCOSTCENTER = New adodb.Recordset
    
        If rsCOSTCENTER!Materials = "Y" Then
                documentNoGotFocusCLAIM
        ElseIf rsCOSTCENTER!Machinery = "Y" Then
        
        ElseIf rsCOSTCENTER!Licenses = "Y" Then
    
        ElseIf rsCOSTCENTER!AdminCost = "Y" Then
        ElseIf rsCOSTCENTER!ManPower = "Y" Then
        ElseIf rsCOSTCENTER!Reinsurer = "Y" Then
        
        End If


Exit Sub

err:
    ErrorMessage
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
                
                DF = rsLIST.RecordCount
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


Public Sub SelectCostCenter()
        
        Set rsCOSTCENTER = New ADODB.Recordset
        
        rsCOSTCENTER.Open "SELECT * FROM ODASPCostCentre WHERE CostCentre = '" & frmODASMVoucher.txtCostCenter.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsCOSTCENTER.EOF Or rsCOSTCENTER.BOF Then Exit Sub
       
       ' Set rsCOSTCENTER = New adodb.Recordset
    
        If rsCOSTCENTER!AdminCost = "Y" Then
                documentNoGotFocusCLAIM
        ElseIf rsCOSTCENTER!Machinery = "Y" Then
        
        ElseIf rsCOSTCENTER!Materials = "Y" Then
                  GetInvoicesNotPaid
        ElseIf rsCOSTCENTER!Rent = "Y" Then
                  GetRentNotRequisitioned
        ElseIf rsCOSTCENTER!ManPower = "Y" Then
        
        ElseIf rsCOSTCENTER!rate = "Y" Then
                  GetRateNotPaid
        ElseIf rsCOSTCENTER!OtherCosts = "Y" Then
                  showALLOTHERREQUISITIONS
                  'GetEachREQUISITIONS
                
        End If
End Sub
Public Sub SelectParticularCostCenter()
        
        Set rsCOSTCENTER = New ADODB.Recordset
        
        rsCOSTCENTER.Open "SELECT * FROM ODASPCostCentre WHERE CostCentre = '" & frmODASMVoucher.txtCostCenter.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsCOSTCENTER.EOF Or rsCOSTCENTER.BOF Then Exit Sub
       
       ' Set rsCOSTCENTER = New adodb.Recordset
    
        If rsCOSTCENTER!AdminCost = "Y" Then
                documentNoGotFocusCLAIM
        ElseIf rsCOSTCENTER!Machinery = "Y" Then
        
        ElseIf rsCOSTCENTER!Materials = "Y" Then
                  GetInvoicesNotPaid
        ElseIf rsCOSTCENTER!Rent = "Y" Then
                  GetRentRequisitionedforSpecificAccount
        ElseIf rsCOSTCENTER!ManPower = "Y" Then
        
        ElseIf rsCOSTCENTER!rate = "Y" Then
                  GetRateNotPaid

        End If
End Sub


Private Sub loadINVOICEDETAILS()
On Error GoTo err

 Dim rsINVOICE As ADODB.Recordset, strInvoiceNo As String
 Set rsINVOICE = New ADODB.Recordset
 
 strInvoiceNo = " select * from ODASMInvoice I, ODASMRequisition R, ODASPAccount A where I.InvoiceNo = '" & frmODASMVoucher.cboDocumentNo & "' and I.LPONo = R.LPONo and R.AccountNo =  A.AccountNo; "
 rsINVOICE.Open strInvoiceNo, cnCOMMON, adOpenKeyset, adLockOptimistic
 
 With rsINVOICE
 
         If .BOF Or .EOF Then Exit Sub
         frmODASMVoucher.txtInvoiceAmount.Text = FormatNumber(rsINVOICE!PriceInclusive)
         frmODASMVoucher.txtVoucherAmount.Text = FormatNumber(rsINVOICE!PriceInclusive)

         frmODASMVoucher.txtStatus = "REQ-PREP"
         frmODASMVoucher.txtReference = Trim(frmODASMVoucher.txtPaymentCodeDescription) + "- " + Trim(frmODASMVoucher.cboDocumentNo.Text)
         frmODASMVoucher.txtAccountNo.Text = !AccountNo
'         frmODASMVoucher.txtProductCode.Text = !ProductCode
         
         frmALISMPayee.txtPayeeDetails = Trim(!CompanyName)
         frmALISMPayee.txtPayeeAddress = !PostalAddress
         frmALISMPayee.txtPostalCode = ""
         frmALISMPayee.cboTownCode = !Towncity
 End With

strInvoiceNo = ""
rsINVOICE.Close

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub loadLANDLORDDETAILS()
On Error GoTo err

 Dim rsINVOICE As ADODB.Recordset, strInvoiceNo As String
 Set rsINVOICE = New ADODB.Recordset
 
 strInvoiceNo = " select * from ODASMInstallment I, ODASPAccount A where I.InvoiceNo = '" & frmODASMVoucher.cboDocumentNo & "' and I.AccountNo =  A.AccountNo; "

 rsINVOICE.Open strInvoiceNo, cnCOMMON, adOpenKeyset, adLockOptimistic
 
 With rsINVOICE
 
         If .BOF Or .EOF Then Exit Sub
         frmODASMVoucher.txtInvoiceAmount.Text = FormatNumber(rsINVOICE!PaymentDue)
         frmODASMVoucher.txtAmountPaid.Text = CDbl(FormatNumber(rsINVOICE!PaymentDue)) - CDbl(frmODASMVoucher.txtVoucherAmount.Text)
         frmODASMVoucher.txtInvoiceBalance = 0

         frmODASMVoucher.txtStatus = "REQ-PREP"
         If frmODASMVoucher.txtCostCenter = "RENT" Then
         Else
             frmODASMVoucher.txtReference = Trim(frmODASMVoucher.txtPaymentCodeDescription) + "- " + Trim(frmODASMVoucher.cboDocumentNo.Text)
         End If
         frmODASMVoucher.txtAccountNo.Text = !AccountNo
         frmODASMVoucher.txtProductCode.Text = "N/A"
         
         frmODASMVoucher.txtPayeeDetails = Trim(!CompanyName)
'         frmODASMVoucher.txtPayeeAddress = !PostalAddress
'         frmODASMVoucher.txtPostalCode = ""
'         frmODASMVoucher.cboTownCode = !towncity
 End With

strInvoiceNo = ""
rsINVOICE.Close

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub loadCOUNCILDETAILS()
On Error GoTo err
 Dim strInvoiceNo, strJOBBRIEF As String
 Set rsFindRecord = New ADODB.Recordset
 
 rsFindRecord.Open "Select * From ODASMCouncilRateDue Where ReferenceNo like '" & frmODASMVoucher.cboDocumentNo & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
 
 If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub
 
         
             Set rsINVOICE = New ADODB.Recordset
             
             If rsFindRecord!Face = "Y" Then
                strInvoiceNo = " select * from ODASPPlotSite S, ODASPPlot P,ODASPCouncil C  where P.PlotNo = S.PlotNo and P.CouncilCode = C.CouncilCode and S.SiteNo  = '" & rsFindRecord!SiteNo & "'; "
                rsINVOICE.Open strInvoiceNo, cnCOMMON, adOpenKeyset, adLockOptimistic
             ElseIf rsFindRecord!BillBoard = "Y" Then
                strInvoiceNo = " select * from ODASPPlotMast S, ODASPPlot P,ODASPCouncil C  where P.PlotNo = S.PlotNo and P.CouncilCode = C.CouncilCode and S.MastNo  = '" & rsFindRecord!SiteNo & "'; "
                rsINVOICE.Open strInvoiceNo, cnCOMMON, adOpenKeyset, adLockOptimistic
             ElseIf (IsNull(rsFindRecord!BillBoard) And IsNull(rsFindRecord!Face)) Then Exit Sub
             End If
             
             With rsINVOICE
             
                     If rsINVOICE.BOF Or rsINVOICE.EOF Then Exit Sub
                     frmODASMVoucher.txtInvoiceAmount.Text = FormatNumber(rsFindRecord!AmountDue)
                     frmODASMVoucher.txtVoucherAmount.Text = FormatNumber(rsFindRecord!AmountDue)
                     frmODASMVoucher.txtInvoiceBalance.Text = 0
                     
                     frmODASMVoucher.txtStatus = "REQ-PREP"
                     frmODASMVoucher.txtReference = Trim(frmODASMVoucher.txtPaymentCodeDescription) + "- " + Trim(frmODASMVoucher.cboDocumentNo.Text)
                     frmODASMVoucher.txtAccountNo.Text = !CouncilAccountNo
                     frmODASMVoucher.txtCouncilCode = !CouncilCode
                     
                     '/ Procedure to load Product
                    
                        Set rsJOBBRIEF = New ADODB.Recordset
            
                        strJOBBRIEF = " select * from ODASMjobBrief JB, ODASMjobBriefItems JBI where JB.jobBriefNo = JBI.JobBriefNo and JBI.JobBriefitemNo = '" & rsINVOICE!JobBriefItemNo & "'; "
                        rsJOBBRIEF.Open strJOBBRIEF, cnCOMMON, adOpenKeyset, adLockOptimistic
                        
                        If rsJOBBRIEF.EOF And rsJOBBRIEF.BOF Then
                                 frmODASMVoucher.txtProductCode.Text = "N/A"
'                                frmODASMVoucher.txtJobBriefNo.Text = !JobBriefItemNo & ""
                        Else
                                frmODASMVoucher.txtProductCode.Text = rsJOBBRIEF!ProductCode
'                                frmODASMVoucher.txtJobBriefNo.Text = rsJOBBRIEF!JobBriefNo
                        End If
                        
                     
                     frmALISMPayee.txtPayeeDetails = Trim(!Council)
'                     frmALISMPayee.txtPayeeAddress = !PostalAddress
                     frmALISMPayee.txtPostalCode = ""
                     frmALISMPayee.cboTownCode = !TownCode
             End With
strInvoiceNo = ""
rsINVOICE.Close

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub DocumentNoLostFocusLOAN()
On Error GoTo err

     Dim rsLOAN As ADODB.Recordset, strLOAN As String
     Set rsLOAN = New ADODB.Recordset
     
     strLOAN = " select * from ALISMLoanManagement, ODASMJobBrief, ODASPAccount where ALISMLoanManagement.LoanNo = '" & frmODASMVoucher.cboDocumentNo & "' and ALISMLoanManagement.JobBriefNo = ODASMJobBrief.JobBriefNo and ODASMJobBrief.AccountNo LIKE ODASPAccount.AccountNo ; "
     rsLOAN.Open strLOAN, cnCOMMON, adOpenKeyset, adLockOptimistic
     
     With rsLOAN
     
             If .BOF Or .EOF Then Exit Sub
             frmODASMVoucher.txtInvoiceAmount.Text = rsLOAN!LoanAmount
             frmODASMVoucher.txtStatus = "REQ-PREP"
             frmODASMVoucher.txtReference = Trim(frmODASMVoucher.txtReference) + "- " + Trim(frmODASMVoucher.cboDocumentNo.Text)
             frmODASMVoucher.txtPayeeDetails = Trim(!othernames) + " " + Trim(!CompanyName)
     End With

    strLOAN = ""
    rsLOAN.Close

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub DocumentNoLostFocusMEDICAL()
On Error GoTo err
        obtainMEDICALFEE
        loadDOCTORSDETAILS

Exit Sub

err:
    ErrorMessage
End Sub


Public Sub DocumentNoLostFocus()
    If frmODASMVoucher.txtCostCenter.Text > "" Then
        Set rsCOSTCENTER = New ADODB.Recordset
        
        rsCOSTCENTER.Open "SELECT * FROM ODASPCostCentre WHERE CostCentre = '" & Screen.ActiveForm.txtCostCenter.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic

        If rsCOSTCENTER!Materials = "Y" Then
                loadINVOICEDETAILS
        ElseIf rsCOSTCENTER!Machinery = "Y" Then
        
        ElseIf rsCOSTCENTER!Rent = "Y" Then
                loadLANDLORDDETAILS
        
        ElseIf rsCOSTCENTER!rate = "Y" Then
                loadCOUNCILDETAILS
        ElseIf rsCOSTCENTER!AdminCost = "Y" Then
                DocumentNoLostFocusLOAN
        ElseIf rsCOSTCENTER!ManPower = "Y" Then
        End If
        
        loadGRID
        
        If bMedicalRequisition = True Then
                Set rsLOADGRID = New clsALISGRID
                rsLOADGRID.loadSpecificMEDICALGRID
                Set rsLOADGRID = Nothing
        End If
        
'        strPAYEE = frmALISMPayee.txtPayeeDetails
'        strADDRESS = frmALISMPayee.txtPayeeAddress
'        frmODASMVoucher.txtPayeeDetails.Text = strPAYEE
'        strPOSTALCODE = frmALISMPayee.txtPostalCode
'        strTOWN = frmALISMPayee.cboTownCode

    End If
            '/ If Payment Type > ""

End Sub
   

Public Sub loadPaymentGRID()
On Error GoTo err

    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset

    If frmODASMVoucher.cboDocumentNo.Text = Empty Then Exit Sub
    strSQL = "SELECT ALISPClaimType.ClaimTypeDescription, ALISMClaim.Amount FROM ALISMClaim, ALISPClaimType WHERE ALISMClaim.claimNo =  '" & frmODASMVoucher.cboDocumentNo & "' and ALISMClaim.type = 'A' and ALISMClaim.ClaimType = ALISPClaimType.ClaimType;"
    GetProceeds
    
    
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub LoadDeductionGRID()
On Error GoTo err
    Dim rsGRID1 As ADODB.Recordset
    Set rsGRID1 = New ADODB.Recordset
    
    If frmODASMVoucher.cboDocumentNo.Text = Empty Then Exit Sub

    rsGRID1.Open "SELECT ALISPClaimType.ClaimTypeDescription, ALISMClaim.Amount FROM ALISMClaim, ALISPClaimType WHERE ALISMClaim.claimNo =  '" & frmODASMVoucher.cboDocumentNo & "' and ALISMClaim.type = 'D' and ALISMClaim.ClaimType = ALISPClaimType.ClaimType;", cnCOMMON, adOpenKeyset, adLockOptimistic
    Set Screen.ActiveForm.DeductionGrid.DataSource = rsGRID1
    
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub LoadRejectedGRID()
On Error GoTo err
    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset
    
    If frmODASMVoucher.cboDocumentNo.Text = Empty Then Exit Sub

    rsGRID.Open "SELECT * FROM ODASMVoucher WHERE Status =  'REQ-PREP' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    Set Screen.ActiveForm.RejectionGrid.DataSource = rsGRID
    
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub loadPendingGRID()
On Error GoTo err
    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset
    
    If frmODASMVoucher.cboDocumentNo.Text = Empty Then Exit Sub

    rsGRID.Open "SELECT * FROM ODASMVoucher WHERE Status =  'REQ-PREP' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    Set Screen.ActiveForm.PendingGrid.DataSource = rsGRID
    
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub loadApprovedGRID()
On Error GoTo err
    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset
    
    If frmODASMVoucher.cboDocumentNo.Text = Empty Then Exit Sub

    rsGRID.Open "SELECT * FROM ODASMVoucher WHERE Status =  'REQ APPROVAL' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    Set Screen.ActiveForm.ApprovedGrid.DataSource = rsGRID
    
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub loadAuthorizedGRID()
On Error GoTo err
    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset
    
    If frmODASMVoucher.cboDocumentNo.Text = Empty Then Exit Sub

    rsGRID.Open "SELECT * FROM ODASMVoucher WHERE Status =  'REQ AUTHORIZATION' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    Set Screen.ActiveForm.AuthorizedGrid.DataSource = rsGRID
    
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub loadTOTALS()
On Error GoTo err
    Dim rsTOTALS As ADODB.Recordset
    Set rsTOTALS = New ADODB.Recordset
    
'    rsTOTALS.Open "SELECT * FROM ALISMClaimTotal WHERE claimNo =  '" & frmODASMVoucher.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
'
'    With rsTOTALS
'            If .BOF Or .EOF = True Then Exit Sub
'            frmODASMVoucher.txtTotalPayments.Text = !proceeds
'            frmODASMVoucher.txtTotalDeductions.Text = !deductions
'
'    End With

'rsTOTALS.Close

Exit Sub
err:
    ErrorMessage
End Sub
Private Sub loadGRID()
    loadTOTALS
End Sub
Public Sub selectCostCenterLostFocus()
On Error GoTo err
        
        Set rsCOSTCENTER = New ADODB.Recordset
        
        rsCOSTCENTER.Open "SELECT * FROM ODASPCostCentre WHERE Description = '" & Screen.ActiveForm.txtCostCenter.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCOSTCENTER
                If .EOF And .BOF Then Exit Sub
                frmODASMVoucher.txtCostCenter.Text = !CostCenter
                frmODASMVoucher.txtPaymentDescription = !Description
        End With
        
        frmODASMVoucher.txtRequisitionDate = Date
        frmODASMVoucher.cboPaymentCode.SetFocus
        
rsCOSTCENTER.Close

Exit Sub

err:
        ErrorMessage
End Sub
Public Sub loadClaimDescription()
On Error GoTo err
    
    Set rsCONTROL = New ADODB.Recordset
    rsCONTROL.Open "SELECT * FROM ODASPPaymentCode WHERE PaymentCode = '" & frmODASMVoucher.cboPaymentCode.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With frmODASMVoucher
                
            If rsCONTROL.EOF Or rsCONTROL.BOF = True Then Exit Sub
            .txtCostCenter.Text = rsCONTROL!CostCenter
            frmODASMVoucher.txtPaymentCodeDescription.Text = rsCONTROL!PaymentCodeDescription
    
    End With
    
    rsCONTROL.Close

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadPaymentDescription()
On Error GoTo err
        
        Dim rsCOSTCENTER  As ADODB.Recordset
        
        Set rsCOSTCENTER = New ADODB.Recordset
        
        rsCOSTCENTER.Open "SELECT * FROM ODASPCostCentre WHERE CostCentre = '" & frmODASMVoucher.txtCostCenter.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCOSTCENTER
                If .EOF And .BOF Then Exit Sub
                frmODASMVoucher.txtPaymentDescription = !COSTCENTREDescription
        End With
        
rsCOSTCENTER.Close

Exit Sub

err:
        ErrorMessage
End Sub

Private Sub checkAPPROVALSTATUS()
On Error GoTo err
                
                Dim rsAUTHORIZATION As ADODB.Recordset, strAuthorization As String
                Set rsAUTHORIZATION = New Recordset
                    
                strAuthorization = "SELECT * FROM ALISPLoanOperationType  WHERE PaymentApproval = '1' ;"
                rsAUTHORIZATION.Open strAuthorization, cnCOMMON, adOpenKeyset, adLockOptimistic

                With rsAUTHORIZATION
                        If .EOF Or .BOF Then Exit Sub
                        
                        Dim rsAPPROVED As ADODB.Recordset, strAPPROVED As String
                        Set rsAPPROVED = New Recordset
                            
                        strAPPROVED = "SELECT * FROM ALISMLoanOperation  WHERE ApplicationNo = '" & frmODASMVoucher.txtVoucherNo & "' and operationType = '" & !OperationType & "' ;"
                        rsAPPROVED.Open strAPPROVED, cnCOMMON, adOpenKeyset, adLockOptimistic
        
                        With rsAPPROVED
                                If .BOF Or .EOF Then
                                        MsgBox "Authorization can only take place immediately after Approval", vbOKOnly
                                                bExitSub = True
                                                Exit Sub
                                ElseIf !Accept = "N" Then
                                        MsgBox "Cannot Approve payment That has been Rejected", vbOKOnly
                                                bExitSub = True
                                                Exit Sub
                                End If
                        
                        End With

            End With
                    '/ Authorization
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub checkAUTHORIZATIONSTATUS()
On Error GoTo err
                
                Dim rsAUTHORIZATION As ADODB.Recordset, strAuthorization As String
                Set rsAUTHORIZATION = New Recordset
                    
                strAuthorization = "SELECT * FROM ALISPLoanOperationType  WHERE PaymentAuthorization = '1' ;"
                rsAUTHORIZATION.Open strAuthorization, cnCOMMON, adOpenKeyset, adLockOptimistic

                With rsAUTHORIZATION
                        If .EOF Or .BOF Then Exit Sub
                        
                        Dim rsAPPROVED As ADODB.Recordset, strAPPROVED As String
                        Set rsAPPROVED = New Recordset
                            
                        strAPPROVED = "SELECT * FROM ALISMLoanOperation  WHERE ApplicationNo = '" & frmODASMVoucher.txtVoucherNo & "' and operationType = '" & !OperationType & "' ;"
                        rsAPPROVED.Open strAPPROVED, cnCOMMON, adOpenKeyset, adLockOptimistic
        
                        With rsAPPROVED
                                If .BOF Or .EOF Then
                                        Exit Sub
                                Else
                                        MsgBox "This Record has Already been Authorized", vbOKOnly
                                                bExitSub = True
                                End If
                        
                        End With

                        Dim rsCHKAPPROVAL As ADODB.Recordset, strCHKAPPROVAL As String
                        Set rsCHKAPPROVAL = New Recordset
                            
                        strCHKAPPROVAL = "SELECT * FROM ALISPLoanOperationType  WHERE PaymentApproval = '1' ;"
                        rsCHKAPPROVAL.Open strCHKAPPROVAL, cnCOMMON, adOpenKeyset, adLockOptimistic
                        
                        With rsCHKAPPROVAL
                                If .BOF Or .EOF Then Exit Sub
                                
                                    Dim rsAPPROVE As ADODB.Recordset, strAPPROVE As String
                                    Set rsAPPROVE = New Recordset
                                        
                                    strAPPROVE = "SELECT * FROM ALISMLoanOperation  WHERE ApplicationNo = '" & frmODASMVoucher.txtVoucherNo & "' and operationType = '" & !OperationType & "' ;"
                                    rsAPPROVE.Open strAPPROVE, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                                    With rsAPPROVE
                                            If .BOF Or .EOF Then
                                                    Exit Sub
                                            Else
                                                    MsgBox "This Record has Already been Approved", vbOKOnly
                                                            bExitSub = True
                                            End If
                                    
                                    End With
                        
                        End With

            End With
                    '/ Authorization
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub Form_Load()
    Call OpenODBCConnection

End Sub

Public Sub loadPAYMENTRECORD()
On Error GoTo err
    
    Set rsCONTROL = New ADODB.Recordset
    rsCONTROL.Open "SELECT * FROM ODASPPaymentCode WHERE PaymentCode = '" & frmODASMVoucher.cboPaymentCode.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With frmODASMVoucher
                
            If rsCONTROL.EOF Or rsCONTROL.BOF = True Then Exit Sub
            .txtCostCenter.Text = rsCONTROL!CostCenter
            
            If rsCONTROL!Payee = "Y" Then bPayee = True
            frmODASMVoucher.txtReference = Trim(rsCONTROL!PaymentCodeDescription)
            frmODASMVoucher.txtPaymentCodeDescription.Text = rsCONTROL!PaymentCodeDescription
            frmODASMVoucher.txtRequisitionDate.Text = Date
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub Cancelrecord()
        enableButtons
        clearALLRECORD
        disableALLRECORD
End Sub

Public Sub deleteRECORD()
'''On Error GoTo Myerr

    If frmODASMVoucher.txtVoucherNo.Text = "" Then
                MsgBox "There is no current record to delete", vbInformation, "Delete Information"
    ElseIf frmODASMVoucher.txtRequisitionDate = "" Then
                MsgBox "There is no current record", vbInformation, "Delete Information"
    Else
        If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            
            Dim rsREQ As ADODB.Recordset, strREQ As String
            Set rsREQ = New ADODB.Recordset
            
            strREQ = "Select * from ODASMVoucher Where VoucherNo = '" & frmODASMVoucher.txtVoucherNo.Text & "';"
            With rsREQ
                .Delete
                .Requery
                clearALLRECORD
                                
                End With
    End If
        '/* End if Msg Box
        
End If
        '/* If txt = ""
        
Exit Sub

Myerr:
    ErrorMessage


End Sub

Private Sub cmdEdit_Click()
    beditRECORD = True
    Set rsClaimApproval = New clsALISApproval
    rsClaimApproval.switchCOMMANDBUTTONS
    rsClaimApproval.loadAPPROVALDETAILS
    Set rsClaimApproval = Nothing

End Sub
Public Sub editRECORD()
On Error GoTo err
Dim strQRE As Variant
Dim rsFIND As ADODB.Recordset, edit As Boolean

        Set rsFIND = New ADODB.Recordset

        Select Case cmdEdit.Caption
                Case "&Edit"
                        enableALLRECORD

                        strQRE = InputBox("Enter Requisition No  to search.", "Search Value")
    
                        rsFIND.Open "SELECT * FROM PaymentRequisition WHERE VoucherNo = '" & strQRE & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

                        With rsFIND
                                If .EOF And .BOF Then
                                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                               Else
                                        frmODASMVoucher.txtInvoiceAmount.Text = !Amount
                                        frmODASMVoucher.txtRequisitionDate.Text = !RequisitionDate
                                        frmODASMVoucher.txtVoucherNo.Text = !vOUCHERnO
                                        frmODASMVoucher.txtStatus.Text = !Status
                                        frmODASMVoucher.cboPaymentCode.Text = !PaymentCode
                                        frmODASMVoucher.cboDocumentNo.Text = !DocumentNo
                                        frmODASMVoucher.txtReference.Text = !Reference
                                        frmODASMVoucher.txtPayeeDetails.Text = !PayeeDetails
                                        frmODASMVoucher.txtCostCenter = !CostCenter
                                                                          edit = True
                                End If
                        End With
                        
                        
                        If edit Then
                                Screen.ActiveForm.cmdEdit.Caption = "Save &Changes"
                        End If
    
                Case "Save &Changes"
                        Dim rsFinder As ADODB.Recordset
                        Set rsFinder = New ADODB.Recordset

                        rsFinder.Open "SELECT * FROM ODASMVoucher WHERE VoucherNo = '" & frmODASMVoucher.txtVoucherNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                        With rsFinder
                            !Amount = frmODASMVoucher.txtInvoiceAmount
                            !Preparedby = CurrentUserName
                            !RequisitionDate = frmODASMVoucher.txtRequisitionDate
                            !vOUCHERnO = frmODASMVoucher.txtVoucherNo
                            !Status = frmODASMVoucher.txtStatus
                            !PaymentCode = frmODASMVoucher.cboPaymentCode
                            !DocumentNo = frmODASMVoucher.cboDocumentNo
                            !Reference = frmODASMVoucher.txtReference
                            !PayeeDetails = frmODASMVoucher.txtPayeeDetails
                            !dateprepared = Date
                            !CostCenter = frmODASMVoucher.txtCostCenter
                            !RequisitionPrepared = "Y"

                            .Update
                            .Requery
                            edit = False
                    End With
                
                    clearALLRECORD
                    Screen.ActiveForm.cmdEdit.Caption = "&Edit"
            Case Else
        
            Exit Sub

        End Select

Exit Sub

err:
    UpdateErrorMessage
End Sub



Public Sub SearchRecord()
'''On Error GoTo Myerr

        Dim strQRE As Variant
        Dim rsFIND As ADODB.Recordset, edit As Boolean

        Set rsFIND = New ADODB.Recordset
        
        If bUseClaimNo = True Then
                strQRE = InputBox("Enter The Claim No to search.", "Search Value")
                rsFIND.Open "SELECT * FROM ODASMVoucher WHERE DocumentNo = '" & strQRE & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        Else
                strQRE = InputBox("Enter The Requisition No to search.", "Search Value")
                rsFIND.Open "SELECT * FROM ODASMVoucher WHERE VoucherNo = '" & strQRE & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        End If
        
        With rsFIND
                If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                Else
                            frmODASMVoucher.txtInvoiceAmount.Text = !Amount
                            frmODASMVoucher.txtRequisitionDate.Text = !RequisitionDate
                            frmODASMVoucher.txtVoucherNo.Text = !vOUCHERnO
                            frmODASMVoucher.txtStatus.Text = !Status
                            frmODASMVoucher.cboPaymentCode.Text = !PaymentCode
                            frmODASMVoucher.cboDocumentNo.Text = !DocumentNo
                            frmODASMVoucher.txtReference.Text = !Reference
                            frmODASMVoucher.txtPayeeDetails.Text = !PayeeDetails
                            frmODASMVoucher.txtCostCenter = !CostCenter
                            
                     edit = True
                End If
            
loadGRID

            End With

        Exit Sub

Myerr:
            ErrorMessage

End Sub
Public Sub ValidateRECORD()
On Error GoTo err
        With frmODASMVoucher
            
            bSaveRECORD = False
            
            If CDbl(.txtInvoiceAmount) <= 0 Then
                    MsgBox "The Requisition Amount cannot be left Blank", vbOKOnly
                    .txtInvoiceAmount.SetFocus
            
            ElseIf .txtRequisitionDate <= "" Then
                    MsgBox "The Payment Requisition Date cannot be left Blank", vbOKOnly
                    .txtRequisitionDate.SetFocus

            ElseIf .cboPaymentCode <= "" Then
                    MsgBox "The Claim Code is Required for all the Transaction", vbOKOnly
                    .cboPaymentCode.SetFocus

            ElseIf .cboDocumentNo <= "" Then
                    MsgBox "The Document No is Used to Determine the Payees details", vbOKOnly
                    .cboDocumentNo.SetFocus
                    
            ElseIf .txtCostCenter.Text = "OTP" And .txtRemark.Text = Empty Then
                    MsgBox "Please Give Detailed Remarks on This Payment", vbOKOnly
                    .txtRemark.SetFocus
           ElseIf .txtCostCenter.Text = "OTP" And .txtPayeeDetails.Text = Empty Then
                    MsgBox "Enter the Payee Details", vbOKOnly
                    .txtPayeeDetails.SetFocus
                                
            Else: bSaveRECORD = True
            
            End If
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub updateMEDICAL()
On Error GoTo err
    
    Dim rsCLAIM As ADODB.Recordset, strCLAIM As String
    
    Set rsCLAIM = New Recordset
    
    strCLAIM = "SELECT * from ALISMMedicalExamination where ProposalNo = '" & frmODASMVoucher.cboDocumentNo.Text & " '; "
    rsCLAIM.Open strCLAIM, cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsCLAIM
            If .EOF Or .BOF Then
                    MsgBox "Record cannot be Processed", vbOKOnly
                    Exit Sub
            End If
            
            !vOUCHERnO = frmODASMVoucher.txtVoucherNo
            !RequisitionDate = frmODASMVoucher.txtRequisitionDate
            !RequisitionPrepared = "Y"
            .Update
            .Requery
    End With
Exit Sub

rsCLAIM.Close
strCLAIM = ""

err:
    UpdateErrorMessage
End Sub
Private Sub saveINVOICE()
On Error GoTo err

    With frmODASMVoucher
    
                Set rsSAVE = New Recordset
                
                strSQL = "SELECT * from ODASMInvoice where InvoiceNo = '" & .cboDocumentNo.Text & " ' and LPONo = '" & .txtLPONo.Text & "'; "
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                With rsSAVE
                        If .EOF Or .BOF Then Exit Sub
                        
                        !Status = "REQ-PREPARED"
                        !StatusDate = Date
                        
                        !InvoiceBalance = CDbl(!InvoiceBalance) - CDbl(frmODASMVoucher.txtInvoiceAmount.Text)
                             
                        If !InvoiceBalance <= 0 Then
                                !Requisitioned = "Y"
                                !DateRequisitioned = Date
                            Else: !Requisitioned = "N"
                                !DateRequisitioned = Date
                        End If
                        
                        .Update
                        .Requery
                End With
            Exit Sub
            
        rsSAVE.Close
        strSQL = ""
    
   End With
err:
    UpdateErrorMessage
End Sub
Private Sub saveINSTALLMENT()
On Error GoTo err

    With frmODASMVoucher
    
            Set rsSAVE = New Recordset
            strSQL = "SELECT * from ODASMInstallment where InvoiceNo = '" & .cboDocumentNo.Text & " ' and ContractNo = '" & .txtLPONo.Text & "'; "
            rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
            If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
            
            rsSAVE!Status = "REQ-PREPARED"
            rsSAVE!StatusDate = Date
            
            If CDbl(.txtInvoiceBalance.Text) = 0 Then
                rsSAVE!Requisitioned = "Y"
                rsSAVE!DateRequisitioned = Date
            Else:
                rsSAVE!Requisitioned = "P"
                rsSAVE!DateRequisitioned = Date
            End If
            rsSAVE!vOUCHERnO = .txtVoucherNo
            rsSAVE!CurrentPeriod = .txtCurrentPeriod.Text
            rsSAVE!VoucherDate = CDate(.txtRequisitionDate.Text)
        
            rsSAVE.Update
            rsSAVE.Requery
    End With
Exit Sub
err:
    UpdateErrorMessage
End Sub
Private Sub saveRATESCHEDULE()
On Error GoTo err

    With frmODASMVoucher
    
                Set rsSAVE = New Recordset
                
                strSQL = "SELECT * from ODASMCouncilRateDue where ReferenceNo = '" & .cboDocumentNo.Text & " ' ; "
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                With rsSAVE
                        If .EOF Or .BOF Then Exit Sub
                        
                        !Status = "REQ-PREPARED"
                        !StatusDate = Date
                        
                        !Balance = CDbl(!AmountDue) - CDbl(frmODASMVoucher.txtInvoiceAmount.Text)
                             
                        If !Balance <= 0 Then
                                !Requisitioned = "Y"
                                !DateRequisitioned = Date
                            Else: !Requisitioned = "N"
                                !DateRequisitioned = Date
                        End If
                        
                        .Update
                        .Requery
                End With
            Exit Sub
            
        rsSAVE.Close
        strSQL = ""
    
   End With
err:
    UpdateErrorMessage
End Sub


Public Sub enableRECORD()
On Error GoTo err
        With frmODASMVoucher
            .txtVoucherAmount.Locked = False
            .txtRemark.Locked = False
            .txtAmountPaid.Locked = False
            .txtPayeeDetails.Enabled = True
            .txtPayeeDetails.Locked = False
        End With
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub clearRECORD()
On Error GoTo err
        With frmODASMVoucher
            .txtAmountPaid.Text = 0
            .txtInvoiceAmount.Text = 0
            .txtInvoiceBalance.Text = 0
        End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub updatePaymentDetails()
On Error GoTo errMSG
        If Screen.ActiveForm.txtTotalVoucherAmount < Screen.ActiveForm.txtInvoiceAmount.Text Then
               
                
                strSQL = "UPDATE ODASMInstallment SET PaymentFlag='P',PaymentDue='" & Screen.ActiveForm.txtInvoiceBalance.Text & "',Balance='" & Screen.ActiveForm.txtInvoiceBalance.Text & "',AmountPaid='" & Screen.ActiveForm.txtTotalVoucherAmount & "',PaymentDate='" & Format(Date, "yyyy/MM/dd") & "',ChequeDate='" & Format(Screen.ActiveForm.DTPChequeDate, "yyyy/MM/dd") & "',ChequeNo='" & Screen.ActiveForm.txtChequeNo & "'  WHERE InstallmentNo LIKE '" & Screen.ActiveForm.txtInstallmentNo & "'"
                Debug.Print strSQL
                Set rsSAVE = cnCOMMON.Execute(strSQL)
        Else
                
        strSQL = "UPDATE ODASMInstallment SET PaymentFlag='Y',PaymentDate='" & Format(Date, "yyyy/MM/dd") & "',ChequeDate='" & Format(Screen.ActiveForm.DTPChequeDate, "yyyy/MM/dd") & "',ChequeNo='" & Screen.ActiveForm.txtChequeNo & "'  WHERE InstallmentNo LIKE '" & Screen.ActiveForm.txtInstallmentNo & "'"
        Debug.Print strSQL
        Set rsSAVE = cnCOMMON.Execute(strSQL)
        End If
        
Exit Sub
errMSG:
        ErrorMessage
End Sub

Public Sub updateRECORD()

        ValidateRECORD
        If bSaveRECORD = True Then
                If Trim(Screen.ActiveForm.txtVoucherNo.Text) <= Empty Then
                        GenerateVoucherNo
                End If
                If frmODASMVoucher.cboPaymentCode.Text = "OTP" Then
                   GenerateVoucherItem
                   saveVOUCHERITEMS
                Else:
                     GenerateVoucherItem
                     computeVOUCHERTOTAL
                     saveVOUCHERITEMS
                     computeVOUCHERTOTAL
                End If
                countVOUCHERITEMS
                saveRecord
                GetInvoicesREQUISITIONED

                Set rsCOSTCENTER = New ADODB.Recordset
                rsCOSTCENTER.Open "SELECT * FROM ODASPCostCentre WHERE CostCentre = '" & Screen.ActiveForm.txtCostCenter.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
                If rsCOSTCENTER!Materials = "Y" Then
                        saveINVOICE
                        GetInvoicesNotPaid
                ElseIf rsCOSTCENTER!Machinery = "Y" Then
                        saveINVOICE
                        GetInvoicesNotPaid
                ElseIf rsCOSTCENTER!Rent = "Y" Then
                        saveINSTALLMENT
                        'saveINSTALLMENTISSUED
'                        GetRentRequisitionedforSpecificAccount
                ElseIf rsCOSTCENTER!rate = "Y" Then
                        saveRATESCHEDULE
                        GetRateNotPaid
                ElseIf rsCOSTCENTER!OtherCosts = "Y" Then
                        'saveREQUISITIONITEMS
                        UpdateCostsAdded
                        UpdateAdministrationCosts
                        showALLOTHERREQUISITIONS
                ElseIf rsCOSTCENTER!AdminCost = "Y" Then
                
                ElseIf rsCOSTCENTER!ManPower = "Y" Then
                
                End If
    
                BSave = False
                disableALLRECORD
                NewRecord = False
                beditRECORD = False

        End If
Exit Sub

Myerr:
UpdateErrorMessage
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            MsgBox "Update Cancelled! You cannot save this record because some required fields are blank or have invalid data!", vbCritical, "Cancel Update"
            RsCode.CancelUpdate
            RsCode.Requery

    Else
           ErrorMessage
    End If
End Sub

Public Sub saveINSTALLMENTISSUED()
On Error GoTo err

    With frmODASMPaymentConfirmation
    
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
        ElseIf IsNull(rsINSTALL!Totals) = True Then
                numAmountDue = 0
        Else: numAmountDue = CDbl(rsINSTALL!Totals)
        End If
                
        numBALANCE = CDbl(.txtTotalVoucherAmount)
        numAmountDue = CDbl(numAmountDue) - CDbl(.txtTotalVoucherAmount)
        Do While Not rsCONTROL.EOF

                Set rsSAVE = New Recordset
                strSQL = "SELECT * from ODASMInstallment where Installment = '" & rsCONTROL!Installment & " ' and  ContractNo = '" & rsCONTROL!ContractNo & " ' and ContractYear = '" & rsCONTROL!ContractYear & " ' and  PaymentMode = '" & rsCONTROL!PaymentMode & " '; "
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                 Debug.Print strSQL
                If rsSAVE.EOF Or rsSAVE.BOF Then
                Else
                        rsSAVE!Status = "CHK-ISSUED"
                        rsSAVE!StatusDate = Date
                        rsSAVE!CurrentPeriod = .txtCurrentPeriod.Text
                        
                        If CDbl(numAmountDue) >= CDbl(.txtInvoiceAmount.Text) Or CDbl(numAmountDue) = CDbl(.txtInvoiceAmount.Text) Then
                                rsSAVE!PaymentFlag = "Y"
'                                rsSAVE!PaymentDue = 0
                                rsSAVE!Balance = 0
                                rsSAVE!AmountPaid = CDbl(rsCONTROL!TotalRent)
                        Else:
                                If CDbl(rsCONTROL!TotalRent) <= numBALANCE Then
'                                        rsSAVE!PaymentDue = 0
                                        rsSAVE!PaymentFlag = "N"
                                        rsSAVE!Balance = 0
                                        rsSAVE!AmountPaid = CDbl(rsCONTROL!TotalRent)
                                        numBALANCE = numBALANCE - CDbl(rsCONTROL!TotalRent)
                                Else
'                                        rsSAVE!PaymentDue = rsCONTROL!TotalRent - numBALANCE
                                        rsSAVE!PaymentFlag = "P"
                                        rsSAVE!Balance = rsCONTROL!TotalRent - numBALANCE
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
Public Sub saveREQUISITIONITEMS()
On Error GoTo err
        With frmODASMVoucher
                Set rsSAVE = New ADODB.Recordset

                strSQL = "SELECT * FROM ODASMRequisitionItems WHERE ItemNo = '" & .cboDocumentNo.Text & "' "
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                        
                        If rsSAVE.EOF Or rsSAVE.BOF Then
                        End If
                                    rsSAVE!Authorized = "Y"
                                    rsSAVE!Issued = "Y"
                                    rsSAVE!returned = "N"
                                    rsSAVE!Used = "Y"
                                    rsSAVE!Updated = "Y"
                                    rsSAVE!Issued = "Y"
                                    rsSAVE!DateIssued = Date
                                    rsSAVE!IssuedBy = CurrentUserName
                                    rsSAVE!Closed = "Y"
                                    rsSAVE!RevisedBy = .txtPayeeDetails.Text
                            
                            rsSAVE.Update

        End With

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub UpdateCostsAdded()
On Error GoTo err

    With frmODASMVoucher
    
                Set rsSAVE = New Recordset
                
                strSQL = "SELECT * from ODASMJobCard where JobCardNo = '" & .txtJobCardNo.Text & " '; "
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                With rsSAVE
                        If .EOF Or .BOF Then Exit Sub
                          If IsNull(rsSAVE!TotalCost) = True Then
                             rsSAVE!TotalCost = 0
                          End If
                            !TotalCost = CDbl(!TotalCost) + CDbl(frmODASMVoucher.txtTotalVoucherAmount.Text)
                        .Update
                        .Requery
                End With
            Exit Sub
            
        rsSAVE.Close
        strSQL = ""
    
   End With
err:
    UpdateErrorMessage
End Sub
Private Sub UpdateAdministrationCosts()
On Error GoTo err

    With frmODASMVoucher
    
                Set rsSAVE = New Recordset
                Set rsFindRecord = New ADODB.Recordset
                
                rsFindRecord.Open "SELECT * FROM ODASPAdminCosting WHERE CostItem = '" & .txtReference & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Set rsSAVE = New ADODB.Recordset
                        rsSAVE.Open "Select * From ODASMAdministrationCosts Where CostItem = '" & frmODASMVoucher.txtReference.Text & "' and JobBriefNo = '" & .txtJobCardNo.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
                         
                            If rsSAVE.EOF And rsSAVE.BOF Then
                                rsSAVE.AddNew
                                rsSAVE!CostItem = .txtCouncilCode.Text
                                rsSAVE!JobBriefNo = .txtJobCardNo.Text
                                rsSAVE!Approved = "N"
                                rsSAVE!Authorized = "N"
                                rsSAVE!valid = "Y"
                                rsSAVE!Preparedby = CurrentUserName
                                rsSAVE!dateprepared = Date
                            End If
                                rsSAVE!Amount = CDbl(.txtAmountPaid.Text)
                                rsSAVE.Update
                                rsSAVE.Requery
                End With
            Exit Sub
            
        rsSAVE.Close
        strSQL = ""
    
err:
    UpdateErrorMessage
End Sub
Public Sub BatchReqItemUpdate()
  On Error GoTo err
    With frmODASMVoucher
    Dim search As String
        j = .ListView1.ListItems.Count
        For i = 1 To j
            If .ListView1.ListItems(i).Checked = True Then
            search = .ListView1.ListItems(i).Text
            
            Set rsSAVE = New ADODB.Recordset
            rsSAVE.Open "SELECT * FROM ODASMRequisitionItems WHERE ItemNo = '" & search & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
            If rsSAVE.RecordCount = 0 Then Exit Sub
                 rsSAVE!Authorized = "Y"
                 rsSAVE!Issued = "Y"
                 rsSAVE!returned = "N"
                 rsSAVE!Used = "Y"
                 rsSAVE!Updated = "Y"
                 rsSAVE!Issued = "Y"
                 rsSAVE!DateIssued = Date
                 rsSAVE!IssuedBy = CurrentUserName
                 rsSAVE!Closed = "Y"
                 rsSAVE!RevisedBy = .txtPayeeDetails.Text
            rsSAVE.Update
            End If
        Next i
    End With
Exit Sub
err:
ErrorMessage

End Sub
Public Sub BatchItemTotalsUpdate()
  On Error GoTo err
    With frmODASMVoucher
    Dim search As String
        j = .ListView1.ListItems.Count
        For i = 1 To j
            If .ListView1.ListItems(i).Checked = True Then
            search = .ListView1.ListItems(i).Text
            
            Set rsSAVE = New ADODB.Recordset
            rsSAVE.Open "SELECT * FROM ODASMRequisitionItems WHERE ItemNo = '" & search & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
            If rsSAVE.RecordCount = 0 Then Exit Sub
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        .txtTotalVoucherAmount.Text = 0
                ElseIf IsNull(rsSAVE!AllowanceTotals) = True Then
                        .txtTotalVoucherAmount.Text = 0
                Else
                       If .txtTotalVoucherAmount.Text = "" Then
                          .txtTotalVoucherAmount.Text = 0
                       End If
                        .txtTotalVoucherAmount.Text = CDbl(.txtTotalVoucherAmount.Text) + CDbl(rsSAVE!AllowanceTotals)
                End If
            rsSAVE.Update
            End If
        Next i
    End With
Exit Sub
err:
ErrorMessage

End Sub
