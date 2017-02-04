VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASReceiptSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub calculateDEPOSIT()
On Error GoTo err
    With frmODASMReceiptSchedule
        .txtAmount.Text = FormatNumber(CDbl(.txtPercentage) * CDbl(.txtBalance.Text) / 100)
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub clearINSTALLMENT()
On Error GoTo err
    With frmODASMReceiptSchedule
        
        .txtPaymentDueDate.Text = Date
        .DTPickerDueDate.Value = Date
        .txtAmount.Text = FormatNumber(CDbl(.txtBalance) - CDbl(.txtAmountCreated))
        .txtInvoiceReference.Text = Trim(.txtJobBriefNo) + "-" + Trim(.txtInstallmentNo)
        .txtPercentage.Text = 0
    End With
Exit Sub

err:
    ErrorMessage
End Sub


Public Sub enableINSTALLMENT()
On Error GoTo err
    With frmODASMReceiptSchedule
        
        .txtPaymentDueDate.Locked = True
        .txtPercentage.Locked = False
        .DTPickerDueDate.Enabled = True
        .txtAmount.Locked = False
        .txtInvoiceReference.Locked = True
    
    End With
Exit Sub

err:
    ErrorMessage
End Sub



Public Sub checkCASHSTATUS()
On Error GoTo err
    With Screen.ActiveForm
    
        Set rsCONTROL = New Recordset
        
        rsCONTROL.Open "SELECT * FROM ODASPGuarantor WHERE GuarantorType = '" & .cboGuarantorType.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
                .txtPercentAuthorized.Text = 0
                .txtAmount.Text = 0
                .txtPercentAuthorized.Locked = True
                .txtAmount.Locked = True
                .txtPercentAuthorized.BackColor = &HFFFFC0
                .txtAmount.BackColor = &HFFFFC0
                .txtRemarkRequired.Text = rsCONTROL!RequireREMARK
                
                .txtPercentAuthorized.Text = rsCONTROL!PercentAUTHORIZED & ""
                .txtAmount.Text = FormatNumber(CDbl(.txtPercentAuthorized) * CDbl(.txtPriceInclusive.Text) / 100)

                If CDbl(.txtAmount.Text) >= CDbl(.txtBalance.Text) Then
                        .txtStatus.Text = "AUTHORIZED"
                ElseIf CDbl(.txtAmount.Text) = 0 Then
                         .txtStatus.Text = "AUTHORIZED"

                Else
                        .txtStatus.Text = "NOT-AUTHORIZED"
                End If
    End With
        
rsCONTROL.Close
strPAYMETHOD = ""

Exit Sub

err:
    ErrorMessage

End Sub
Public Sub processAUTHORIZATION()
On Error GoTo err
        
        
        With frmODASMReceiptSchedule

        End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub GenerateAuthorizationNo()
On Error GoTo err

        Set rsCONTROL = New Recordset
        
        strSQL = "SELECT * FROM ODASPLastNumbers WHERE AutoAuthorizationNo = 'Y'"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
        With rsCONTROL
                If .BOF Or .EOF = True Then Exit Sub
                        Screen.ActiveForm.txtAuthorizationNo.Text = !AuthorizationNo & ""
                
                Select Case Len(Trim(Screen.ActiveForm.txtAuthorizationNo))
                        Case 1: Screen.ActiveForm.txtAuthorizationNo.Text = Trim(!AuthorizationPrefix) + "00000" + Trim(Screen.ActiveForm.txtAuthorizationNo)
                        Case 2: Screen.ActiveForm.txtAuthorizationNo.Text = Trim(!AuthorizationPrefix) + "0000" + Trim(Screen.ActiveForm.txtAuthorizationNo)
                        Case 3: Screen.ActiveForm.txtAuthorizationNo.Text = Trim(!AuthorizationPrefix) + "000" + Trim(Screen.ActiveForm.txtAuthorizationNo)
                        Case 4: Screen.ActiveForm.txtAuthorizationNo.Text = Trim(!AuthorizationPrefix) + "00" + Trim(Screen.ActiveForm.txtAuthorizationNo)
                        Case 5: Screen.ActiveForm.txtAuthorizationNo.Text = Trim(!AuthorizationPrefix) + "0" + Trim(Screen.ActiveForm.txtAuthorizationNo)
                        Case 6: Screen.ActiveForm.txtAuthorizationNo.Text = Trim(!AuthorizationPrefix) + Trim(Screen.ActiveForm.txtAuthorizationNo)
                End Select
                
                !AuthorizationNo = !AuthorizationNo + 1
                .Update
                .Requery

        End With
        
rsCONTROL.Close
strTRANS = ""

Exit Sub
err:
    ErrorMessage
End Sub

Public Sub updateRECORD()
        validateRECORD
        If bsaveRECORD = True Then
                saveRecord
                updateINSTALLMENTS
                If frmODASMReceiptSchedule.chkEqualInstallment.Value = 1 Then calculateDUEDATE
                updateJOBBRIEF
                disableALLRECORD
                NewRecord = False
                beditRECORD = False
        End If
        showALLInstallments
Exit Sub

End Sub
Public Sub saveINSTALLMENT()
On Error GoTo err
    
       With frmODASMReceiptSchedule

                strSQL = "select * from ODASMJobBriefInstallment Where InvoiceReference = '" & .txtInvoiceReference & "' ;"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        rsSAVE.AddNew
                        rsSAVE!InvoiceReference = .txtInvoiceReference.Text
                        rsSAVE!InstallmentNo = .txtInstallmentNo.Text
                        rsSAVE!dateprepared = Date
                        rsSAVE!AmountPaid = 0
                        rsSAVE!Preparedby = CurrentUserName
                        rsSAVE!Invoiced = "N"
                        rsSAVE!JobBriefNo = .txtJobBriefNo.Text
                End If
                
                rsSAVE!PaymentDueDate = .txtPaymentDueDate.Text
                rsSAVE!Amount = CDbl(.txtAmount.Text)
                rsSAVE!Balance = CDbl(.txtAmount.Text)
                rsSAVE!DurationCode = .cboPaymentMode
                rsSAVE!Percentage = CDbl(.txtPercentage.Text)
                rsSAVE.Update
                rsSAVE.Requery
                
                
                Set rsSAVE = Nothing
                strSQL = Empty
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub updateINSTALLMENTS()
On Error GoTo err
    With frmODASMReceiptSchedule
      Dim InstallmentNo As String
      Dim PaymentMode, DurationMode As String
    If .txtPriceInclusive.Text = 0 Then Exit Sub
        PaymentMode = Trim(.cboPaymentMode)
        DurationMode = Trim(.cboDurationMode)
        
        
        Set rsfindrecord2 = New ADODB.Recordset
        Set rsFindRecord = New ADODB.Recordset
        rsFindRecord.Open "SELECT * FROM ODASMJobBrief WHERE JobBriefNo = '" & .txtJobBriefNo.Text & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub
        
        rsfindrecord2.Open "SELECT * FROM ODASPPaymentMode WHERE PaymentMode = '" & PaymentMode & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        PaymentsInAYear = rsfindrecord2!PaymentsInAYear
Verify:
        Set rsFindRecord1 = New ADODB.Recordset
        rsFindRecord1.Open "SELECT * FROM ODASPInstallment WHERE ODASPInstallment.PaymentMode = '" & PaymentMode & "' and ODASPInstallment.LeasePeriod = '1';", cnCOMMON, adOpenKeyset, adLockOptimistic
     
        If rsFindRecord1.RecordCount = 0 Then
            
            Set rsfindrecord2 = New ADODB.Recordset
            rsfindrecord2.Open "SELECT * FROM ODASPPaymentMode WHERE PaymentMode = '" & PaymentMode & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

            Dim PInts, Dur As String
            PInts = 1: Dur = 0
            Do While PInts <> (12 / rsfindrecord2!CoverPeriod) + 1
                rsFindRecord1.AddNew
                rsFindRecord1!LeasePeriod = 1
                rsFindRecord1!PaymentMode = .cboPaymentMode.Text
                rsFindRecord1!Installment = PInts
                rsFindRecord1!Duration = Dur
                rsFindRecord1!InstallmentDescription = "Installment" & " " & PInts
                rsFindRecord1!dateprepared = Date
                rsFindRecord1!Preparedby = CurrentUserName
                
                rsFindRecord1.Update
                PInts = PInts + 1: Dur = Dur + rsfindrecord2!CoverPeriod
            Loop
        GoTo Verify
        End If
        Dim installmentsAmount As String, LeaseP, ProjectAmount As Variant
        LeaseP = 1:
        ProjectAmount = .txtPriceInclusive.Text
        ComDate = rsFindRecord!expectedDOC

        If Trim(DurationMode) = "M" Then
             InstallmentNo = 1
             Do While LeaseP < (CInt(.txtDuration.Text) + 1)
             PaymentsInAMonth = CInt(.txtDuration.Text)
             installmentsAmount = ProjectAmount / PaymentsInAMonth
            
                Set rsSAVE = New ADODB.Recordset
                
                rsSAVE.Open "SELECT * FROM ODASMJobBriefInstallment WHERE ContractYear = '" & LeaseP & "' and JobBriefNo = '" & rsFindRecord!JobBriefNo & "' and InstallmentNo = '" & rsFindRecord1!Installment & "' and PaymentMode = '" & PaymentMode & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                If rsSAVE.EOF And rsSAVE.BOF Then
                   rsSAVE.AddNew
                    rsSAVE!JobBriefNo = .txtJobBriefNo
                    rsSAVE!AccountNo = rsFindRecord!AccountNo
                    rsSAVE!DurationCode = .cboDurationMode
                    rsSAVE!InstallmentNo = InstallmentNo
                    rsSAVE!Duration = .txtDuration
                    rsSAVE!PaymentMode = .cboPaymentMode.Text
                    rsSAVE!InvoiceReference = GenerateInvoiceNo
                    rsSAVE!Invoiced = "Y"
                    rsSAVE!InvoiceNo = GenerateInvoiceNo
                    rsSAVE!InvoiceDate = Date
                    rsSAVE!ContractYear = LeaseP
                    rsSAVE!Amount = ProjectAmount
                    rsSAVE!CurrentPeriod = CurrentPeriod
                    rsSAVE!PaymentDueDate = DateAdd("M", rsFindRecord1!Duration, ComDate)
                    rsSAVE!Percentage = rsFindRecord1!InstallmentPercent
                    rsSAVE!PaymentDue = CDbl(installmentsAmount)
                    rsSAVE!Balance = CDbl(installmentsAmount)
                    rsSAVE!PaymentFlag = "N"
                End If
                rsSAVE.Update
                InstallmentNo = InstallmentNo + 1
            LeaseP = LeaseP + 1: ComDate = DateAdd("m", 1, ComDate)
            Loop
:
        Else
            Period = (CInt(.txtDuration.Text) + 1)
            
           Do While LeaseP < (CInt(.txtDuration.Text) + 1)
                 If Trim(DurationMode) = "Y" And (PaymentMode = "M") Then
                     PaymentsInAYear = CInt(.txtDuration.Text) * 12
                     installmentsAmount = ProjectAmount / PaymentsInAYear
                 ElseIf Trim(DurationMode) = "Y" And (PaymentMode = "Q") Then
                     PaymentsInAYear = CInt(.txtDuration.Text) * 4
                     installmentsAmount = ProjectAmount / PaymentsInAYear
                 ElseIf Trim(DurationMode) = "Y" And (PaymentMode = "S") Then
                     PaymentsInAYear = CInt(.txtDuration.Text) * 1
                     installmentsAmount = ProjectAmount / PaymentsInAYear
                 ElseIf Trim(DurationMode) = "Y" And (PaymentMode = "H") Then
                     PaymentsInAYear = CInt(.txtDuration.Text) * 2
                     installmentsAmount = ProjectAmount / PaymentsInAYear
                 Else
                    installmentsAmount = ProjectAmount / CInt(.txtDuration.Text)
                End If
                
                rsFindRecord1.MoveFirst
                While rsFindRecord1.EOF <> True
                
                    Set rsSAVE = New ADODB.Recordset
                    rsSAVE.Open "SELECT * FROM ODASMJobBriefInstallment WHERE ContractYear = '" & LeaseP & "' and JobBriefNo = '" & rsFindRecord!JobBriefNo & "' and InstallmentNo = '" & rsFindRecord1!Installment & "' and PaymentMode = '" & .cboPaymentMode.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                    If rsSAVE.EOF And rsSAVE.BOF Then
                        rsSAVE.AddNew
                        rsSAVE!JobBriefNo = .txtJobBriefNo
                        rsSAVE!AccountNo = rsFindRecord!AccountNo
                        rsSAVE!InstallmentNo = rsFindRecord1!Installment
                        rsSAVE!PaymentMode = .cboPaymentMode.Text
                        rsSAVE!DurationCode = .cboDurationMode
                        rsSAVE!InvoiceReference = GenerateInvoiceNo
                        rsSAVE!Duration = .txtDuration
                        rsSAVE!Invoiced = "Y"
                        rsSAVE!InvoiceNo = GenerateInvoiceNo
                        rsSAVE!InvoiceDate = Date
                        rsSAVE!ContractYear = LeaseP
                        rsSAVE!Amount = ProjectAmount
                        rsSAVE!CurrentPeriod = CurrentPeriod
                        rsSAVE!Percentage = rsFindRecord1!InstallmentPercent
                        rsSAVE!PaymentDueDate = DateAdd("M", rsFindRecord1!Duration, ComDate)
                        rsSAVE!PaymentDue = CDbl(installmentsAmount)
                        rsSAVE!Balance = CDbl(installmentsAmount)
                        rsSAVE!PaymentFlag = "N"
                    End If
                    rsSAVE.Update
                    rsFindRecord1.MoveNext
                Wend
                    LeaseP = LeaseP + 1: ComDate = DateAdd("yyyy", 1, ComDate)
                Loop
         End If
        Set rsSAVE = Nothing: Set rsfindrecord2 = Nothing: Set rsFindRecord = Nothing: Set rsFindRecord1 = Nothing

     End With
Exit Sub
err:
ErrorMessage
End Sub

Function GenerateInvoiceNo()
    With frmODASMReceiptSchedule
        Set rsFindRecord4 = New ADODB.Recordset
        rsFindRecord4.Open "SELECT * FROM ODASMJobBrief WHERE JobBriefNo = '" & .txtJobBriefNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsFindRecord4.RecordCount = 0 Then Exit Function
            Set rsfindrecord2 = New ADODB.Recordset
            rsfindrecord2.Open "SELECT * FROM ODASMJobBriefInstallment WHERE JobBriefNo = '" & rsFindRecord4!JobBriefNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
            If rsfindrecord2.EOF And rsfindrecord2.BOF Then
                GenerateInvoiceNo = rsFindRecord4!JobBriefNo & "-" & 1
            Else
                GenerateInvoiceNo = rsfindrecord2!JobBriefNo & "-" & (rsfindrecord2.RecordCount + 1)
            End If
    End With
End Function
Public Sub saveRecord()
On Error GoTo err
    
       With frmODASMReceiptSchedule

                strSQL = "select * from ODASMJobBriefSchedule Where jobBriefNo = '" & .txtJobBriefNo & "' ;"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        rsSAVE.AddNew
                        rsSAVE!JobBriefNo = .txtJobBriefNo.Text
                        rsSAVE!TransactionType = .txtTransactionType.Text
                        rsSAVE!Preparedby = CurrentUserName
                        rsSAVE!dateprepared = Date
                        rsSAVE!CurrentPeriod = .txtCurrentPeriod
                        rsSAVE!InstallmentsPaid = 0
                        rsSAVE!DurationMode = .cboDurationMode.Text
                        rsSAVE!Duration = CInt(.txtDuration.Text)
                        rsSAVE!Installments = CInt(.txtInstallments)
                        rsSAVE!PaymentMode = .cboPaymentMode.Text
                        rsSAVE!NoofMonths = .txtNoOfMonths.Text
                        rsSAVE!InstallmentType = .cboInstallmentType
                        
                        If .chkEqualInstallment.Value = 1 Then
                                    rsSAVE!EqualInstallment = "Y"
                            Else: rsSAVE!EqualInstallment = "N"
                        End If
                
                        If .chkAfterCommencementDate.Value = 1 Then
                                rsSAVE!AfterDOC = "Y"
                        Else: rsSAVE!AfterDOC = "N"
                        End If
                End If
                
                rsSAVE!Amount = CDbl(.txtAmount.Text)
                rsSAVE!remark = .txtRemark.Text
                rsSAVE!Status = Trim(.txtStatus.Text)
                rsSAVE!PaymentDueDate = .txtPaymentDueDate.Text
                
                
                rsSAVE.Update
                rsSAVE.Requery
                
                Set rsSAVE = Nothing
                strSQL = Empty
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub enableRECORD()
On Error GoTo err
        With frmODASMReceiptSchedule
            .cboDurationMode.Locked = False
            .cboPaymentMode.Locked = False
            .chkAfterCommencementDate.Enabled = True
            .UpDownInstallment.Enabled = True
            .UpDownDuration.Enabled = True
            .chkEqualInstallment.Enabled = True
            .txtRemark.Locked = False
            .cboInstallmentType.Locked = False
            .DTPickerDueDate.Enabled = True
            .ChkRestore.Enabled = True
            .txtPriceInclusive.Locked = False
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub validateRECORD()
On Error GoTo err
        
    With frmODASMReceiptSchedule
            
            bsaveRECORD = False
                       
            If CDbl(.txtBalance.Text) <= 0 Then
                    MsgBox "Job Brief Fully paid hence No need for Approval"
                    .txtBalance.SetFocus
            
            ElseIf CDbl(.txtAmountCreated.Text) + CDbl(.txtAmount.Text) > CDbl(.txtBalance.Text) Then
                    MsgBox "The Installments Have been Fully Created "
                    .txtAmount.SetFocus
            
            ElseIf CDbl(.txtAmountCreated.Text) > CDbl(.txtBalance.Text) Then
                    MsgBox "The Installments Have been Fully Created "
                    .txtPercentage.SetFocus
                    
            ElseIf .txtDuration.Text <= Empty Then
                    MsgBox "The Duration is Required"
                    .txtDuration.SetFocus
            
            ElseIf .cboDurationMode.Text = Empty Then
                    MsgBox "The Duration Mode is Required to Determine the start Date of the Contract"
                    .cboDurationMode.SetFocus
            
            ElseIf .cboPaymentMode.Text = Empty Then
                    MsgBox "The Payment Mode is Necessary"
                    .cboPaymentMode.SetFocus
                    
            ElseIf CInt(.txtInstallmentNo.Text) < 1 Then
                    MsgBox "The Number of Installment Cannot be Less Than One"
                    .txtInstallmentNo.SetFocus
            
            ElseIf CInt(.txtInstallments.Text) < 1 Then
                    MsgBox "The Number of Installment MUST be Greater than 1"
                    .txtInstallments.SetFocus
                    
            ElseIf CDbl(.txtAmount.Text) < 0 Then
                    MsgBox "The Amount Owed MUST be Greater Than Zero"
                    .txtAmount.SetFocus
                    
            ElseIf .txtInvoiceReference.Text = Empty Then
                    MsgBox "The Invoice Reference is Required "
                    .txtInvoiceReference.SetFocus
                    
            Else
                    bsaveRECORD = True
            End If
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadRECORD()
On Error GoTo err
    With frmODASMReceiptSchedule
    
       
        Set rsCONTROL = New ADODB.Recordset
        strSQL = "SELECT * FROM ODASMJobBrief JB, ODASPAccount AC  WHERE JB.AccountNo = AC.AccountNo and JB.JobBriefNo = '" & .txtJobBriefNo.Text & "' ; "
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub

                .txtCurrentPeriod.Text = CurrentPeriod
                .txtCompanyName.Text = rsCONTROL!CompanyName
                .txtDescriptionOfOrder.Text = rsCONTROL!descriptionOfOrder
                .txtProductCode.Text = rsCONTROL!ProductCode & ""
                .txtJobBriefDate.Text = rsCONTROL!JobBriefDate
                
                If Not IsNull(rsCONTROL!CommencementDate) Then
                    .txtCommencementDate.Text = rsCONTROL!CommencementDate
                Else
                    .txtCommencementDate.Text = Date
                End If
                
                .txtPriceInclusive.Text = FormatNumber(rsCONTROL!TotalPrice) & ""
                .txtBalance.Text = FormatNumber(rsCONTROL!Balance) & ""
                .txtExpiryDate.Text = rsCONTROL!expirydate & ""
                .cboPaymentMode.Text = rsCONTROL!PaymentMode & ""
                
                If IsNull(rsCONTROL!expirydate) = True Then
                    .txtNoOfMonths.Text = 1
                Else: .txtNoOfMonths.Text = DateDiff("M", .txtCommencementDate.Text, .txtExpiryDate)
                End If
                
                If CDbl(.txtBalance) <= Empty Then
                        .txtBalance.Text = 0
                    Else: .txtAmount.Text = FormatNumber(CDbl(.txtBalance.Text) - CDbl(.txtAmountCreated)) & ""
                End If
                
    End With

Exit Sub
err:
    ErrorMessage
End Sub
Public Sub loadAMOUNTCREATED()
On Error GoTo err
    
       With frmODASMReceiptSchedule

                strCONTROL = "select sum(amount) as TotalBalance from ODASMjobBriefInstallment Where JobBriefNo = '" & .txtJobBriefNo & "' ;"
                Set rsCONTROL = New ADODB.Recordset
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                If rsCONTROL.EOF Or rsCONTROL.BOF Then
                        .txtAmountCreated.Text = 0
                ElseIf IsNull(rsCONTROL!TotalBalance) = True Then
                        .txtAmountCreated.Text = 0
                Else: .txtAmountCreated.Text = FormatNumber(rsCONTROL!TotalBalance)
                End If
                
                
                '/* Procedure to Count the Number of Installments
                
                strCONTROL = "select * from ODASMjobBriefInstallment Where JobBriefNo = '" & .txtJobBriefNo & "' ;"
                Set rsCONTROL = New ADODB.Recordset
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsCONTROL.EOF Or rsCONTROL.BOF Then
                         .txtInstallmentNo.Text = rsCONTROL.RecordCount + 1
                Else: .txtInstallmentNo.Text = rsCONTROL.RecordCount + 1
                End If
        End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub updateJOBBRIEF()
On Error GoTo err
    
       With frmODASMReceiptSchedule

                strSQL = "select * from ODASMJobBrief Where JobBriefNo = '" & .txtJobBriefNo & "' ;"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
                
                '/ Calculate the Total Amount scheduled So Far
                
                strCONTROL = "select sum(amount) as TotalBalance from ODASMjobBriefInstallment Where JobBriefNo = '" & .txtJobBriefNo & "' ;"
                Set rsCONTROL = New ADODB.Recordset
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                If rsCONTROL.EOF Or rsCONTROL.BOF Then
                        rsSAVE!Scheduled = "N"
                        
                ElseIf rsCONTROL!TotalBalance >= rsSAVE!Balance Then
                        rsSAVE!Scheduled = "Y"
                        rsSAVE!TotalPrice = .txtPriceInclusive.Text
                        rsSAVE!VATAmount = CDbl(.txtPriceInclusive.Text) * 0.16
                        rsSAVE!PriceExclusive = CDbl(.txtPriceInclusive.Text) - CDbl(rsSAVE!VATAmount)
                Else
                        rsSAVE!Scheduled = "N"
                End If
                
                .txtAmountCreated.Text = FormatNumber(rsCONTROL!TotalBalance)
                         
                rsSAVE.Update
                rsSAVE.Requery
                
                
                Set rsSAVE = Nothing
                strSQL = Empty
    End With
Exit Sub

err:
    ErrorMessage
End Sub


Public Sub calculateDUEDATE()
On Error GoTo err
        
        Dim Wtemp As Double
    
        With frmODASMReceiptSchedule
            If .chkEqualInstallment.Value = 1 Then
                    
                    Dim strMODE As String
                    Dim rsMODE As ADODB.Recordset
                    
                    '/ Access The Payment Mode
                    
                    strMODE = "select * from ODASPPaymentMode Where PaymentMode = '" & .cboPaymentMode & "';"
                    Set rsMODE = New ADODB.Recordset
                    rsMODE.Open strMODE, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                    If rsMODE.BOF Or rsMODE.EOF Then Exit Sub
                    
                    Wtemp = CDbl(.txtBalance.Text)
                    
                    strCONTROL = "select * from ODASMJobBriefInstallment Where JobBriefNo = '" & .txtJobBriefNo & "' order by InstallmentNo ;"
                    Set rsCONTROL = New ADODB.Recordset
                    rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                    If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
                    
                    Wtemp = CDbl(.txtBalance) - CDbl(.txtAmount.Text)
                    rsCONTROL.MoveFirst
                    Do While Wtemp > 0
                            
                            .txtInstallmentNo.Text = CInt(.txtInstallmentNo) + 1
                            .txtInvoiceReference.Text = Trim(.txtJobBriefNo.Text) + "-" + Trim(.txtInstallmentNo.Text)
                            .txtPaymentDueDate.Text = DateAdd("M", rsMODE!CoverPeriod, .txtPaymentDueDate)
                            
                            saveINSTALLMENT

                            strSQL = "select sum(Amount) as TotalInstallment from ODASMJobBriefInstallment Where JobBriefNo = '" & .txtJobBriefNo & "'  ;"
                            Set rsSAVE = New ADODB.Recordset
                            rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                            
                            If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
                            
                            Wtemp = CDbl(.txtBalance) - CDbl(rsSAVE!TotalInstallment)

                            'rsCONTROL.MoveNext
                    Loop
                End If
        End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub calculateAMOUNT()
On Error GoTo err
        With frmODASMReceiptSchedule
        
            
            If .chkEqualInstallment.Value = 1 Then
                    .txtAmount.Text = FormatNumber(CDbl(.txtBalance.Text) / CDbl(.txtInstallments))
                    .txtPercentage.Text = 100 / CDbl(.txtInstallments.Text)
            Else
                    .txtAmount.Text = 0
            End If
            
            If .chkAfterCommencementDate.Value = 1 Then
                     .txtPaymentDueDate.Text = DateAdd(Trim(.cboDurationMode.Text), CDbl(.txtDuration), .txtCommencementDate.Text)
            Else
                    .txtPaymentDueDate.Text = Date
            End If
            
            .txtInvoiceReference.Text = Trim(.txtJobBriefNo.Text) + "-" + Trim(.txtInstallmentNo.Text)

        End With
Exit Sub

err:
    ErrorMessage
End Sub


Public Sub LoadDEFAULT()
On Error GoTo err
    With frmODASMReceiptSchedule
        .cboPaymentMode.Text = "M"
        .cboDurationMode.Text = "D"
        .txtStatus.Text = "CREATED"
        .txtDuration.Text = 0
        .txtInstallments.Text = 1
        .chkAfterCommencementDate.Value = 1
        .chkEqualInstallment.Value = 1
        .cboInstallmentType.Text = "P"
        .txtTransactionType.Text = "NWB"
        .txtInvoiceReference.Text = Trim(.txtJobBriefNo.Text) + "-" + Trim(.txtInstallmentNo.Text)
        .txtPaymentDueDate.Text = .txtCommencementDate.Text
        .UpDownDuration.Value = 1

    End With
Exit Sub

err:
    ErrorMessage
End Sub
