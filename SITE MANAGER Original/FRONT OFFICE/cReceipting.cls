VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cReceipting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim rsreceipt As ADODB.Recordset, bVAL As Boolean
Dim rsDEP As ADODB.Recordset, strDEP As String, BPolicyHolder As Boolean
Public strCalculateCOMMISSION As String
Private MySName As String, MyOName As String


Sub loadpolicyGRID()

Dim rsPOLICYGRID As ADODB.Recordset
Set rsPOLICYGRID = New Recordset

   rsPOLICYGRID.Open "SELECT JobBriefNo, ExpectedPremium, DateOfCommencement, DateOfLastPayment FROM ODASMJobBrief WHERE AccountNo LIKE  '" & Screen.ActiveForm.txtAccountNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
   Set Screen.ActiveForm.policiesGRID.DataSource = rsPOLICYGRID

End Sub


Sub loadloanGRID()

Dim rsLOANGRID As ADODB.Recordset
Set rsLOANGRID = New Recordset

   rsLOANGRID.Open "SELECT LoanNo, RepaymentAmount,CommencementDate, DateofLastPayment FROM ALISMLoanManagement WHERE AccountNo LIKE  '" & Screen.ActiveForm.txtAccountNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
   Set Screen.ActiveForm.loanGRID.DataSource = rsLOANGRID

End Sub

Sub loadproposalGRID()

Dim rsPROPOSALGRID As ADODB.Recordset
Set rsPROPOSALGRID = New Recordset

   rsPROPOSALGRID.Open "SELECT ProposalNo, ExpectedPremium, ProposalDate, DateOfLastPayment FROM ALISMProposal WHERE AccountNo LIKE  '" & Screen.ActiveForm.txtAccountNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
   Set Screen.ActiveForm.proposalGRID.DataSource = rsPROPOSALGRID

End Sub

Sub LoadReceiptGrid()

Dim rsGRID As ADODB.Recordset
Set rsGRID = New Recordset

   rsGRID.Open "SELECT * FROM ALISMReceipt WHERE DocumentNo =  '" & Screen.ActiveForm.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
   Set Screen.ActiveForm.ReceiptGrid.DataSource = rsGRID

End Sub

Public Sub UpdatePolicyLedger()

On Error GoTo err

With Screen.ActiveForm
    If .cboReceiptType.Text = "BIP" Then Exit Sub
    
    Set rsFindRecord = New ADODB.Recordset
    rsFindRecord.Open "SELECT * FROM ODASPAccount WHERE AccountNo='" & .txtAccountNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        MySName = Empty: MyOName = Empty
    ElseIf rsFindRecord!CompanyName = "" And rsFindRecord!OtherNames = "" Then
        MySName = Empty: MyOName = Empty
    ElseIf IsNull(rsFindRecord!CompanyName) = True And IsNull(rsFindRecord!OtherNames) = True Then
        MySName = Empty: MyOName = Empty
    Else
        MySName = rsFindRecord!CompanyName & "": MyOName = rsFindRecord!OtherNames & ""
    End If
    
    If rsRCPT!Deposit = "1" Then Exit Sub
    
    Dim rsDEP As ADODB.Recordset, strDEP As String
    Set rsDEP = New ADODB.Recordset

    rsDEP.Open "SELECT * FROM ODASMJobBriefLedger;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
        rsDEP.AddNew
            rsDEP!ReceiptNo = .txtReceiptNo.Text
            rsDEP!ReceiptDate = .txtReceiptDate.Text
            rsDEP!ReceiptType = .cboReceiptType.Text
            rsDEP!DocumentNo = .cboDocumentNo.Text
            rsDEP!AccountNo = .txtAccountNo.Text
            rsDEP!ReceiptAmount = .txtReceiptAmount.Text
            rsDEP!AccountingPeriod = CurrentPeriod
            rsDEP!PaymentStatus = "PAID"
            rsDEP!UnitsPaid = .txtUnitsPaid.Text
            rsDEP!AccountCode = .cboAccountNo.Text
            rsDEP!DueDate = .txtDueDate.Text
            rsDEP!SuspenseAccount = .txtSuspenseAccount.Text
            rsDEP!PaymentMethod = "SO"
            rsDEP!BankNo = .cboBankNo.Text
            rsDEP!ChequeNo = .txtChequeNo.Text
            rsDEP!DateofLastPayment = .txtDateOfLastPayment.Text
            rsDEP!ReceivedToDate = .txtReceivedTodate.Text
            rsDEP!UnitCount = .txtUnitCount.Text
            rsDEP!UnitCountBeforePayment = .txtUnitCountBeforePayment
            rsDEP!StatusCode = .txtStatusCode.Text
            rsDEP!TransactionDate = Date
            rsDEP!TransactionNo = .txtTransactionNo.Text
            rsDEP!CompanyName = MySName
            rsDEP!OtherNames = MyOName
            rsDEP!Preparedby = CurrentUserName
        rsDEP.Update
    
    Set rsDEP = Nothing
    
End With
Exit Sub
err:
UpdateErrorMessage
If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
rsDEP.CancelUpdate
rsDEP.Requery
End If

End Sub

Public Sub processUPDATE()

On Error GoTo err

With Screen.ActiveForm
Dim CCOm As Variant
Set rsRCPT = New Recordset
strSQL = "SELECT * FROM ALISPReceipt where receiptType = '" & Screen.ActiveForm.cboReceiptType.Text & "';"
rsRCPT.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
If rsRCPT!Deposit <> "1" Then
    Dim rsPOLICY As ADODB.Recordset
    Set rsPOLICY = New Recordset
    rsPOLICY.Open "SELECT * FROM ODASMJobBrief Where JobBriefNo = '" & Trim(.cboDocumentNo.Text) & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    .cmdUpdate.Enabled = False
    If .cboReceiptType.Text = "BIP" Then GoTo StartCommission
    Dim rsPRODUCT As ADODB.Recordset, strPRODUCT As String
    Set rsPRODUCT = New Recordset
    strPRODUCT = "SELECT * FROM ALISPProduct where productCode = '" & rsPOLICY!ProductCode & "';"
    rsPRODUCT.Open strPRODUCT, cnCOMMON, adOpenKeyset, adLockOptimistic
    If rsPRODUCT.EOF Or rsPRODUCT.BOF = True Then Exit Sub
 End If

StartCommission:
    createRECEIPT
    validateRECORD
    If bsaveRECORD = True Then
        saveRecord
        enableCBRECEIPT
        disableDATAENTRY
        LoadReceiptGrid
    End If
    
    If .cboReceiptType.Text = "BIP" Then
        CCOm = 1
    Else
        CCOm = rsRCPT!CalculateCommission
    End If
    
    
    Set rsPOLICY = Nothing
    Set rsPRODUCT = Nothing
    strPRODUCT = ""
End With
    Exit Sub
err:
    UpdateErrorMessage
End Sub

Public Sub LoadReceiptGrid1()
On Error GoTo err:

Dim rsGRID As ADODB.Recordset, StrGRID As String
Set rsGRID = New Recordset

   rsGRID.Open "SELECT * FROM ALISMReceipt WHERE DocumentNo =  '" & Screen.ActiveForm.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
   Set Screen.ActiveForm.ReceiptGrid.DataSource = rsGRID

Exit Sub

err:
    UpdateErrorMessage
End Sub

Public Sub clearRECORD()
On Error GoTo err:

        With Screen.ActiveForm
            .txtReceiptNo.Text = ""
            .txtReceiptDate.Text = ""
            .cboReceiptType.Text = ""
            .txtCostCenter.Text = ""
            .cboDocumentNo.Text = ""
            .txtExpectedAmount.Text = ""
            .txtReceiptAmount.Text = ""
            .txtAccountNo.Text = ""
            .txtAccountingPeriod.Text = ""
            .txtPaymentStatus.Text = ""
            .cboAccountNo.Text = ""
            .cboEmployer.Text = ""
            .txtDateOfLastPayment.Text = ""
            .txtReceivedTodate.Text = ""
            .txtNames.Text = ""
            .txtUnitCount.Text = ""
            .txtSuspenseAccount.Text = ""
            .txtUnitsPaid.Text = ""
            .txtUnitCountBeforePayment.Text = ""
            .txtDueDate.Text = ""
            .txtStatusCode.Text = ""
            .txtReceivedTodate.Text = ""
            .txtAccountNo.Text = ""
            .cboPaymentMethod.Text = ""
            .txtChequeNo.Text = ""
            .cboBankNo.Text = ""
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub enableDATAENTRY()
On Error GoTo err

    bVAL = False
    
    Call EnableReceipt
    
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub disableDATAENTRY()
On Error GoTo err

    bVAL = True
    
    Call EnableReceipt
    
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub EnableReceipt()
On Error GoTo err


        With Screen.ActiveForm
            .txtReceiptNo.Locked = True
            .txtReceiptDate.Locked = True
            .cboReceiptType.Locked = bVAL
            '.txtCostcenter.Locked = Bval
            .txtReceiptAmount.Locked = bVAL
            .cboDocumentNo.Locked = bVAL
            .txtAccountNo.Locked = bVAL
            .txtExpectedAmount.Locked = True
            .txtAccountingPeriod.Locked = True
            .txtPaymentStatus.Locked = True
            .txtUnitCount.Locked = True
            .txtReceivedTodate.Locked = True
            .txtUnitCountBeforePayment.Locked = True
            .txtDueDate.Locked = True
            .txtStatusCode.Locked = True
            .txtSuspenseAccount.Locked = True
            .txtUnitsPaid.Locked = True
            .txtAccountNo.Locked = bVAL
            .cboAccountNo.Locked = bVAL
            .cboEmployer.Locked = bVAL
            .txtNames.Locked = True
            .txtDateOfLastPayment.Locked = True
            .cboPaymentMethod.Locked = bVAL
            .txtChequeNo.Locked = bVAL
            .cboBankNo.Locked = bVAL

        End With
    
Exit Sub

err:
    ErrorMessage
End Sub



Public Sub calculateLOANAMOUNTDUE()
On Error GoTo err
'/ Subtract todays date from the commencement date of the Policy then
'/ calculate the expected Premiums
        
        Dim rsLOANDUE As ADODB.Recordset
        Set rsLOANDUE = New Recordset
            
        rsLOANDUE.Open "SELECT * FROM ALISMLoanManagement Where LoanNo = '" & Screen.ActiveForm.cboDocumentNo & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsLOANDUE
                If .EOF Or .BOF Then Exit Sub
                
                !UnitsDue = DateDiff("M", !CommencementDate, Date)
                
                If IsNull(!UnitsDue) = True Then !UnitsDue = 0
                If IsNull(!UnitCount) = True Then !UnitCount = 0
                
                If !UnitsDue < 0 Then
                        !DueDate = DateAdd("M", !UnitCount, !CommencementDate)
                        !AmountDue = 0
                    
                    Else: !DueDate = DateAdd("M", (!UnitCount + !UnitsDue), !CommencementDate)
                    
                    !AmountDue = !UnitsDue * !PrincipalAmount
                    !CurrentBalance = !AmountDue
                
                End If
                
                .Update
        End With
        
        rsLOANDUE.Close
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub calculateAMOUNTDUE()
On Error GoTo err
        

        '/ Subtract todays date from the commencement date of the Policy then
        '/ calculate the expected Premiums
        
        Dim rsPOLICY As ADODB.Recordset
        Set rsPOLICY = New Recordset
            
        rsPOLICY.Open "SELECT * FROM ODASMJobBrief Where JobBriefNo = '" & Screen.ActiveForm.cboDocumentNo & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        '/Obtain the payment Method
        
        Dim rsPM As ADODB.Recordset, strPM As String
        Set rsPM = New Recordset
            
        rsPM.Open "SELECT * FROM ODASPPaymentMode where paymentMode = '" & rsPOLICY!PaymentMode & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsPOLICY
                If .EOF Or .BOF Then Exit Sub
                
                Dim daysm As Double
                
                If DateDiff("d", Date, rsPOLICY!MaturityDate) > 0 Then
                        !UnitsDue = DateDiff("M", rsPOLICY!DateOfCommencement, Date)
                Else
                        !UnitsDue = DateDiff("M", rsPOLICY!DateOfCommencement, rsPOLICY!MaturityDate)
                End If
                        '/ * Check for Maturity Date
                        
                !UnitsDue = (Fix(!UnitsDue / rsPM!CoverPeriod) * rsPM!CoverPeriod) - !UnitCount
                
                If !UnitsDue < 0 Then
                        
                        !DueDate = DateAdd("M", !UnitCount, !DateOfCommencement)
                        !AmountDue = 0
                    
                    Else: !DueDate = DateAdd("M", (!UnitCount + !UnitsDue), !DateOfCommencement)
                    
                    !AmountDue = (!UnitsDue / rsPM!CoverPeriod) * !ExpectedPremium
                   ' !Currentbalance = !AmountDue
                
                End If
                
                .Update
        End With
        
        rsPM.Close
        rsPOLICY.Close
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub updateRECORD()
On Error GoTo err
        
        Dim rsPOL As ADODB.Recordset, strPOL As String
        Set rsPOL = New Recordset
            
        rsPOL.Open "SELECT * FROM ODASMJobBrief WHERE JobBriefNo =  '" & Screen.ActiveForm.cboDocumentNo & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsPOL
                If .EOF Or .BOF = True Then
                            MsgBox "The Policy No Entered Does not exist within the Database", vbOKOnly
                            Screen.ActiveForm.cboDocumentNo.SetFocus
                            Exit Sub
                End If
                Dim rsPM As ADODB.Recordset, strPM As String
                Set rsPM = New Recordset
                    
                rsPM.Open "SELECT * FROM ODASPPaymentMode Where paymentMode = '" & !PaymentMode & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

    
                '/* this procedure is used to intialize the variable
                
                Screen.ActiveForm.txtSuspenseAccount.Text = !SuspenseAccount
                Screen.ActiveForm.txtUnitCount.Text = !UnitCount
                Screen.ActiveForm.txtReceivedTodate.Text = !ReceivedToDate & ""
                
                                
                '/* this Procedure is used to calculate the number of Premiums
                '/* paid by the Policyholder
                
                If (CDbl(Screen.ActiveForm.txtReceiptAmount.Text) >= !ExpectedPremium) And (!UnitCount < !TermOfPolicy * 12) And (Trim(rsPOL!StatusCode) = "10" Or Trim(rsPOL!StatusCode) = "INFORCE") Then

                        Screen.ActiveForm.txtUnitsPaid = Fix(Screen.ActiveForm.txtReceiptAmount / Screen.ActiveForm.txtExpectedAmount) * rsPM!CoverPeriod
                        
                        If Screen.ActiveForm.txtUnitCount <= "" Then
                            Screen.ActiveForm.txtUnitCount = 0
                        End If
                        
                        If Screen.ActiveForm.txtSuspenseAccount <= "" Then
                            Screen.ActiveForm.txtSuspenseAccount = 0
                        End If
                        
                        Screen.ActiveForm.txtUnitCountBeforePayment = CDbl(Screen.ActiveForm.txtUnitCount)
                        Screen.ActiveForm.txtUnitCount = !UnitCount + CDbl(Screen.ActiveForm.txtUnitsPaid.Text)
                        Screen.ActiveForm.txtDateOfLastPayment = Screen.ActiveForm.txtReceiptDate
                        Screen.ActiveForm.txtSuspenseAccount = !SuspenseAccount + CDbl(Screen.ActiveForm.txtReceiptAmount - (Screen.ActiveForm.txtUnitsPaid * (Screen.ActiveForm.txtExpectedAmount / rsPM!CoverPeriod)))
                        
                        If Screen.ActiveForm.txtReceivedTodate <= "" Then
                                        Screen.ActiveForm.txtReceivedTodate = CDbl(Screen.ActiveForm.txtReceiptAmount)
                                Else: Screen.ActiveForm.txtReceivedTodate = !ReceivedToDate + CDbl(Screen.ActiveForm.txtReceiptAmount)
                        End If
                        
                        '/If the total premiums exceed the expected number of Premiums
                
                        If CDbl(Screen.ActiveForm.txtUnitCount) >= !TermOfPolicy * 12 Then
                                Screen.ActiveForm.txtUnitsPaid = CDbl(Screen.ActiveForm.txtUnitsPaid) - (Screen.ActiveForm.txtUnitCount - (!TermOfPolicy * 12))
                                Screen.ActiveForm.txtUnitCount = !TermOfPolicy * 12
                                Screen.ActiveForm.txtStatusCode = "MATURE"
                        End If
                        
                        Screen.ActiveForm.txtDueDate = DateAdd("M", CDbl(Screen.ActiveForm.txtUnitCount), !DateOfCommencement)
                                               
                        '/* check if the Maturity date been reached
                
                ElseIf (Screen.ActiveForm.txtReceiptAmount < Screen.ActiveForm.txtExpectedAmount) And (Screen.ActiveForm.txtUnitCount < !TermOfPolicy * 12) Then
                        Screen.ActiveForm.txtSuspenseAccount = CDbl(!SuspenseAccount) + CDbl(Screen.ActiveForm.txtReceiptAmount)
                        Screen.ActiveForm.txtUnitsPaid.Text = 0
                        Screen.ActiveForm.txtStatusCode.Text = !StatusCode
                        Screen.ActiveForm.txtUnitCount.Text = !UnitCount
                        Screen.ActiveForm.txtReceivedTodate = !ReceivedToDate + Screen.ActiveForm.txtReceiptAmount
                        Screen.ActiveForm.txtDueDate.Text = !DueDate
                        
                Else:   Screen.ActiveForm.txtSuspenseAccount = !SuspenseAccount + CDbl(Screen.ActiveForm.txtReceiptAmount)
                        Screen.ActiveForm.txtUnitsPaid.Text = 0
                        Screen.ActiveForm.txtStatusCode.Text = !StatusCode
                        Screen.ActiveForm.txtUnitCount.Text = !UnitCount
                        Screen.ActiveForm.txtReceivedTodate.Text = !ReceivedToDate + CDbl(Screen.ActiveForm.txtReceiptAmount)
                        Screen.ActiveForm.txtDueDate.Text = !DueDate
      End If

rsPOL.Close
strPOL = ""
rsPM.Close
strPM = ""

End With


Exit Sub


err:
    ErrorMessage
End Sub

Public Sub updateLOANRECORD()
On Error GoTo err
        
        Dim rsPM As ADODB.Recordset, strPM As String
        Set rsPM = New Recordset
            
        rsPM.Open "SELECT * FROM ODASPPaymentMode;", cnCOMMON, adOpenKeyset, adLockOptimistic

                
        Dim rsPOL As ADODB.Recordset, strPOL As String
        Set rsPOL = New Recordset
            
        rsPOL.Open "SELECT * FROM ALISMLoanManagement WHERE LoanNo =  '" & Screen.ActiveForm.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsPOL
                '/* this procedure is used to intialize the variable
                
                Screen.ActiveForm.txtSuspenseAccount.Text = !SuspenseAccount & ""
                Screen.ActiveForm.txtUnitCount.Text = !CurrentBalance
                Screen.ActiveForm.lblUnitCount.Caption = "Interest/Loan Paid"
                Screen.ActiveForm.txtReceivedTodate.Text = !TotalReceivedTodate & ""
                
                                
                '/* this Procedure is used to calculate the Amount to be charge to the
                '/* interest and Principal account
                
                If (CDbl(Screen.ActiveForm.txtReceiptAmount.Text) >= !RepaymentAmount) Then '/ And (!CurentBalance > 0) Then

                        Screen.ActiveForm.txtUnitsPaid = Fix(Screen.ActiveForm.txtReceiptAmount / Screen.ActiveForm.txtExpectedAmount) * rsPM!CoverPeriod
                        
                        If Screen.ActiveForm.txtUnitCount <= "" Then
                            Screen.ActiveForm.txtUnitCount = 0
                        End If
                        
                        If Screen.ActiveForm.txtSuspenseAccount <= "" Then
                            Screen.ActiveForm.txtSuspenseAccount = 0
                        End If
                        
                        Screen.ActiveForm.txtUnitCountBeforePayment = CDbl(Screen.ActiveForm.txtUnitCount)
                     '   Screen.Activeform.txtUnitCount = !UnitCount + cdbl(Screen.Activeform.txtUnitsPaid.Text)
                        Screen.ActiveForm.txtDateOfLastPayment = Screen.ActiveForm.txtReceiptDate
                        Screen.ActiveForm.txtSuspenseAccount = !SuspenseAccount & ""
                        
                        If Screen.ActiveForm.txtReceivedTodate <= "" Then
                                Screen.ActiveForm.txtReceivedTodate = CDbl(Screen.ActiveForm.txtReceiptAmount)
                        Else: Screen.ActiveForm.txtReceivedTodate = !TotalReceivedTodate + CDbl(Screen.ActiveForm.txtReceiptAmount)
                        End If
                        '/If the total Units exceed the expected number of Units
                
                        If CDbl(Screen.ActiveForm.txtUnitCount) >= !RepaymentPeriod Then
                                Screen.ActiveForm.txtUnitsPaid = Screen.ActiveForm.txtUnitsPaid - (Screen.ActiveForm.txtUnitCount - !RepaymentPeriod)
                                Screen.ActiveForm.txtUnitCount = !RepaymentPeriod
                                Screen.ActiveForm.txtStatusCode = "COMPLETE"
                        End If
                        
                        Screen.ActiveForm.txtDueDate = DateAdd("M", CDbl(Screen.ActiveForm.txtUnitCount), !CommencementDate)
                                               
                        '/* check if the Maturity date been reached
                
                ElseIf (Screen.ActiveForm.txtReceiptAmount < Screen.ActiveForm.txtExpectedAmount) And (Screen.ActiveForm.txtUnitCount < !TermOfPolicy * 12) Then
                        Screen.ActiveForm.txtSuspenseAccount = CDbl(!SuspenseAccount) + CDbl(Screen.ActiveForm.txtReceiptAmount)
                        Screen.ActiveForm.txtUnitsPaid.Text = 0
                        Screen.ActiveForm.txtStatusCode.Text = !StatusCode
                        Screen.ActiveForm.txtUnitCount.Text = !UnitCount
                        Screen.ActiveForm.txtReceivedTodate = !ReceivedToDate + Screen.ActiveForm.txtReceiptAmount
                        Screen.ActiveForm.txtDueDate.Text = !DueDate
                        
                Else:   Screen.ActiveForm.txtSuspenseAccount = !SuspenseAccount + CDbl(Screen.ActiveForm.txtReceiptAmount)
                        Screen.ActiveForm.txtSuspenseAccount = !SuspenseAccount + CDbl(Screen.ActiveForm.txtReceiptAmount)
                        Screen.ActiveForm.txtUnitsPaid.Text = 0
                        Screen.ActiveForm.txtStatusCode.Text = !StatusCode
                        Screen.ActiveForm.txtUnitCount.Text = !UnitCount
                        Screen.ActiveForm.txtReceivedTodate.Text = !ReceivedToDate + CDbl(Screen.ActiveForm.txtReceiptAmount)
                        Screen.ActiveForm.txtDueDate.Text = !DueDate
      End If

End With


Exit Sub

err:
   ErrorMessage
End Sub

Public Sub updateEMPLOYER()
On Error GoTo err
                
    Screen.ActiveForm.txtSuspenseAccount.Text = 0
    Screen.ActiveForm.txtUnitCount.Text = 0
    Screen.ActiveForm.txtReceivedTodate.Text = 0
    Screen.ActiveForm.txtUnitsPaid = 1
    Screen.ActiveForm.txtUnitCountBeforePayment = 0
    Screen.ActiveForm.txtUnitCount = 0
    Screen.ActiveForm.txtDateOfLastPayment = 0
    Screen.ActiveForm.txtReceivedTodate = CDbl(Screen.ActiveForm.txtReceiptAmount)
                           
Exit Sub


err:
    ErrorMessage
End Sub

Public Sub loadRECORD()
On Error GoTo err
    With rsDEP
            Screen.ActiveForm.txtReceiptNo.Text = !ReceiptNo
            Screen.ActiveForm.txtReceiptDate.Text = !ReceiptDate
            Screen.ActiveForm.cboReceiptType.Text = !ReceiptType
            Screen.ActiveForm.txtCostCenter.Text = !CostCenter
            Screen.ActiveForm.cboDocumentNo.Text = !DocumentNo
            Screen.ActiveForm.txtReceiptAmount.Text = !ReceiptAmount
            Screen.ActiveForm.txtAccountNo.Text = !AccountNo
            Screen.ActiveForm.txtAccountingPeriod.Text = !AccountingPeriod
            Screen.ActiveForm.txtPaymentStatus.Text = !PaymentStatus
            Screen.ActiveForm.cboAccountNo.Text = !AccountCode & ""
            Screen.ActiveForm.txtDateOfLastPayment.Text = !DateofLastPayment
            Screen.ActiveForm.txtReceivedTodate.Text = !ReceivedToDate
            Screen.ActiveForm.txtUnitCount.Text = !UnitCount
            Screen.ActiveForm.txtSuspenseAccount.Text = !SuspenseAccount & ""
            Screen.ActiveForm.txtUnitsPaid.Text = !UnitsPaid
            Screen.ActiveForm.txtUnitCountBeforePayment.Text = !UnitCountBeforePayment
            Screen.ActiveForm.txtDueDate.Text = !DueDate & ""
            Screen.ActiveForm.txtStatusCode.Text = !StatusCode
            Screen.ActiveForm.cboPaymentMethod.Text = !PaymentMethod & ""
            Screen.ActiveForm.txtChequeNo.Text = !ChequeNo & ""
            Screen.ActiveForm.cboBankNo.Text = !BankNo & ""

    End With
    
    LoadNAMES
    loadEMPLOYER
    
rsDEP.Close
strDEP = ""

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub LoadNAMES()
On Error GoTo err
    
    Dim rsREF As ADODB.Recordset, strREF As String
    Set rsREF = New Recordset
    
    strREF = "SELECT * FROM ODASPAccount Where AccountNo LIKE '" & Screen.ActiveForm.txtAccountNo.Text & "';"
    rsREF.Open strREF, cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsREF
            If .BOF Or .EOF Then
                    MsgBox "The Policyholder Does not Exist", vbOKOnly
                    Exit Sub
            End If
    
            Screen.ActiveForm.txtNames = Trim(!OtherNames) + " " + Trim(!CompanyName)
            Screen.ActiveForm.txtAccountNo.Text = !AccountNo & ""
    End With
    
    rsREF.Close

Exit Sub

err:
        ErrorMessage
End Sub

Public Sub createRECEIPT()
On Error GoTo err
        Dim rsLAST As ADODB.Recordset, strLAST As String
        
        Set rsLAST = New Recordset
      
        strLAST = "SELECT * FROM ALISPLastNumber;"
        rsLAST.Open strLAST, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsLAST
            Screen.ActiveForm.txtReceiptNo = !ReceiptNo
            !ReceiptNo = !ReceiptNo + 1
            .Update
        End With

Exit Sub

err:
    UpdateErrorMessage
End Sub
Public Sub selectAccountNoLostFocus()
On Error GoTo err:

        If Screen.ActiveForm.txtAccountNo.Text <= "" Then Exit Sub
        
        Dim rsEMPLOYEE As ADODB.Recordset, strEMPLOYEE As String
        Set rsEMPLOYEE = New Recordset

        strEMPLOYEE = "SELECT * FROM ODASPAccount Where AccountNo = '" & Screen.ActiveForm.cboAccountNo & "' and AccountNo = '" & Screen.ActiveForm.txtAccountNo & "';"
        rsEMPLOYEE.Open strEMPLOYEE, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If Screen.ActiveForm.cboAccountNo.Text > "" Then
        With rsEMPLOYEE
        
                If .EOF Or .BOF = True Then
                        MsgBox "The Employee entered does not exist within the database", vbOKOnly
                                Screen.ActiveForm.txtAccountNo.SetFocus
                        Exit Sub
                Else
                        Dim rsPROPOSAL As ADODB.Recordset, strPROPOSAL As String
                
                        Dim rsPOLICY As ADODB.Recordset, strPOLICY As String
                        Set rsPOLICY = New Recordset
                        
                        strPOLICY = "SELECT * FROM ODASMJobBrief Where AccountNo = '" & rsEMPLOYEE!AccountNo & "' ;"
                        rsPOLICY.Open strPOLICY, cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                        With rsPOLICY
                                If .EOF Or .BOF = True Then
                                        
                                        '/ Check whether the policy holder has any proposal that has not been issues
                                        Set rsPROPOSAL = New Recordset
                                        
                                        strPROPOSAL = "SELECT * FROM ALISMProposal Where AccountNo = '" & rsEMPLOYEE!AccountNo & "' ;"
                                        rsPROPOSAL.Open strPROPOSAL, cnCOMMON, adOpenKeyset, adLockOptimistic
                                        
                                        With rsPROPOSAL
                                        
                                            If .EOF Or .BOF = True Then
                                                       MsgBox "This Employee does not have either Policy or a Proposal", vbOKOnly
                                                                Screen.ActiveForm.txtAccountNo.SetFocus
                                                       Exit Sub
                                            Else
                                                    MsgBox "This is the Proposal No and not the Policy number change the Receipt Type to Deposit", vbOKOnly
                                                        Screen.ActiveForm.cboDocumentNo.Text = rsPROPOSAL!ProposalNo
                                                        Exit Sub
                                            End If
                                            
                                        End With
                                                
                                        Exit Sub
                                
                                End If
                                
                                If .RecordCount > 1 Then
                                        MsgBox "The PolicyHolder has more than one Policy That is inforce", vbOKOnly
                                        Screen.ActiveForm.cboDocumentNo = rsPOLICY!JobBriefNo
                                
                                ElseIf .RecordCount = 1 Then
                                        Screen.ActiveForm.cboDocumentNo = rsPOLICY!JobBriefNo
                                End If
                                
                        End With
                                 '/ With Policy
                End If
                
        End With
        End If
                        '/ End with Employee
rsEMPLOYEE.Close
strEMPLOYEE = ""
rsPOLICY.Close
strPOLICY = ""

Exit Sub

err:
    UpdateErrorMessage
End Sub

Public Sub enableCBRECEIPT()
On Error GoTo err
    With Screen.ActiveForm
        .cmdAddNew.Enabled = True
        .cmdUpdate.Enabled = False
        .cmdSearch.Enabled = True
        .cmdCancel.Enabled = True
        .cmdprintreceipt.Enabled = True
    End With
    Exit Sub
err:
ErrorMessage
   
End Sub

Public Sub DisableCBRECEIPT()
    With Screen.ActiveForm
        .cmdAddNew.Enabled = False
        .cmdUpdate.Enabled = True
        .cmdSearch.Enabled = False
        .cmdCancel.Enabled = True
        .cmdprintreceipt.Enabled = False
    End With
End Sub

Public Sub SelectAGENTKEYPRESS(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub

Public Sub selectPaymentMethodGotFocus()
On Error GoTo err

        Dim rsPAYMETHOD As ADODB.Recordset, strPAYMETHOD As String
        Set rsPAYMETHOD = New Recordset
      
        strPAYMETHOD = "SELECT * FROM ALISPPaymentMethod;"
        rsPAYMETHOD.Open strPAYMETHOD, cnCOMMON, adOpenKeyset, adLockOptimistic

        Screen.ActiveForm.cboPaymentMethod.Clear

            With rsPAYMETHOD
                    Do Until .EOF
                        Screen.ActiveForm.cboPaymentMethod.AddItem !PaymentMethodDescription
                        .MoveNext
                    Loop
            End With

rsPAYMETHOD.Close
strPAYMETHOD = ""

Exit Sub

err:
    UpdateErrorMessage
End Sub

Public Sub selectPaymentMethodKeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Public Sub selectPaymentMethodLostFocus()
On Error GoTo err

        Dim rsPAYMETHOD As ADODB.Recordset, strPAYMETHOD As String
        Set rsPAYMETHOD = New Recordset
        
        rsPAYMETHOD.Open "SELECT * FROM ALISPPaymentMethod WHERE PaymentMethodDescription= '" & Screen.ActiveForm.cboPaymentMethod.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsPAYMETHOD
                If .EOF And .BOF Then Exit Sub
                        Screen.ActiveForm.cboPaymentMethod.Text = !PaymentMethod
                        
                        If !RequireAccountDetails = "Yes" Then
                                bRequireAccountNo = True
                        Else: bRequireAccountNo = False
                        End If
        End With
        
        If bRequireAccountNo = True Then
            Screen.ActiveForm.txtChequeNo.Locked = False
            Screen.ActiveForm.cboBankNo.Locked = False
            Screen.ActiveForm.cboBankNo.SetFocus
        Else
            Screen.ActiveForm.txtChequeNo.Locked = True
            Screen.ActiveForm.cboBankNo.Locked = True
        
        End If


rsPAYMETHOD.Close
strPAYMETHOD = ""

Exit Sub

err:
    ErrorMessage

End Sub

Private Sub loadEMPLOYER()
On Error GoTo err
       Dim rsEMPLOYER As ADODB.Recordset, strEMPLOYER As String
        Set rsEMPLOYER = New Recordset
        
        rsEMPLOYER.Open "SELECT * FROM ODASPAccount WHERE AccountNo = '" & Screen.ActiveForm.cboAccountNo.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
 
        With rsEMPLOYER
                If .EOF And .BOF Then Exit Sub
                            Screen.ActiveForm.cboEmployer.Text = !CompanyName & ""
        End With
       

rsEMPLOYER.Close
strEMPLOYER = ""

Exit Sub

err:
    UpdateErrorMessage
End Sub

Public Sub selectBankNOGotFocus()
On Error GoTo err
      Dim rsECGF As ADODB.Recordset, strECGF As String
      Set rsECGF = New Recordset
      strECGF = "SELECT * FROM ALISPBank;"
      rsECGF.Open strECGF, cnCOMMON, adOpenKeyset, adLockOptimistic
      Screen.ActiveForm.cboBankNo.Clear
        With rsECGF
        Do Until .EOF
        If Not IsNull(!CompanyName) And !CompanyName <> "" Then
            Screen.ActiveForm.cboBankNo.AddItem !CompanyName
        End If
        .MoveNext
        Loop
        End With
      Exit Sub
err:
    UpdateErrorMessage
End Sub

Public Sub selectBankNoKeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Public Sub selectBankNoLostFocus()
On Error GoTo err

        Dim rsECLF As ADODB.Recordset, strECLF As String
        Set rsECLF = New Recordset
        
        rsECLF.Open "SELECT * FROM ALISPBank WHERE CompanyName = '" & Screen.ActiveForm.cboBankNo.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsECLF
                If .EOF And .BOF Then Exit Sub
                            Screen.ActiveForm.cboBankNo.Text = !BankNo
        End With
        
        Screen.ActiveForm.txtChequeNo.SetFocus
        
Exit Sub

err:
    ErrorMessage

End Sub

Public Sub selectCostCenterGotFocus()
    
    Dim rsPAY As ADODB.Recordset, strPAY As String
    Set rsPAY = New Recordset
    
    strPAY = "SELECT * FROM ODASPCostCentres;"
    rsPAY.Open strPAY, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Screen.ActiveForm.txtCostCenter.Clear

    With rsPAY
            Do Until .EOF
                    Screen.ActiveForm.txtCostCenter.AddItem !CostCenterDescription
                    .MoveNext
            Loop
    
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub selectCOSTCENTERKEYPRESS(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub

Public Sub selectCostCenterLostFocus()
On Error GoTo err
    Dim rsCOSTCENTER As ADODB.Recordset, strCOSTCENTER As String
    Set rsCOSTCENTER = New Recordset
    rsCOSTCENTER.Open "SELECT * FROM ODASPCostCentres WHERE CostCenterDescription= '" & Screen.ActiveForm.txtCostCenter.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
    With rsCOSTCENTER
        If .EOF And .BOF Then Exit Sub
        Screen.ActiveForm.txtCostCenter.Text = !CostCenterCode
    End With
    Exit Sub
err:
    ErrorMessage
End Sub

Public Sub SelectEmployerGotFocus()
On Error GoTo err
If Screen.ActiveForm.cboEmployer.ListCount <> 0 Then Exit Sub
    Dim rsECGF As ADODB.Recordset, strECGF As String
    Set rsECGF = New Recordset
    strECGF = "SELECT * FROM ODASPAccount;"
    rsECGF.Open strECGF, cnCOMMON, adOpenKeyset, adLockOptimistic
    Screen.ActiveForm.cboEmployer.Clear
    With rsECGF
        Do Until .EOF
            Screen.ActiveForm.cboEmployer.AddItem !CompanyName
            .MoveNext
        Loop
    End With
rsECGF.Close
strECGF = ""
Exit Sub
err:
    UpdateErrorMessage
End Sub

Public Sub selectEmployerLostFocus()
On Error GoTo err
    Dim rsECLF As ADODB.Recordset, strECLF As String
    Set rsECLF = New Recordset
    rsECLF.Open "SELECT * FROM ODASPAccount WHERE CompanyName= '" & Screen.ActiveForm.cboEmployer.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
    With rsECLF
    If .EOF And .BOF Then Exit Sub
        Screen.ActiveForm.cboAccountNo.Text = !AccountCode & ""
        Screen.ActiveForm.cboEmployer.Text = !CompanyName & ""
        Screen.ActiveForm.txtPayer.Text = !CompanyName & ""
    If bEmployerUpdate = True Then
        Screen.ActiveForm.cboDocumentNo.Text = Screen.ActiveForm.cboAccountNo
    End If
    End With
rsECLF.Close
strECLF = ""
Exit Sub
err:
    UpdateErrorMessage
End Sub

Public Sub selectRECEIPTTYPEGOTFOCUS()
On Error GoTo err
    Set rsRCPT = New Recordset
    strSQL = "SELECT * FROM ALISPReceipt;"
    rsRCPT.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    Screen.ActiveForm.cboReceiptType.Clear
    With rsRCPT
        Do Until .EOF
            Screen.ActiveForm.cboReceiptType.AddItem !Description
            .MoveNext
        Loop
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub selectRECEIPTTYPEKEYPRESS(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Public Sub selectRECEIPTTYPELOSTFOCUS()
On Error GoTo err
    Set rsRCPT = New ADODB.Recordset
    rsRCPT.Open "SELECT * FROM ALISPReceipt WHERE Description = '" & Screen.ActiveForm.cboReceiptType.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
    With rsRCPT
            If .EOF And .BOF Then Exit Sub
            Screen.ActiveForm.cboReceiptType.Text = !ReceiptType
            '/ Enable the Employer Code Control if the receipt type is set to
            '/ Employer = "1" else locks it up
            If !Account = "1" Then
                With Screen.ActiveForm
                    .cboAccountNo.Locked = True
                    .txtAccountNo.Locked = False
                    .cboEmployer.Locked = False
                    bEmployerUpdate = False
                    .lblDocumentNo.Caption = "Employer Code"
                End With
            ElseIf !BreakDown = "1" Then
                With Screen.ActiveForm
                    .cboAccountNo.Locked = True
                    .txtAccountNo.Locked = False
                    bEmployerUpdate = True
                    .cboEmployer.Locked = False
                    .lblDocumentNo.Caption = "Employer Code"
                End With
            Else
                With Screen.ActiveForm
                    .cboAccountNo.Locked = True
                    .txtAccountNo.Locked = True
                    bEmployerUpdate = False
                    .cboEmployer.Locked = False
                End With
            End If
    End With
    '/ This portion is used to set the default details for the receipting
    '/ Module
    Screen.ActiveForm.txtPaymentStatus.Text = "PAID"
    Screen.ActiveForm.txtPaymentStatus.FontBold = True
    Screen.ActiveForm.txtPaymentStatus.Locked = True
    Screen.ActiveForm.txtAccountingPeriod = CurrentPeriod
    Screen.ActiveForm.txtAccountingPeriod.FontBold = True
    Screen.ActiveForm.txtAccountingPeriod.Locked = True
    Screen.ActiveForm.cboAccountNo.FontBold = True
    Screen.ActiveForm.txtReceiptDate.Text = Date
    Screen.ActiveForm.txtReceiptDate.FontBold = True
    Screen.ActiveForm.txtReceiptDate.Locked = True
    Screen.ActiveForm.txtReceiptNo.FontBold = True
loadRECEIPTTYPE
Exit Sub
err:
    ErrorMessage
End Sub


Public Sub addRECORD()
''''On Error GoTo Myerr
    clearRECORD
    enableDATAENTRY
    DisableCBRECEIPT
    Screen.ActiveForm.txtReceiptNo.SetFocus
    Call LoadReceiptGrid

Exit Sub

Myerr:
    ErrorMessage
End Sub


Public Sub selectCURRENCYCODEKEYPRESS(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub


Public Sub selectPAYMENTMODEKEYPRESS(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Public Sub selectPlanProposalNo_KeyPress(KeyAscii As Integer)

   KeyAscii = 0
    
End Sub

Public Sub processDOCUMENTNO()
On Error GoTo err

        If Screen.ActiveForm.cboDocumentNo.Text <= "" Then Exit Sub
        
        If Screen.ActiveForm.cboReceiptType.Text <= "" Then
                            MsgBox "The Receipt Type Is Required Before executive This Function", vbOKOnly
                            Screen.ActiveForm.cboReceiptType.SetFocus
                            Exit Sub
                            
        ElseIf Screen.ActiveForm.txtCostCenter.Text <= "" Then
                            MsgBox "The Cost Center Is Required Before executive This Function", vbOKOnly
                            Screen.ActiveForm.txtCostCenter.SetFocus
                            Exit Sub
        Else

                If Screen.ActiveForm.cboReceiptType > "" Then
                        
                        
                        With rsRCPT
                            
                                If !Deposit = 1 Then
                            
                                    Dim rsDNLF As ADODB.Recordset, strDNLF As String
                                    Set rsDNLF = New Recordset
                            
                                    rsDNLF.Open "SELECT * FROM ALISMProposal WHERE ProposalNo= '" & Screen.ActiveForm.cboDocumentNo.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
                      
                                    With rsDNLF
                                        If .EOF And .BOF Then Exit Sub
                                        
                                        Screen.ActiveForm.cboDocumentNo.Text = !ProposalNo
                                        Screen.ActiveForm.txtAccountNo.Text = !AccountNo
                                        Screen.ActiveForm.txtDateOfLastPayment.Text = !DateofLastPayment & ""
                                        Screen.ActiveForm.txtReceiptAmount.Text = !ExpectedPremium
                                        Screen.ActiveForm.txtExpectedAmount.Text = !ExpectedPremium
                                        Screen.ActiveForm.txtDueDate.Text = Date
                                        
                                        If IsNull(!SuspenseAccount) = True Then
                                                Screen.ActiveForm.txtSuspenseAccount = 0
                                                Else: Screen.ActiveForm.txtSuspenseAccount.Text = !SuspenseAccount
                                        End If
                                        
                                        If IsNull(!DepositReceived) = True Then
                                                Screen.ActiveForm.txtReceivedTodate.Text = 0
                                            Else: Screen.ActiveForm.txtReceivedTodate.Text = !DepositReceived
                                        End If
                                        
                                        Screen.ActiveForm.txtStatusCode.Text = !StatusCode
                                        BPolicyHolder = True
                                    End With
                            
                                    rsDNLF.Close
                                    strDNLF = ""
                                          '/* rsRCPT is deposit */
                            
                            ElseIf !Jobbrief = 1 Or !Account = 1 Then
                    
                                    Dim rsPOL As ADODB.Recordset, strPOL As String
                                    Set rsPOL = New Recordset
                            
                                    rsPOL.Open "SELECT * FROM ODASMJobBrief WHERE JobBriefNo= '" & Screen.ActiveForm.cboDocumentNo.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
                      
                                    With rsPOL
                                        If .EOF And .BOF = True Then
                                                MsgBox "The Policy Does not exsits", vbOKOnly
                                                        Screen.ActiveForm.cboDocumentNo.SetFocus
                                                Exit Sub
                                        Else
                                                calculateAMOUNTDUE
                                                Screen.ActiveForm.cboDocumentNo.Text = !JobBriefNo
                                                Screen.ActiveForm.txtReceiptAmount = 0
                                                Screen.ActiveForm.txtExpectedAmount = !ExpectedPremium
                                                Screen.ActiveForm.txtAccountNo.Text = !AccountNo
                                                Screen.ActiveForm.txtDateOfLastPayment = !DateofLastPayment & ""
                                                Screen.ActiveForm.txtUnitCountBeforePayment = !UnitCount & ""
                                                Screen.ActiveForm.txtDueDate = !DueDate & ""
                                                Screen.ActiveForm.txtSuspenseAccount = !SuspenseAccount & ""
                                                Screen.ActiveForm.txtUnitCount = !UnitCount & ""
                                                Screen.ActiveForm.txtUnitsPaid = 0
                                                Screen.ActiveForm.txtStatusCode = !StatusCode & ""
                                                Screen.ActiveForm.txtReceivedTodate = !ReceivedToDate & ""
                                                BPolicyHolder = True
                                        End If
                                    End With
                            
                                    rsPOL.Close
                                    strPOL = ""
                                            '/* rsRCPT is deposit */
                                            
                            ElseIf !IssuePolicy = 1 Then
                    
                                    Dim rsISSUE As ADODB.Recordset, strISSUE As String
                                    Set rsISSUE = New Recordset
                            
                                    rsISSUE.Open "SELECT * FROM ODASMJobBrief WHERE JobBriefNo= '" & Screen.ActiveForm.cboDocumentNo.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
                      
                                    With rsISSUE
                                        If .EOF And .BOF Then Exit Sub
                                        Screen.ActiveForm.cboDocumentNo.Text = !JobBriefNo
                                        Screen.ActiveForm.txtReceiptAmount = !DepositReceived
                                        Screen.ActiveForm.txtReceiptAmount.Locked = True
                                        Screen.ActiveForm.txtExpectedAmount = !ExpectedPremium
                                        Screen.ActiveForm.txtAccountNo.Text = !AccountNo
                                        Screen.ActiveForm.txtDateOfLastPayment = !DateofLastPayment & ""
                                        Screen.ActiveForm.txtReceivedTodate = !ReceivedToDate & ""
                                        Screen.ActiveForm.txtUnitCountBeforePayment = 0
                                        Screen.ActiveForm.txtDueDate = DateAdd("M", !UnitCount, !DateOfCommencement)
                                        Screen.ActiveForm.txtSuspenseAccount = 0
                                        Screen.ActiveForm.txtUnitCount = 0
                                        Screen.ActiveForm.txtUnitsPaid = !UnitCount & ""
                                        Screen.ActiveForm.txtStatusCode = "INFORCE"
                                        Screen.ActiveForm.txtReceivedTodate = !DepositReceived & ""
                                        BPolicyHolder = True

                                    End With
                            
                                    rsISSUE.Close
                                    strISSUE = ""
                                            '/* rsRCPT is deposit */
                            ElseIf !loan = 1 Then
                                    
                                    calculateLOANAMOUNTDUE
                                    
                                    Dim rsLOAN As ADODB.Recordset, strLOAN As String
                                    Set rsLOAN = New Recordset
                            
                                    rsLOAN.Open "SELECT * FROM ALISMLoanManagement WHERE LoanNo= '" & Screen.ActiveForm.cboDocumentNo.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
                      
                                    With rsLOAN
                                        If .EOF And .BOF Then Exit Sub
                                        
                                        Screen.ActiveForm.cboDocumentNo.Text = !LoanNo
                                        Screen.ActiveForm.txtReceiptAmount = !RepaymentAmount
                                        Screen.ActiveForm.txtExpectedAmount = !RepaymentAmount
                                        Screen.ActiveForm.txtDateOfLastPayment = !DateofLastPayment & ""
                                        Screen.ActiveForm.txtReceivedTodate = !TotalReceivedTodate & ""
                                        Screen.ActiveForm.txtDueDate = !DueDate & ""
                                        Screen.ActiveForm.txtSuspenseAccount = !SuspenseAccount & ""
                                        Screen.ActiveForm.txtStatusCode = !Status & ""
                                        Screen.ActiveForm.txtAccountNo = !AccountNo
                                        BPolicyHolder = True

                                    End With
                            
                                    rsLOAN.Close
                                    strLOAN = ""
                                    
                                    Else
                                        BPolicyHolder = False

                                    
                                            '/* rsRCPT is Loan */
                            End If
                                    '/* end if the Statement Statement
                        End With
                        
                                    '/* End with Receipt
                                    
                End If
                '/* if Receipt Type is Greater than Zero
                
                Call loadEMPLOYER
                Call ReferenceDetails
                Call LoadReceiptGrid
                Call loadloanGRID
                Call loadproposalGRID
                Call loadpolicyGRID

                Screen.ActiveForm.txtReceiptAmount.SetFocus

        End If
                '/* end if Document Number is Greater Than Zero

Exit Sub

err:
        ErrorMessage

End Sub

Public Sub loadRECEIPTTYPE()
On Error GoTo err
    
    If Screen.ActiveForm.cboReceiptType > "" Then
        
        With rsRCPT
           
                If !Deposit = 1 Then
                        Screen.ActiveForm.lblDocumentNo.Caption = "Proposal No"
                ElseIf !Jobbrief = 1 Then
                        Screen.ActiveForm.lblDocumentNo.Caption = "Policy No"
                ElseIf !IssuePolicy = 1 Then
                        Screen.ActiveForm.lblDocumentNo.Caption = "Policy No"
                ElseIf !loan = 1 Then
                        Screen.ActiveForm.lblDocumentNo.Caption = "Loan No"
                ElseIf !interest Then
                        Screen.ActiveForm.lblDocumentNo.Caption = "Loan No"
                ElseIf !Client = 1 Then
                        Screen.ActiveForm.lblDocumentNo.Caption = "Policy No"
                ElseIf !BreakDown = 1 Then
                        Screen.ActiveForm.lblDocumentNo.Caption = "Employer Code"
                ElseIf !RevivalFee = 1 Then
                        Screen.ActiveForm.lblDocumentNo.Caption = "Revival Fee"
                ElseIf !Account = 1 Then
                        Screen.ActiveForm.lblDocumentNo.Caption = "Policy No"
                ElseIf !Ghost = 1 Then
                        Screen.ActiveForm.lblDocumentNo.Caption = "Proposal No"
                ElseIf !Miscellaneous = 1 Then
                        Screen.ActiveForm.lblDocumentNo.Caption = "Account No"
                Else
                        Screen.ActiveForm.lblDocumentNo.Caption = "Policy No"
                End If
                
        End With
 End If
        '/* receipt Type > ""

Exit Sub

err:
        ErrorMessage

End Sub
      
Private Sub ReferenceDetails()
On Error GoTo err

        Dim rsREF As ADODB.Recordset, strREF As String
        Set rsREF = New Recordset

        rsREF.Open "SELECT * FROM ODASPAccount WHERE AccountNo LIKE '" & Screen.ActiveForm.txtAccountNo & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
                    
        Screen.ActiveForm.txtNames = rsREF!OtherNames + " " + rsREF!CompanyName
        
        If Screen.ActiveForm.cboAccountNo.Text <= "" Then
                Screen.ActiveForm.txtAccountNo.Text = rsREF!AccountNo & ""
                Screen.ActiveForm.cboAccountNo.Text = rsREF!AccountCode & ""
        End If
        
        With Screen.ActiveForm
                
                If BPolicyHolder = True Then
                        .txtPayer.Text = .txtNames
                Else
                        .txtPayer = .cboEmployer
                End If
        
        End With
        
        rsREF.Close
        strREF = ""
        
    
        Exit Sub

err:
        ErrorMessage

End Sub

Public Sub cmdFirstCode_Click(Index As Integer)
''''On Error GoTo Myerr

Screen.ActiveForm.cmdUpdate.Enabled = False

        With rsDEP
            If .EOF And .BOF Then Exit Sub
            Select Case Index
                Case 0
                    .MoveFirst
                Case 1
                    .MovePrevious
                    If .BOF Then .MoveFirst
                Case 2
                    .MoveNext
                    If .EOF Then .MoveLast
                Case 3
                    .MoveLast
            End Select
        End With
        Call LoadReceiptGrid

Exit Sub
Myerr:
    ErrorMessage
End Sub


Public Sub validateRECORD()


    If Screen.ActiveForm.txtReceiptNo.Text = "" Then
        MsgBox "Receipt No is Required"
        Screen.ActiveForm.txtReceiptNo.SetFocus
    
    ElseIf Screen.ActiveForm.txtPayer = "" Then
        MsgBox "The Details of the Person Paying are Necessary and cannot be Left Blank"
        Screen.ActiveForm.txtPayer.SetFocus
 
    ElseIf Screen.ActiveForm.txtReceiptDate.Text Like 0 Then
        MsgBox "The Receipt Date is  required"
        Screen.ActiveForm.txtReceiptDate.SetFocus
    
    ElseIf Screen.ActiveForm.txtReceiptDate.Text > Date Then
        MsgBox "Payment cannot be accepted in advance"
        Screen.ActiveForm.txtReceiptDate.SetFocus
    
    ElseIf Screen.ActiveForm.txtCostCenter.Text = "" Then
        MsgBox "The Cost Center cannot be left blank"
        Screen.ActiveForm.txtCostCenter.SetFocus
    
    ElseIf Screen.ActiveForm.cboDocumentNo.Text = "" And bEmployerUpdate = False Then
        MsgBox "The Document Number is Mandatory for this Process"
        Screen.ActiveForm.cboDocumentNo.SetFocus
    
    ElseIf Screen.ActiveForm.txtReceiptAmount.Text Like 0 Then
        MsgBox "The Receipt Amount is required and There cannot be 0"
        Screen.ActiveForm.txtReceiptAmount.SetFocus
    
    ElseIf Screen.ActiveForm.cboReceiptType.Text = "" Then
        MsgBox "The Receipt is required", vbOKOnly
        Screen.ActiveForm.cboReceiptType.SetFocus
    
    ElseIf Screen.ActiveForm.cboPaymentMethod = "" Then
        MsgBox "The Payment Method is required", vbOKOnly
        Screen.ActiveForm.cboPaymentMethod.SetFocus

    Else
            bsaveRECORD = True
    End If

Exit Sub

End Sub

Private Sub saveRecord()
On Error GoTo err

    Dim rsDEP As ADODB.Recordset, strDEP As String
    Set rsDEP = New Recordset
    
    rsDEP.Open "select * from ALISMReceipt ; ", cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsDEP
            .AddNew
            !ReceiptNo = Screen.ActiveForm.txtReceiptNo
            !ReceiptDate = Screen.ActiveForm.txtReceiptDate
            !ReceiptType = Screen.ActiveForm.cboReceiptType
            !CostCenter = Screen.ActiveForm.txtCostCenter & ""
            !DocumentNo = Screen.ActiveForm.cboDocumentNo
            !AccountNo = CDbl(Screen.ActiveForm.txtAccountNo) & ""
            !AccountingPeriod = Screen.ActiveForm.txtAccountingPeriod
            !PaymentStatus = Screen.ActiveForm.txtPaymentStatus
            !UnitsPaid = CDbl(Screen.ActiveForm.txtUnitsPaid) & ""
            !AccountCode = Screen.ActiveForm.cboAccountNo
            !Payer = Screen.ActiveForm.txtPayer
            
            If Screen.ActiveForm.txtDueDate <= "" Then
                    Screen.ActiveForm.txtDueDate = Date
            End If
            
            !DueDate = Screen.ActiveForm.txtDueDate & ""
            !SuspenseAccount = CDbl(Screen.ActiveForm.txtSuspenseAccount) & ""
            !PaymentMethod = Screen.ActiveForm.cboPaymentMethod & ""
            !BankNo = Screen.ActiveForm.cboBankNo & ""

            If Screen.ActiveForm.cboPaymentMethod.Text = 3 Then
            
                 !ReceiptAmount = CDbl(Screen.ActiveForm.txtReceiptAmount)
                 !Cash = CDbl(Screen.ActiveForm.txtReceiptAmount)

            Else
                 If Screen.ActiveForm.cboBankNo.Text <> "" Then
                     Dim rsbanktype As ADODB.Recordset
                     Set rsbanktype = New ADODB.Recordset
                     
                     rsbanktype.Open "Select * from ALISPBank where BankNo= '" & Screen.ActiveForm.cboBankNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                     If rsbanktype!LocalBank = "Y" Then
                        !Local = CDbl(Screen.ActiveForm.txtReceiptAmount)
                        !ReceiptAmount = CDbl(Screen.ActiveForm.txtReceiptAmount)

                    ElseIf rsbanktype!LocalBank = "N" Then
                        !upcountry = CDbl(Screen.ActiveForm.txtReceiptAmount)
                        !ReceiptAmount = CDbl(Screen.ActiveForm.txtReceiptAmount)

                    End If
                End If
            End If


            !ChequeNo = Screen.ActiveForm.txtChequeNo & ""
            
            If Screen.ActiveForm.txtDateOfLastPayment <= "" Or CDbl(Screen.ActiveForm.txtDateOfLastPayment) <= 0 Then
                    Screen.ActiveForm.txtDateOfLastPayment = Date
                Else: !DateofLastPayment = Screen.ActiveForm.txtDateOfLastPayment
            End If
            
            !ReceivedToDate = CDbl(Screen.ActiveForm.txtReceivedTodate) & ""
            !UnitCount = CDbl(Screen.ActiveForm.txtUnitCount)
            !UnitCountBeforePayment = CDbl(Screen.ActiveForm.txtUnitCount) & ""

            If Screen.ActiveForm.txtDueDate.Text = "" Or Screen.ActiveForm.txtDueDate.Text = "0" Then
                    Screen.ActiveForm.txtDueDate.Text = Date
                    
                    Else: !DueDate = Screen.ActiveForm.txtDueDate & ""
            End If
            
            If Screen.ActiveForm.txtStatusCode <= "" Or Screen.ActiveForm.txtStatusCode = "10" Then
                    Screen.ActiveForm.txtStatusCode = "INFORCE"
            End If
            
            !StatusCode = Screen.ActiveForm.txtStatusCode.Text
            !TransactionNo = Screen.ActiveForm.txtTransactionNo.Text
            !TransactionDate = Date
            !Preparedby = CurrentUserName
            .Update
    End With
    
    
    If rsRCPT!Deposit = 1 Then
        
        Dim rsPROP As ADODB.Recordset, strPROP As String
        Set rsPROP = New Recordset
            
                rsPROP.Open "SELECT * FROM ALISMProposal WHERE proposalNo =  '" & Screen.ActiveForm.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                With rsPROP
                    If .EOF Or .BOF Then Exit Sub
                    
                    !ProposalNo = Screen.ActiveForm.cboDocumentNo
                    !DepositReceived = !DepositReceived + CDbl(Screen.ActiveForm.txtReceiptAmount)
                    !DateofLastPayment = Screen.ActiveForm.txtReceiptDate
                    !StatusCode = "DepRecvd"
                    .Update
                End With
        
        rsPROP.Close
        strPROP = ""
    
    ElseIf rsRCPT!Jobbrief = 1 Then
        
        Dim rsPOL As ADODB.Recordset, strPOL As String
        Set rsPOL = New Recordset
            
                rsPOL.Open "SELECT * FROM ODASMJobBrief WHERE JobBriefNo =  '" & Screen.ActiveForm.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                With rsPOL
                    !JobBriefNo = Screen.ActiveForm.cboDocumentNo
                    !DepositReceived = 0
                    !DateofLastPayment = Screen.ActiveForm.txtReceiptDate
                    !UnitCount = Screen.ActiveForm.txtUnitCount
                    !SuspenseAccount = Screen.ActiveForm.txtSuspenseAccount
                    !DateofLastPayment = Screen.ActiveForm.txtDateOfLastPayment & ""
                    
                    If Screen.ActiveForm.txtReceivedTodate <= "" Then
                    !ReceivedToDate = 0#
                    Else
                    !ReceivedToDate = Screen.ActiveForm.txtReceivedTodate
                    End If

                    !UnitCountBeforePayment = Screen.ActiveForm.txtUnitCountBeforePayment
                    !DueDate = Screen.ActiveForm.txtDueDate
                    
                    If Screen.ActiveForm.txtStatusCode <= "" Or Screen.ActiveForm.txtStatusCode = "10" Then
                        Screen.ActiveForm.txtStatusCode = "INFORCE"
                    End If
                    
                    !StatusCode = Screen.ActiveForm.txtStatusCode
                    .Update
                End With
        
        rsPOL.Close
        strPOL = ""
    ElseIf rsRCPT!IssuePolicy = 1 Then
        
        Dim rsISSUE As ADODB.Recordset, strISSUE As String
        Set rsISSUE = New Recordset
            
                rsISSUE.Open "SELECT * FROM ODASMJobBrief WHERE JobBriefNo =  '" & Screen.ActiveForm.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                With rsISSUE
                    !JobBriefNo = Screen.ActiveForm.cboDocumentNo
                    !DepositReceived = 0
                    !IssueDate = Date
                    !DateofLastPayment = Screen.ActiveForm.txtReceiptDate
                    !DueDate = Screen.ActiveForm.txtDueDate
                    !UnitCount = Screen.ActiveForm.txtUnitCount
                    !SuspenseAccount = Screen.ActiveForm.txtSuspenseAccount
                    !DateofLastPayment = Screen.ActiveForm.txtDateOfLastPayment & ""
                    !ReceivedToDate = Screen.ActiveForm.txtReceivedTodate & ""
                    !UnitCount = Screen.ActiveForm.txtUnitCount
                    !UnitCountBeforePayment = Screen.ActiveForm.txtUnitCountBeforePayment
                    !DueDate = Screen.ActiveForm.txtDueDate
            
                    If Screen.ActiveForm.txtStatusCode <= "" Then
                        Screen.ActiveForm.txtStatusCode = "INFORCE"
                    End If
                    
                    !StatusCode = Screen.ActiveForm.txtStatusCode
                    .Update
                End With
        
        rsISSUE.Close
        strISSUE = ""

        ElseIf rsRCPT!loan = 1 Then
        
            Dim rsLOAN As ADODB.Recordset, strLOAN As String
            Set rsLOAN = New Recordset
                
            rsLOAN.Open "SELECT * FROM ALISMLoanManagement WHERE LoanNo =  '" & Screen.ActiveForm.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                    With rsLOAN
                    
                        !JobBriefNo = Screen.ActiveForm.cboDocumentNo
                        
                        If Screen.ActiveForm.txtReceiptAmount > !RepaymentAmount Then
                                !SuspenseAccount = (Screen.ActiveForm.txtReceiptAmount - !RepaymentAmount)
                        ElseIf (Screen.ActiveForm.txtReceiptAmount <= !RepaymentAmount) And (Screen.ActiveForm.txtReceiptAmount > !PrincipalAmount) Then
                                Screen.ActiveForm.txtReceiptAmount = 0
                        ElseIf (Screen.ActiveForm.txtReceiptAmount <= !PrincipalAmount) Then
                                Screen.ActiveForm.txtReceiptAmount = 0
                        End If
                        
                        !TotalReceivedTodate = !TotalReceivedTodate + CDbl(Screen.ActiveForm.txtReceiptAmount)
                        !DateofLastPayment = Screen.ActiveForm.txtReceiptDate
                        !DueDate = Screen.ActiveForm.txtDueDate
             '           !SuspenseAccount = Screen.Activeform.txtSuspenseAccount & ""
                        
                        !CurrentBalance = !CurrentBalance - rsLOAN!PrincipalAmount
                        !InterestReceived = !InterestReceived + (Screen.ActiveForm.txtReceiptAmount.Text - !PrincipalAmount)
                        !DateofLastPayment = Screen.ActiveForm.txtDateOfLastPayment & ""
                        !DueDate = Screen.ActiveForm.txtDueDate
                
                        If Screen.ActiveForm.txtStatusCode <= "" Then
                            Screen.ActiveForm.txtStatusCode = "INF"
                        End If
                        
                        !Status = Screen.ActiveForm.txtStatusCode
                        .Update
                    End With
            
            rsLOAN.Close
            strLOAN = ""
    End If


Exit Sub

err:
        ErrorMessage
End Sub

Public Sub searchRECORD()
''''On Error GoTo Myerr

        Dim strQRE As Variant
        Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset
            strQRE = InputBox("Enter Receipt No to search.", "Search Value")
            rsFind.Open "SELECT * FROM ALISMReceipt WHERE ReceiptNo = '" & strQRE & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsFind
            If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
            Else:
                    Screen.ActiveForm.txtReceiptNo.Text = !ReceiptNo
                    Screen.ActiveForm.txtReceiptDate.Text = !ReceiptDate
                    Screen.ActiveForm.cboReceiptType.Text = !ReceiptType
                    Screen.ActiveForm.txtCostCenter.Text = !CostCenter
                    Screen.ActiveForm.cboDocumentNo.Text = !DocumentNo
                    Screen.ActiveForm.txtReceiptAmount.Text = !ReceiptAmount
                    Screen.ActiveForm.txtAccountNo.Text = !AccountNo
                    Screen.ActiveForm.txtAccountingPeriod.Text = !AccountingPeriod
                    Screen.ActiveForm.txtPaymentStatus.Text = !PaymentStatus
                    Screen.ActiveForm.cboAccountNo.Text = !AccountCode & ""
                    Screen.ActiveForm.txtDateOfLastPayment.Text = !DateofLastPayment
                    Screen.ActiveForm.txtReceivedTodate.Text = !ReceivedToDate
                    Screen.ActiveForm.txtUnitCount.Text = !UnitCount
                    Screen.ActiveForm.txtSuspenseAccount.Text = !SuspenseAccount & ""
                    Screen.ActiveForm.txtUnitsPaid.Text = !UnitsPaid
                    Screen.ActiveForm.txtUnitCountBeforePayment.Text = !UnitCountBeforePayment
                    Screen.ActiveForm.txtDueDate.Text = !DueDate & ""
                    Screen.ActiveForm.txtStatusCode.Text = !StatusCode
                    Screen.ActiveForm.txtChequeNo.Text = !ChequeNo & ""
                    Screen.ActiveForm.cboPaymentMethod.Text = !PaymentMethod & ""
                    Screen.ActiveForm.cboBankNo.Text = !BankNo & ""
                    Screen.ActiveForm.txtPayer.Text = !Payer
                    Edit = True
            End If
        
            If Screen.ActiveForm.txtReceiptNo.Text > "" Then
                LoadNAMES
                loadEMPLOYER
                loadRECEIPTTYPE
            End If
        
        End With

Exit Sub

Myerr:
    ErrorMessage
End Sub


Public Sub processReceipt()
On Error GoTo err
        
    If Screen.ActiveForm.txtReceiptAmount.Text <= "" Or Screen.ActiveForm.cboDocumentNo.Text <= "" Then Exit Sub

    If CDbl(Screen.ActiveForm.txtReceiptAmount) <= 0 Then
            MsgBox "The Receipt Amount is Mandatory and is required for all the calculations", vbOKCancel
            Screen.ActiveForm.txtReceiptAmount.SetFocus
            Exit Sub

    ElseIf CDbl(Screen.ActiveForm.cboDocumentNo.Text) <= 0 And bEmployerUpdate = False Then
            MsgBox "This Document No is Required prior to entering the Receipt Amount", vbOKCancel
            Screen.ActiveForm.cboDocumentNo.SetFocus
            Exit Sub
    
    ElseIf Screen.ActiveForm.cboEmployer.Text <= "" And bEmployerUpdate = True Then
            MsgBox "The Employer Code is Mandatory", vbOKCancel
            Screen.ActiveForm.cboDocumentNo.SetFocus
            Exit Sub

    ElseIf Screen.ActiveForm.cboReceiptType.Text <= "" Then
            MsgBox "The Receipt Type Is Required Before executive This Function", vbOKOnly
            Screen.ActiveForm.cboReceiptType.SetFocus
            Exit Sub
                
    ElseIf Screen.ActiveForm.txtCostCenter.Text <= "" Then
            MsgBox "The Cost Center Is Required Before executive This Function", vbOKOnly
            Screen.ActiveForm.txtCostCenter.SetFocus
            Exit Sub
    
    Else
        If CDbl(Screen.ActiveForm.txtReceiptAmount) > 0 Then
            
            Set rsRCPT = New Recordset
            strSQL = "SELECT * FROM ALISPReceipt where receiptType = '" & Screen.ActiveForm.cboReceiptType.Text & "';"
            rsRCPT.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

            If rsRCPT!Jobbrief = "1" Or rsRCPT!IssuePolicy = "1" Then
                Call updateRECORD
            ElseIf rsRCPT!loan = "1" Then
                Call updateLOANRECORD
            ElseIf rsRCPT!BreakDown = "1" Then
                Call updateEMPLOYER
            End If
        End If

        Screen.ActiveForm.cboPaymentMethod.SetFocus
    End If
Exit Sub

err:
    UpdateErrorMessage
End Sub

Private Sub verifyTHEAMOUNTRECEIVED()
On Error GoTo err
Exit Sub

err:
    ErrorMessage
End Sub


Private Sub Class_Initialize()
    Set rsRCPT = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
    Set rsRCPT = Nothing
End Sub
