VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReceiptDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private MySName As String, MyOName As String
Dim rsPOLICY As ADODB.Recordset, strRECPT As String
Dim rsreceipt As ADODB.Recordset, bVAL, BPolicyHolder, bsearchRECORD As Boolean
Dim rsDEP As ADODB.Recordset, strDEP As String
Public strCalculateCOMMISSION As String
Private Sub Form_Unload(cancel As Integer)
On Error GoTo err
    If addpen = True Then
        cancel = True
        MsgBox "Data entry in progress. Click Refresh to Cancel", vbCritical
    Else
        cancel = False
    End If
Exit Sub
err:
ErrorMessage
End Sub

Public Sub SelectEmployerGotFocus()
On Error GoTo err

        Set rsCONTROL = New Recordset

        strSQL = "SELECT * FROM ODASPAccount where StatusCode = 'A';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        frmODASMReceipt.txtAccountNo.Clear

            With rsCONTROL
                    Do Until .EOF
                             Screen.ActiveForm.cboEmployer.AddItem !CompanyName
                             .MoveNext
                    Loop
            End With
rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    UpdateErrorMessage
End Sub

Public Sub selectEmployerLostFocus()
On Error GoTo err
        Set rsCONTROL = New Recordset
        
        rsCONTROL.Open "SELECT * FROM ODASPAccount WHERE CompanyName= '" & Screen.ActiveForm.cboEmployer.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
 
        With rsCONTROL
                If .EOF And .BOF Then Exit Sub
                            Screen.ActiveForm.cboAccountNo.Text = !AccountCode & ""
                            Screen.ActiveForm.cboEmployer.Text = !CompanyName & ""
                
                Dim rsEMPLOYER As ADODB.Recordset, strEMPLOYER As String
                Set rsEMPLOYER = New Recordset
        
                rsEMPLOYER.Open "SELECT * FROM ALISPReceipt WHERE employer = '1'", cnCOMMON, adOpenKeyset, adLockOptimistic

                With rsEMPLOYER
                        If .EOF Or .BOF = True Then Exit Sub
                        
                        Screen.ActiveForm.cboDocumentNo.Text = Screen.ActiveForm.cboAccountNo
                        Screen.ActiveForm.txtTransactionAmount.Text = Screen.ActiveForm.txtReceiptAmount.Text
                End With
                
                rsEMPLOYER.Close
                
    End With
rsCONTROL.Close

Exit Sub

err:
    UpdateErrorMessage
End Sub


Public Sub processReceipt()
On Error GoTo err
                If Screen.ActiveForm.cboDocumentNo.Text <= "" Then Exit Sub
                
                Screen.ActiveForm.txtDateOfLastPayment.Text = Screen.ActiveForm.txtReceiptDate.Text

                If Screen.ActiveForm.txtTransactionAmount.Text <= "" Or Screen.ActiveForm.cboDocumentNo.Text <= "" Then Exit Sub
        
                If CDbl(Screen.ActiveForm.txtTransactionAmount) <= 0 Then
                        MsgBox "The Transaction Amount is Mandatory and is required for all the calculations", vbOKCancel
                        Screen.ActiveForm.txtTransactionAmount.SetFocus
                        Exit Sub
                
                ElseIf Screen.ActiveForm.cboDocumentNo.Text <= "" And bEmployerUpdate = False Then
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
                ElseIf CDbl(Screen.ActiveForm.txtTransactionAmount.Text) > CDbl(Screen.ActiveForm.txtReceiptAmount.Text) Then
                        MsgBox "The Transaction Amount MUST be Less Than the Receipt Amount", vbOKOnly
                        Screen.ActiveForm.txtTransactionAmount.SetFocus
                ElseIf CDbl(Screen.ActiveForm.txtTransactionAmount.Text) > CDbl(Screen.ActiveForm.txtBalance.Text) Then
                        MsgBox "The Transaction Amount MUST be Less Than the Balance", vbOKOnly
                        Screen.ActiveForm.txtTransactionAmount.SetFocus
                ElseIf Screen.ActiveForm.txtDateOfLastPayment.Text <= "" Then
                        MsgBox "The Date of Last Payment is Blank, Testing ..... "
                        Screen.ActiveForm.txtDateOfLastPayment.SetFocus
                Else
                
                        If CDbl(Screen.ActiveForm.txtTransactionAmount.Text) > 0 Then
                                
                                Set rsRCPT = New ADODB.Recordset
                                rsRCPT.Open "SELECT * FROM ALISPReceipt Where receiptType = '" & Screen.ActiveForm.cboReceiptType.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

                                If rsRCPT!Jobbrief = "1" Or rsRCPT!IssuePolicy = "1" Then
                                    Call updateRECORD
                                ElseIf rsRCPT!Deposit = "1" Then
                                    updateACCOUNTS
                                ElseIf rsRCPT!accounts = "1" Then
                                    Call updateACCOUNTS
                                End If
                        End If
    
                End If
Exit Sub

err:
    UpdateErrorMessage
End Sub

Private Sub calcBalance()
            Screen.ActiveForm.txtBalance.Text = CDbl(Screen.ActiveForm.txtReceiptAmountDetails) - CDbl(Screen.ActiveForm.txtTotalAmount.Text)
End Sub

Public Sub updateEMPLOYER()
On Error GoTo err
                
    Screen.ActiveForm.txtSuspenseAccount.Text = 0
    Screen.ActiveForm.txtUnitCount.Text = 0
    Screen.ActiveForm.txtReceivedTodate.Text = 0
    Screen.ActiveForm.txtUnitsPaid = 1
    Screen.ActiveForm.txtUnitCountBeforePayment = 0
    Screen.ActiveForm.txtUnitCount = 0
    Screen.ActiveForm.txtDateOfLastPayment = Date
    Screen.ActiveForm.txtReceivedTodate = CDbl(Screen.ActiveForm.txtReceiptAmount)
                           
Exit Sub


err:
    ErrorMessage
End Sub

Sub loadRECEIPTDETAILSGRID()

Dim rsreceipt1 As ADODB.Recordset
Set rsreceipt1 = New Recordset

   rsreceipt1.Open "SELECT DocumentNo, ReceiptType, TransactionAmount, transactionDate FROM ALISMreceiptdetails WHERE DocumentNo =  '" & Screen.ActiveForm.cboDocumentNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
   Set Screen.ActiveForm.receiptDetailsGRID.DataSource = rsreceipt1
   Set Screen.ActiveForm.ReceiptMGRID.DataSource = rsreceipt1

End Sub

Sub loadpolicyGRID()

Dim rsPOLICYGRID As ADODB.Recordset
Set rsPOLICYGRID = New Recordset

   rsPOLICYGRID.Open "SELECT JobBriefNo, ExpectedPremium, DateOfCommencement, DateOfLastPayment FROM ODASMJobBrief WHERE AccountNo LIKE  '" & Screen.ActiveForm.txtAccountNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
   Set Screen.ActiveForm.policiesGRID.DataSource = rsPOLICYGRID

End Sub

Private Sub loadloanGRID()

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


Public Sub clearRECORD()
On Error GoTo err:

        With Screen.ActiveForm
            .txtTransactionNODetails.Text = ""
            .txtTransactionDate.Text = ""
            .cboReceiptType.Text = ""
            .cboDocumentNo.Text = ""
            .txtExpectedAmount.Text = ""
            .txtAccountNo.Text = ""
            .txtCurrentPeriodDetails.Text = ""
            .txtPaymentStatusDetails.Text = ""
            .cboAccountNo.Text = ""
            .txtDateOfLastPayment.Text = ""
            .txtReceivedTodate.Text = ""
            .txtNames.Text = ""
            .txtUnitCount.Text = ""
            .txtSuspenseAccount.Text = ""
            .txtUnitsPaid.Text = ""
            .txtUnitCountBeforePayment.Text = ""
            .txtDueDate.Text = ""
            .txtStatusCode.Text = ""
            .txtAccountNo.Text = ""
            .txtTransactionAmount.Text = ""
            .cboEmployer.Text = ""
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
On Error GoTo err:


        With Screen.ActiveForm
            .txtReceiptNoDetails.Locked = bVAL
            .txtReceiptDate.Locked = True
            .cboReceiptType.Locked = bVAL
            .txtReceiptAmount.Locked = bVAL
            .cboDocumentNo.Locked = bVAL
            .txtAccountNo.Locked = bVAL
            .txtExpectedAmount.Locked = True
            .txtCurrentPeriodDetails.Locked = True
            .txtPaymentStatus.Locked = True
            .txtUnitCount.Locked = True
            .txtReceivedTodate.Locked = True
            .txtUnitCountBeforePayment.Locked = True
            .txtDueDate.Locked = True
            .txtStatusCode.Locked = True
            .txtSuspenseAccount.Locked = True
            .txtUnitsPaid.Locked = True
            .txtAccountNo.Locked = bVAL
            .cboAccountNo.Locked = False
            .txtNames.Locked = True
            .txtDateOfLastPayment.Locked = True
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
        Set rsLOANDUE = New ADODB.Recordset
            
        rsLOANDUE.Open "SELECT * FROM ALISMLoanManagement Where ALISMLoanManagement.LoanNo = '" & Screen.ActiveForm.cboDocumentNo & "'  ;", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsLOANDUE
                If .EOF Or .BOF Then Exit Sub
                
                !UnitsDue = DateDiff("M", !CommencementDate, Date)
                
                If !UnitsDue > !RepaymentPeriod Then
                    !UnitsDue = !RepaymentPeriod
                End If
                
                !UnitsDue = !UnitsDue - !UnitCount
                
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
                
                If DateDiff("M", Date, rsPOLICY!MaturityDate) > 0 Then
                        !UnitsDue = DateDiff("M", rsPOLICY!DateOfCommencement, Date)
                Else
                        !UnitsDue = DateDiff("M", rsPOLICY!DateOfCommencement, rsPOLICY!MaturityDate)
                End If
                        '/ * Check for Maturity Date
                        
                !UnitsDue = (Fix(!UnitsDue / rsPM!CoverPeriod) * rsPM!CoverPeriod) - !UnitCount
                
                If !UnitsDue < 0 Then
                        
                        !DueDate = DateAdd("M", !UnitCount, !DateOfCommencement)
                        !AmountDue = 0
                    
                    Else: !DueDate = DateAdd("M", (!UnitCount), !DateOfCommencement)
                    
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

                Dim rsPM As ADODB.Recordset, strPM As String
                Set rsPM = New Recordset
                    
                rsPM.Open "SELECT * FROM ODASPPaymentMode Where paymentMode = '" & !PaymentMode & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

    
                '/* this procedure is used to intialize the variable
                
                Screen.ActiveForm.txtSuspenseAccount.Text = !SuspenseAccount
                Screen.ActiveForm.txtUnitCount.Text = !UnitCount
                Screen.ActiveForm.txtReceivedTodate.Text = !ReceivedToDate & ""
                
                                
                '/* this Procedure is used to calculate the number of Premiums
                '/* paid by the Policyholder
                
                If (CDbl(Screen.ActiveForm.txtTransactionAmount.Text) >= !ExpectedPremium) And (!UnitCount < !TermOfPolicy * 12) And ((Trim(rsPOL!StatusCode) = "10" Or Trim(rsPOL!StatusCode) = "INFORCE") Or Trim(rsPOL!StatusCode = "UNISSUED")) Then

                        Screen.ActiveForm.txtUnitsPaid = Fix(Screen.ActiveForm.txtTransactionAmount / Screen.ActiveForm.txtExpectedAmount) * rsPM!CoverPeriod
                        
                        If Screen.ActiveForm.txtUnitCount <= "" Then
                            Screen.ActiveForm.txtUnitCount = 0
                        End If
                        
                        If Screen.ActiveForm.txtSuspenseAccount <= "" Then
                            Screen.ActiveForm.txtSuspenseAccount = 0
                        End If
                        
                        Screen.ActiveForm.txtUnitCountBeforePayment = CDbl(Screen.ActiveForm.txtUnitCount)
                        Screen.ActiveForm.txtUnitCount = !UnitCount + CDbl(Screen.ActiveForm.txtUnitsPaid.Text)
                        Screen.ActiveForm.txtDateOfLastPayment = Screen.ActiveForm.txtReceiptDate
                        Screen.ActiveForm.txtSuspenseAccount = !SuspenseAccount + CDbl(Screen.ActiveForm.txtTransactionAmount - (Screen.ActiveForm.txtUnitsPaid * (Screen.ActiveForm.txtExpectedAmount / rsPM!CoverPeriod)))
                        
                        If Screen.ActiveForm.txtReceivedTodate <= "" Then
                                        Screen.ActiveForm.txtReceivedTodate = CDbl(Screen.ActiveForm.txtTransactionAmount)
                                Else: Screen.ActiveForm.txtReceivedTodate = !ReceivedToDate + CDbl(Screen.ActiveForm.txtTransactionAmount)
                        End If
                        
                        '/If the total premiums exceed the expected number of Premiums
                
                        If CDbl(Screen.ActiveForm.txtUnitCount) >= !TermOfPolicy * 12 Then
                                Screen.ActiveForm.txtUnitsPaid = CDbl(Screen.ActiveForm.txtUnitsPaid) - (Screen.ActiveForm.txtUnitCount - (!TermOfPolicy * 12))
                                Screen.ActiveForm.txtUnitCount = !TermOfPolicy * 12
                                Screen.ActiveForm.txtStatusCode = "MATURED"
                        End If
                        
                        Screen.ActiveForm.txtDueDate = DateAdd("M", CDbl(Screen.ActiveForm.txtUnitCount), !DateOfCommencement)
                                               
                        '/* check if the Maturity date been reached
                
                ElseIf (Screen.ActiveForm.txtTransactionAmount < Screen.ActiveForm.txtExpectedAmount) And (Screen.ActiveForm.txtUnitCount < !TermOfPolicy * 12) Then
                        Screen.ActiveForm.txtSuspenseAccount = CDbl(!SuspenseAccount) + CDbl(Screen.ActiveForm.txtTransactionAmount)
                        Screen.ActiveForm.txtUnitsPaid.Text = 0
                        Screen.ActiveForm.txtStatusCode.Text = !StatusCode
                        Screen.ActiveForm.txtUnitCount.Text = !UnitCount
                        Screen.ActiveForm.txtReceivedTodate = !ReceivedToDate + Screen.ActiveForm.txtTransactionAmount
                        Screen.ActiveForm.txtDueDate.Text = !DueDate
                        
                Else:   Screen.ActiveForm.txtSuspenseAccount = !SuspenseAccount + CDbl(Screen.ActiveForm.txtTransactionAmount)
                        Screen.ActiveForm.txtUnitsPaid.Text = 0
                        Screen.ActiveForm.txtStatusCode.Text = !StatusCode
                        Screen.ActiveForm.txtUnitCount.Text = !UnitCount
                        
                        If IsNull(!ReceivedToDate) = True Then
                                Screen.ActiveForm.txtReceivedTodate.Text = CDbl(Screen.ActiveForm.txtTransactionAmount)
                        Else: Screen.ActiveForm.txtReceivedTodate.Text = !ReceivedToDate + CDbl(Screen.ActiveForm.txtTransactionAmount)
                        End If
                        
                        Screen.ActiveForm.txtDueDate.Text = !DueDate & ""
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

Public Sub updateREVIVAL()
On Error GoTo err
        
        Dim rsPOL As ADODB.Recordset, strPOL As String
        Set rsPOL = New Recordset
            
        rsPOL.Open "SELECT * FROM ODASMJobBrief, ALISMReinstatement WHERE ODASMJobBrief.JobBriefNo =  '" & Screen.ActiveForm.cboDocumentNo & "' and ODASMJobBrief.statusCode = 'LAPSED' and ODASMJobBrief.JobBriefNo = ALISMReinstatement.JobBriefNo and ALISMReinstatement.Approved = 'Y' and ALISMReinstatement.Reinstated ='N' ;", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsPOL
        
                If .BOF Or .EOF Then
                            MsgBox "This policy is not Lapsed, The System is abandoning the Exercise", vbOKOnly
                            Screen.ActiveForm.cmdUpdateDetails.Enabled = False
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
                
                If CDbl(Screen.ActiveForm.txtTransactionAmount.Text) >= 0 And (!UnitCount < !TermOfPolicy * 12) Then
                        Screen.ActiveForm.txtUnitsPaid = 0
                        
                        Screen.ActiveForm.txtUnitCountBeforePayment = CDbl(Screen.ActiveForm.txtUnitCount)
                        Screen.ActiveForm.txtUnitCount = !UnitCount + CDbl(Screen.ActiveForm.txtUnitsPaid.Text)
                        Screen.ActiveForm.txtDateOfLastPayment = Screen.ActiveForm.txtReceiptDate
                        Screen.ActiveForm.txtSuspenseAccount = !SuspenseAccount
                        
                        If Screen.ActiveForm.txtReceivedTodate <= "" Then
                                        Screen.ActiveForm.txtReceivedTodate = CDbl(Screen.ActiveForm.txtTransactionAmount)
                                Else: Screen.ActiveForm.txtReceivedTodate = !ReceivedToDate + CDbl(Screen.ActiveForm.txtTransactionAmount)
                        End If
                        
                        '/If the total premiums exceed the expected number of Premiums
                
                        If CDbl(Screen.ActiveForm.txtUnitCount) >= !TermOfPolicy * 12 Then
                                Screen.ActiveForm.txtUnitsPaid = CDbl(Screen.ActiveForm.txtUnitsPaid) - (Screen.ActiveForm.txtUnitCount - (!TermOfPolicy * 12))
                                Screen.ActiveForm.txtUnitCount = !TermOfPolicy * 12
                                Screen.ActiveForm.txtStatusCode = "MATURED"
                        End If
                        
                        Screen.ActiveForm.txtDueDate = DateAdd("M", CDbl(Screen.ActiveForm.txtUnitCount), !DateOfCommencement)
                                               
                        '/* check if the Maturity date been reached
                
                ElseIf (Screen.ActiveForm.txtTransactionAmount < Screen.ActiveForm.txtExpectedAmount) And (Screen.ActiveForm.txtUnitCount < !TermOfPolicy * 12) Then
                        Screen.ActiveForm.txtSuspenseAccount = CDbl(!SuspenseAccount) + CDbl(Screen.ActiveForm.txtTransactionAmount)
                        Screen.ActiveForm.txtUnitsPaid.Text = 0
                        Screen.ActiveForm.txtStatusCode.Text = !StatusCode
                        Screen.ActiveForm.txtUnitCount.Text = !UnitCount
                        Screen.ActiveForm.txtReceivedTodate = !ReceivedToDate + Screen.ActiveForm.txtTransactionAmount
                        Screen.ActiveForm.txtDueDate.Text = !DueDate
                        
                Else:   Screen.ActiveForm.txtSuspenseAccount = !SuspenseAccount + CDbl(Screen.ActiveForm.txtTransactionAmount)
                        Screen.ActiveForm.txtUnitsPaid.Text = 0
                        Screen.ActiveForm.txtStatusCode.Text = !StatusCode
                        Screen.ActiveForm.txtUnitCount.Text = !UnitCount
                        
                        If IsNull(!ReceivedToDate) = True Then
                                Screen.ActiveForm.txtReceivedTodate.Text = CDbl(Screen.ActiveForm.txtTransactionAmount)
                        Else: Screen.ActiveForm.txtReceivedTodate.Text = !ReceivedToDate + CDbl(Screen.ActiveForm.txtTransactionAmount)
                        End If
                        
                        Screen.ActiveForm.txtDueDate.Text = !DueDate & ""
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

Public Sub updateISSUEDPOLICY()
On Error GoTo err
        
        Dim rsPOL As ADODB.Recordset, strPOL As String
        Set rsPOL = New Recordset
            
        rsPOL.Open "SELECT * FROM ODASMJobBrief WHERE JobBriefNo =  '" & Screen.ActiveForm.cboJobBriefNo & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsPOL

                Dim rsPM As ADODB.Recordset, strPM As String
                Set rsPM = New Recordset
                    
                rsPM.Open "SELECT * FROM ODASPPaymentMode Where paymentMode = '" & !PaymentMode & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

    
                '/* this procedure is used to intialize the variable
                
                Screen.ActiveForm.txtSuspenseAccount.Text = !SuspenseAccount
                Screen.ActiveForm.txtUnitCount.Text = 0
                Screen.ActiveForm.txtReceivedTodate.Text = !ReceivedToDate & ""
                Screen.ActiveForm.txtTransactionAmount.Text = !ReceivedToDate
                                
                '/* this Procedure is used to calculate the number of Premiums
                '/* paid by the Policyholder
                
                If (CDbl(Screen.ActiveForm.txtTransactionAmount.Text) >= !ExpectedPremium) Then

                        Screen.ActiveForm.txtUnitCount = Fix(Screen.ActiveForm.txtTransactionAmount / Screen.ActiveForm.txtexpectedpremium) * rsPM!CoverPeriod
                        Screen.ActiveForm.txtUnitCountBeforePayment = 0
                        Screen.ActiveForm.txtSuspenseAccount = !SuspenseAccount + CDbl(Screen.ActiveForm.txtTransactionAmount - (Screen.ActiveForm.txtUnitCount * (Screen.ActiveForm.txtexpectedpremium / rsPM!CoverPeriod)))
                        
                        If Screen.ActiveForm.txtReceivedTodate <= "" Then
                                        Screen.ActiveForm.txtReceivedTodate = CDbl(Screen.ActiveForm.txtTransactionAmount)
                                Else: Screen.ActiveForm.txtReceivedTodate = !ReceivedToDate + CDbl(Screen.ActiveForm.txtTransactionAmount)
                        End If
                        
                        Screen.ActiveForm.txtDueDate = DateAdd("M", CDbl(Screen.ActiveForm.txtUnitCount), !DateOfCommencement)
                                               
                        '/* check if the Maturity date been reached
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

Public Sub updateACCOUNTS()
On Error GoTo err
        
        Dim rsACCOUNT As ADODB.Recordset, strPOL As String
        Set rsACCOUNT = New ADODB.Recordset
            
        rsACCOUNT.Open "SELECT * FROM MEGAMAccount WHERE AccountNo =  '" & Screen.ActiveForm.cboDocumentNo & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsACCOUNT
                If .BOF Or .EOF Then Exit Sub
                Screen.ActiveForm.txtDateOfLastPayment = Screen.ActiveForm.txtReceiptDate
                Screen.ActiveForm.txtStatusCode = "PAID"
        End With
        
rsACCOUNT.Close
strPOL = ""

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub processDOCUMENTNO()

        If Screen.ActiveForm.cboDocumentNo.Text <= "" Then Exit Sub
        
        If Screen.ActiveForm.cboReceiptType.Text <= "" Then
                            MsgBox "The Receipt Type Is Required Before executive This Function", vbOKOnly
                            Screen.ActiveForm.cboReceiptType.SetFocus
        Else
                If Screen.ActiveForm.cboReceiptType > "" Then
                        
                        Set rsRCPT = New ADODB.Recordset
                        rsRCPT.Open "SELECT * FROM ALISPReceipt Where receiptType = '" & Screen.ActiveForm.cboReceiptType.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

                        With rsRCPT
                                If .EOF Or .BOF Then
                                Else
                                            If !Deposit = 1 Then
                                                verifyACCOUNTdetails
                                            ElseIf !Jobbrief = 1 Or !Account = 1 Then
                                                verifyJOBBRIEFDETAILS
                                            ElseIf !RevivalFee = 1 Then
                                                verifyJOBBRIEFDETAILS
                                                obtainREVIVALFEE
                                            ElseIf !IssuePolicy = 1 Then
                                                verifyJOBBRIEFDETAILS
                                            ElseIf !loan = 1 Then
                                                calculateLOANAMOUNTDUE
                                                verifyLOANDETAILS
                                            ElseIf !accountsreceivable = 1 Then
                                                verifyACCOUNTdetails
                                            Else
                                                    BPolicyHolder = False
                                                '/* rsRCPT is Loan */
                                            End If
                                    '/* end if the Statement Statement
                                        Call loadEMPLOYER
                            End If

                                        Screen.ActiveForm.txtTransactionAmount.SetFocus
                        End With
                                    '/* End with Receipt
                End If
                '/* if Receipt Type is Greater than Zero
        End If
                '/* end if Document Number is Greater Than Zero


End Sub
Private Sub obtainREVIVALFEE()
On Error GoTo err

        Dim rsREVIVAL As ADODB.Recordset
        Set rsREVIVAL = New ADODB.Recordset
        rsREVIVAL.Open "SELECT * FROM ALISMReinstatement Where JobBriefNo = '" & Screen.ActiveForm.cboDocumentNo.Text & "' and Reinstated = 'N' and Approved = 'Y';", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsREVIVAL
                If .BOF Or .EOF Then
                        MsgBox "You can only revive a Lapsed Policy, Have You Processed the Reinstatement Prior to executing this function ", vbOKOnly
                        Screen.ActiveForm.cmdUpdateDetails = False
                        Exit Sub
                End If
                
                Screen.ActiveForm.txtTransactionAmount.Text = !ReinstatementCharge
                
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
                
                Screen.ActiveForm.txtSuspenseAccount.Text = !suspense & ""
                'screen.activeform.txtUnitCount.Text = !CurrentBalance
                Screen.ActiveForm.lblUnitCount.Caption = "Interest/Loan Paid"
                Screen.ActiveForm.txtReceivedTodate.Text = !TotalReceivedTodate & ""
                
                                
                '/* this Procedure is used to calculate the Amount to be charge to the
                '/* interest and Principal account
                
                If (CDbl(Screen.ActiveForm.txtReceiptAmount.Text) >= !RepaymentAmount) Then '/ And (!CurentBalance > 0) Then

                        Screen.ActiveForm.txtUnitsPaid = Fix(Screen.ActiveForm.txtReceiptAmount / Screen.ActiveForm.txtExpectedAmount)
                        Screen.ActiveForm.txtUnitCount.Text = CDbl(Screen.ActiveForm.txtUnitsPaid.Text) + CDbl(Screen.ActiveForm.txtUnitCount.Text)
                        Screen.ActiveForm.txtDateOfLastPayment = Screen.ActiveForm.txtReceiptDate
                        Screen.ActiveForm.txtSuspenseAccount = !suspense & ""
                        
                        If Screen.ActiveForm.txtReceivedTodate <= "" Then
                                Screen.ActiveForm.txtReceivedTodate = CDbl(Screen.ActiveForm.txtReceiptAmount)
                        Else: Screen.ActiveForm.txtReceivedTodate = !TotalReceivedTodate + CDbl(Screen.ActiveForm.txtTransactionAmount)
                        End If
                        
                        '/If the total Units exceed the expected number of Units
                
                        If CDbl(Screen.ActiveForm.txtUnitCount) >= !RepaymentPeriod Then
                                Screen.ActiveForm.txtUnitsPaid = Screen.ActiveForm.txtUnitsPaid - (Screen.ActiveForm.txtUnitCount - !RepaymentPeriod)
                                Screen.ActiveForm.txtUnitCount = !RepaymentPeriod
                                Screen.ActiveForm.txtStatusCode = "COMPLETE"
                        End If
                        
                        Screen.ActiveForm.txtDueDate = DateAdd("M", CDbl(Screen.ActiveForm.txtUnitCount), !CommencementDate)
                                               
                        '/* check if the Maturity date been reached
                
      End If

End With


Exit Sub


err:
   ErrorMessage
End Sub

Public Sub loadPolicy()
On Error GoTo err
'    Set rsPOLICY = New ADODB.Recordset
'
'    With rsPOLICY
'            frmALISMLedgerDetails.txtReceiptNo.Text = !ReceiptNo
'            frmALISMLedgerDetails.txtReceiptDate.Text = !ReceiptDate
'            frmALISMLedgerDetails.txtReceiptAmount.Text = !ReceiptAmount
'            frmALISMLedgerDetails.txtAccountNo.Text = !AccountNo
'            frmALISMLedgerDetails.txtPaymentStatus.Text = !PaymentStatus
'            frmALISMLedgerDetails.txtDateOfLastPayment.Text = !DateofLastPayment
'            frmALISMLedgerDetails.txtReceivedTodate.Text = !ReceivedTodate
'            frmALISMLedgerDetails.txtUnitCount.Text = !UnitCount
'            frmALISMLedgerDetails.txtSuspenseAccount.Text = !SuspenseAcCount & ""
'            frmALISMLedgerDetails.txtunitsPaid.Text = !UnitsPaid
'            frmALISMLedgerDetails.txtUnitCountBeforePayment.Text = !UnitCountBeforePayment
'            frmALISMLedgerDetails.txtDueDate.Text = !DueDate & ""
'            frmALISMLedgerDetails.txtStatusCode.Text = !StatusCode
'    End With
Exit Sub
err:
ErrorMessage
End Sub

Private Sub LoadNAMES()
On Error GoTo err
    
    Dim rsREF As ADODB.Recordset, strREF As String
    Set rsREF = New Recordset
    
    strREF = "SELECT * FROM ODASPAccount,ODASMJobBrief,ALISMReceipt Where ODASPAccount.AccountNo=ODASMJobBrief.AccountNo and  ODASMJobBrief.JobBriefNo= ALISMReceipt.DocumentNo ;"
    rsREF.Open strREF, cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsREF
            If .BOF Or .EOF Then
                    MsgBox "The Policyholder Does not Exist", vbOKOnly
                    Exit Sub
            End If
    
            Screen.ActiveForm.txtNames = Trim(!OtherNames) + " " + Trim(!CompanyName)
    End With
    
    Dim rsAgent As ADODB.Recordset, strAgent As String
    Set rsAgent = New Recordset
    
    strAgent = "SELECT * FROM ALISPAgent,ODASMJobBrief,ALISMReceipt Where ALISPAgent.AgentNo=ODASMJobBrief.AgentNo and  ODASMJobBrief.JobBriefNo= ALISMReceipt.DocumentNo;"
    rsAgent.Open strAgent, cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsAgent
            If .BOF Or .EOF Then
                    MsgBox "The agent Does not Exist", vbOKOnly
                    Exit Sub
            End If
    
            'frmALISMLedgerDetails.txtagent = Trim(!otherNames) + " " + Trim(!CompanyName)
    End With
    rsAgent.Close
    rsREF.Close

Exit Sub

err:
        ErrorMessage
End Sub

Public Sub enableCBRECEIPT()
    Screen.ActiveForm.cmdAddNewDetails.Enabled = True
    Screen.ActiveForm.cmdUpdateDetails.Enabled = False
    Screen.ActiveForm.cmdSearchDetails.Enabled = True
    Screen.ActiveForm.cmdCancelDetails.Enabled = True
End Sub

Public Sub DisableCBRECEIPT()
    Screen.ActiveForm.cmdAddNewDetails.Enabled = False
    Screen.ActiveForm.cmdUpdateDetails.Enabled = True
    Screen.ActiveForm.cmdSearchDetails.Enabled = False
    Screen.ActiveForm.cmdCancelDetails.Enabled = True
End Sub

Public Sub SelectAGENTKEYPRESS(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub


Public Sub selectRECEIPTTYPEGOTFOCUS()
On Error GoTo err
    
    
    Dim rsPAY As ADODB.Recordset, strPAY As String
    Set rsPAY = New Recordset
    
    strPAY = "SELECT * FROM ALISPReceipt;"
    rsPAY.Open strPAY, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Screen.ActiveForm.cboReceiptType.Clear

    With rsPAY
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

Public Sub processDOCUMENT()

        Set rsRCPT = New Recordset
        
        rsRCPT.Open "SELECT * FROM ALISPReceipt WHERE receiptType = '" & Screen.ActiveForm.cboReceiptType.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsRCPT
                If .EOF And .BOF Then Exit Sub
                
                '/ Enable the Employer Code Control if the receipt type is set to
                '/ Employer = "1" else locks it up
                
                If !Jobbrief = "1" Then
                        verifyJOBBRIEFDETAILS
                        loadEMPLOYER
                ElseIf !Proposal = "1" Then
                        verifyACCOUNTdetails
                        loadEMPLOYER
                ElseIf !loan = "1" Then
                
                ElseIf !AccountCode = "1" Then
                    
                End If
                
        End With
        
rsRCPT.Close

        Exit Sub

err:
        ErrorMessage

End Sub
Private Sub verifyACCOUNTdetails()
On Error GoTo err

        Dim rsDNLF As ADODB.Recordset, strDNLF As String
        Set rsDNLF = New Recordset

        rsDNLF.Open "SELECT ALISMProposal.*, ODASPAccount.* FROM ALISMProposal, ODASPAccount WHERE ALISMProposal.ProposalNo= '" & Screen.ActiveForm.cboDocumentNo.Text & "' and ALISMProposal.AccountNo LIKE ODASPAccount.AccountNo ", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsDNLF
            If .EOF And .BOF Then Exit Sub
            
            Screen.ActiveForm.cboDocumentNo.Text = !ProposalNo
            Screen.ActiveForm.txtAccountNo.Text = !AccountNo
                        
            If !DateofLastPayment > "" Then
                    Screen.ActiveForm.txtDateOfLastPayment.Text = !DateofLastPayment & ""
            End If
            
            Screen.ActiveForm.txtTransactionAmount.Text = !ExpectedPremium
            Screen.ActiveForm.txtExpectedAmount.Text = !ExpectedPremium
            Screen.ActiveForm.txtDueDate.Text = Date
            Screen.ActiveForm.cboAccountNo = !AccountCode & ""
            Screen.ActiveForm.txtAccountNo = !AccountNo & ""
            
            If IsNull(!SuspenseAccount) = True Then
                    Screen.ActiveForm.txtSuspenseAccount = 0
                    Else: Screen.ActiveForm.txtSuspenseAccount.Text = !SuspenseAccount
            End If
            Screen.ActiveForm.txtNames.Text = Trim(!OtherNames) + " " + Trim(!CompanyName)
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
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub verifyJOBBRIEFDETAILS()
        
        Dim rsPOLICY As ADODB.Recordset, strPOLICY As String
        Set rsPOLICY = New Recordset
                         
        rsPOLICY.Open "SELECT ODASMJobBrief.*, ODASPAccount.* FROM ODASMJobBrief, ODASPAccount WHERE ODASMJobBrief.JobBriefNo = '" & Screen.ActiveForm.cboDocumentNo.Text & "' and ODASMJobBrief.AccountNo LIKE ODASPAccount.AccountNo ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsPOLICY
                If .EOF Or .BOF Then
                        MsgBox "The Policy Does not Exist within the Database", vbOKOnly
                        Exit Sub
                End If
                
                Screen.ActiveForm.txtExpectedAmount = !ExpectedPremium
                Screen.ActiveForm.cboAccountNo = !AccountCode & ""
                Screen.ActiveForm.txtReceivedTodate = !ReceivedToDate & ""
                Screen.ActiveForm.txtStatusCode.Text = !StatusCode & ""
                Screen.ActiveForm.txtUnitCountBeforePayment = !UnitCount & ""
                Screen.ActiveForm.txtSuspenseAccount = !SuspenseAccount & ""
                Screen.ActiveForm.txtAccountNo.Text = !AccountNo & ""
                Screen.ActiveForm.txtDueDate.Text = !DueDate & ""
                
                If Screen.ActiveForm.txtDateOfLastPayment.Text > "" Then
                        Screen.ActiveForm.txtDateOfLastPayment.Text = !DateofLastPayment
                End If
                
                Screen.ActiveForm.txtCurrentPeriodDetails = CurrentPeriod
                Screen.ActiveForm.txtNames.Text = !OtherNames + " " + !CompanyName
                Screen.ActiveForm.txtAccountNo.Text = !AccountNo
        End With
    rsPOLICY.Close

End Sub
Private Sub verifyLOANDETAILS()
On Error GoTo err

        Dim rsPOLICY As ADODB.Recordset, strPOLICY As String
        Set rsPOLICY = New Recordset
                         
        rsPOLICY.Open "SELECT * FROM ODASPAccount, ALISMLoanManagement WHERE ALISMLoanManagement.LoanNo = '" & Screen.ActiveForm.cboDocumentNo.Text & "' and ALISMLoanManagement.AccountNo LIKE ODASPAccount.AccountNo ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsPOLICY
                If .EOF Or .BOF Then
                        MsgBox "The Policy Does not Exist within the Database", vbOKOnly
                        Exit Sub
                End If
                
                Screen.ActiveForm.cboDocumentNo.Text = !LoanNo
                Screen.ActiveForm.txtExpectedAmount = !RepaymentAmount
                Screen.ActiveForm.cboAccountNo = !AccountCode & ""
                Screen.ActiveForm.txtReceivedTodate = !TotalReceivedTodate & ""
                Screen.ActiveForm.txtStatusCode.Text = !Status & ""
                Screen.ActiveForm.txtUnitCount = !UnitCount & ""
                Screen.ActiveForm.txtSuspenseAccount = !suspense & ""
                Screen.ActiveForm.txtAccountNo.Text = !AccountNo & ""
                Screen.ActiveForm.txtDueDate.Text = !DueDate & ""
                
                If Screen.ActiveForm.txtDateOfLastPayment.Text > "" Then
                        Screen.ActiveForm.txtDateOfLastPayment.Text = !DateofLastPayment
                End If
                
                Screen.ActiveForm.txtCurrentPeriodDetails = CurrentPeriod
                Screen.ActiveForm.txtNames.Text = !OtherNames + " " + !CompanyName
                Screen.ActiveForm.txtAccountNo.Text = !AccountNo
        End With
    rsPOLICY.Close

Exit Sub

err:
    ErrorMessage
End Sub


Public Sub UpdatePolicyLedger()

On Error GoTo err

With Screen.ActiveForm
    If .cboReceiptType.Text = "BIP" Or .txtReceiptNoDetails.Text <= "" Then Exit Sub
    
    Set rsFindRecord = New ADODB.Recordset
    rsFindRecord.Open "SELECT * FROM ODASPAccount WHERE AccountNo LIKE '" & .txtAccountNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
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
    If rsRCPT!accountsreceivable = "1" Then Exit Sub
    
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
            rsDEP!SuspenseAccount = CDbl(.txtSuspenseAccount.Text)
            rsDEP!PaymentMethod = "SO"
            rsDEP!BankNo = .txtBankNo.Text
            rsDEP!ChequeNo = .txtChequeNo.Text
            If .txtDateOfLastPayment.Text > "" Then
                rsDEP!DateofLastPayment = .txtDateOfLastPayment.Text
            End If
            rsDEP!ReceivedToDate = .txtReceivedTodate.Text
            rsDEP!UnitCount = .txtUnitCount.Text
            rsDEP!UnitCountBeforePayment = .txtUnitCountBeforePayment
            rsDEP!StatusCode = .txtStatusCode.Text
            rsDEP!TransactionDate = Date
            rsDEP!TransactionNo = .txtTransactionNODetails.Text
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

Public Sub createRECEIPT()
On Error GoTo err
        Dim rsTRANS As ADODB.Recordset, strTRANS As String
        Set rsTRANS = New Recordset
        
        rsTRANS.Open "SELECT * FROM ALISMReceiptNew WHERE ReceiptNo = '" & Screen.ActiveForm.txtReceiptNoDetails.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
                
        With rsTRANS
                If .BOF Or .EOF = True Then Exit Sub
                        Screen.ActiveForm.txtTransactionNODetails = !TransactionNo & ""
                        
                With Screen.ActiveForm
                        .txtTransactionNODetails.Text = CDbl(.txtTransactionNODetails.Text) + 1
                End With
        
        End With
rsTRANS.Close
strTRANS = ""
Exit Sub

err:
    UpdateErrorMessage
End Sub

Public Sub processUPDATE()
        createRECEIPT
        validateRECORD
        If bsaveRECORD = True Then
                updateRECEIPTDATA
                updateRECEIPTBALANCE
                saveRecord
                enableCBRECEIPT
                disableDATAENTRY
                
                'LoadReceiptGrid
        End If

End Sub


Public Sub selectRECEIPTTYPELOSTFOCUS()
On Error GoTo err

        Set rsRCPT = New Recordset
        
        rsRCPT.Open "SELECT * FROM ALISPReceipt WHERE Description= '" & Screen.ActiveForm.cboReceiptType.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsRCPT
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboReceiptType.Text = !ReceiptType
                Screen.ActiveForm.txtTransactionDate.Text = Date
                Screen.ActiveForm.txtPaymentStatusDetails.Text = "PAID"
                
                '/ Enable the Employer Code Control if the receipt type is set to
                '/ Employer = "1" else locks it up
                'enableALLRECORD
                
                If !Account = "1" Then
                        Screen.ActiveForm.cboAccountNo.Locked = False
                        Screen.ActiveForm.txtAccountNo.Locked = False
                        
                        bEmployerUpdate = True
                        
                ElseIf !accountsreceivable = "1" Then
                        Screen.ActiveForm.lblDocumentNo.Caption = "Account No"
                ElseIf !IssuePolicy = "1" Then
                        Screen.ActiveForm.lblDocumentNo.Caption = "Policy No"

                Else
                        Screen.ActiveForm.cboAccountNo.Locked = True
                        Screen.ActiveForm.txtAccountNo.Locked = True
                End If
                
        End With
        
        '/ This portion is used to set the default details for the receipting
        '/ Module
        
        Screen.ActiveForm.txtPaymentStatus.Text = "PAID"
        Screen.ActiveForm.txtCurrentPeriodDetails = CurrentPeriod
        Screen.ActiveForm.txtReceiptDate.Text = Date
       
        Exit Sub

err:
        ErrorMessage

End Sub

Public Sub selectISSUEDRECEIPT()
On Error GoTo err

        Set rsRCPT = New Recordset
        
        rsRCPT.Open "SELECT * FROM ALISPReceipt WHERE IssuePolicy = '1'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsRCPT
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboReceiptType.Text = !ReceiptType
                Screen.ActiveForm.txtTransactionDate.Text = Date
        End With
        
        '/ This portion is used to set the default details for the receipting
        '/ Module
        
        Screen.ActiveForm.txtPaymentStatus.Text = "PAID"
        Screen.ActiveForm.txtCurrentPeriodDetails = CurrentPeriod
        Screen.ActiveForm.txtReceiptDate.Text = Date
       
        Exit Sub

err:
        ErrorMessage

End Sub


Private Sub loadEMPLOYER()
On Error GoTo err

        Dim rsEMP As ADODB.Recordset, strEMP As String
        Set rsEMP = New Recordset
        
        rsEMP.Open "SELECT * FROM ODASPAccount WHERE AccountNo = '" & Screen.ActiveForm.cboAccountNo.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsEMP
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboEmployer = !CompanyName
                Screen.ActiveForm.cboAccountNo = !AccountCode
        End With

rsEMP.Close

Exit Sub

err:
    UpdateErrorMessage
End Sub

Public Sub addRECORD()
    addpen = True
    clearRECORD
    loadRECORD
    enableDATAENTRY
    DisableCBRECEIPT
    Screen.ActiveForm.txtReceiptNo.SetFocus
End Sub
Public Sub loadRECORD()
On Error GoTo err
    With Screen.ActiveForm
            .txtUnitCount.Text = 0
    End With
Exit Sub

err:
    ErrorMessage
End Sub


Public Sub selectCURRENCYCODEKEYPRESS(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub

Public Sub selectPaymentMethodKeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Public Sub selectPAYMENTMODEKEYPRESS(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Public Sub selectPlanProposalNo_KeyPress(KeyAscii As Integer)

   KeyAscii = 0
    
End Sub
Public Sub selectDOCUMENTLOSTFOCUS()
On Error GoTo err
    
    If Screen.ActiveForm.cboReceiptType > "" Then
        
        Set rsRCPT = New Recordset
    
        strSQL = "SELECT * FROM ALISPReceipt Where ReceiptType = '" & Screen.ActiveForm.cboReceiptType & "';"
        rsRCPT.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        
        With rsRCPT
            If .EOF Or .BOF = True Then
                    MsgBox "The Receipt Selected is Invalid", vbOKCancel
                    Exit Sub
            
            ElseIf !Deposit = 1 Then
                    verifyACCOUNTdetails
                    loadEMPLOYER
            ElseIf !Jobbrief = 1 Then
                    verifyJOBBRIEFDETAILS
                    loadEMPLOYER
            ElseIf !IssuePolicy = 1 Then
                    verifyJOBBRIEFDETAILS
                    loadEMPLOYER
            ElseIf !loan = 1 Then
                    calculateLOANAMOUNTDUE
                    verifyLOANDETAILS
            End If
                    '/* end if the Statement Statement
                    
        End With
        
        Call ReferenceDetails
        Call loadloanGRID
        Call loadproposalGRID
        Call loadpolicyGRID
    End If

Exit Sub

err:
        ErrorMessage

End Sub
Private Sub loadPAYMENTMETHOD()
On Error GoTo err
        Dim rsPAY As ADODB.Recordset, strPAY As String
        Set rsPAY = New Recordset

        rsPAY.Open "SELECT * FROM ALISPPaymentMethod WHERE PaymentMethod = '" & Screen.ActiveForm.txtPaymentMethod.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        With rsPAY
                If .BOF Or .EOF Then Exit Sub
                Screen.ActiveForm.cboPaymentMethod.Text = !PaymentMethodDescription
        End With
    
rsPAY.Close
strPAY = ""

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub loadBANK()
On Error GoTo err
        Dim rsPAY As ADODB.Recordset, strPAY As String
        Set rsPAY = New Recordset

        rsPAY.Open "SELECT * FROM ALISPBank WHERE BankNo = '" & Screen.ActiveForm.txtBankNo & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        With rsPAY
                If .BOF Or .EOF Then Exit Sub
                Screen.ActiveForm.cboBankNo.Text = !CompanyName
        End With
        
Exit Sub

err:
    ErrorMessage
End Sub



Private Sub ReferenceDetails()
On Error GoTo err

        Dim rsREF As ADODB.Recordset, strREF As String
        Set rsREF = New Recordset

        rsREF.Open "SELECT * FROM ODASPAccount WHERE AccountNo LIKE '" & Screen.ActiveForm.txtAccountNo & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
                    
        Screen.ActiveForm.txtNames = "Names:  " + rsREF!OtherNames + " " + rsREF!CompanyName
        
        If Screen.ActiveForm.cboAccountNo.Text <= "" Then
                Screen.ActiveForm.txtAccountNo.Text = rsREF!AccountNo & ""
                Screen.ActiveForm.cboAccountNo.Text = rsREF!AccountCode & ""
        End If
        
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

Exit Sub
Myerr:
    ErrorMessage
End Sub

Public Sub validateRECORD()

On Error GoTo err

    If Screen.ActiveForm.txtReceiptNoDetails.Text = "" Then
        MsgBox "Receipt No is Required"
        Screen.ActiveForm.txtReceiptNoDetails.SetFocus
    
    ElseIf Screen.ActiveForm.txtTransactionNODetails.Text = "" Then
        MsgBox "Transaction No is Required"
        Screen.ActiveForm.txtTransactionNODetails.SetFocus

    ElseIf Screen.ActiveForm.txtReceiptDate.Text Like 0 Then
        MsgBox "The Receipt Date is  required"
        Screen.ActiveForm.txtReceiptDate.SetFocus
    
    ElseIf DateDiff("d", Screen.ActiveForm.txtReceiptDate.Text, Date) < 0 Then
        MsgBox "Payment cannot be accepted in advance"
        Screen.ActiveForm.txtReceiptDate.SetFocus
    
    ElseIf Screen.ActiveForm.txtTransactionDate.Text Like 0 Then
        MsgBox "The Transaction Date is  required"
        Screen.ActiveForm.txtTransactionDate.SetFocus
    
    ElseIf DateDiff("d", Screen.ActiveForm.txtTransactionDate.Text, Date) < 0 Then
        MsgBox "Transaction Date cannot be accepted in advance"
        Screen.ActiveForm.txtTransactionDate.SetFocus

    ElseIf Screen.ActiveForm.cboDocumentNo.Text = "" Then
        MsgBox "The Document Number is Mandatory for this Process"
        Screen.ActiveForm.cboDocumentNo.SetFocus
    
    ElseIf Screen.ActiveForm.txtTransactionAmount.Text Like 0 Then
        MsgBox "The Receipt Amount is required and There cannot be 0"
        Screen.ActiveForm.txtTransactionAmount.SetFocus
    
    ElseIf Screen.ActiveForm.cboReceiptType.Text = "" Then
        MsgBox "The Receipt is required"
        Screen.ActiveForm.cboReceiptType.SetFocus
    Else
        bsaveRECORD = True

    End If

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub saveRecord()
On Error GoTo err

    Set rsRCPT = New Recordset
    
    strSQL = "SELECT * FROM ALISPReceipt Where ReceiptType = '" & Screen.ActiveForm.cboReceiptType & "';"
    rsRCPT.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
 
    If rsRCPT!Deposit = 1 Then
        updateACCOUNTS
    ElseIf rsRCPT!Jobbrief = 1 Then
        updateJOBBRIEFDATA
    ElseIf rsRCPT!accounts = 1 Then
        updateACCOUNTS
    End If


Exit Sub

err:
        ErrorMessage
End Sub
Private Sub updateRECEIPTDATA()
On Error GoTo err

    Dim rsDEP As ADODB.Recordset, strDEP As String
    Set rsDEP = New Recordset
    
    rsDEP.Open "select * from ALISMReceiptDetails ; ", cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsDEP
            .AddNew
            !ReceiptNo = Screen.ActiveForm.txtReceiptNoDetails
            !TransactionNo = Screen.ActiveForm.txtTransactionNODetails & ""
            !TransactionDate = Screen.ActiveForm.txtTransactionDate
            !ReceiptType = Screen.ActiveForm.cboReceiptType
            !DocumentNo = Screen.ActiveForm.cboDocumentNo
            !AccountNo = CDbl(Screen.ActiveForm.txtAccountNo.Text)
            !TransactionAmount = CDbl(Screen.ActiveForm.txtTransactionAmount)
            !AccountingPeriod = Screen.ActiveForm.txtCurrentPeriodDetails
            !PaymentStatus = Screen.ActiveForm.txtPaymentStatus
            !DateofLastPayment = Screen.ActiveForm.txtReceiptDate.Text
            
            If Screen.ActiveForm.txtUnitsPaid <= "" Then
                !UnitsPaid = 0
                Else: !UnitsPaid = Screen.ActiveForm.txtUnitsPaid
            End If
            
            !AccountCode = Screen.ActiveForm.cboAccountNo & ""
            
            If Screen.ActiveForm.txtDueDate.Text > "" Then
                !DueDate = Screen.ActiveForm.txtDueDate
            End If
            
            !SuspenseAccount = CDbl(Screen.ActiveForm.txtSuspenseAccount)
            
            If Screen.ActiveForm.txtDateOfLastPayment <= "" Then
                    Screen.ActiveForm.txtDateOfLastPayment = frmODASMReceipt.txtReceiptDate.Text
                Else:
                        !DateofLastPayment = Screen.ActiveForm.txtDateOfLastPayment
            End If
            
            !ReceivedToDate = CDbl(Screen.ActiveForm.txtReceivedTodate) & ""
            If Screen.ActiveForm.txtUnitCount.Text <= "" Then
                    !UnitCount = 0
            Else: !UnitCount = CDbl(Screen.ActiveForm.txtUnitCount)
            End If
            
            !UnitCountBeforePayment = CDbl(Screen.ActiveForm.txtUnitCount) & ""

            If Screen.ActiveForm.txtDueDate.Text = "" Or Screen.ActiveForm.txtDueDate.Text = "0" Then
                    Screen.ActiveForm.txtDueDate.Text = Date
                    
                    Else: !DueDate = Screen.ActiveForm.txtDueDate & ""
            End If
            
            If Screen.ActiveForm.txtStatusCode <= "" Or Screen.ActiveForm.txtStatusCode = "10" Then
                    Screen.ActiveForm.txtStatusCode = "INFORCE"
            End If
            
            !StatusCode = Screen.ActiveForm.txtStatusCode
            !TransactionDate = Date
            
            .Update
    End With
rsDEP.Close

Exit Sub

err:
    UpdateErrorMessage
End Sub

Private Sub updateRECEIPTBALANCE()
On Error GoTo err

        Dim strDATA, strCALC As String
        Dim rsDATA, rsCALC As ADODB.Recordset
        
        Set rsCALC = New ADODB.Recordset
        
        strCALC = "SELECT sum(ALISMReceiptDetails.TransactionAmount) as totals from ALISMReceiptDetails where ReceiptNo = '" & Screen.ActiveForm.txtReceiptNoDetails.Text & "' ;"
        rsCALC.Open strCALC, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        Screen.ActiveForm.txtTotalAmount.Text = rsCALC!TOTALS
        
        Set rsDATA = New Recordset
        
        strDATA = "SELECT * FROM ALISMReceiptNew Where ReceiptNo = '" & Screen.ActiveForm.txtReceiptNoDetails.Text & "';"
        rsDATA.Open strDATA, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsDATA
                If .BOF Or .EOF = True Then Exit Sub
                
                Screen.ActiveForm.txtBalance = !ReceiptAmount - CDbl(Screen.ActiveForm.txtTotalAmount.Text)
                
                If CDbl(Screen.ActiveForm.txtBalance.Text) <= 0 Then
                    addpen = False
                    enableCBRECEIPT
                End If
                
                !TotalAmount = Screen.ActiveForm.txtTotalAmount.Text
                !TransactionNo = Screen.ActiveForm.txtTransactionNODetails
                .Update
                .Requery
        End With

rsCALC.Close
rsDATA.Close

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub updatePROPOSALDATA()

        Dim rsPROP As ADODB.Recordset, strPROP As String
        Set rsPROP = New Recordset
            
                rsPROP.Open "SELECT * FROM ALISMProposal WHERE proposalNo =  '" & Screen.ActiveForm.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                With rsPROP
                    !ProposalNo = Screen.ActiveForm.cboDocumentNo
                    !DepositReceived = !DepositReceived + CDbl(Screen.ActiveForm.txtReceiptAmount)
                    !DateofLastPayment = Screen.ActiveForm.txtReceiptDate
                    !StatusCode = "DepRecvd"
                    .Update
                End With
        
        rsPROP.Close
        strPROP = ""
End Sub

Private Sub updateISSUANCE()
        
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

End Sub
Private Sub updateLOANDATA()
    Dim rsLOAN As ADODB.Recordset, strLOAN As String
    Set rsLOAN = New Recordset
        
    rsLOAN.Open "SELECT * FROM ALISMLoanManagement WHERE LoanNo =  '" & Screen.ActiveForm.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
            
            With rsLOAN
            
                !JobBriefNo = Screen.ActiveForm.cboDocumentNo
                
                If Screen.ActiveForm.txtReceiptAmount > !RepaymentAmount Then
                        !suspense = (Screen.ActiveForm.txtReceiptAmount - !RepaymentAmount)
                ElseIf (Screen.ActiveForm.txtReceiptAmount <= !RepaymentAmount) And (Screen.ActiveForm.txtReceiptAmount > !PrincipalAmount) Then
                        Screen.ActiveForm.txtReceiptAmount = 0
                ElseIf (Screen.ActiveForm.txtReceiptAmount <= !PrincipalAmount) Then
                        Screen.ActiveForm.txtReceiptAmount = 0
                End If
                
                !TotalReceivedTodate = !TotalReceivedTodate + CDbl(Screen.ActiveForm.txtTransactionAmount)
                !DateofLastPayment = Screen.ActiveForm.txtReceiptDate
                !DueDate = Screen.ActiveForm.txtDueDate
                !suspense = Screen.ActiveForm.txtSuspenseAccount & ""
                
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
End Sub
Private Sub updateJOBBRIEFDATA()
        Dim rsPOL As ADODB.Recordset, strPOL As String
        Set rsPOL = New Recordset
            
                rsPOL.Open "SELECT * FROM ODASMJobBrief WHERE JobBriefNo =  '" & Screen.ActiveForm.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                With rsPOL
                    !JobBriefNo = Screen.ActiveForm.cboDocumentNo
                    !DepositReceived = 0
                    !DateofLastPayment = Screen.ActiveForm.txtReceiptDate
                    !UnitCount = CDbl(Screen.ActiveForm.txtUnitCount)
                    !SuspenseAccount = Screen.ActiveForm.txtSuspenseAccount
                    !ReceivedToDate = Screen.ActiveForm.txtReceivedTodate & ""
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

End Sub
Private Sub updateREVIVALDATA()
        Dim rsPOL As ADODB.Recordset, strPOL As String
        Set rsPOL = New Recordset
            
        rsPOL.Open "SELECT * FROM ODASMJobBrief WHERE JobBriefNo =  '" & Screen.ActiveForm.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                With rsPOL
                    !JobBriefNo = Screen.ActiveForm.cboDocumentNo
                    !DepositReceived = 0
                    !DateofLastPayment = Screen.ActiveForm.txtReceiptDate
                    !UnitCount = Screen.ActiveForm.txtUnitCount
                    !SuspenseAccount = Screen.ActiveForm.txtSuspenseAccount
                    !ReceivedToDate = Screen.ActiveForm.txtReceivedTodate & ""
                    !UnitCountBeforePayment = Screen.ActiveForm.txtUnitCountBeforePayment
                    !DueDate = Screen.ActiveForm.txtDueDate
                    !ReinstatementCharge = Screen.ActiveForm.txtTransactionAmount.Text
                    !StatusCode = "INFORCE"
                    !StatusDate = Date
                    .Update
                End With
        
        rsPOL.Close
 
End Sub
Private Sub updateREINSTATEMENTFILE()
        Dim rsREINSTATEMENT As ADODB.Recordset, strPOL As String
        Set rsREINSTATEMENT = New Recordset
            
        rsREINSTATEMENT.Open "SELECT * FROM ALISMReinstatement WHERE JobBriefNo =  '" & Screen.ActiveForm.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                With rsREINSTATEMENT
                    !Reinstated = "Y"
                    !ReceiptDate = Date
                    .Update
                    .Requery
                End With
        
        rsREINSTATEMENT.Close
 
End Sub

Private Sub updateACCOUNTSDATA()

        Dim rsACCOUNT As ADODB.Recordset
        Set rsACCOUNT = New Recordset
            
        rsACCOUNT.Open "SELECT * FROM ODASMACCOUNT  WHERE AccountNo =  '" & Screen.ActiveForm.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
        With rsACCOUNT
            .AddNew
            !AccountNo = Screen.ActiveForm.cboDocumentNo.Text
            !TransactionType = Screen.ActiveForm.cboReceiptType.Text
            !TransactionDate = Screen.ActiveForm.txtTransactionDate.Text
            !InclusiveAmount = CDbl(Screen.ActiveForm.txtTransactionAmount.Text)
            !Amount = CDbl(Screen.ActiveForm.txtTransactionAmount.Text)
            !VAT = 0
            !VATRate = 0
            !Balance = !InclusiveAmount
            !Status = "PAID"
            !DueDate = Date
            !ReceiptNo = Screen.ActiveForm.txtReceiptNo.Text
            !InvoiceNo = Screen.ActiveForm.cboEmployer.Text
            !DateofLastPayment = Screen.ActiveForm.txtReceiptDate
            !ReceivedToDate = Screen.ActiveForm.txtReceivedTodate
            !dateprepared = Date
            !Preparedby = CurrentUserName
            .Update
            .Requery
        End With
        
        rsACCOUNT.Close

End Sub


Public Sub searchRECORD()
    clearRECORD
    locateRECORD
    If bsearchRECORD = True Then
        bsearchRECORD = False
        loadBANK
        loadPAYMENTMETHOD
        loadEMPLOYER
        loadRECEIPTDETAILSGRID
    End If
End Sub

Private Sub locateRECORD()
''''On Error GoTo Myerr

        Dim strQRE As Variant
        Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset
            strQRE = InputBox("Enter Transaction No to search.", "Search Value")
            rsFind.Open "SELECT ALISMReceiptDetails.*, ALISPReceipt.* FROM ALISMReceiptdetails, ALISPReceipt WHERE ALISMReceiptDetails.ReceiptNo = '" & Screen.ActiveForm.txtReceiptNo.Text & "' and ALISMReceiptDetails.transactionNO LIKE '" & strQRE & "' and ALISPReceipt.ReceiptType = ALISMReceiptDetails.ReceiptType;", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsFind
            If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
            Else:
                    Screen.ActiveForm.txtTransactionNo.Text = !TransactionNo
                    Screen.ActiveForm.txtTransactionDate.Text = !TransactionDate
                    Screen.ActiveForm.txtTransactionAmount.Text = !TransactionAmount
                    
                    Screen.ActiveForm.cboReceiptType.Text = !Description
                    Screen.ActiveForm.cboDocumentNo.Text = !DocumentNo
                    Screen.ActiveForm.txtAccountNo.Text = !AccountNo
                    Screen.ActiveForm.txtCurrentPeriodDetails.Text = !AccountingPeriod
                    Screen.ActiveForm.txtPaymentStatusDetails.Text = !PaymentStatus
                    Screen.ActiveForm.cboAccountNo.Text = !AccountCode & ""
                    
                    If IsNull(!DateofLastPayment) = False Then
                        Screen.ActiveForm.txtDateOfLastPayment.Text = !DateofLastPayment
                    End If
            
                    Screen.ActiveForm.txtReceivedTodate.Text = !ReceivedToDate
                    Screen.ActiveForm.txtUnitCount.Text = !UnitCount
                    Screen.ActiveForm.txtSuspenseAccount.Text = !SuspenseAccount & ""
                    Screen.ActiveForm.txtUnitsPaid.Text = !UnitsPaid
                    Screen.ActiveForm.txtUnitCountBeforePayment.Text = !UnitCountBeforePayment
                    Screen.ActiveForm.txtDueDate.Text = !DueDate & ""
                    Screen.ActiveForm.txtStatusCode.Text = !StatusCode
  
                    If rsFind!Jobbrief = "1" Then
                            
                            Dim rsPOL As ADODB.Recordset
                            Set rsPOL = New ADODB.Recordset
                            rsPOL.Open "SELECT ODASMJobBrief.*, ODASPAccount.* FROM ODASMJobBrief, ODASPAccount WHERE JobBriefNo = '" & Screen.ActiveForm.cboDocumentNo.Text & "' and ODASPAccount.AccountNo LIKE ODASMJobBrief.AccountNo ; ", cnCOMMON, adOpenKeyset, adLockOptimistic

                            Screen.ActiveForm.txtExpectedAmount.Text = rsPOL!ExpectedPremium
                            Screen.ActiveForm.txtNames.Text = Trim(rsPOL!OtherNames) + " " + Trim(rsPOL!CompanyName)
                            Screen.ActiveForm.txtAccountNo.Text = rsPOL!AccountNo
                            
                            rsPOL.Close
                    
                    ElseIf rsFind!Deposit = "1" Then
                            
                            Dim rsPRO As ADODB.Recordset
                            Set rsPRO = New ADODB.Recordset
                            rsPRO.Open "SELECT ALISMProposal.*, ODASPAccount.* FROM ALISMProposal, ODASPAccount WHERE ProposalNo = '" & Screen.ActiveForm.cboDocumentNo.Text & "' and ODASPAccount.AccountNo LIKE ALISMProposal.AccountNo ; ", cnCOMMON, adOpenKeyset, adLockOptimistic

                            Screen.ActiveForm.txtExpectedAmount.Text = rsPRO!ExpectedPremium
                            Screen.ActiveForm.txtNames.Text = Trim(rsPRO!OtherNames) + " " + Trim(rsPRO!CompanyName)
                            Screen.ActiveForm.txtAccountNo.Text = rsPRO!AccountNo
                            
                    rsPRO.Close

                    End If
                    
                    bsearchRECORD = True
                    Edit = True
            End If
        
        End With

Exit Sub

Myerr:
    ErrorMessage
End Sub

Private Sub Class_Initialize1()
On Error GoTo err
    'Open the connection
    Set cnCOMMON = New ADODB.Connection
    cnCOMMON.ConnectionString = "PROVIDER=MSDASQL;DSN=DSNALIS;UID=;PWD=;"
    
    cnCOMMON.Open
    'Begin a transaction
    ''cnCOMMON.BeginTrans
       
    ''Set rsReceipt = New ADODB.Recordset
    ''rsReceipt.Open "SELECT * FROM ALISPReceipt;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    ' 'Call EnableCBPolicyControls
    
    'Commit the transaction
    cnCOMMON.CommitTrans

Exit Sub
err:
'Incase of an error the connection object shoud call its rollback object
'if an error occured, roll back the changes
'cnCOMMON.RollbackTrans
MsgBox "An error occured changing the records.", vbExclamation
ErrorMessage
End Sub

Private Sub Class_Terminate1()
On Error GoTo err
    'Close the instance of a recordset
    rsreceipt.Close
    
    'Destroy the instance of the recordset
    Set rsreceipt = Nothing
Exit Sub

err:
    ErrorMessage
End Sub



Public Sub verifyTHEAMOUNTRECEIVED()
On Error GoTo err
    If CDbl(Screen.ActiveForm.txtReceiptAmount) <= 0 Or Screen.ActiveForm.txtReceiptAmount <= "" Then
            MsgBox "The Receipt Amount Must be > 0 "
                    Screen.ActiveForm.txtReceiptAmount.SetFocus
            Exit Sub
            
    Else:
        Set rsRCPT = New Recordset
        
        rsRCPT.Open "SELECT * FROM ALISPReceipt WHERE ReceiptType= '" & Screen.ActiveForm.cboReceiptType.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
      
            
           If rsRCPT!Jobbrief = 1 Or rsRCPT!IssuePolicy = 1 Then
                 Call updateRECORD
            ElseIf rsRCPT!loan = 1 Then
                Call updateLOANRECORD
            End If
    End If
Exit Sub

err:
    ErrorMessage
End Sub


