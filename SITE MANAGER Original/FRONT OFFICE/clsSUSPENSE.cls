VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSUSPENSE"
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
Public Sub processReceipt()
On Error GoTo err
                If frmALISMSuspense.cboDocumentNo.Text <= "" Then Exit Sub
                
                frmALISMSuspense.txtDateOfLastPayment.Text = frmALISMSuspense.txtReceiptDate.Text

        
'                     If CDbl(frmALISMSuspense.txtTransactionAmount.Text) = " " Then
'                             frmALISMSuspense.txtTransactionAmount.Text = 0
'                     End If
                        If CDbl(frmALISMSuspense.txtTransactionAmount.Text) >= CDbl(frmALISMSuspense.txtExpectedAmount.Text) Then
                                
                                Set rsRCPT = New ADODB.Recordset
                                rsRCPT.Open "SELECT * FROM ALISPReceipt Where receiptType = '" & frmALISMSuspense.cboReceiptType.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                                
                                If rsRCPT.BOF Or rsRCPT.EOF Then Exit Sub
'                                updateRECORD
                        End If
    
Exit Sub

err:
    UpdateErrorMessage
End Sub
Public Sub updateEMPLOYER()
On Error GoTo err
                
    frmALISMSuspense.txtSuspenseAccount.Text = 0
    frmALISMSuspense.txtReceivedTodate.Text = 0
    frmALISMSuspense.txtDateOfLastPayment = Date
    frmALISMSuspense.txtReceivedTodate = CDbl(frmALISMSuspense.txtReceiptAmount)
                           
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
            
        rsLOANDUE.Open "SELECT * FROM ALISMLoanManagement Where ALISMLoanManagement.LoanNo = '" & frmALISMSuspense.cboDocumentNo & "'  ;", cnCOMMON, adOpenKeyset, adLockOptimistic

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
            
        rsPOLICY.Open "SELECT * FROM ODASMJobBrief Where JobBriefNo = '" & frmALISMSuspense.cboDocumentNo & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
        
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
        Set rsPOL = New ADODB.Recordset
            
        rsPOL.Open "SELECT * FROM ODASMJobBrief, ODASPPaymentMode WHERE ODASMJobBrief.JobBriefNo =  '" & frmALISMSuspense.cboDocumentNo & "' and ODASMJobBrief.PaymentMode = ODASPPaymentMode.PaymentMode ;", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsPOL

                                
                '/* this Procedure is used to calculate the number of Premiums
                '/* paid by the Policyholder
                
                If (CDbl(frmALISMSuspense.txtTransactionAmount.Text) >= !ExpectedPremium) And (!UnitCount < !TermOfPolicy * 12) And ((Trim(rsPOL!StatusCode) = "10" Or Trim(rsPOL!StatusCode) = "INFORCE") Or Trim(rsPOL!StatusCode = "UNISSUED")) Then

                        frmALISMSuspense.txtUnitsPaid = Fix(frmALISMSuspense.txtTransactionAmount / frmALISMSuspense.txtExpectedAmount) * rsPOL!CoverPeriod
                        frmALISMSuspense.txtDateOfLastPayment = frmALISMSuspense.txtReceiptDate
                        frmALISMSuspense.txtSuspenseAccount = CDbl(frmALISMSuspense.txtTransactionAmount - (frmALISMSuspense.txtUnitsPaid * (frmALISMSuspense.txtExpectedAmount / rsPOL!CoverPeriod)))
                        
                              
                        '/If the total premiums exceed the expected number of Premiums
                        
                                               
                        '/* check if the Maturity date been reached
                
                        frmALISMSuspense.txtSuspenseAccount = CDbl(frmALISMSuspense.txtTransactionAmount)
                        frmALISMSuspense.txtStatusCode.Text = !StatusCode
                        frmALISMSuspense.txtDueDate.Text = !DueDate
                        
      End If

rsPOL.Close

End With


Exit Sub


err:
    ErrorMessage
End Sub

Public Sub updateREVIVAL()
On Error GoTo err
        
        Dim rsPOL As ADODB.Recordset, strPOL As String
        Set rsPOL = New Recordset
            
        rsPOL.Open "SELECT * FROM ODASMJobBrief, ALISMReinstatement WHERE ODASMJobBrief.JobBriefNo =  '" & frmALISMSuspense.cboDocumentNo & "' and ODASMJobBrief.statusCode = 'LAPSED' and ODASMJobBrief.JobBriefNo = ALISMReinstatement.JobBriefNo and ALISMReinstatement.Approved = 'Y' and ALISMReinstatement.Reinstated ='N' ;", cnCOMMON, adOpenKeyset, adLockOptimistic

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
                
                frmALISMSuspense.txtSuspenseAccount.Text = !SuspenseAccount
                frmALISMSuspense.txtReceivedTodate.Text = !ReceivedToDate & ""
                
                                
                '/* this Procedure is used to calculate the number of Premiums
                '/* paid by the Policyholder
                
                If CDbl(frmALISMSuspense.txtTransactionAmount.Text) >= 0 And (!UnitCount < !TermOfPolicy * 12) Then
                        frmALISMSuspense.txtUnitsPaid = 0
                        
                        frmALISMSuspense.txtDateOfLastPayment = frmALISMSuspense.txtReceiptDate
                        frmALISMSuspense.txtSuspenseAccount = !SuspenseAccount
                        
                        If frmALISMSuspense.txtReceivedTodate <= "" Then
                                        frmALISMSuspense.txtReceivedTodate = CDbl(frmALISMSuspense.txtTransactionAmount)
                                Else: frmALISMSuspense.txtReceivedTodate = !ReceivedToDate + CDbl(frmALISMSuspense.txtTransactionAmount)
                        End If
                        
                        '/If the total premiums exceed the expected number of Premiums
                
                        
                                                       
                        '/* check if the Maturity date been reached
                
                        frmALISMSuspense.txtSuspenseAccount = CDbl(!SuspenseAccount) + CDbl(frmALISMSuspense.txtTransactionAmount)
                        frmALISMSuspense.txtUnitsPaid.Text = 0
                        frmALISMSuspense.txtStatusCode.Text = !StatusCode
                        frmALISMSuspense.txtReceivedTodate = !ReceivedToDate + frmALISMSuspense.txtTransactionAmount
                        frmALISMSuspense.txtDueDate.Text = !DueDate
                        
                        
                        frmALISMSuspense.txtDueDate.Text = !DueDate & ""
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
            
        rsACCOUNT.Open "SELECT * FROM MEGAMAccount WHERE AccountNo =  '" & frmALISMSuspense.cboDocumentNo & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsACCOUNT
                If .BOF Or .EOF Then Exit Sub
                'frmALISMSuspense.txtTransactionAmount.Text = !InclusiveAmount
                frmALISMSuspense.txtDateOfLastPayment = frmALISMSuspense.txtReceiptDate
                frmALISMSuspense.txtStatusCode = "PAID"
        End With
        
rsACCOUNT.Close
strPOL = ""

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub processDOCUMENTNO()

        If frmALISMSuspense.cboDocumentNo.Text <= "" Then Exit Sub
        
        If frmALISMSuspense.cboReceiptType.Text <= "" Then
                            MsgBox "The Receipt Type Is Required Before executive This Function", vbOKOnly
                            frmALISMSuspense.cboReceiptType.SetFocus
        Else
                If frmALISMSuspense.cboReceiptType > "" Then
                        
                        Set rsRCPT = New ADODB.Recordset
                        rsRCPT.Open "SELECT * FROM ALISPReceipt Where receiptType = '" & frmALISMSuspense.cboReceiptType.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

                        With rsRCPT
                                If !Deposit = 1 Then
                                    verifyACCOUNTdetails
                                ElseIf !Jobbrief = 1 Or !Account = 1 Then
                                    verifyJOBBRIEFDETAILS
                                ElseIf !accounts = 1 Then
                                    verifyACCOUNTdetails
                                Else
                                        BPolicyHolder = False
                                    '/* rsRCPT is Loan */
                                End If
                                    '/* end if the Statement Statement
                                        Call loadEMPLOYER
                        
                                        frmALISMSuspense.txtTransactionAmount.SetFocus
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
        rsREVIVAL.Open "SELECT * FROM ALISMReinstatement Where JobBriefNo = '" & frmALISMSuspense.cboDocumentNo.Text & "' and Reinstated = 'N' and Approved = 'Y';", cnCOMMON, adOpenKeyset, adLockOptimistic
        
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



Public Sub processDOCUMENT()

        Set rsRCPT = New Recordset
        
        rsRCPT.Open "SELECT * FROM ALISPReceipt WHERE receiptType = '" & frmALISMSuspense.cboReceiptType.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
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

        rsDNLF.Open "SELECT ALISMProposal.*, ODASPAccount.* FROM ALISMProposal, ODASPAccount WHERE ALISMProposal.ProposalNo= '" & frmALISMSuspense.cboDocumentNo.Text & "' and ALISMProposal.AccountNo LIKE ODASPAccount.AccountNo ", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsDNLF
            If .EOF And .BOF Then Exit Sub
            
            frmALISMSuspense.cboDocumentNo.Text = !ProposalNo
            frmALISMSuspense.txtAccountNo.Text = !Accountno
                        
            If !DateofLastPayment > "" Then
                    frmALISMSuspense.txtDateOfLastPayment.Text = !DateofLastPayment & ""
            End If
            
            frmALISMSuspense.txtTransactionAmount.Text = !ExpectedPremium
            frmALISMSuspense.txtExpectedAmount.Text = !ExpectedPremium
            frmALISMSuspense.txtDueDate.Text = Date
            'frmALISMSuspense.cboAccountNo = !AccountCode & ""
            frmALISMSuspense.txtAccountNo = !Accountno & ""
            
            If IsNull(!SuspenseAccount) = True Then
                    frmALISMSuspense.txtSuspenseAccount = 0
                    Else: frmALISMSuspense.txtSuspenseAccount.Text = !SuspenseAccount
            End If
            frmALISMSuspense.txtNames.Text = Trim(!OtherNames) + " " + Trim(!CompanyName)
            If IsNull(!DepositReceived) = True Then
                    frmALISMSuspense.txtReceivedTodate.Text = 0
                Else: frmALISMSuspense.txtReceivedTodate.Text = !DepositReceived
            End If
            
            frmALISMSuspense.txtStatusCode.Text = !StatusCode
            BPolicyHolder = True
        End With
      
rsDNLF.Close
strDNLF = ""
      '/* rsRCPT is deposit */
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub LoadDEFAULT()
On Error GoTo err

    With frmALISMSuspense
        .txtPaymentStatus.Text = "PAID"
        .txtReceiptDate.Text = Date
        .cboReceiptType.Text = "SUS"
        .txtSuspenseAccount.Text = 0
    End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub verifyJOBBRIEFDETAILS()
        
        Dim rsPOLICY As ADODB.Recordset, strPOLICY As String
        Set rsPOLICY = New Recordset
                         
        rsPOLICY.Open "SELECT ODASMJobBrief.*, ODASPAccount.* FROM ODASMJobBrief, ODASPAccount WHERE ODASMJobBrief.JobBriefNo = '" & frmALISMSuspense.cboDocumentNo.Text & "' and ODASMJobBrief.AccountNo LIKE ODASPAccount.AccountNo ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsPOLICY
                If .EOF Or .BOF Then
                        MsgBox "The Policy Does not Exist within the Database", vbOKOnly
                        Exit Sub
                End If
                
                frmALISMSuspense.txtExpectedAmount = !Balance
                frmALISMSuspense.cboDocumentNo = !Accountno & ""
                frmALISMSuspense.txtReceivedTodate = !ReceivedToDate & ""
                frmALISMSuspense.txtStatusCode.Text = !Status & ""
                frmALISMSuspense.txtSuspenseAccount = !SuspenseAccount & ""
                frmALISMSuspense.txtAccountNo.Text = !Accountno & ""
                frmALISMSuspense.txtDueDate.Text = !DueDate & ""
                frmALISMSuspense.txtDateOfCommencement.Text = !CommencementDate & ""
                'frmALISMSuspense.txtMaturityDate.Text = !ExpiryDate
                
                If frmALISMSuspense.txtDateOfLastPayment.Text > "" Then
                        frmALISMSuspense.txtDateOfLastPayment.Text = !DateofLastPayment
                End If
                
                frmALISMSuspense.txtAccountingPeriod = CurrentPeriod
                frmALISMSuspense.txtNames.Text = !CompanyName
                frmALISMSuspense.txtAccountNo.Text = !Accountno
        End With
    rsPOLICY.Close

End Sub




Public Sub createRECEIPT()
On Error GoTo err
        Set rsCONTROL = New ADODB.Recordset
        
        strSQL = "SELECT * FROM ALISPLastNumber WHERE AutoSuspenseNo = 'Y';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
      
        With rsCONTROL
                If .BOF Or .EOF = True Then Exit Sub
                        frmALISMSuspense.txtReceiptNo.Text = !SuspenseNo & ""
                        !SuspenseNo = !SuspenseNo + 1
                        
                        .Update
                        .Requery
                 
        
        End With
rsCONTROL.Close
Exit Sub

err:
    UpdateErrorMessage
End Sub

Public Sub processUPDATE()
        createRECEIPT
        validateRECORD
        
        If bsaveRECORD = True Then
                updateRECEIPTDATA
                saveRecord
                enableButtons
                disableALLRECORD
                getRECEIPTS
        End If

End Sub


Private Sub loadEMPLOYER()
On Error GoTo err

        Set rsCONTROL = New Recordset
        
        rsCONTROL.Open "SELECT * FROM ODASPAccount WHERE AccountNo = '" & frmALISMSuspense.txtAccountNo.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .EOF And .BOF Then Exit Sub
                frmALISMSuspense.txtNames.Text = !CompanyName
        End With

rsCONTROL.Close

Exit Sub

err:
    UpdateErrorMessage
End Sub


Public Sub selectDOCUMENTLOSTFOCUS()
On Error GoTo err
    
    If frmALISMSuspense.cboReceiptType > "" Then
        
        Set rsRCPT = New Recordset
    
        strSQL = "SELECT * FROM ALISPReceipt Where ReceiptType = '" & frmALISMSuspense.cboReceiptType & "';"
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
            End If
                    '/* end if the Statement Statement
                    
        End With
        
    End If

Exit Sub

err:
        ErrorMessage

End Sub

Private Sub ReferenceDetails()
On Error GoTo err

        Dim rsREF As ADODB.Recordset, strREF As String
        Set rsREF = New Recordset

        rsREF.Open "SELECT * FROM ODASPAccount WHERE AccountNo LIKE '" & frmALISMSuspense.txtAccountNo & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
                    
        frmALISMSuspense.txtNames = "Names:  " + rsREF!OtherNames + " " + rsREF!CompanyName
        
        If frmALISMSuspense.txtAccountNo.Text <= "" Then
                frmALISMSuspense.txtAccountNo.Text = rsREF!Accountno & ""
        End If
        
        rsREF.Close
        strREF = ""
        
    
        Exit Sub

err:
        ErrorMessage

End Sub

Public Sub cmdFirstCode_Click(Index As Integer)
''''On Error GoTo Myerr

frmALISMSuspense.cmdUpdate.Enabled = False

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

    If frmALISMSuspense.txtReceiptNo.Text = "" Then
        MsgBox "Receipt No is Required"
        frmALISMSuspense.txtReceiptNo.SetFocus
    
    ElseIf frmALISMSuspense.txtReceiptDate.Text = "" Then
        MsgBox "The Receipt Date is  required"
        frmALISMSuspense.txtReceiptDate.SetFocus
    
    ElseIf DateDiff("d", frmALISMSuspense.txtReceiptDate.Text, Date) < 0 Then
        MsgBox "Payment cannot be accepted in advance"
        frmALISMSuspense.txtReceiptDate.SetFocus
    
    ElseIf frmALISMSuspense.cboDocumentNo.Text = "" Then
        MsgBox "The Document Number is Mandatory for this Process"
        frmALISMSuspense.cboDocumentNo.SetFocus
    
    ElseIf frmALISMSuspense.txtTransactionAmount.Text Like 0 Then
        MsgBox "The Receipt Amount is required and There cannot be 0"
        frmALISMSuspense.txtTransactionAmount.SetFocus
    
    ElseIf frmALISMSuspense.cboReceiptType.Text = "" Then
        MsgBox "The Receipt is required"
        frmALISMSuspense.cboReceiptType.SetFocus
    Else
        bsaveRECORD = True

    End If

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub updateRECEIPTDATA()
On Error GoTo err

    Set rsSAVE = New ADODB.Recordset
    
    strSQL = "select * from ALISMReceiptDetails Where ReceiptNo like '" & frmALISMSuspense.txtReceiptNo.Text & "'"
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsSAVE
            If .BOF Or .EOF Then
            .AddNew
                    !ReceiptNo = frmALISMSuspense.txtReceiptNo.Text
                    !TransactionNo = 1
                    !TransactionDate = Date
            End If
            
            !ReceiptType = frmALISMSuspense.cboReceiptType
            !DocumentNo = frmALISMSuspense.cboDocumentNo
            !Accountno = CDbl(frmALISMSuspense.txtAccountNo.Text)
            !TransactionAmount = CDbl(frmALISMSuspense.txtTransactionAmount)
            !AccountingPeriod = frmALISMSuspense.txtAccountingPeriod
            !PaymentStatus = frmALISMSuspense.txtPaymentStatus
            !DateofLastPayment = frmALISMSuspense.txtReceiptDate.Text
            
            If frmALISMSuspense.txtUnitsPaid <= "" Then
                !UnitsPaid = 0
                Else: !UnitsPaid = frmALISMSuspense.txtUnitsPaid
            End If
            
            !AccountCode = frmALISMSuspense.txtAccountNo & ""
            
            If frmALISMSuspense.txtDueDate.Text > "" Then
                !DueDate = frmALISMSuspense.txtDueDate
            End If
            
            !SuspenseAccount = CDbl(frmALISMSuspense.txtSuspenseAccount)
            
            If frmALISMSuspense.txtDateOfLastPayment <= "" Then
                    frmALISMSuspense.txtDateOfLastPayment = frmODASMReceipt.txtReceiptDate.Text
                Else:
                        !DateofLastPayment = frmALISMSuspense.txtDateOfLastPayment
            End If
            
            !ReceivedToDate = CDbl(frmALISMSuspense.txtReceivedTodate) & ""
            

            If frmALISMSuspense.txtDueDate.Text = "" Or frmALISMSuspense.txtDueDate.Text = "0" Then
                    frmALISMSuspense.txtDueDate.Text = Date
                    
                    Else: !DueDate = frmALISMSuspense.txtDueDate & ""
            End If
            
            If frmALISMSuspense.txtStatusCode <= "" Or frmALISMSuspense.txtStatusCode = "10" Then
                    frmALISMSuspense.txtStatusCode = "INFORCE"
            End If
            
            !StatusCode = frmALISMSuspense.txtStatusCode
            
            .Update
            .Requery
            
    End With
rsSAVE.Close

Exit Sub

err:
    UpdateErrorMessage
End Sub

Private Sub updatePROPOSALDATA()

        Dim rsPROP As ADODB.Recordset, strPROP As String
        Set rsPROP = New Recordset
            
                rsPROP.Open "SELECT * FROM ALISMProposal WHERE proposalNo =  '" & frmALISMSuspense.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                With rsPROP
                    !ProposalNo = frmALISMSuspense.cboDocumentNo
                    !DepositReceived = !DepositReceived + CDbl(frmALISMSuspense.txtReceiptAmount)
                    !DateofLastPayment = frmALISMSuspense.txtReceiptDate
                    !StatusCode = "DepRecvd"
                    .Update
                End With
        
        rsPROP.Close
        strPROP = ""
End Sub

Private Sub updateISSUANCE()
        
        Dim rsISSUE As ADODB.Recordset, strISSUE As String
        Set rsISSUE = New Recordset
            
                rsISSUE.Open "SELECT * FROM ODASMJobBrief WHERE JobBriefNo =  '" & frmALISMSuspense.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                With rsISSUE
                    !JobBriefNo = frmALISMSuspense.cboDocumentNo
                    !DepositReceived = 0
                    !IssueDate = Date
                    !DateofLastPayment = frmALISMSuspense.txtReceiptDate
                    !DueDate = frmALISMSuspense.txtDueDate
                    !SuspenseAccount = frmALISMSuspense.txtSuspenseAccount
                    !DateofLastPayment = frmALISMSuspense.txtDateOfLastPayment & ""
                    !ReceivedToDate = frmALISMSuspense.txtReceivedTodate & ""
                    !DueDate = frmALISMSuspense.txtDueDate
            
                    If frmALISMSuspense.txtStatusCode <= "" Then
                        frmALISMSuspense.txtStatusCode = "INFORCE"
                    End If
                    
                    !StatusCode = frmALISMSuspense.txtStatusCode
                    .Update
                End With
        
        rsISSUE.Close
        strISSUE = ""

End Sub
Private Sub updateLOANDATA()
    Dim rsLOAN As ADODB.Recordset, strLOAN As String
    Set rsLOAN = New Recordset
        
    rsLOAN.Open "SELECT * FROM ALISMLoanManagement WHERE LoanNo =  '" & frmALISMSuspense.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
            
            With rsLOAN
            
                !JobBriefNo = frmALISMSuspense.cboDocumentNo
                
                If frmALISMSuspense.txtReceiptAmount > !RepaymentAmount Then
                        !suspense = (frmALISMSuspense.txtReceiptAmount - !RepaymentAmount)
                ElseIf (frmALISMSuspense.txtReceiptAmount <= !RepaymentAmount) And (frmALISMSuspense.txtReceiptAmount > !PrincipalAmount) Then
                        frmALISMSuspense.txtReceiptAmount = 0
                ElseIf (frmALISMSuspense.txtReceiptAmount <= !PrincipalAmount) Then
                        frmALISMSuspense.txtReceiptAmount = 0
                End If
                
                !TotalReceivedTodate = !TotalReceivedTodate + CDbl(frmALISMSuspense.txtTransactionAmount)
                !DateofLastPayment = frmALISMSuspense.txtReceiptDate
                !DueDate = frmALISMSuspense.txtDueDate
                !suspense = frmALISMSuspense.txtSuspenseAccount & ""
                
                !CurrentBalance = !CurrentBalance - rsLOAN!PrincipalAmount
                !InterestReceived = !InterestReceived + (frmALISMSuspense.txtReceiptAmount.Text - !PrincipalAmount)
                !DateofLastPayment = frmALISMSuspense.txtDateOfLastPayment & ""
                !DueDate = frmALISMSuspense.txtDueDate
        
                If frmALISMSuspense.txtStatusCode <= "" Then
                    frmALISMSuspense.txtStatusCode = "INF"
                End If
                
                !Status = frmALISMSuspense.txtStatusCode
                .Update
            End With
    
    rsLOAN.Close
    strLOAN = ""
End Sub

Private Sub saveRecord()
On Error GoTo err

        Set rsSAVE = New ADODB.Recordset
        strSQL = "SELECT * FROM ODASMJobBrief WHERE JobBriefNo =  '" & frmALISMSuspense.cboDocumentNo & "';"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
        With rsSAVE
            If .EOF Or .BOF Then Exit Sub
            
                    !DepositReceived = 0
                    !DateofLastPayment = frmALISMSuspense.txtReceiptDate
                    !SuspenseAccount = frmALISMSuspense.txtSuspenseAccount
                    !ReceivedToDate = frmALISMSuspense.txtReceivedTodate & ""
                    !UnitCountBeforePayment = frmALISMSuspense.txtUnitCountBeforePayment
                    !DueDate = frmALISMSuspense.txtDueDate
                    
                    If frmALISMSuspense.txtStatusCode <= "" Or frmALISMSuspense.txtStatusCode = "10" Then
                        frmALISMSuspense.txtStatusCode = "INFORCE"
                    End If
                    
                    !StatusCode = frmALISMSuspense.txtStatusCode
                    
                    .Update
        End With
        
rsSAVE.Close

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub updateREVIVALDATA()
        Dim rsPOL As ADODB.Recordset, strPOL As String
        Set rsPOL = New Recordset
            
        rsPOL.Open "SELECT * FROM ODASMJobBrief WHERE JobBriefNo =  '" & frmALISMSuspense.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                With rsPOL
                    !JobBriefNo = frmALISMSuspense.cboDocumentNo
                    !DepositReceived = 0
                    !DateofLastPayment = frmALISMSuspense.txtReceiptDate
                    !SuspenseAccount = frmALISMSuspense.txtSuspenseAccount
                    !ReceivedToDate = frmALISMSuspense.txtReceivedTodate & ""
                    !UnitCountBeforePayment = frmALISMSuspense.txtUnitCountBeforePayment
                    !DueDate = frmALISMSuspense.txtDueDate
                    !ReinstatementCharge = frmALISMSuspense.txtTransactionAmount.Text
                    !StatusCode = "INFORCE"
                    !StatusDate = Date
                    .Update
                End With
        
        rsPOL.Close
 
End Sub
Private Sub updateREINSTATEMENTFILE()
        Dim rsREINSTATEMENT As ADODB.Recordset, strPOL As String
        Set rsREINSTATEMENT = New Recordset
            
        rsREINSTATEMENT.Open "SELECT * FROM ALISMReinstatement WHERE JobBriefNo =  '" & frmALISMSuspense.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
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
            
        rsACCOUNT.Open "SELECT * FROM MEGAMAcTransaction  WHERE AccountNo =  '" & frmALISMSuspense.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
        With rsACCOUNT
            .AddNew
            !Accountno = frmALISMSuspense.cboDocumentNo.Text
            !TransactionType = frmALISMSuspense.cboReceiptType.Text
            !TransactionDate = frmALISMSuspense.txtReceiptDate.Text
            !InclusiveAmount = CDbl(frmALISMSuspense.txtTransactionAmount.Text)
            !Amount = CDbl(frmALISMSuspense.txtTransactionAmount.Text)
            !VAT = 0
            !VATRate = 0
            !Balance = !InclusiveAmount
            !Status = "PAID"
            !DueDate = Date
            !ReceiptNo = frmALISMSuspense.txtReceiptNo.Text
            !DateofLastPayment = frmALISMSuspense.txtReceiptDate
            !ReceivedToDate = frmALISMSuspense.txtReceivedTodate
            !dateprepared = Date
            !Preparedby = CurrentUserName
            .Update
            .Requery
        End With
        
        rsACCOUNT.Close

End Sub


Private Sub loadRECORD()
''''On Error GoTo Myerr

        Dim strQRE As Variant
        Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset
            strQRE = InputBox("Enter Transaction No to search.", "Search Value")
            rsFind.Open "SELECT ALISMReceiptDetails.*, ALISPReceipt.* FROM ALISMReceiptdetails, ALISPReceipt WHERE ALISMReceiptDetails.txtReceiptNo = '" & frmALISMSuspense.txtReceiptNo.Text & "' and ALISMReceiptDetails.transactionNO LIKE '" & strQRE & "' and ALISPReceipt.ReceiptType = ALISMReceiptDetails.ReceiptType;", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsFind
            If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
            Else:
                    frmALISMSuspense.txtReceiptNo.Text = !TransactionNo
                    frmALISMSuspense.txtReceiptDate.Text = !TransactionDate
                    frmALISMSuspense.txtTransactionAmount.Text = !TransactionAmount
                    
                    frmALISMSuspense.cboReceiptType.Text = !Description
                    frmALISMSuspense.cboDocumentNo.Text = !DocumentNo
                    frmALISMSuspense.txtAccountNo.Text = !Accountno
                    frmALISMSuspense.txtAccountingPeriod.Text = !AccountingPeriod
                    frmALISMSuspense.txtPaymentStatus.Text = !PaymentStatus
                    frmALISMSuspense.txtAccountNo.Text = !Accountno & ""
                    
                    If IsNull(!DateofLastPayment) = False Then
                        frmALISMSuspense.txtDateOfLastPayment.Text = !DateofLastPayment
                    End If
            
                    frmALISMSuspense.txtReceivedTodate.Text = !ReceivedToDate
                    frmALISMSuspense.txtSuspenseAccount.Text = !SuspenseAccount & ""
                    frmALISMSuspense.txtDueDate.Text = !DueDate & ""
                    frmALISMSuspense.txtStatusCode.Text = !StatusCode
  
                    If rsFind!Jobbrief = "1" Then
                            
                            Dim rsPOL As ADODB.Recordset
                            Set rsPOL = New ADODB.Recordset
                            rsPOL.Open "SELECT ODASMJobBrief.*, ODASPAccount.* FROM ODASMJobBrief, ODASPAccount WHERE JobBriefNo = '" & frmALISMSuspense.cboDocumentNo.Text & "' and ODASPAccount.AccountNo LIKE ODASMJobBrief.AccountNo ; ", cnCOMMON, adOpenKeyset, adLockOptimistic

                            frmALISMSuspense.txtExpectedAmount.Text = rsPOL!ExpectedPremium
                            frmALISMSuspense.txtNames.Text = Trim(rsPOL!OtherNames) + " " + Trim(rsPOL!CompanyName)
                            frmALISMSuspense.txtAccountNo.Text = rsPOL!Accountno
                            
                            rsPOL.Close
                    
                    ElseIf rsFind!Deposit = "1" Then
                            
                            Dim rsPRO As ADODB.Recordset
                            Set rsPRO = New ADODB.Recordset
                            rsPRO.Open "SELECT ALISMProposal.*, ODASPAccount.* FROM ALISMProposal, ODASPAccount WHERE ProposalNo = '" & frmALISMSuspense.cboDocumentNo.Text & "' and ODASPAccount.AccountNo LIKE ALISMProposal.AccountNo ; ", cnCOMMON, adOpenKeyset, adLockOptimistic

                            frmALISMSuspense.txtExpectedAmount.Text = rsPRO!ExpectedPremium
                            frmALISMSuspense.txtNames.Text = Trim(rsPRO!OtherNames) + " " + Trim(rsPRO!CompanyName)
                            frmALISMSuspense.txtAccountNo.Text = rsPRO!Accountno
                            
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
    If CDbl(frmALISMSuspense.txtReceiptAmount) <= 0 Or frmALISMSuspense.txtReceiptAmount <= "" Then
            MsgBox "The Receipt Amount Must be > 0 "
                    frmALISMSuspense.txtReceiptAmount.SetFocus
            Exit Sub
            
    Else:
        Set rsRCPT = New Recordset
        
        rsRCPT.Open "SELECT * FROM ALISPReceipt WHERE ReceiptType= '" & frmALISMSuspense.cboReceiptType.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
      
            
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


