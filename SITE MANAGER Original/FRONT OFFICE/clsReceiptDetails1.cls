VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReceiptDetails1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private MySName As String, MyOName As String
Dim rsPOLICY As ADODB.Recordset
Dim rsreceipt As ADODB.Recordset, bVAL, BPolicyHolder, bsearchRECORD As Boolean
Dim rsDEP As ADODB.Recordset, strDEP As String


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

        Dim rsEMPLOYERGF As ADODB.Recordset, strEMPLOYERGF As String
        Set rsEMPLOYERGF = New Recordset

        strEMPLOYERGF = "SELECT * FROM ODASPAccount;"
        rsEMPLOYERGF.Open strEMPLOYERGF, cnCOMMON, adOpenKeyset, adLockOptimistic

        frmODASMReceipt.txtAccountNo.Clear

            With rsEMPLOYERGF
                    Do Until .EOF
                             Screen.ActiveForm.cboEmployer.AddItem !CompanyName
                             .MoveNext
                    Loop
            End With
rsEMPLOYERGF.Close
strEMPLOYERGF = ""

Exit Sub

err:
    UpdateErrorMessage
End Sub

Public Sub selectEmployerLostFocus()
On Error GoTo err
        Dim rsEMPLOYERLF As ADODB.Recordset, strEMPLOYERLF As String
        Set rsEMPLOYERLF = New Recordset
        
        rsEMPLOYERLF.Open "SELECT * FROM ODASPAccount WHERE CompanyName= '" & Screen.ActiveForm.cboEmployer.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
 
        With rsEMPLOYERLF
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
rsEMPLOYERLF.Close
strEMPLOYERLF = ""

Exit Sub

err:
    UpdateErrorMessage
End Sub
Public Sub processReceipt()
On Error GoTo err
        With frmODASMReceipt
                If .cboDocumentNo.Text <= Empty Then Exit Sub
                .txtDateOfLastPayment.Text = .txtReceiptDate.Text

                If .txtTransactionAmount.Text <= Empty Or .cboDocumentNo.Text <= Empty Then Exit Sub
        
                If CDbl(.txtTransactionAmount) <= 0 Then
                        MsgBox "The Receipt Amount is Mandatory and is required for all the calculations", vbOKCancel
                        .txtTransactionAmount.SetFocus
                
                ElseIf .cboDocumentNo.Text <= Empty And bEmployerUpdate = False Then
                        MsgBox "This Document No is Required prior to entering the Receipt Amount", vbOKCancel
                        .cboDocumentNo.SetFocus
                
                ElseIf .cboReceiptType.Text <= Empty Then
                        MsgBox "The Receipt Type Is Required Before executive This Function", vbOKOnly
                        .cboReceiptType.SetFocus
                
                ElseIf .txtDateOfLastPayment.Text <= Empty Then
                        MsgBox "The Date of Last Payment is Blank, Testing ..... "
                        .txtDateOfLastPayment.SetFocus
                Else
                
                        If CDbl(.txtTransactionAmount.Text) > 0 Then
                                
                                If .txtReceiptAmount.Text = Empty And baddRECORD = True Then
                                        .txtReceiptAmount.Text = .txtTransactionAmount.Text
                                End If
                                
                                Set rsRCPT = New ADODB.Recordset
                                rsRCPT.Open "SELECT * FROM ALISPReceipt Where receiptType = '" & .cboReceiptType.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                                
                                If rsRCPT.EOF Or rsRCPT.BOF Then Exit Sub
                                
                                If rsRCPT!Jobbrief = "1" Then
                                    Call updateRECORD
                                ElseIf rsRCPT!Deposit = "1" Then
                                    Call updateACCOUNT
                                ElseIf rsRCPT!Account = "1" Then
                                    Call updateACCOUNT
                                ElseIf rsRCPT!Account = "1" Then
                                    Call updateACCOUNT
                                ElseIf rsRCPT!Invoice = "1" Then
                                    Call updateINVOICE
                                End If
                        End If
    
                End If
        End With
Exit Sub

err:
    UpdateErrorMessage
End Sub

Private Sub calcBalance()
            Screen.ActiveForm.txtBalance.Text = CDbl(Screen.ActiveForm.txtReceiptAmountDetails) - CDbl(Screen.ActiveForm.txtTotalAmount.Text)
End Sub

Public Sub updateACCOUNT()
On Error GoTo err
                
    Screen.ActiveForm.txtSuspenseAccount.Text = 0
    Screen.ActiveForm.txtReceivedTodate.Text = 0
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
            .txtTransactionNo.Text = Empty
            .cboReceiptType.Clear
            .cboDocumentNo.Text = Empty
            .txtExpectedAmount.Text = 0
            .txtPaymentStatusDetails.Text = Empty
            .txtDateOfLastPayment.Text = Empty
            .txtReceivedTodate.Text = 0
            .txtNames.Text = Empty
            .txtSuspenseAccount.Text = 0
            .txtDueDate.Text = Empty
            .txtStatusCode.Text = Empty
            .txtTransactionAmount.Text = 0
            .txtReceiptNo.Text = Empty
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub enableDATAENTRY()
On Error GoTo err
        With Screen.ActiveForm
            .cboReceiptType.Locked = False
            .cboDocumentNo.Locked = False
            .txtTransactionAmount.Locked = False
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
            
        rsPOL.Open "SELECT * FROM ODASMJobBrief WHERE JobBriefNo =  '" & frmODASMReceipt.cboDocumentNo & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsPOL
                If rsPOL.EOF Or rsPOL.BOF Then Exit Sub
                
                '/* this procedure is used to intialize the variable
                
                If IsNull(!SuspenseAccount) = True Then
                        frmODASMReceipt.txtSuspenseAccount.Text = 0
                Else: frmODASMReceipt.txtSuspenseAccount = FormatNumber(!SuspenseAccount)
                End If
                
                frmODASMReceipt.txtStatusCode.Text = !Status
                
                If IsNull(!ReceivedToDate) = True Then
                        frmODASMReceipt.txtReceivedTodate = FormatNumber(CDbl(frmODASMReceipt.txtTransactionAmount.Text))
                Else: frmODASMReceipt.txtReceivedTodate = FormatNumber(CDbl(!ReceivedToDate) + CDbl(frmODASMReceipt.txtTransactionAmount))
                End If
                
                If IsNull(!DueDate) = True Then
                        frmODASMReceipt.txtDueDate.Text = Date
                Else: frmODASMReceipt.txtDueDate.Text = !DueDate
                End If
                
       End With

rsPOL.Close
strPOL = ""



Exit Sub


err:
    ErrorMessage
End Sub
Public Sub updateINVOICE()
On Error GoTo err
        
        Dim rsPOL, rsSUM As ADODB.Recordset, strPOL, StrSUM As String
        Set rsPOL = New ADODB.Recordset
            
        rsPOL.Open "SELECT * FROM ODASMINvoiceSent WHERE InvoiceNo =  '" & frmODASMReceipt.cboDocumentNo & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsPOL
                
                If .BOF Or .EOF Then Exit Sub
                
                '/* this procedure is used to intialize the variable
                
                If IsNull(!SuspenseAccount) = True Then frmODASMReceipt.txtSuspenseAccount.Text = 0
                frmODASMReceipt.txtReceivedTodate.Text = !ReceivedToDate & ""
                
                frmODASMReceipt.txtSuspenseAccount = CDbl(frmODASMReceipt.txtSuspenseAccount.Text)

                frmODASMReceipt.txtStatusCode.Text = !Status
                
                '/* Calcululating the Received Todate Amount
                
                Set rsSUM = New ADODB.Recordset
                StrSUM = "SELECT Sum(TransactionAmount) as TotalAmount FROM ALISMReceiptDetails WHERE InvoiceNo =  '" & frmODASMReceipt.cboDocumentNo & "' ;"
                rsSUM.Open StrSUM, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsPOL.EOF And rsPOL.BOF = True Then
                        frmODASMReceipt.txtReceivedTodate = FormatNumber(frmODASMReceipt.txtTransactionAmount.Text)
                ElseIf IsNull(!ReceivedToDate) = True Then
                        frmODASMReceipt.txtReceivedTodate = FormatNumber(frmODASMReceipt.txtTransactionAmount.Text)
                Else: frmODASMReceipt.txtReceivedTodate = FormatNumber(CDbl(rsSUM!TotalAmount))
                End If
                
                If IsNull(!InvoiceDate) = True Then
                        frmODASMReceipt.txtDueDate.Text = Date
                Else: frmODASMReceipt.txtDueDate.Text = !InvoiceDate
                End If
                
       End With

rsPOL.Close
strPOL = ""
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub processDOCUMENTNO()

        If frmODASMReceipt.cboDocumentNo.Text <= Empty Then Exit Sub
        
        If frmODASMReceipt.cboReceiptType.Text <= Empty Then
            MsgBox "The Receipt Type Is Required Before executive This Function", vbOKOnly
            frmODASMReceipt.cboReceiptType.SetFocus
        Else
                If frmODASMReceipt.cboReceiptType > "" Then
                        
                        Set rsRCPT = New ADODB.Recordset
                        rsRCPT.Open "SELECT * FROM ALISPReceipt Where receiptType = '" & frmODASMReceipt.cboReceiptType.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

                        With rsRCPT
                                If .BOF Or .EOF Then Exit Sub
                                If !Deposit = 1 Then
                                    verifyACCOUNTdetails
                                ElseIf !Invoice = 1 Then
                                    verifyINVOICEDETAILS
                                ElseIf !Jobbrief = 1 Then
                                    verifyJOBBRIEFDETAILS
                                ElseIf !Account = 1 Then
                                    verifyACCOUNTdetails
                                End If
                                    '/* end if the Statement Statement
                        
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
        rsREVIVAL.Open "SELECT * FROM ALISMReinstatement Where JobBriefNo = '" & Screen.ActiveForm.cboDocumentNo.Text & "' and Reinstated = 'N' and Authorized = 'Y';", cnCOMMON, adOpenKeyset, adLockOptimistic
        
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

strPAY = Empty
rsPAY.Close

Exit Sub

err:
    ErrorMessage
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

Private Sub verifyINVOICEDETAILS()
On Error GoTo err

        Dim rsCONTROL As ADODB.Recordset, strPOLICY As String
        Set rsCONTROL = New Recordset
        
        
        strPOLICY = "SELECT ODASMInvoiceSENT.*, ODASPAccount.* FROM ODASMInvoiceSent, ODASPAccount WHERE ODASMInvoiceSent.InvoiceNo = '" & frmODASMReceipt.cboDocumentNo.Text & "' and ODASMInvoiceSent.AccountNo = ODASPAccount.AccountNo ;"
        rsCONTROL.Open strPOLICY, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsCONTROL
                If .EOF Or .BOF Then
                        MsgBox "The Job Brief Does not Exist within the Database", vbOKOnly
                        Exit Sub
                End If
                
                frmODASMReceipt.txtExpectedAmount = FormatNumber(!PriceInclusive) & ""
                If Not IsNull(!ReceivedToDate) Then
                    frmODASMReceipt.txtReceivedTodate = FormatNumber(!ReceivedToDate)
                Else
                    frmODASMReceipt.txtReceivedTodate = FormatNumber(!PriceInclusive) & ""
                End If
                frmODASMReceipt.txtStatusCode.Text = !Status & ""
                frmODASMReceipt.txtSuspenseAccount = 0
                frmODASMReceipt.txtDueDate.Text = !InvoiceDate & ""
                frmODASMReceipt.txtAccountNo.Text = !Accountno & ""
                frmODASMReceipt.txtJobBriefNo.Text = !JobBriefNo
                frmODASMReceipt.txtAccountBalance.Text = FormatNumber(!Balance)
                
                If frmODASMReceipt.txtDateOfLastPayment.Text > "" Then
                        frmODASMReceipt.txtDateOfLastPayment.Text = !DateofLastPayment & ""
                ElseIf IsNull(!DateofLastPayment) = True Then
                        frmODASMReceipt.txtDateOfLastPayment.Text = Date

                End If
                
                frmODASMReceipt.txtNames.Text = !CompanyName
                
                If frmODASMReceipt.txtPayer.Text = Empty Then
                    frmODASMReceipt.txtPayer.Text = frmODASMReceipt.txtNames.Text
                End If
                
                frmODASMReceipt.txtTransactionAmount.Text = FormatNumber(!PriceInclusive)
                frmODASMReceipt.txtReceiptAmount.Text = FormatNumber(!PriceInclusive)
        End With
        
        rsCONTROL.Close

Exit Sub
err:
    ErrorMessage
End Sub
Private Sub verifyJOBBRIEFDETAILS()
On Error GoTo err

        Dim rsCONTROL As ADODB.Recordset, strPOLICY As String
        Set rsCONTROL = New Recordset
                         
        rsCONTROL.Open "SELECT ODASMJobBrief.*, ODASPAccount.* FROM ODASMJobBrief, ODASPAccount WHERE ODASMJobBrief.JobBriefNo = '" & frmODASMReceipt.cboDocumentNo.Text & "' and ODASMJobBrief.AccountNo = ODASPAccount.AccountNo ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsCONTROL
                If .EOF Or .BOF Then
                        MsgBox "The Job Brief Does not Exist within the Database", vbOKOnly
                        Exit Sub
                End If
                
                frmODASMReceipt.txtExpectedAmount = FormatNumber(!TotalPrice) & ""
                frmODASMReceipt.txtReceivedTodate = FormatNumber(!ReceivedToDate) & ""
                frmODASMReceipt.txtStatusCode.Text = !Status & ""
                frmODASMReceipt.txtSuspenseAccount = FormatNumber(!SuspenseAccount) & ""
                frmODASMReceipt.txtDueDate.Text = !DueDate & ""
                frmODASMReceipt.txtAccountNo.Text = !Accountno & ""
                frmODASMReceipt.txtAccountBalance.Text = FormatNumber(!Balance)
                frmODASMReceipt.txtJobBriefNo.Text = !JobBriefNo
                
                If frmODASMReceipt.txtDateOfLastPayment.Text > "" Then
                        frmODASMReceipt.txtDateOfLastPayment.Text = !DateofLastPayment & ""
                ElseIf IsNull(!DateofLastPayment) = True Then
                        frmODASMReceipt.txtDateOfLastPayment.Text = Date

                End If
                
                frmODASMReceipt.txtNames.Text = !CompanyName
                
                If frmODASMReceipt.txtPayer.Text = Empty Then
                    frmODASMReceipt.txtPayer.Text = frmODASMReceipt.txtNames.Text
                End If

        
        End With
        
        rsCONTROL.Close

Exit Sub
err:
    ErrorMessage
End Sub


Private Sub verifyACCOUNTdetails1()
        
        Set rsCONTROL = New Recordset
                         
        rsCONTROL.Open "SELECT ODASPAccount.* FROM ODASPAccount WHERE AccountNo = '" & Screen.ActiveForm.cboDocumentNo.Text & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsCONTROL
                If .EOF Or .BOF Then
                        MsgBox "The Employer Does not Exist within the Database", vbOKOnly
                        Exit Sub
                End If
                
                Screen.ActiveForm.txtExpectedAmount = 0
                Screen.ActiveForm.txtReceivedTodate = 0
                Screen.ActiveForm.txtStatusCode.Text = !StatusCode & ""
                Screen.ActiveForm.txtUnitCountBeforePayment = 0
                Screen.ActiveForm.txtSuspenseAccount = 0
                Screen.ActiveForm.txtDueDate.Text = Date
                
                If Screen.ActiveForm.txtDateOfLastPayment.Text > "" Then
                        Screen.ActiveForm.txtDateOfLastPayment.Text = !DateofLastPayment
                End If
                
                Screen.ActiveForm.txtNames.Text = !CompanyName
                
                If Screen.ActiveForm.txtPayer.Text = Empty Then
                    Screen.ActiveForm.txtPayer.Text = Screen.ActiveForm.txtNames.Text
                End If

        
        End With
    rsCONTROL.Close

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
                Screen.ActiveForm.txtDueDate.Text = !DueDate & ""
                
                
                If Screen.ActiveForm.txtDateOfLastPayment.Text > "" Then
                        Screen.ActiveForm.txtDateOfLastPayment.Text = !DateofLastPayment
                End If
                
                Screen.ActiveForm.txtNames.Text = !OtherNames + " " + !CompanyName
                
                If Screen.ActiveForm.txtPayer.Text = Empty Then
                    Screen.ActiveForm.txtPayer.Text = Screen.ActiveForm.txtNames.Text
                End If

        
        End With
    rsPOLICY.Close

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub verifyACCOUNTdetails()
        
        Set rsCONTROL = New Recordset
                         
        rsCONTROL.Open "SELECT * FROM ODASPAccount WHERE AccountNo = '" & frmODASMReceipt.cboDocumentNo.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsCONTROL
                If .EOF Or .BOF Then
                        MsgBox "The Account Entered Does not Exist within the Database", vbOKOnly
                        Exit Sub
                End If
                
                frmODASMReceipt.txtExpectedAmount = 0
                frmODASMReceipt.txtReceivedTodate = 0
                frmODASMReceipt.txtStatusCode.Text = 0
                frmODASMReceipt.txtSuspenseAccount = 0
                frmODASMReceipt.txtAccountNo.Text = !Accountno
                frmODASMReceipt.txtAccountingPeriod = CurrentPeriod
                frmODASMReceipt.txtNames.Text = !CompanyName
                frmODASMReceipt.txtJobBriefNo.Text = "XX"
                bloadINVOICE = True
        End With
    rsCONTROL.Close

End Sub
Public Sub createRECEIPT()
On Error GoTo err
        Dim rsTRANS As ADODB.Recordset, strTRANS As String
        Set rsTRANS = New Recordset
        
        rsTRANS.Open "SELECT * FROM ALISMReceiptNew WHERE ReceiptNo = '" & Screen.ActiveForm.txtReceiptNo.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
                
        With rsTRANS
            If .BOF Or .EOF = True Then Exit Sub
            
            
            Screen.ActiveForm.txtTransactionNo = !TransactionNo & ""
                    
            With Screen.ActiveForm
                    If .txtTransactionNo <= Empty Then .txtTransactionNo = 0
                    .txtTransactionNo.Text = CDbl(.txtTransactionNo.Text) + 1
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
            
    End If
End Sub

Public Sub selectRECEIPTTYPELOSTFOCUS()
On Error GoTo err
    Dim rsRCPT As ADODB.Recordset
    Set rsRCPT = New Recordset
    rsRCPT.Open "SELECT * FROM ALISPReceipt WHERE Description= '" & Screen.ActiveForm.cboReceiptType.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsRCPT
            If .EOF And .BOF Then Exit Sub
            
            Screen.ActiveForm.cboReceiptType.Text = !ReceiptType
            Screen.ActiveForm.txtPaymentStatusDetails.Text = "PAID"
            Screen.ActiveForm.txtPaymentStatus.Text = "PAID"

            If !Account = "1" Then
                    showALLCUSTOMERS
                    bEmployerUpdate = True
                    Screen.ActiveForm.lblDocumentNo.Caption = "AccountNo"
            
            ElseIf !Jobbrief = "1" Then
                    ShowALLJOBBRIEFS
                    Screen.ActiveForm.lblDocumentNo.Caption = "JBrief No"
            
            ElseIf !Invoice = "1" Then
                    showINVOICES
                    Screen.ActiveForm.lblDocumentNo.Caption = "Invoice No"

            ElseIf !Deposit = "1" Then
                    showALLCUSTOMERS
                    Screen.ActiveForm.lblDocumentNo.Caption = "Account No"
            
            ElseIf !Miscellaneous = "1" Then
                    Screen.ActiveForm.txtPayer.Locked = False
                    Screen.ActiveForm.lblDocumentNo.Caption = "Account No"

            Else
                    ShowALLJOBBRIEFS
                    Screen.ActiveForm.lblDocumentNo.Caption = "JBrief No"
            End If
            
    End With
    
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
        Screen.ActiveForm.txtAccountingPeriod = CurrentPeriod
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
    enableDATAENTRY
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
                Screen.ActiveForm.txtAccountNo.Text = rsREF!Accountno & ""
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

    If Screen.ActiveForm.txtReceiptNo.Text = Empty Then
        MsgBox "Receipt No is Required"
        Screen.ActiveForm.txtReceiptNo.SetFocus
    
    ElseIf Screen.ActiveForm.txtTransactionNo.Text = "" Then
        MsgBox "Transaction No is Required"
        Screen.ActiveForm.txtTransactionNo.SetFocus

    ElseIf Screen.ActiveForm.txtReceiptDate.Text Like 0 Then
        MsgBox "The Receipt Date is  required"
        Screen.ActiveForm.txtReceiptDate.SetFocus
    
    ElseIf DateDiff("d", Screen.ActiveForm.txtReceiptDate.Text, Date) < 0 Then
        MsgBox "Payment cannot be accepted in advance"
        Screen.ActiveForm.txtReceiptDate.SetFocus
    

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
        updateACCOUNTDATA
    
    ElseIf rsRCPT!Invoice = 1 Then
        updateJBBALANCE
        updateINVOICEBALANCE
        updateInvoiceDATA

    ElseIf rsRCPT!Jobbrief = 1 Then
        updateJBBALANCE
        updateINVOICEBALANCE
        updateJOBBRIEFDATA

    ElseIf rsRCPT!Account = 1 Then
        updateACCOUNTDATA
    ElseIf rsRCPT!Account = 1 Then
        updateACCOUNTDATA
        
    End If


Exit Sub

err:
        ErrorMessage
End Sub
Private Sub updateRECEIPTDATA()
On Error GoTo err

    Set rsRCPT = New Recordset
    rsRCPT.Open "select * from ALISPReceipt where ReceiptType = '" & frmODASMReceipt.cboReceiptType & "' ; ", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    If rsRCPT.EOF Or rsRCPT.BOF Then Exit Sub
    
    Set rsSAVE = New Recordset
    rsSAVE.Open "select * from ALISMReceiptDetails ; ", cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsSAVE
                    .AddNew
                    !ReceiptNo = frmODASMReceipt.txtReceiptNo
                    !TransactionNo = frmODASMReceipt.txtTransactionNo
                    !dateprepared = Date
                    !Preparedby = CurrentUserName
                    !Banked = "N"
                    !AmountDue = 0
            
            !TransactionDate = frmODASMReceipt.txtReceiptDate
            !ReceiptType = frmODASMReceipt.cboReceiptType
            !DocumentNo = frmODASMReceipt.cboDocumentNo
            !TransactionAmount = CDbl(frmODASMReceipt.txtTransactionAmount)
            !AccountingPeriod = frmODASMReceipt.txtAccountingPeriod
            !PaymentStatus = frmODASMReceipt.txtPaymentStatus
            !DateofLastPayment = frmODASMReceipt.txtReceiptDate.Text
            !Accountno = frmODASMReceipt.txtAccountNo.Text
            
            If rsRCPT!Jobbrief = "1" Then
                    !JobBriefNo = frmODASMReceipt.txtJobBriefNo.Text
                    !InvoiceNo = "XX"
            ElseIf rsRCPT!Invoice = "1" Then
                    !JobBriefNo = frmODASMReceipt.txtJobBriefNo.Text
                    !InvoiceNo = frmODASMReceipt.cboDocumentNo.Text
            Else
                    !InvoiceNo = "XX"
                    !JobBriefNo = "XX"
            End If
            
            If frmODASMReceipt.txtDueDate.Text > Empty Then
                !DueDate = frmODASMReceipt.txtDueDate
            End If
            
            !SuspenseAccount = 0
            
            If frmODASMReceipt.txtDateOfLastPayment <= Empty Then
                    frmODASMReceipt.txtDateOfLastPayment = frmODASMReceipt.txtReceiptDate.Text
                Else:
                    !DateofLastPayment = frmODASMReceipt.txtDateOfLastPayment
            End If
            
            !ReceivedToDate = CDbl(frmODASMReceipt.txtReceivedTodate) & Empty
            
            If frmODASMReceipt.txtStatusCode <= Empty Or frmODASMReceipt.txtStatusCode = "10" Then
                    frmODASMReceipt.txtStatusCode = "OPEN"
            End If
            
            !StatusCode = frmODASMReceipt.txtStatusCode
            !TransactionDate = Date
            
            .Update
    End With
rsSAVE.Close

Exit Sub

err:
    UpdateErrorMessage
End Sub
Private Sub updateJBBALANCE()

On Error GoTo err

        '/ Check whether the Payment Type is Invoice
        
        
        Dim strDATA, strCALC As String
        
        Set rsCONTROL = New ADODB.Recordset
        
        strCONTROL = "SELECT sum(ALISMReceiptDetails.TransactionAmount) as totals from ALISMReceiptDetails where JobBriefNo = '" & frmODASMReceipt.txtJobBriefNo.Text & "' ;"
        rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF Or rsCONTROL.BOF Then
                frmODASMReceipt.txtJobBriefTotal.Text = 0
        ElseIf IsNull(rsCONTROL!TOTALS) = True Then
                frmODASMReceipt.txtJobBriefTotal.Text = 0
        Else
                frmODASMReceipt.txtJobBriefTotal.Text = FormatNumber(rsCONTROL!TOTALS)
        End If
                        
                    
        Set rsSAVE = New Recordset
        
        strSQL = "SELECT * FROM ODASMJobBrief Where JobBriefNo = '" & frmODASMReceipt.txtJobBriefNo.Text & "';"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsSAVE
                If .BOF Or .EOF = True Then Exit Sub
                
                !ReceivedToDate = frmODASMReceipt.txtJobBriefTotal.Text
                
                If Not IsNull(!TotalOverallCost) Then
                    !Balance = CDbl(!TotalOverallCost) - CDbl(frmODASMReceipt.txtJobBriefTotal.Text)
                ElseIf Not IsNull(!TotalPrice) Then
                    !Balance = CDbl(!TotalPrice) - CDbl(frmODASMReceipt.txtJobBriefTotal.Text)
                End If
                If CDbl(!Balance) < 0 Then
                        !SuspenseAccount = -1 * CDbl(!Balance)
                        !Balance = 0
                End If
                
                frmODASMReceipt.txtAccountBalance.Text = FormatNumber(!Balance)
                .Update
                .Requery
        End With

rsCONTROL.Close
rsSAVE.Close

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub updateINVOICEBALANCE()

On Error GoTo err
        
        
        '/ Check whether the Payment Type is Invoice

        Dim strDATA, strCALC As String
        
        Set rsCONTROL = New ADODB.Recordset
        
        strCONTROL = "SELECT sum(ALISMReceiptDetails.TransactionAmount) as totals from ALISMReceiptDetails where InvoiceNo = '" & frmODASMReceipt.cboDocumentNo.Text & "' ;"
        rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF Or rsCONTROL.BOF Then
                frmODASMReceipt.txtInvoiceTotal.Text = 0
        ElseIf IsNull(rsCONTROL!TOTALS) = True Then
                frmODASMReceipt.txtInvoiceTotal.Text = 0
        Else
                frmODASMReceipt.txtInvoiceTotal.Text = FormatNumber(rsCONTROL!TOTALS)
        End If
                    
        Set rsSAVE = New Recordset
        
        strSQL = "SELECT * FROM ODASMInvoiceSent Where InvoiceNo = '" & frmODASMReceipt.cboDocumentNo.Text & "';"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsSAVE
                If .BOF Or .EOF = True Then Exit Sub
                !ReceivedToDate = frmODASMReceipt.txtInvoiceTotal.Text
                
                !Balance = CDbl(!PriceInclusive) - CDbl(frmODASMReceipt.txtJobBriefTotal.Text)
                
                If CDbl(!Balance) < 0 Then
                        !SuspenseAccount = -1 * CDbl(!Balance)
                        !Balance = 0
                End If
                
                frmODASMReceipt.txtAccountBalance.Text = FormatNumber(!Balance)
                frmODASMReceipt.txtSuspenseAccount.Text = FormatNumber(!SuspenseAccount)
                .Update
                .Requery
        End With

rsCONTROL.Close
rsSAVE.Close

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub updateRECEIPTBALANCE()
On Error GoTo err

        Dim strDATA, strCALC As String
        
        Set rsCONTROL = New ADODB.Recordset
        
        strCONTROL = "SELECT sum(ALISMReceiptDetails.TransactionAmount) as totals from ALISMReceiptDetails where ReceiptNo = '" & Screen.ActiveForm.txtReceiptNo.Text & "' ;"
        rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF Or rsCONTROL.BOF Then
                Screen.ActiveForm.txtTotalAmount.Text = 0
                Screen.ActiveForm.txtReceiptAmount.Text
                
        ElseIf IsNull(rsCONTROL!TOTALS) = True Then
                Screen.ActiveForm.txtTotalAmount.Text = 0
                Screen.ActiveForm.txtReceiptAmount.Text = 0
        Else
                Screen.ActiveForm.txtTotalAmount.Text = FormatNumber(rsCONTROL!TOTALS)
                Screen.ActiveForm.txtReceiptAmount.Text = FormatNumber(rsCONTROL!TOTALS)
        End If
        

        Set rsSAVE = New Recordset
        
        strSQL = "SELECT * FROM ALISMReceiptNew Where ReceiptNo = '" & Screen.ActiveForm.txtReceiptNo.Text & "';"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsSAVE
                If .BOF Or .EOF = True Then Exit Sub
                !TotalAmount = Screen.ActiveForm.txtTotalAmount.Text
                !ReceiptAmount = Screen.ActiveForm.txtTotalAmount.Text
                !TransactionNo = Screen.ActiveForm.txtTransactionNo
                .Update
                .Requery
        End With

rsCONTROL.Close
rsSAVE.Close

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub updateJOBBRIEFDATA()
On Error GoTo err
        
        
                Set rsSAVE = New Recordset
            
                rsSAVE.Open "SELECT * FROM ODASMJobBrief WHERE JobBriefNo =  '" & Screen.ActiveForm.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                With rsSAVE
                    If .BOF Or .EOF Then Exit Sub
                    
                    !DateofLastPayment = Screen.ActiveForm.txtReceiptDate
                    !DueDate = Screen.ActiveForm.txtDueDate
                    
                    If !Balance <= 0 Then
                            !Paid = "Y"
                    ElseIf CDbl(!Balance) > 0 And CDbl(!Balance) < CDbl(!TotalPrice) Then
                            !Paid = "P"
                    Else
                            !Paid = "N"
                    End If

                    .Update
                    .Requery
                End With
        
        rsSAVE.Close
Exit Sub

err:
    UpdateErrorMessage
End Sub
Private Sub totalJOBBRIEFDATA()
On Error GoTo err
        
        
                Set rsSAVE = New Recordset
            
                rsSAVE.Open "SELECT * FROM ODASMJobBrief WHERE JobBriefNo =  '" & Screen.ActiveForm.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                With rsSAVE
                    If .BOF Or .EOF Then Exit Sub
                    
                    !JobBriefNo = Screen.ActiveForm.cboDocumentNo
                    !DateofLastPayment = Screen.ActiveForm.txtReceiptDate
                    !ReceivedToDate = CDbl(Screen.ActiveForm.txtReceivedTodate) & ""
                    !SuspenseAccount = CDbl(Screen.ActiveForm.txtSuspenseAccount)
                    
                    If IsNull(!Balance) = True Then
                            !Balance = 0
                    Else
                            !Balance = CDbl(!Balance) - CDbl(Screen.ActiveForm.txtReceivedTodate)
                    End If
                    
                    !DueDate = Screen.ActiveForm.txtDueDate
                    .Update
                    .Requery
                End With
        
        rsSAVE.Close
Exit Sub

err:
    UpdateErrorMessage
End Sub

Private Sub updateInvoiceDATA()
On Error GoTo err
        
        
                Set rsSAVE = New Recordset
            
                rsSAVE.Open "SELECT * FROM ODASMInvoiceSent WHERE InvoiceNo =  '" & Screen.ActiveForm.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                With rsSAVE
                    If .BOF Or .EOF Then Exit Sub
                    
                    !DateofLastPayment = Screen.ActiveForm.txtReceiptDate
                    !Balance = Screen.ActiveForm.txtAccountBalance
                    
                    If !Balance <= 0 Then
                            !Paid = "Y"
                    ElseIf CDbl(!Balance) > 0 And CDbl(!Balance) < CDbl(!PriceInclusive) Then
                            !Paid = "P"
                    ElseIf !Balance <= frmODASMReceipt.txtAccountBalance.Text Then
                            !Paid = "N"
                    Else
                            !Paid = "N"
                    End If
                    
                    .Update
                    .Requery
                End With
        
        rsSAVE.Close
Exit Sub

err:
    UpdateErrorMessage
End Sub


Private Sub updateACCOUNTDATA()
On Error GoTo err
        
        
                Set rsSAVE = New Recordset
            
                rsSAVE.Open "SELECT * FROM ODASMAccount WHERE AccountNo =  '" & Screen.ActiveForm.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                With rsSAVE
                    If .BOF Or .EOF Then
                            .AddNew
                            !Accountno = frmODASMReceipt.txtAccountNo
                            !Preparedby = CurrentUserName
                            !dateprepared = Date
                    End If
                    
                    !TransactionDate = frmODASMReceipt.txtReceiptDate.Text
                    !DateofLastPayment = Screen.ActiveForm.txtReceiptDate
                    !ReceivedToDate = Screen.ActiveForm.txtReceivedTodate & ""
                    !SuspenseAccount = Screen.ActiveForm.txtSuspenseAccount
                    !ReceiptNo = frmODASMReceipt.txtReceiptNo.Text
                    .Update
                    .Requery
                End With
        
        rsSAVE.Close
Exit Sub

err:
    UpdateErrorMessage
End Sub



Public Sub searchRECORD()
    locateRECORD
End Sub

Private Sub locateRECORD()
'''''On Error GoTo Myerr

        Dim strQRE As Variant
        Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset
            strSQL = "SELECT ALISMReceiptDetails.*, ALISPReceipt.* FROM ALISMReceiptdetails, ALISPReceipt WHERE ALISMReceiptDetails.ReceiptNo = '" & Screen.ActiveForm.txtReceiptNo.Text & "' and ALISMReceiptDetails.transactionNO LIKE '" & Screen.ActiveForm.txtTransactionNo.Text & "';"
            rsFind.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsFind
            If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
            Else:
                    Screen.ActiveForm.txtTransactionNo.Text = !TransactionNo
                    Screen.ActiveForm.txtTransactionAmount.Text = !TransactionAmount
                    
                    Screen.ActiveForm.cboReceiptType.Text = !Description
                    Screen.ActiveForm.cboDocumentNo.Text = !DocumentNo
                    Screen.ActiveForm.txtAccountingPeriod.Text = !AccountingPeriod
                    Screen.ActiveForm.txtPaymentStatusDetails.Text = !PaymentStatus
                     
                    If IsNull(!DateofLastPayment) = False Then
                        Screen.ActiveForm.txtDateOfLastPayment.Text = !DateofLastPayment
                    End If
            
                    Screen.ActiveForm.txtReceivedTodate.Text = FormatNumber(!ReceivedToDate)
                    Screen.ActiveForm.txtSuspenseAccount.Text = !SuspenseAccount & ""
                    Screen.ActiveForm.txtDueDate.Text = !DueDate & ""
                    Screen.ActiveForm.txtStatusCode.Text = !StatusCode
  
                    If rsFind!Jobbrief = "1" Then
                            
                            Dim rsPOL As ADODB.Recordset
                            Set rsPOL = New ADODB.Recordset
                            rsPOL.Open "SELECT ODASMJobBrief.*, ODASPAccount.* FROM ODASMJobBrief, ODASPAccount WHERE JobBriefNo = '" & Screen.ActiveForm.cboDocumentNo.Text & "' and ODASPAccount.AccountNo LIKE ODASMJobBrief.AccountNo ; ", cnCOMMON, adOpenKeyset, adLockOptimistic

                            Screen.ActiveForm.txtExpectedAmount.Text = FormatNumber(rsPOL!ExpectedPremium)
                            Screen.ActiveForm.txtNames.Text = Trim(rsPOL!CompanyName)
                            Screen.ActiveForm.txtAccountNo.Text = rsPOL!Accountno
                            
                            rsPOL.Close
                    
                    ElseIf rsFind!Deposit = "1" Then
                            
                            Dim rsPRO As ADODB.Recordset
                            Set rsPRO = New ADODB.Recordset
                            rsPRO.Open "SELECT M.Balance, P.CompanyName, P.AccountNo FROM ODASPAccount P, ODASMAccount M WHERE P.AccountNo = M.AccountNo and M.AccountNo = '" & Screen.ActiveForm.cboDocumentNo.Text & "' ; ", cnCOMMON, adOpenKeyset, adLockOptimistic
                            
                            With frmODASMReceipt
                                    '.txtExpectedAmount.Text = rsPRO!Balance
                                    '.txtNames.Text = Trim(rsPRO!CompanyName)
                                    '.txtAccountNo.Text = rsPRO!AccountNo
                            End With
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
            ElseIf rsRCPT!accounts = 1 Then
                Call updateRECORD
            End If
    End If
Exit Sub

err:
    ErrorMessage
End Sub


