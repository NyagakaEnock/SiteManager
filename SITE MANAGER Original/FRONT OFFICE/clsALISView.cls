VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsALISView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub loadPolicy()
'On Error GoTo err
    
    Dim rsPOL As ADODB.Recordset, strPOL As String
    Set rsPOL = New ADODB.Recordset
    
    
    strSQL = "SELECT * FROM ODASMJobBrief, ODASPAccount where ODASMJobBrief.JobBriefNo  = '" & Screen.ActiveForm.txtJobBriefNo.Text & "' and ODASMJobBrief.AccountNo LIKE ODASPAccount.AccountNo ;"
    
    rsPOL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
    With rsPOL
        If .BOF Or .EOF = True Then Exit Sub
        
        Screen.ActiveForm.txtJobBriefNo = !JobBriefNo
        Screen.ActiveForm.txtLifeAssuredP = Trim(!OtherNames) + " " + Trim(!CompanyName)
        Screen.ActiveForm.txtAccountNoP = !AccountCode
        Screen.ActiveForm.txtAccountNoP = !AccountNo & ""
        Screen.ActiveForm.txtAccountNoP = !AccountNo
        Screen.ActiveForm.txtDateOfCommencementP = !DateofCommencement
        Screen.ActiveForm.txtMaturityDateP = !MaturityDate
        Screen.ActiveForm.txtTermOfPolicyP = !TermOfPolicy
        Screen.ActiveForm.txtPaymentPeriodP = !PaymentPeriod
        Screen.ActiveForm.txtProductCodeP = !ProductCode
        Screen.ActiveForm.txtexpectedpremium = !ExpectedPremium
        Screen.ActiveForm.txtSumAssuredP = !sumassured
        Screen.ActiveForm.txtReceivedTodateP = !ReceivedTodate & ""
        Screen.ActiveForm.txtSuspenseAccountP = !SuspenseAccount
        Screen.ActiveForm.txtPaymentMethodP = !PaymentMethod
        Screen.ActiveForm.txtDateOfLastPaymentP = !DateofLastPayment & ""
        Screen.ActiveForm.txtPremiumCountP = !UnitCount
        Screen.ActiveForm.txtStatusCodeP = !StatusCode
        Screen.ActiveForm.txtSurrenderValueP = !SurrenderValue & ""
        Screen.ActiveForm.txtAccruedBonusP = !accruedBonus & ""
    End With

rsPOL.Close
strPOL = ""

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadLOAN()
'On Error GoTo err
    
    Set rsCONTROL = New Recordset
    
    strSQL = "SELECT * FROM ALISMLoanManagement where ALISMLoanManagement.JobBriefNo = '" & Screen.ActiveForm.txtJobBriefNo & " ';"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    

    With rsCONTROL
            If .EOF Or .BOF Then
                    Exit Sub
            Else
                        Screen.ActiveForm.txtCurrentBalance = !CurrentBalance
                        Screen.ActiveForm.txtDueDate = !DueDate
                        Screen.ActiveForm.txtInterestDue = !InterestDue
                        Screen.ActiveForm.txtInterestRate = !InterestRate
                        Screen.ActiveForm.txtInterestReceived = !InterestReceived
                        Screen.ActiveForm.txtInterestReceivedTodate = !InterestReceivedTodate & ""
                        Screen.ActiveForm.txtLoanAmount = !LoanAmount
                        Screen.ActiveForm.cboLoanNo = !LoanNo
                        Screen.ActiveForm.txtJobBriefNo = !JobBriefNo
                        Screen.ActiveForm.txtPrincipalAmount = !PrincipalAmount
                        Screen.ActiveForm.txtPrincipalReceived = !PrincipalReceived
                        Screen.ActiveForm.txtPrincipalReceivedTodate = !PrincipalReceivedTodate & ""
                        Screen.ActiveForm.txtRepaymentAmount = !RepaymentAmount
                        Screen.ActiveForm.txtRepaymentPeriod = !RepaymentPeriod
                        Screen.ActiveForm.txtStatus = !Status
                        Screen.ActiveForm.txtTotalReceived = !TotalReceived
                        Screen.ActiveForm.txtTotalReceivedTodate = !TotalReceivedTodate & ""
           End If
        
        End With
    
Exit Sub

err:
    ErrorMessage
End Sub
