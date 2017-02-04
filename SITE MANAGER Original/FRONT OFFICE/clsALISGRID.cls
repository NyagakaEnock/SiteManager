VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsALISGRID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub LoadChequeNotScheduled()
On Error GoTo err

    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset
    
    strSQL = ""
    strSQL = "SELECT * FROM ALISMCheque WHERE BankNo =  '" & frmALISMBankAccountCfgPay.cboBankNo.Text & " ' and (scheduled = '' or scheduled is null ) and authorized = 'Y';"
    rsGRID.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Set frmALISMBankAccountCfgPay.DataGrid1.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub LoadChequeScheduled()
On Error GoTo err

    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset
    
    strSQL = ""
    strSQL = "SELECT * FROM ALISMCheque WHERE BankNo =  '" & frmALISMBankAccountCfgPay.cboBankNo.Text & " ' and reference = '" & frmALISMBankAccountCfgPay.txtAccountNo.Text & "';"
    rsGRID.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Set frmALISMBankAccountCfgPay.DataGrid1.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub LoadDUPGRID()
On Error GoTo err

    Dim rsGRID As ADODB.Recordset, StrGRID As String
    Set rsGRID = New Recordset

    StrGRID = "SELECT * FROM ODASMInvoice where JobBriefNo = '" & Screen.ActiveForm.cboJobBriefNo.Text & "' and PaymentCode = '" & Screen.ActiveForm.cboPaymentCode & "';"
    rsGRID.Open StrGRID, cnCOMMON, adOpenKeyset, adLockOptimistic
    Set Screen.ActiveForm.ClaimGrid.DataSource = rsGRID

Exit Sub

err:
    UpdateErrorMessage
End Sub

Public Sub loadPOLICYRIDERGRID()
On Error GoTo err
    
    Dim rsGRID As ADODB.Recordset, StrGRID As String
    Set rsGRID = New ADODB.Recordset

    rsGRID.Open "SELECT * FROM ODASMJobBriefRider WHERE JobBriefNo =  '" & Screen.ActiveForm.cboRiderJobBriefNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    Set Screen.ActiveForm.RiderGRID.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadNWBGRID()
On Error GoTo err
    
    Dim rsGRID As ADODB.Recordset, StrGRID As String
    Set rsGRID = New ADODB.Recordset
    
    StrGRID = "SELECT JobBriefNo, DateofCommencement, IssueDate, ProductCode,SumAssured, MaturityDate, DateOfLastPayment, RenewalDate FROM ODASMJobBrief WHERE month(IssueDate) =  '" & CurrentMonth & "' and Statuscode = 'INFORCE' and (year(RenewalDate) <= '" & UnderwritingYear & "' or Year(RenewalDate) <= '' );"
    rsGRID.Open StrGRID, cnCOMMON, adOpenKeyset, adLockOptimistic
    Set Screen.ActiveForm.OtherReinsuranceGRID.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub LoadPolicyPlanGrid()
On Error GoTo err

        Dim rsGRID As ADODB.Recordset, StrGRID As String
        Set rsGRID = New Recordset
    
        rsGRID.Open "SELECT * FROM ODASMJobBriefPlan WHERE JobBriefNo =  '" & Screen.ActiveForm.cboPlanJobBriefNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        Set Screen.ActiveForm.PlanGrid.DataSource = rsGRID

Exit Sub

err:
ErrorMessage

End Sub

Public Sub loadRETENTIONGRID()
On Error GoTo err
    
    Dim rsGRID As ADODB.Recordset, StrGRID As String
    Set rsGRID = New ADODB.Recordset

    rsGRID.Open "SELECT JobBriefNo, DateofCommencement, MaturityDate, SumAssured, StatusCode FROM ODASMJobBrief WHERE AccountNo LIKE  '" & Screen.ActiveForm.cboAccountNo.Text & "' and StatusCode = 'INFORCE' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    Set Screen.ActiveForm.RetentionGRID.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadSUPPRETENTIONGRID()
On Error GoTo err
    
    Dim rsGRID As ADODB.Recordset, StrGRID As String
    Set rsGRID = New ADODB.Recordset

    rsGRID.Open "SELECT ODASMJobBriefRider.JobBriefNo, ODASMJobBrief.DateofCommencement, ODASMJobBrief.MaturityDate, ODASMJobBriefRider.SumAssured, ODASMJobBrief.StatusCode, ODASMJobBriefRider.RiderCode FROM ODASMJobBrief, ODASMJobBriefRider WHERE ODASMJobBrief.AccountNo LIKE  '" & Screen.ActiveForm.cboAccountNo.Text & "' and ODASMJobBrief.StatusCode = 'INFORCE' and ODASMJobBrief.JobBriefNo = ODASMJobBriefRider.JobBriefNo ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    Set Screen.ActiveForm.SuppRetentionGrid.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadAccountNoGRID()
On Error GoTo err
    
    Dim rsGRID As ADODB.Recordset, StrGRID As String
    Set rsGRID = New ADODB.Recordset

    rsGRID.Open "SELECT * FROM ODASPAccount WHERE AccountNo =  '" & Screen.ActiveForm.txtAccountNo.Text & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    With rsGRID
            If .EOF Or .BOF Then Exit Sub
            
            If Screen.ActiveForm.txtAccountNo >= "" Then
                    MsgBox "Kindly Check the Details of the Policyholder Listed below prior to continuing with Data Entry", vbOKOnly
            End If
    End With

    Set Screen.ActiveForm.ReferenceGRID.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadIDENTITYGRID()
On Error GoTo err
    
    Dim rsGRID As ADODB.Recordset, StrGRID As String
    Set rsGRID = New ADODB.Recordset

    rsGRID.Open "SELECT * FROM ODASPAccount WHERE IdentityNo =  '" & Screen.ActiveForm.txtIdentityNo.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
    With rsGRID
            If .EOF Or .BOF Then Exit Sub
            
            If Screen.ActiveForm.txtIdentityNo.Text > "" Then
                    MsgBox "Kindly Check the Details of the Policyholder Listed below prior to continuing with Data Entry", vbOKOnly
                    Screen.ActiveForm.cmdUpdate.Enabled = False
            End If
    End With

    Set Screen.ActiveForm.ReferenceGRID.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadCompanyNameGRID()
On Error GoTo err
    
    Dim rsGRID As ADODB.Recordset, StrGRID As String
    Set rsGRID = New ADODB.Recordset

    rsGRID.Open "SELECT * FROM ODASPAccount WHERE CompanyName =  '" & Screen.ActiveForm.txtCompanyName.Text & "' or OtherNames = '" & Screen.ActiveForm.txtCompanyName.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    With rsGRID
            If .EOF Or .BOF Then Exit Sub
            MsgBox "Kindly Check the Details of the Policyholder Listed on the First tab prior to continuing with Data Entry", vbOKOnly
    End With
    Set Screen.ActiveForm.SimilarCompanyNameGRID.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadNAMESGRID()
On Error GoTo err
    
    Dim rsGRID As ADODB.Recordset, StrGRID As String
    Set rsGRID = New ADODB.Recordset
    
    StrGRID = "SELECT ODASMJobBrief.JobBriefNo, ODASMJobBrief.ExpectedPremium , ODASPAccount.* FROM ODASMJobBrief, ODASPAccount WHERE ODASPAccount.CompanyName =  '" & Screen.ActiveForm.txtCompanyName.Text & "' and  ODASPAccount.OtherNames = '" & Screen.ActiveForm.txtOtherNames.Text & "' and ODASPAccount.AccountNo LIKE ODASMJobBrief.AccountNo ;"
    rsGRID.Open StrGRID, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsGRID
            If .EOF Or .BOF Then Exit Sub
            MsgBox "Kindly Check the Details of the Policyholder Listed Below prior to continuing with Data Entry", vbOKOnly
    End With
    Set Screen.ActiveForm.ReferenceGRID.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadOTHERNAMESGRID()
On Error GoTo err
    
    Dim rsGRID As ADODB.Recordset, StrGRID As String
    Set rsGRID = New ADODB.Recordset
    
    StrGRID = "SELECT * FROM ODASPAccount WHERE CompanyName =  '" & Screen.ActiveForm.txtCompanyName.Text & "' or OtherNames = '" & Screen.ActiveForm.txtOtherNames.Text & "' OR CompanyName =  '" & Screen.ActiveForm.txtOtherNames.Text & "' OR OtherNames = '" & Screen.ActiveForm.txtCompanyName.Text & "' ;"
    rsGRID.Open StrGRID, cnCOMMON, adOpenKeyset, adLockOptimistic
    With rsGRID
            If .EOF Or .BOF Then Exit Sub
            MsgBox "Kindly Check the Details of the Policyholder Listed second Tab prior to continuing with Data Entry", vbOKOnly
    End With

    Set Screen.ActiveForm.SimilarNamesGRID.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadREFERENCEGRID()
On Error GoTo err
    
    Dim rsGRID As ADODB.Recordset, StrGRID As String
    Set rsGRID = New ADODB.Recordset

    rsGRID.Open "SELECT * FROM ODASMJobBrief WHERE AccountNo LIKE  '" & Screen.ActiveForm.cboAccountNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    Set Screen.ActiveForm.ReferenceGRID.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadGRID()

Dim rsGRID As ADODB.Recordset, StrGRID As String
Set rsGRID = New Recordset

   rsGRID.Open "SELECT * FROM ODASMJobBriefPlan WHERE JobBriefNo =  '" & Screen.ActiveForm.cboPlanJobBriefNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
   Set Screen.ActiveForm.PlanGrid.DataSource = rsGRID

End Sub

Public Sub LoadTakeonGrid()
On Error GoTo err

Dim rsGRID As ADODB.Recordset, StrGRID As String
Set rsGRID = New Recordset

   rsGRID.Open "SELECT * FROM ODASMJobBrief WHERE takeonDate =  '" & Screen.ActiveForm.txtDatePrepared.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
   Set Screen.ActiveForm.TAKEONGRID.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadPendingDischarge()
On Error GoTo err

        Dim rsPENDINGGRID As ADODB.Recordset, strPENDINGGRID As String
        Set rsPENDINGGRID = New Recordset
        
        strPENDINGGRID = "SELECT *  FROM ODASMInvoice where dischargeprepared = 'N';"
        rsPENDINGGRID.Open strPENDINGGRID, cnCOMMON, adOpenKeyset, adLockOptimistic

        Set Screen.ActiveForm.PendingGrid.DataSource = rsPENDINGGRID
        
        
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadProcessedDischarge()
On Error GoTo err

        Dim rsPENDINGGRID As ADODB.Recordset, strPENDINGGRID As String
        Set rsPENDINGGRID = New Recordset
        
        strPENDINGGRID = "SELECT *  FROM ODASMInvoice where dischargeprepared = 'Y';"
        rsPENDINGGRID.Open strPENDINGGRID, cnCOMMON, adOpenKeyset, adLockOptimistic

        Set Screen.ActiveForm.ProcessedGRID.DataSource = rsPENDINGGRID
        
        
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadSimilarDischarge()
On Error GoTo err

        Dim rsPENDINGGRID As ADODB.Recordset, strPENDINGGRID As String
        Set rsPENDINGGRID = New Recordset
        
        strPENDINGGRID = "SELECT *  FROM ODASMInvoice where JobBriefNo = '" & frmALISMClaimDischarge.txtJobBriefNo.Text & "';"
        rsPENDINGGRID.Open strPENDINGGRID, cnCOMMON, adOpenKeyset, adLockOptimistic

        Set Screen.ActiveForm.SimilarGRID.DataSource = rsPENDINGGRID
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadClaimRegistrationGRID()
On Error GoTo err

        Dim rsGRID As ADODB.Recordset, StrGRID As String
        Set rsGRID = New Recordset
        
        StrGRID = "SELECT *  FROM ODASMInvoice where JobBriefNo = '" & Screen.ActiveForm.txtJobBriefNo.Text & "';"
        rsGRID.Open StrGRID, cnCOMMON, adOpenKeyset, adLockOptimistic

        Set Screen.ActiveForm.ClaimRegistrationGRID.DataSource = rsGRID
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadClaimDischargeGRID()
On Error GoTo err

        Dim rsGRID As ADODB.Recordset, StrGRID As String
        Set rsGRID = New Recordset
        
        StrGRID = "SELECT *  FROM ALISMClaimTotal, ODASMInvoice where ODASMInvoice.JobBriefNo = '" & Screen.ActiveForm.txtJobBriefNo.Text & "' and ODASMInvoice.claimNo = ALISMClaimTotal.ClaimNo ;"
        rsGRID.Open StrGRID, cnCOMMON, adOpenKeyset, adLockOptimistic

        Set Screen.ActiveForm.ClaimDischargeGRID.DataSource = rsGRID
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadClaimPaymentGRID()
On Error GoTo err

        Dim rsGRID As ADODB.Recordset, StrGRID As String
        Set rsGRID = New Recordset
        
        StrGRID = "SELECT *  FROM ODASMVoucher, ODASMInvoice where ODASMInvoice.JobBriefNo = '" & Screen.ActiveForm.txtJobBriefNo.Text & "' and ODASMInvoice.claimNo = ODASMVoucher.DocumentNo ;"
        rsGRID.Open StrGRID, cnCOMMON, adOpenKeyset, adLockOptimistic

        Set Screen.ActiveForm.ClaimPaymentGRID.DataSource = rsGRID
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadClaimChequeGRID()
On Error GoTo err

        Dim rsGRID As ADODB.Recordset, StrGRID As String
        Set rsGRID = New Recordset
        
        StrGRID = "SELECT *  FROM ALISMChequeEntry, ODASMInvoice, ODASMVoucher where ODASMInvoice.JobBriefNo = '" & Screen.ActiveForm.txtJobBriefNo.Text & "' and ALISMChequeEntry.RequisitionNo = ODASMVoucher.RequisitionNo and ODASMVoucher.DocumentNo = ODASMInvoice.claimNO;"
        rsGRID.Open StrGRID, cnCOMMON, adOpenKeyset, adLockOptimistic

        Set Screen.ActiveForm.ClaimChequeGRID.DataSource = rsGRID
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub LoadLoanApplicationGrid()
On Error GoTo err
    Dim rsGRID As ADODB.Recordset, StrGRID As String
    Set rsGRID = New Recordset
       
       rsGRID.Open "SELECT * FROM ODASMInvoiceSENT where JobBriefNo = '" & frmALISMLoanEnquiry.txtJobBriefNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
       Set frmALISMLoanEnquiry.LoanApplicationGRID.DataSource = rsGRID

Exit Sub


err:
ErrorMessage

End Sub

Public Sub LoadLoanCheckedGrid()
On Error GoTo err
    Dim rsGRID As ADODB.Recordset, StrGRID As String
    Set rsGRID = New Recordset
       
       rsGRID.Open "SELECT * FROM ODASMInvoiceSENT where JobBriefNo = '" & frmALISMLoanEnquiry.txtJobBriefNo.Text & "' and applicationChecked = 'Y';", cnCOMMON, adOpenKeyset, adLockOptimistic
       Set frmALISMLoanEnquiry.LoanCheckedGRID.DataSource = rsGRID

Exit Sub


err:
ErrorMessage

End Sub

Public Sub LoadLoanApprovedGrid()
On Error GoTo err
    Dim rsGRID As ADODB.Recordset, StrGRID As String
    Set rsGRID = New Recordset
       
       rsGRID.Open "SELECT * FROM ODASMInvoiceSENT where JobBriefNo = '" & Screen.ActiveForm.txtJobBriefNo.Text & "' and ApplicationApproved = 'Y';", cnCOMMON, adOpenKeyset, adLockOptimistic
       Set Screen.ActiveForm.LoanApprovedGRID.DataSource = rsGRID

Exit Sub

err:
ErrorMessage

End Sub

Public Sub LoadLoanAuthorizedGrid()
On Error GoTo err
    Dim rsGRID As ADODB.Recordset, StrGRID As String
    Set rsGRID = New Recordset
       
       rsGRID.Open "SELECT * FROM ODASMInvoiceSENT where JobBriefNo = '" & Screen.ActiveForm.txtJobBriefNo.Text & "' and ApplicationAuthorized = 'Y';", cnCOMMON, adOpenKeyset, adLockOptimistic
       Set Screen.ActiveForm.LoanAuthorizedGRID.DataSource = rsGRID

Exit Sub

err:
ErrorMessage

End Sub
