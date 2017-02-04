VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReceipting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim rsreceipt As ADODB.Recordset, bVAL As Boolean
Dim rsDEP As ADODB.Recordset, strDEP As String, BPolicyHolder As Boolean
Private MySName As String, MyOName As String

Public Sub loadEMPLOYER()
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

Public Sub loadRECEIPTDETAILS()
On Error GoTo err

        Dim rsDETAILS As ADODB.Recordset

        Set rsDETAILS = New ADODB.Recordset
        rsDETAILS.Open "SELECT ALISMReceiptDetails.*, ALISPReceipt.* FROM ALISMReceiptdetails, ALISPReceipt WHERE ALISMReceiptDetails.ReceiptNo = '" & Screen.ActiveForm.txtReceiptNo.Text & "' and ALISMReceiptDetails.transactionNO LIKE '1' and ALISPReceipt.ReceiptType = ALISMReceiptDetails.ReceiptType;", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsDETAILS
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

                    If rsDETAILS!Jobbrief = "1" Then
                            
                            Dim rsPOL As ADODB.Recordset
                            Set rsPOL = New ADODB.Recordset
                            rsPOL.Open "SELECT ODASMJobBrief.*, ODASPAccount.* FROM ODASMJobBrief, ODASPAccount WHERE JobBriefNo = '" & Screen.ActiveForm.cboDocumentNo.Text & "' and ODASPAccount.AccountNo LIKE ODASMJobBrief.AccountNo ; ", cnCOMMON, adOpenKeyset, adLockOptimistic

                            Screen.ActiveForm.txtExpectedAmount.Text = rsPOL!ExpectedPremium
                            Screen.ActiveForm.txtNames.Text = Trim(rsPOL!OtherNames) + " " + Trim(rsPOL!CompanyName)
                            Screen.ActiveForm.txtAccountNo.Text = rsPOL!AccountNo
                            
                            rsPOL.Close
                    
                    ElseIf rsDETAILS!Deposit = "1" Then
                            
                            Dim rsPRO As ADODB.Recordset
                            Set rsPRO = New ADODB.Recordset
                            rsPRO.Open "SELECT ALISMProposal.*, ODASPAccount.* FROM ALISMProposal, ODASPAccount WHERE ProposalNo = '" & Screen.ActiveForm.cboDocumentNo.Text & "' and ODASPAccount.AccountNo LIKE ALISMProposal.AccountNo ; ", cnCOMMON, adOpenKeyset, adLockOptimistic

                            Screen.ActiveForm.txtExpectedAmount.Text = rsPRO!ExpectedPremium
                            Screen.ActiveForm.txtNames.Text = Trim(rsPRO!OtherNames) + " " + Trim(rsPRO!CompanyName)
                            Screen.ActiveForm.txtAccountNo.Text = rsPRO!AccountNo
                            
                    rsPRO.Close

                    End If
                    
            End If
        
        End With

Exit Sub

err:
    ErrorMessage
End Sub


Public Sub clearRECORD()
On Error GoTo err:

        With Screen.ActiveForm
            .txtReceiptNo.Text = ""
            .txtReceiptDate.Text = ""
            .txtReceiptAmount.Text = ""
            .txtAccountingPeriod.Text = ""
            .txtPaymentStatus.Text = ""
            .cboPaymentMethod.Text = ""
            .txtChequeNo.Text = ""
            .cboBankNo.Text = ""
            .txtLocal.Text = ""
            .txtBankNo.Text = ""
            .txtPaymentMethod.Text = ""
            .txtPayer.Text = ""
            .cboCurrencyCode.Text = ""
            Set .ReceiptMGRID.DataSource = Nothing
            Set .receiptDetailsGRID.DataSource = Nothing
            Set .policiesGRID.DataSource = Nothing
            Set .proposalGRID.DataSource = Nothing
            Set .loanGRID.DataSource = Nothing
            
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
            .txtReceiptNo.Locked = bVAL
            .txtReceiptDate.Locked = bVAL
            .txtReceiptAmount.Locked = bVAL
            .txtAccountingPeriod.Locked = True
            .txtPaymentStatus.Locked = True
            .cboPaymentMethod.Locked = bVAL
            .txtChequeNo.Locked = bVAL
            .cboBankNo.Locked = bVAL
            .txtLocal.Locked = True
            .txtPayer.Locked = bVAL
            .txtBankNo.Locked = True
            .txtPaymentMethod.Locked = True
            .cboCurrencyCode.Locked = bVAL
        End With
    
Exit Sub

err:
    ErrorMessage
End Sub



Public Sub loadRECORD()
On Error GoTo err
    With rsDEP
            Screen.ActiveForm.txtReceiptNo.Text = !ReceiptNo
            Screen.ActiveForm.txtReceiptDate.Text = !ReceiptDate
            Screen.ActiveForm.txtReceiptAmount.Text = !ReceiptAmount
            Screen.ActiveForm.txtAccountingPeriod.Text = !AccountingPeriod
            Screen.ActiveForm.txtPaymentStatus.Text = !PaymentStatus
            Screen.ActiveForm.txtStatusCode.Text = !StatusCode
            Screen.ActiveForm.txtPaymentMethod.Text = !PaymentMethod & ""
            Screen.ActiveForm.txtChequeNo.Text = !ChequeNo & ""
            Screen.ActiveForm.txtBankNo.Text = !BankNo & ""
            Screen.ActiveForm.txtLocal.Text = !LocalCheque & ""
    End With
    
    
rsDEP.Close
strDEP = ""

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub lockRECEIPTNO()
On Error GoTo err
        Dim rsLAST As ADODB.Recordset, strLAST As String
        
        Set rsLAST = New Recordset
      
        strLAST = "SELECT * FROM ALISPLastNumber;"
        rsLAST.Open strLAST, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsLAST
            If .BOF Or .EOF Then Exit Sub
                If !AutoReceiptNo = "Yes" Then
                        Screen.ActiveForm.txtReceiptNo.Locked = True
                        Screen.ActiveForm.txtReceiptNo.BackColor = &HFFFFC0
                End If
        End With

Exit Sub

err:
    UpdateErrorMessage
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
                        Screen.ActiveForm.txtPaymentMethod.Text = !PaymentMethod
                        Screen.ActiveForm.cboPaymentMethod.Text = !PaymentMethodDescription

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

Public Sub selectBankNOGotFocus()
On Error GoTo err

        Dim rsECGF As ADODB.Recordset, strECGF As String
        Set rsECGF = New Recordset
      
        strECGF = "SELECT * FROM ALISPBank;"
        rsECGF.Open strECGF, cnCOMMON, adOpenKeyset, adLockOptimistic

        Screen.ActiveForm.cboBankNo.Clear

            With rsECGF
                    Do Until .EOF
                            Screen.ActiveForm.cboBankNo.AddItem !CompanyName
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

        Dim rsBANKLF As ADODB.Recordset, strBANKLF As String
        Set rsBANKLF = New Recordset
        
        rsBANKLF.Open "SELECT * FROM ALISPBank WHERE CompanyName = '" & Screen.ActiveForm.cboBankNo.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsBANKLF
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboBankNo.Text = !CompanyName
                Screen.ActiveForm.txtBankNo.Text = !BankNo
                Screen.ActiveForm.txtLocal = !LocalBank
        End With
        
        Screen.ActiveForm.txtChequeNo.SetFocus
  
rsBANKLF.Close
 
Exit Sub

err:
    ErrorMessage

End Sub

Public Sub selectCURRENCYGOTFOCUS()
    
    Dim rsPAY As ADODB.Recordset, strPAY As String
    Set rsPAY = New Recordset
    
    strPAY = "SELECT * FROM ALISPCurrency;"
    rsPAY.Open strPAY, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Screen.ActiveForm.cboCurrencyCode.Clear

    With rsPAY
            Do Until .EOF
                    Screen.ActiveForm.cboCurrencyCode.AddItem !CurrencyCode
                    .MoveNext
            Loop
    
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub selectCURRENCYKEYPRESS(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub

Public Sub selectCURRENCYLOSTFOCUS()
On Error GoTo err

        Dim rsCOSTCENTER As ADODB.Recordset, strCOSTCENTER As String
        Set rsCOSTCENTER = New Recordset
        
        rsCOSTCENTER.Open "SELECT * FROM ALISPCurrency WHERE CurrencyDescription= '" & Screen.ActiveForm.cboCurrencyCode.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCOSTCENTER
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboCurrencyCode.Text = !CurrencyCode
        End With
        
        Exit Sub

err:
        ErrorMessage

End Sub
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

Public Sub addRECORD()
'''On Error GoTo Myerr
    addpen = True
    clearRECORD
    enableDATAENTRY
    DisableCBRECEIPT
    initRECORD
Exit Sub

Myerr:
    ErrorMessage
End Sub
Private Sub initRECORD()
        Screen.ActiveForm.txtPayer.SetFocus
        Screen.ActiveForm.txtAccountingPeriod.Text = CurrentPeriod
        Screen.ActiveForm.txtReceiptDate.Text = Date
        Screen.ActiveForm.txtPaymentStatus.Text = "PAID"
        Screen.ActiveForm.cboCurrencyCode.Text = "KSH"
End Sub

Public Sub initNEXT()
On Error GoTo err

        With Screen.ActiveForm
            .txtReceiptNoDetails.Text = .txtReceiptNo.Text
            .txtReceiptAmountDetails.Text = .txtReceiptAmount.Text
            .txtTransactionNODetails.Text = .txtTransactionNo.Text
            .txtBalance.Text = .txtReceiptAmount.Text
            .SSTReceipt.Tab = 4
        End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub cmdFirstCode_Click(Index As Integer)
'''On Error GoTo Myerr

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

    
    If Screen.ActiveForm.txtPayer = "" Then
        MsgBox "The Details of the Person Paying are Necessary and cannot be Left Blank"
        Screen.ActiveForm.txtPayer.SetFocus
    
    ElseIf Screen.ActiveForm.cboCurrencyCode.Text = "" Then
        MsgBox "The Currency Code is Required"
        Screen.ActiveForm.cboCurrencyCode.SetFocus

    ElseIf Screen.ActiveForm.txtReceiptDate.Text <= "" Then
        MsgBox "The Receipt Date is  required"
        Screen.ActiveForm.txtReceiptDate.SetFocus
    
    ElseIf DateDiff("D", Screen.ActiveForm.txtReceiptDate.Text, Date) < 0 Then
        MsgBox "Payment cannot be accepted in the future"
        Screen.ActiveForm.txtReceiptDate.SetFocus
    
    ElseIf Screen.ActiveForm.txtReceiptAmount.Text <= "" Then
        MsgBox "The Receipt Amount is required and There cannot be 0"
        Screen.ActiveForm.txtReceiptAmount.SetFocus
    
'    ElseIf Screen.Activeform.txtReceiptNo.Text <= "" Then
'        MsgBox "The Receipt No is required and There cannot be 0"
'        Screen.Activeform.txtReceiptNo.SetFocus

    ElseIf Screen.ActiveForm.cboPaymentMethod = "" Then
        MsgBox "The Payment Method is required", vbOKOnly
        Screen.ActiveForm.cboPaymentMethod.SetFocus

    Else
            bsaveRECORD = True
    End If

End Sub

Public Sub updateRECORD()
        validateRECORD
        If bsaveRECORD = True Then
                createRECEIPT
                saveRecord
                bsaveRECORD = False
                DisableCBRECEIPT
                enableCBRECEIPT
                initNEXT
                NewRecord = False
                beditRECORD = False

        End If
End Sub
Private Sub saveRecord()
On Error GoTo err

    Dim rsDEP As ADODB.Recordset, strDEP As String
    Set rsDEP = New Recordset
    
    rsDEP.Open "select * from ALISMReceiptNew ; ", cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsDEP
            .AddNew
            !ReceiptNo = Screen.ActiveForm.txtReceiptNo
            !ReceiptDate = Screen.ActiveForm.txtReceiptDate
            !AccountingPeriod = Screen.ActiveForm.txtAccountingPeriod
            !PaymentStatus = Screen.ActiveForm.txtPaymentStatus
            !Payer = Screen.ActiveForm.txtPayer
            !PaymentMethod = Screen.ActiveForm.txtPaymentMethod & ""
            !BankNo = Screen.ActiveForm.txtBankNo & ""
            !ReceiptAmount = CDbl(Screen.ActiveForm.txtReceiptAmount)
            !ChequeNo = Screen.ActiveForm.txtChequeNo & ""
            !LocalCheque = Screen.ActiveForm.txtLocal & ""
            !CurrencyCode = Screen.ActiveForm.cboCurrencyCode.Text & ""
            !Preparedby = CurrentUserName
            .Update
    End With

rsDEP.Close
strDEP = ""
    

Exit Sub

err:
        ErrorMessage
End Sub
Sub loadRECEIPTDETAILSGRID()

Dim rsreceipt1 As ADODB.Recordset
Set rsreceipt1 = New Recordset

   rsreceipt1.Open "SELECT DocumentNo, ReceiptType, TransactionAmount FROM ALISMreceiptdetails WHERE ReceiptNo LIKE  '" & Screen.ActiveForm.txtReceiptNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
   Set Screen.ActiveForm.receiptDetailsGRID.DataSource = rsreceipt1
   Set Screen.ActiveForm.ReceiptMGRID.DataSource = rsreceipt1

End Sub

Public Sub searchRECORD()
    
    clearRECORD
    locateRECORD
    If bsearchRECORD = True Then
        loadBANK
        loadPAYMENTMETHOD
        loadRECEIPTDETAILSGRID
        bsearchRECORD = False
    End If
    
End Sub
Public Sub loadBANK()
On Error GoTo err

        Dim rsbank As ADODB.Recordset
        Set rsbank = New ADODB.Recordset
        
        rsbank.Open "SELECT * FROM ALISPBank WHERE bankNo = '" & Screen.ActiveForm.txtBankNo.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsbank
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboBankNo.Text = !CompanyName
 
        End With
rsbank.Close

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadPAYMENTMETHOD()
On Error GoTo err

        Dim rsPAYMETHOD As ADODB.Recordset
        Set rsPAYMETHOD = New Recordset
        
        rsPAYMETHOD.Open "SELECT * FROM ALISPPaymentMethod WHERE PaymentMethod = '" & Screen.ActiveForm.txtPaymentMethod.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsPAYMETHOD
                If .EOF And .BOF Then Exit Sub
                        Screen.ActiveForm.cboPaymentMethod.Text = !PaymentMethodDescription
        End With

rsPAYMETHOD.Close

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub locateRECORD()
'''On Error GoTo Myerr

        Dim strQRE As Variant
        Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset
            strQRE = InputBox("Enter Receipt No to search.", "Search Value")
            rsFind.Open "SELECT * FROM ALISMReceiptNew WHERE ReceiptNo = '" & strQRE & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsFind
            If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
            Else:
                    Screen.ActiveForm.txtReceiptNo.Text = !ReceiptNo
                    Screen.ActiveForm.txtReceiptDate.Text = !ReceiptDate
                    Screen.ActiveForm.txtReceiptAmount.Text = !ReceiptAmount
                    Screen.ActiveForm.txtAccountingPeriod.Text = !AccountingPeriod
                    Screen.ActiveForm.txtPaymentStatus.Text = !PaymentStatus
                    Screen.ActiveForm.txtChequeNo.Text = !ChequeNo & ""
                    Screen.ActiveForm.txtPaymentMethod.Text = !PaymentMethod & ""
                    Screen.ActiveForm.txtBankNo.Text = !BankNo & ""
                    Screen.ActiveForm.txtPayer.Text = !Payer
                    Screen.ActiveForm.txtLocal.Text = !LocalCheque & ""
                    Screen.ActiveForm.cboCurrencyCode.Text = !CurrencyCode & ""
                    Screen.ActiveForm.txtTotalAmount.Text = !TotalAmount & ""
                    Screen.ActiveForm.txtBalance.Text = (!ReceiptAmount - !TotalAmount) & ""
                    Screen.ActiveForm.txtReceiptNoDetails.Text = !ReceiptNo
                    Screen.ActiveForm.txtReceiptAmountDetails.Text = !ReceiptAmount
                    Screen.ActiveForm.txtTransactionNo.Text = !TransactionNo & ""
                    Screen.ActiveForm.txtTransactionNODetails.Text = !TransactionNo & ""
                    Screen.ActiveForm.txtBalance.Text = !ReceiptAmount

                    bsearchRECORD = True
                    Edit = True
            End If
        End With

Exit Sub

Myerr:
    ErrorMessage
End Sub


