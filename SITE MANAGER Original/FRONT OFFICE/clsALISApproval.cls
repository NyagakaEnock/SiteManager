VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsALISApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub showAPPROVAL()
On Error GoTo err
    
    Dim rsAPPROVAL As ADODB.Recordset
    
    Set rsAPPROVAL = New ADODB.Recordset
    
    strSQL = ""
    If bClaimRegApproval = True Then
            strSQL = "SELECT * from ODASMInvoice where ClaimNo = '" & Screen.ActiveForm.txtClaimNo & "'; "
    ElseIf bclaimregAuthorization = True Then
            strSQL = "SELECT * from ODASMInvoice where ClaimNo = '" & Screen.ActiveForm.txtClaimNo & "'; "
    End If
    
    rsAPPROVAL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsAPPROVAL
            If .BOF Or .EOF Then Exit Sub
            Screen.ActiveForm.txtAuthorizedBy.Text = !Preparedby
            Screen.ActiveForm.txtDateAuthorized.Text = !dateprepared
            Screen.ActiveForm.txtAuthorized.Text = !Authorized & ""
            Screen.ActiveForm.txtApproved.Text = !Approved & ""
            Screen.ActiveForm.txtAuthorizedBy = !AuthorizedBy & ""
            Screen.ActiveForm.txtDateAuthorized = !DateAuthorized & ""
            Screen.ActiveForm.txtApprovedBy = !ApprovedBy & ""
            Screen.ActiveForm.txtDateApproved = !DateApproved & ""
    End With

rsAPPROVAL.Close
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub checkAPPROVEDDISCHARGE()
On Error GoTo err

    Dim rsAPPROVAL As ADODB.Recordset
    Set rsAPPROVAL = New ADODB.Recordset
    
    bapproveRECORD = False
    
    strSQL = ""
    If bClaimRegApproval = True Then
                strSQL = "SELECT * from ODASMInvoice where ClaimNo = '" & Screen.ActiveForm.txtClaimNo & "' and (Approved = '' or Approved is null); "
    ElseIf bclaimregAuthorization = True Then
                strSQL = "SELECT * from ODASMInvoice where ClaimNo = '" & Screen.ActiveForm.txtClaimNo & "' and (Authorized = '' or Authorized is null) and Approved = 'Y'; "
    ElseIf bApproveVOUCHER = True Then
                strSQL = "SELECT * from ODASMVoucher where VoucherNo = '" & CurrentRecord & "' and (Approved = 'N' or Approved is null); "
    ElseIf bauthorizeVOUCHER = True Then
                strSQL = "SELECT * from ODASMVoucher where VoucherNo = '" & CurrentRecord & "' and (Authorized = 'N' or Authorized is null) and Approved = 'Y'; "
    ElseIf bapproveREQUISITION = True Then
                strSQL = "SELECT * from ODASMRequisition where RequisitionNo = '" & CurrentRecord & "' and Approved = 'N' and prepared = 'Y'; "
    ElseIf bAuthorizeREQUISITION = True Then
                strSQL = "SELECT * from ODASMRequisition where RequisitionNo = '" & CurrentRecord & "' and Authorized = 'N' and Approved = 'Y'; "
    ElseIf bApproveCheque = True Then
                strSQL = "SELECT * from ALISMCheque where ChequeNo = '" & CurrentRecord & "' and Approved = 'N' and prepared = 'Y'; "
    ElseIf bAuthorizeCheque = True Then
                strSQL = "SELECT * from ALISMCheque where ChequeNo = '" & CurrentRecord & "' and Authorized = 'N' and Approved = 'Y'; "
    ElseIf bapproveINVOICE = True Then
                strSQL = "SELECT * from ODASMInvoiceSENT where InvoiceNo = '" & CurrentRecord & "' and Prepared = 'Y' and (Approved = 'N' or Approved is null) ; "
    ElseIf bLoanCHECKED = True Then
                strSQL = "SELECT * from ODASMInvoiceSENT where InvoiceNo = '" & CurrentRecord & "' and (Approved = 'N' or Approved is null); "
    ElseIf bauthorizeINVOICE = True Then
                strSQL = "SELECT * from ODASMInvoiceSENT where invoiceNo = '" & CurrentRecord & "' and (Authorized = 'N' or Authorized is null) and Approved = 'Y'; "
    ElseIf bApproveREINSTATEMENT = True Then
                strSQL = "SELECT * from ALISMReinstatement where JobBriefNo = '" & Screen.ActiveForm.cboJobBriefNo & "' and Approved = 'N' and prepared = 'Y'; "
    ElseIf bAuthorizeREINSTATEMENT = True Then
                strSQL = "SELECT * from ALISMReinstatement where JobBriefNo = '" & Screen.ActiveForm.cboJobBriefNo & "' and Approved = 'Y' and Prepared = 'Y' and Authorized = 'N'; "
    ElseIf bApprovePAIDUP = True Then
                strSQL = "SELECT * from ALISMPaidup where JobBriefNo = '" & Screen.ActiveForm.txtJobBriefNo & "' and Approved = 'N' and prepared = 'Y'; "
    ElseIf bAuthorizePAIDUP = True Then
                strSQL = "SELECT * from ALISMPaidup where JobBriefNo = '" & Screen.ActiveForm.txtJobBriefNo & "' and Approved = 'Y' and Prepared = 'Y' and Authorized = 'N'; "
    End If
    
    If strSQL = Empty Then Exit Sub
    
    rsAPPROVAL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsAPPROVAL
            If .BOF Or .EOF Then
                    If bApproveVOUCHER = True Or bapproveREQUISITION = True Or bApproveCheque = True Or bapproveINVOICE = True Or bApproveREINSTATEMENT = True Or bApprovePAIDUP = True Then
                        MsgBox "Cannot Approved This Document", vbOKOnly
                    ElseIf bauthorizeVOUCHER = True Or bAuthorizeREQUISITION = True Or bAuthorizeCheque = True Or bauthorizeINVOICE = True And bAuthorizeREINSTATEMENT = True Or bAuthorizePAIDUP = True Then
                        MsgBox "Cannot Authorize this Claim, MUST be Approved First", vbOKOnly
                    End If
                    
                    bApproveVOUCHER = False
                    bauthorizeVOUCHER = False
                    bApproveCheque = False
                    bapproveREQUISITION = False
                    bapproveINVOICE = False
                    bApproveREINSTATEMENT = False
                    bAuthorizeREINSTATEMENT = False
                    bAuthorizePAIDUP = False
                    bApprovePAIDUP = False
                    bapproveINVOICE = False
                    bauthorizeINVOICE = False
            Else
                    bapproveRECORD = True
            End If
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub switchCOMMANDBUTTONS()
On Error GoTo err
        With Screen.ActiveForm
                If .txtApprovedBy.Text > "" Then
                        .cmdApprove.Enabled = False
                        .cmdAuthorize.Enabled = True
                        .cmdPrint.Enabled = False
                End If
                
                If .txtApprovedBy.Text <= "" Then
                        .cmdAuthorize.Enabled = False
                        .cmdApprove.Enabled = True
                        .cmdPrint.Enabled = False
                End If
                
                If .txtAuthorizedBy.Text > "" Then
                        .cmdApprove.Enabled = False
                        .cmdAuthorize.Enabled = False
                        .cmdPrint.Enabled = True
                End If
                
                If .txtAuthorizedBy.Text <= "" And .txtApprovedBy.Text <= "" Then
                        .cmdApprove.Enabled = True
                        .cmdAuthorize.Enabled = False
                        .cmdPrint.Enabled = True
                End If

                
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadAPPROVALDETAILS()
On Error GoTo err

        Dim rsCTOTAL As ADODB.Recordset
        Set rsCTOTAL = New Recordset
            
        If bApproveDischarge = True Or bAuthorizeDischarge = True Or bsearchDISCHARGE = True Then
                rsCTOTAL.Open "SELECT * FROM ALISMClaimTotal Where ClaimNo =  '" & frmALISMClaimDischarge.cboClaimNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        ElseIf bapproveREQUISITION = True Or bAuthorizeREQUISITION = True Or bsearchRequisition = True Then
                rsCTOTAL.Open "SELECT * FROM ODASMVoucher Where RequisitionNo =  '" & Screen.ActiveForm.txtVoucherNo.Text & "'; ", cnCOMMON, adOpenKeyset, adLockOptimistic
        ElseIf bApproveCheque = True Or bAuthorizeCheque = True Then
                rsCTOTAL.Open "SELECT * FROM ALISMCheque Where ChequeNo =  '" & frmALISMCheque.txtChequeNo.Text & "'; ", cnCOMMON, adOpenKeyset, adLockOptimistic
        ElseIf bClaimRegApproval = True Or bclaimregAuthorization = True Then
                rsCTOTAL.Open "SELECT * FROM ODASMInvoice Where ClaimNo =  '" & frmODASMInvoice.txtClaimNo.Text & "'; ", cnCOMMON, adOpenKeyset, adLockOptimistic
        Else
            rsCTOTAL.Open "SELECT * FROM ALISMClaimTotal Where ClaimNo =  '" & frmALISMClaimDischarge.cboClaimNo & "'; ", cnCOMMON, adOpenKeyset, adLockOptimistic
        End If
        
           With rsCTOTAL
                    If .BOF Or .EOF Then Exit Sub
                            Screen.ActiveForm.txtpreparedby.Text = !Preparedby & ""
                            Screen.ActiveForm.txtPrepared.Text = !Prepared & ""
                            Screen.ActiveForm.txtDatePrepared.Text = !dateprepared & ""
                            Screen.ActiveForm.txtDateApproved.Text = !DateApproved & ""
                            Screen.ActiveForm.txtApprovedBy.Text = !ApprovedBy & ""
                            Screen.ActiveForm.txtAuthorizedBy.Text = !AuthorizedBy & ""
                            Screen.ActiveForm.txtDateAuthorized.Text = !DateAuthorized & ""
                            Screen.ActiveForm.txtApproved.Text = !Approved & ""
                            Screen.ActiveForm.txtAuthorized.Text = !Authorized & ""
            End With
            
rsCTOTAL.Close


Exit Sub

err:
    ErrorMessage
End Sub

Public Sub approveCLAIM()
On Error GoTo err
        If bClaimRegApproval = True Or bclaimregAuthorization = True Then
                GlobalApplicationNo = Screen.ActiveForm.txtClaimNo.Text
                GlobalClaimNo = Screen.ActiveForm.txtClaimNo
        
        ElseIf bapproveRECORD = True Then
                GlobalApplicationNo = CurrentRecord
                GlobalClaimNo = CurrentRecord

        End If
        
        '/* Procedure to pick the Operation type that refers to the APPROVED Operation
        
        Dim rsAPPROVED As ADODB.Recordset, strAPPROVED As String
        Set rsAPPROVED = New ADODB.Recordset
        
        If bClaimRegApproval = True Then
                strAPPROVED = "SELECT * FROM ODASPOperationType  WHERE RegistrationApproval = '1' ;"
        ElseIf bclaimregAuthorization = True Then
                strAPPROVED = "SELECT * FROM ODASPOperationType  WHERE RegistrationAuthorization = '1' ;"
        ElseIf bApproveVOUCHER = True Then
                strAPPROVED = "SELECT * FROM ODASPOperationType  WHERE VoucherAPPROVal = '1' ;"
        ElseIf bapproveREQUISITION = True Then
                strAPPROVED = "SELECT * FROM ODASPOperationType  WHERE PaymentAPPROVal = '1' ;"
        ElseIf bAuthorizeREQUISITION = True Then
                strAPPROVED = "SELECT * FROM ODASPOperationType  WHERE PaymentAuthorization = '1' ;"
        ElseIf bApproveCheque = True Then
                strAPPROVED = "SELECT * FROM ODASPOperationType  WHERE ChequeAPPROVal = '1' ;"
        ElseIf bAuthorizeCheque = True Then
                strAPPROVED = "SELECT * FROM ODASPOperationType  WHERE ChequeAuthorization = '1' ;"
        ElseIf bauthorizeVOUCHER = True Then
                strAPPROVED = "SELECT * FROM ODASPOperationType  WHERE VoucherAuthorization = '1' ;"
        ElseIf bauthorizeINVOICE = True Then
                strAPPROVED = "SELECT * FROM ODASPOperationType  WHERE InvoiceAuthorization = '1' ;"
        ElseIf bapproveINVOICE = True Then
                strAPPROVED = "SELECT * FROM ODASPOperationType  WHERE InvoiceApproval = '1' ;"
        ElseIf bApproveREINSTATEMENT = True Then
                strAPPROVED = "SELECT * FROM ODASPOperationType  WHERE ReinstatementApproval = '1' ;"
        ElseIf bAuthorizeREINSTATEMENT = True Then
                strAPPROVED = "SELECT * FROM ODASPOperationType  WHERE ReinstatementAuthorization = '1' ;"
        ElseIf bApprovePAIDUP = True Then
                strAPPROVED = "SELECT * FROM ODASPOperationType  WHERE PaidupApproval = '1' ;"
        ElseIf bAuthorizePAIDUP = True Then
                strAPPROVED = "SELECT * FROM ODASPOperationType  WHERE PaidupAuthorization = '1' ;"

        End If
        
        rsAPPROVED.Open strAPPROVED, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsAPPROVED
                If .EOF Or .BOF Then Exit Sub
                
                GlobalOperationType = rsAPPROVED!OperationType
                GlobalOperationDescription = rsAPPROVED!Description
        End With
        
        Call rsLOANVALUE.operationAPPROVED
        Set rsLOANVALUE = Nothing
        
        GlobalOperationType = ""
        GlobalOperationDescription = ""
        GlobalApplicationNo = ""
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
                            
                        strAPPROVED = "SELECT * FROM ALISMLoanOperation  WHERE ApplicationNo = '" & frmALISMClaimDischarge.cboClaimNo & "' and operationType = '" & !OperationType & "' ;"
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
                                        
                                    strAPPROVE = "SELECT * FROM ALISMLoanOperation  WHERE ApplicationNo = '" & frmALISMClaimDischarge.cboClaimNo & "' and operationType = '" & !OperationType & "' ;"
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
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub AuthorizeClaim()
On Error GoTo err
        If frmALISMClaimDischarge.cboClaimNo.Text <= "" Then
                MsgBox "Kindly Select A Record before executing this Button", vbOKOnly
                Exit Sub
        Else:
                
                GlobalApplicationNo = frmALISMClaimDischarge.cboClaimNo.Text
                
                    If bExitSub = True Then
                        bExitSub = False
                        Exit Sub
                    End If
                    
                '/* Procedure to pick the Operation type that refers to the APPROVED Operation
                
                Dim rsAUTHORIZATION As ADODB.Recordset, strAPPROVED As String
                Set rsAUTHORIZATION = New Recordset
                    
                If bAuthorizeDischarge = True Then
                        strAPPROVED = "SELECT * FROM ALISPLoanOperationType  WHERE DischargeAuthorization = '1' ;"
                ElseIf bAuthorizeREQUISITION = True Then
                        strAPPROVED = "SELECT * FROM ALISPLoanOperationType  WHERE PaymentAuthorization = '1' ;"
                ElseIf bAuthorizeCheque = True Then
                        strAPPROVED = "SELECT * FROM ALISPLoanOperationType  WHERE chequeAuthorization = '1' ;"
                End If
                
                rsAUTHORIZATION.Open strAPPROVED, cnCOMMON, adOpenKeyset, adLockOptimistic

                With rsAUTHORIZATION
                        If .EOF Or .BOF Then Exit Sub
                        
                        GlobalOperationType = rsAUTHORIZATION!OperationType
                        GlobalOperationDescription = rsAUTHORIZATION!Description
                End With

                Call rsLOANVALUE.operationAPPROVED
                Set rsLOANVALUE = Nothing
                
                GlobalOperationType = ""
                GlobalOperationDescription = ""
                
                GlobalApplicationNo = ""
        End If
Exit Sub

err:
    ErrorMessage

End Sub
