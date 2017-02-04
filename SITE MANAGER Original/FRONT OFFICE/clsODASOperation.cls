VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim rsOPERATION As clsODASOperation
Public Sub showAPPROVAL()
On Error GoTo err
    
   
    Set rsCONTROL = New ADODB.Recordset
    
    strSQL = ""
    If bClaimRegApproval = True Then
            strSQL = "SELECT * from ALISMClaimRegistration where ClaimNo = '" & Screen.ActiveForm.txtClaimNo & "'; "
    ElseIf bclaimregAuthorization = True Then
            strSQL = "SELECT * from ALISMClaimRegistration where ClaimNo = '" & Screen.ActiveForm.txtClaimNo & "'; "
    End If
    
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsCONTROL
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

rsCONTROL.Close
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub checkAPPROVEDDISCHARGE()
On Error GoTo err

    Set rsCONTROL = New ADODB.Recordset
    
    strSQL = ""
    
    If bQuotationApproval = True Then
                strSQL = "SELECT * from ODASMQuotation where QuotationNo = '" & CurrentRecord & "' and (Approved = 'N' or Approved is null); "
    ElseIf bQuotationAuthorization = True Then
                strSQL = "SELECT * from ODASMQuotation where QuotationNo = '" & CurrentRecord & "' and (Authorized = 'N' or Authorized is null) and Approved = 'Y'; "
    ElseIf btripAPPROVAL = True Then
                strSQL = "SELECT * from ODASMTripMaster where TripNo = '" & CurrentRecord & "' and (Approved = 'N' or Approved is null); "
    ElseIf btripAUTHORIZATION = True Then
                strSQL = "SELECT * from ODASMTripMaster where TripNo = '" & CurrentRecord & "' and (Authorized = 'N' or Authorized is null) and Approved = 'Y'; "
    ElseIf bCostingsApproval = True Then
                strSQL = "SELECT * from ODASMProjectCost where JobBriefNo = '" & CurrentRecord & "' and (Approved = 'N' or Approved is null); "
    ElseIf bCostingsAuthorization = True Then
                strSQL = "SELECT * from ODASMProjectCost where JobBriefNo = '" & CurrentRecord & "' and (Authorized = 'N' or Authorized is null) and Approved = 'Y'; "
    ElseIf bJobBriefApproval = True Then
                strSQL = "SELECT * from ODASMJobBrief where JobBriefNo = '" & CurrentRecord & "' and (Approved = 'N' or Approved is null); "
    ElseIf bJobBriefAuthorization = True Then
                strSQL = "SELECT * from ODASMJobBrief where JobBriefNo = '" & CurrentRecord & "' and (Authorized = 'N' or Authorized is null) and Approved = 'Y'; "
    ElseIf bSiteAPPROVAL = True Then
                strSQL = "SELECT * from ODASMLeaseAgreement where ContractNo = '" & CurrentRecord & "' and (Approved = 'N' or Approved is null ); "
    ElseIf bSiteAuthorization = True Then
                strSQL = "SELECT * from ODASMLeaseAgreement where ContractNo = '" & CurrentRecord & "' and (Authorized = 'N' or Authorized is null) and Approved = 'Y'; "
    ElseIf bopenJOBBRIEF = True Then
                strSQL = "SELECT * from ODASMJobCard where JobCardNo = '" & CurrentRecord & "' and Prepared = 'Y' and (Opened = 'N' or Opened is null) and Status = 'START' ; "
    ElseIf bcloseJobBrief = True Then
                strSQL = "SELECT * from ODASMJobCard where JobCardNo = '" & CurrentRecord & "' and DepartmentCode = '" & globalDepartmentCode & "'and Used = 'Y' and (Closed = 'N' or closed is null) and Status = 'OPENED' ; "
    ElseIf bRequisitionAPPROVAL = True Then
                strSQL = "SELECT * from ODASMRequisitionItems where ItemNo = '" & CurrentRecord & "' and Prepared = 'Y' and (Approved = 'N' or Approved is null); "
    ElseIf bRequisitionAUTHORIZATION = True Then
                strSQL = "SELECT * from ODASMRequisitionItems where ItemNo = '" & CurrentRecord & "' and Approved = 'Y' and (Authorized = 'N' or Authorized is null); "
    ElseIf bPurchaseOrderAPPROVAL = True Then
                strSQL = "SELECT * from ODASMLPO where LPONo = '" & CurrentRecord & "' and Prepared = 'Y' and (LPOApproved = 'N' or LPOApproved is null); "
    ElseIf bPurchaseOrderAUTHORIZATION = True Then
                strSQL = "SELECT * from ODASMLPO where LPONo = '" & CurrentRecord & "' and LPOApproved = 'Y' and (LPOAuthorized = 'N' or LPOAuthorized is null); "
    ElseIf bsendnoticeAPPROVAL = True Then
                strSQL = "SELECT * from ODASMLeaseAgreement where ContractNo = '" & CurrentRecord & "' and NoticePrepared = 'Y' and (NoticeApproved = 'N' or NoticeApproved is null); "
    ElseIf bsendnoticeAUTHORIZATION = True Then
                strSQL = "SELECT * from ODASMLeaseAgreement where ContractNo = '" & CurrentRecord & "' and NoticeApproved = 'Y' and (NoticeAuthorized = 'N' or NoticeAuthorized is null); "
    ElseIf breceivenoticeAPPROVAL = True Then
                strSQL = "SELECT * from ODASMLeaseAgreement where ContractNo = '" & CurrentRecord & "' and NoticePrepared = 'Y' and (NoticeApproved = 'N' or NoticeApproved is null); "
    ElseIf breceivenoticeAUTHORIZATION = True Then
                strSQL = "SELECT * from ODASMLeaseAgreement where ContractNo = '" & CurrentRecord & "' and NoticeApproved = 'Y' and (NoticeAuthorized = 'N' or NoticeAuthorized is null); "
    ElseIf bCancelREQUISITION = True Then
                strSQL = "SELECT * from ODASMRequisitionItems where ItemNo = '" & CurrentRecord & "' and (cancelled = 'N' or CancelLed is null); "
    ElseIf bApproveVOUCHER = True Then
                strSQL = "SELECT * from ODASMVoucher where VoucherNo = '" & CurrentRecord & "' and (Approved = 'N' or Approved is null); "
    ElseIf bauthorizeVOUCHER = True Then
                strSQL = "SELECT * from ODASMVoucher where VoucherNo = '" & CurrentRecord & "' and (Authorized = 'N' or Authorized is null) and Approved = 'Y'; "
    ElseIf bApproveCheque = True Then
                strSQL = "SELECT * from ALISMCheque where ChequeNo = '" & CurrentRecord & "' and Approved = 'N' and prepared = 'Y'; "
    ElseIf bAuthorizeCheque = True Then
                strSQL = "SELECT * from ALISMCheque where ChequeNo = '" & CurrentRecord & "' and Authorized = 'N' and Approved = 'Y'; "
    ElseIf bapproveINVOICE = True Then
                strSQL = "SELECT * from ODASMInvoiceSENT where InvoiceNo = '" & CurrentRecord & "' and Prepared = 'Y' and (Approved = 'N' or Approved is null) ; "
    ElseIf bauthorizeINVOICE = True Then
                strSQL = "SELECT * from ODASMInvoiceSENT where invoiceNo = '" & CurrentRecord & "' and (Authorized = 'N' or Authorized is null) and Approved = 'Y'; "
    ElseIf bLPOApproval = True Then
                strSQL = "SELECT * from ODASMrequisition where RequisitionNo = '" & CurrentRecord & "' and Prepared = 'Y' and (LPOApproved = 'N' or LPOApproved is null) ; "
    ElseIf bLPOAuthorization = True Then
                strSQL = "SELECT * from ODASMrequisition where RequisitionNo = '" & CurrentRecord & "' and (LPOAuthorized = 'N' or LPOAuthorized is null) and Approved = 'Y'; "
    Else
            bapproveRECORD = False
    End If
    
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsCONTROL
            If .BOF Or .EOF Then
                    
                    bapproveRECORD = False
                    If bCancelREQUISITION = True Or bQuotationApproval = True Or bJobBriefApproval = True Or bSiteAPPROVAL = True Or bopenJOBBRIEF = True Or bLPOApproval Or bcloseJobBrief = True Or bPurchaseOrderAPPROVAL = True Or bsendnoticeAPPROVAL = True Or breceivenoticeAPPROVAL = True Then
                        MsgBox "Cannot Approved This Document", vbOKOnly
                    ElseIf bQuotationAuthorization = True Or bJobBriefAuthorization = True Or bSiteAuthorization = True Or bPurchaseOrderAUTHORIZATION = True Or bsendnoticeAUTHORIZATION = True Or breceivenoticeAUTHORIZATION = True Then
                        MsgBox "Cannot Authorize this Claim, MUST be Approved First", vbOKOnly
                    ElseIf bcloseJobBrief = True Then
                        MsgBox "Cannot Close The Job Brief"
                    End If
                    
                    bQuotationApproval = False
                    bJobBriefApproval = False
                    bSiteAPPROVAL = False
                    bSiteAuthorization = False
                    bopenJOBBRIEF = False
                    bcloseJobBrief = False
                    bRequisitionAPPROVAL = False
                    bRequisitionAUTHORIZATION = False
                    bLPOApproval = True
                    bLPOAuthorization = True
                    bsendnoticeAPPROVAL = False
                    bsendnoticeAUTHORIZATION = False
                    breceivenoticeAPPROVAL = False
                    breceivenoticeAUTHORIZATION = False
                    bCancelREQUISITION = False
                    bApproveVOUCHER = False
                    bauthorizeVOUCHER = False
                    bApproveCheque = False
                    bapproveINVOICE = False
                    bapproveINVOICE = False
                    bauthorizeINVOICE = False
                    bCostingsApproval = False
                    bCostingsAuthorization = False

            Else
                bapproveRECORD = True
            End If
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub approveOPERATION()
On Error GoTo err
       If bQuotationApproval = True Or bQuotationAuthorization = True Then
                GlobalApplicationNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
        ElseIf bPurchaseOrderAPPROVAL = True Or bPurchaseOrderAUTHORIZATION = True Then
                globalpolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord
        ElseIf bCancelREQUISITION = True Then
                globalpolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord
        ElseIf bApproveCheque = True Or bAuthorizeCheque = True Then
                globalpolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord
        ElseIf bApproveVOUCHER = True Or bauthorizeVOUCHER = True Then
                globalpolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord
        ElseIf bapproveINVOICE = True Or bauthorizeINVOICE = True Then
                globalpolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord
        ElseIf bsendnoticeAPPROVAL = True Or bsendnoticeAUTHORIZATION = True Then
                globalpolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord
        
        ElseIf breceivenoticeAPPROVAL = True Or breceivenoticeAUTHORIZATION = True Then
                globalpolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord
        ElseIf btripAPPROVAL = True Or btripAUTHORIZATION = True Then
                globalpolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord
        ElseIf bPurchaseOrderAPPROVAL = True Or bPurchaseOrderAUTHORIZATION = True Then
                globalpolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord

        ElseIf bopenJOBBRIEF = True Or bcloseJobBrief = True Then
                globalpolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord
        
        ElseIf bRequisitionAPPROVAL = True Or bRequisitionAUTHORIZATION = True Then
                globalpolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord

        ElseIf bJobCardAPPROVAL = True Or bJobCardAUTHORIZATION = True Then
                globalpolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord

        ElseIf bSiteAPPROVAL = True Or bSiteAuthorization = True Then
                globalpolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord

        ElseIf bJobBriefApproval = True Or bJobBriefAuthorization = True Then
                globalpolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord

        ElseIf bCostingsApproval = True Or bCostingsAuthorization = True Then
                globalpolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord
        End If
        
        '/* Procedure to pick the Operation type that refers to the APPROVED Operation
        
        Set rsCONTROL = New ADODB.Recordset
        
        If bQuotationApproval = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE QuotationApproval = '1' ;"
        
        ElseIf bCancelREQUISITION = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE cancelrequisition = '1' ;"
        ElseIf bQuotationAuthorization = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE QuotationAuthorization = '1' ;"
        ElseIf bJobCardAPPROVAL = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE JobCardApproval = '1' ;"
        ElseIf bJobCardAUTHORIZATION = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE JobCardAuthorization = '1' ;"
        ElseIf bsendnoticeAPPROVAL = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE SendNoticeApproval = '1' ;"
        ElseIf bsendnoticeAUTHORIZATION = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE SendNoticeAuthorization = '1' ;"
        ElseIf bCostingsApproval = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE CostingsAPPROVAL = '1' ;"
        ElseIf bCostingsAuthorization = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE CostingsAUTHORIZATION = '1' ;"
        
        ElseIf breceivenoticeAPPROVAL = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE ReceiveNoticeApproval = '1' ;"
        ElseIf breceivenoticeAPPROVAL = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE ReceiveNoticeAuthorization = '1' ;"
        ElseIf bJobBriefApproval = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE JobBriefApproval = '1' ;"
        ElseIf bJobBriefAuthorization = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE JobBriefAuthorization = '1' ;"
        ElseIf bSiteAPPROVAL = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE SiteApproval = '1' ;"
        ElseIf bSiteAuthorization = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE SiteAuthorization = '1' ;"
        ElseIf bopenJOBBRIEF = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE openJobCard = '1' ;"
        ElseIf bcloseJobBrief = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE closeJobCard = '1' ;"
        ElseIf bRequisitionAPPROVAL = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE RequisitionApproval = '1' ;"
        ElseIf bRequisitionAUTHORIZATION = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE RequisitionAuthorization = '1' ;"
        ElseIf bPurchaseOrderAPPROVAL = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE PurchaseOrderApproval = '1' ;"
        ElseIf bPurchaseOrderAUTHORIZATION = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE PurchaseOrderAuthorization = '1' ;"
        ElseIf btripAPPROVAL = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE TripAPPROVAL = '1' ;"
        ElseIf btripAUTHORIZATION = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE TripAUTHORIZATION = '1' ;"
        ElseIf bApproveVOUCHER = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE voucherApproval = '1' ;"
        ElseIf bauthorizeVOUCHER = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE VoucherAuthorization = '1' ;"
        ElseIf bapproveINVOICE = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE InvoiceApproval = '1' ;"
        ElseIf bauthorizeINVOICE = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE InvoiceAuthorization = '1' ;"
        ElseIf bApproveCheque = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE ChequeApproval = '1' ;"
        ElseIf bAuthorizeCheque = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE ChequeAuthorization = '1' ;"
        End If
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsCONTROL
                If .EOF Or .BOF Then Exit Sub
                
                GlobalOperationType = rsCONTROL!OperationType
                GlobalOperationDescription = rsCONTROL!Description
                Load frmODASMOperation
                frmODASMOperation.Show 1
        End With
        
        GlobalOperationType = ""
        GlobalOperationDescription = ""
        GlobalApplicationNo = ""


Exit Sub

err:
    ErrorMessage
End Sub
