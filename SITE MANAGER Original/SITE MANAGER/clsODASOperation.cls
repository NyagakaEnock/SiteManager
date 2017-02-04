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
                strSQL = "SELECT * from AdvertQuotation where QuotationNo = '" & CurrentRecord & "' and (Approved = 'N' or Approved is null); "
    ElseIf bQuotationAuthorization = True Then
                strSQL = "SELECT * from AdvertQuotation where QuotationNo = '" & CurrentRecord & "' and (Authorized = 'N' or Authorized is null) and Approved = 'Y'; "
    ElseIf breceivenoticeAPPROVAL = True Then
                strSQL = "SELECT * from ODASMJobBriefItems where JobBriefItemNo = '" & CurrentRecord & "' and (NoticeApproved = 'N' or NoticeApproved is null); "
    
    ElseIf breceivenoticeAUTHORIZATION = True Then
                strSQL = "SELECT * from ODASMJobBriefItems where JobBriefItemNo = '" & CurrentRecord & "' and (NoticeAuthorized = 'N' or NoticeAuthorized is null) and NoticeApproved = 'Y'; "
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
    ElseIf bRequisitionAPPROVAL = True Then
                strSQL = "SELECT * from ODASMRequisition where RequisitionNo = '" & CurrentRecord & "' and Prepared = 'Y' and (Approved = 'N' or Approved is null); "
    ElseIf bRequisitionAUTHORIZATION = True Then
                strSQL = "SELECT * from ODASMRequisition where RequisitionNo = '" & CurrentRecord & "' and Approved = 'Y' and (Authorized = 'N' or Approved is null); "
    ElseIf bPurchaseOrderAPPROVAL = True Then
                strSQL = "SELECT * from ODASMPurchaseOrder where OrderNo = '" & CurrentRecord & "' and Prepared = 'Y' and (Approved = 'N' or Approved is null); "
    ElseIf bPurchaseOrderAUTHORIZATION = True Then
                strSQL = "SELECT * from ODASMPurchaseOrder where OrderNo = '" & CurrentRecord & "' and Approved = 'Y' and (Authorized = 'N' or Authorized is null); "
    ElseIf bsendnoticeAPPROVAL = True Then
                strSQL = "SELECT * from ODASMLeaseAgreement where ContractNo = '" & CurrentRecord & "' and NoticePrepared = 'Y' and (NoticeApproved = 'N' or NoticeApproved is null); "
    ElseIf bsendnoticeAUTHORIZATION = True Then
                strSQL = "SELECT * from ODASMLeaseAgreement where ContractNo = '" & CurrentRecord & "' and NoticeApproved = 'Y' and (NoticeAuthorized = 'N' or NoticeAuthorized is null); "
    ElseIf breceivenoticeAPPROVAL = True Then
                strSQL = "SELECT * from ODASMLeaseAgreement where ContractNo = '" & CurrentRecord & "' and NoticePrepared = 'Y' and (NoticeApproved = 'N' or NoticeApproved is null); "
    ElseIf breceivenoticeAUTHORIZATION = True Then
                strSQL = "SELECT * from ODASMLeaseAgreement where ContractNo = '" & CurrentRecord & "' and NoticeApproved = 'Y' and (NoticeAuthorized = 'N' or NoticeAuthorized is null); "

    End If
    
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsCONTROL
            If .BOF Or .EOF Then
                    If bQuotationApproval = True Or bJobBriefApproval = True Or bSiteAPPROVAL = True Or bopenJOBBRIEF = True Or bcloseJobBrief = True Or bPurchaseOrderAPPROVAL = True Or bsendnoticeAPPROVAL = True Or breceivenoticeAPPROVAL = True Then
                        MsgBox "Cannot Approved This Document", vbOKOnly
                    ElseIf bQuotationAuthorization = True Or bJobBriefAuthorization = True Or bSiteAuthorization = True Or bPurchaseOrderAUTHORIZATION = True Or bsendnoticeAUTHORIZATION = True Or breceivenoticeAUTHORIZATION = True Then
                        MsgBox "Cannot Authorize this Claim, MUST be Approved First", vbOKOnly
                    End If
                    bQuotationApproval = False
                    bJobBriefApproval = False
                    bSiteAPPROVAL = False
                    bSiteAuthorization = False
                    bopenJOBBRIEF = False
                    bcloseJobBrief = False
                    bRequisitionAPPROVAL = False
                    bRequisitionAUTHORIZATION = False
                    bPurchaseOrderAPPROVAL = False
                    bPurchaseOrderAUTHORIZATION = False
                    bsendnoticeAPPROVAL = False
                    bsendnoticeAUTHORIZATION = False
                    breceivenoticeAPPROVAL = False
                    breceivenoticeAUTHORIZATION = False

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
        
        ElseIf bsendnoticeAPPROVAL = True Or bsendnoticeAUTHORIZATION = True Then
                GlobalPolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord
        ElseIf breceivenoticeAPPROVAL = True Or breceivenoticeAUTHORIZATION = True Then
                GlobalPolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord
        ElseIf breceivenoticeAPPROVAL = True Or breceivenoticeAUTHORIZATION = True Then
                GlobalPolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord

        ElseIf bPurchaseOrderAPPROVAL = True Or bPurchaseOrderAUTHORIZATION = True Then
                GlobalPolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord

        ElseIf bopenJOBBRIEF = True Or bcloseJobBrief = True Then
                GlobalPolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord
        
        ElseIf bRequisitionAPPROVAL = True Or bRequisitionAUTHORIZATION = True Then
                GlobalPolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord

        ElseIf bJobCardAPPROVAL = True Or bJobCardAUTHORIZATION = True Then
                GlobalPolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord

        ElseIf bSiteAPPROVAL = True Or bSiteAuthorization = True Then
                GlobalPolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord

        ElseIf bJobBriefApproval = True Or bJobBriefAuthorization = True Then
                GlobalPolicyNo = CurrentRecord
                globalJOBCARDNo = CurrentRecord
                GlobalApplicationNo = CurrentRecord

        End If
        
        '/* Procedure to pick the Operation type that refers to the APPROVED Operation
        
        Set rsCONTROL = New ADODB.Recordset
        
        If bQuotationApproval = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE QuotationApproval = '1' ;"
        ElseIf bQuotationAuthorization = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE QuotationAuthorization = '1' ;"
        ElseIf breceivenoticeAPPROVAL = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE RenewalAPPROVAL = '1' ;"
        ElseIf breceivenoticeAUTHORIZATION = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE RenewalAUTHORIZATION = '1' ;"

        ElseIf bJobCardAPPROVAL = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE JobCardApproval = '1' ;"
        ElseIf bJobCardAUTHORIZATION = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE JobCardAuthorization = '1' ;"
        ElseIf bsendnoticeAPPROVAL = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE SendNoticeApproval = '1' ;"
        ElseIf bsendnoticeAUTHORIZATION = True Then
                strSQL = "SELECT * FROM ODASPOperationType  WHERE SendNoticeAuthorization = '1' ;"
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
        Else
                Exit Sub
        End If
        
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        With rsCONTROL
                If .EOF And .BOF Then Exit Sub
                
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
