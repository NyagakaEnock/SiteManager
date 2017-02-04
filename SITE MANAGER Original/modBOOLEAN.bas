Attribute VB_Name = "modBOOLEAN"
Public GlobalOperationType, GlobalApplicationNo, GlobalOperationDescription As String
Public strSQL As String, strCONTROL As String
Public rsCONTROL As ADODB.Recordset, rsSiteSchedule As ADODB.Recordset, rsSAVE As ADODB.Recordset, rsAnnualRate As ADODB.Recordset, rsMast As ADODB.Recordset
Public bSaveRECORD, baddRECORD, beditRECORD, bsearchRECORD, bDontChange As Boolean, bexportRECORD As Boolean
Public bQuotationApproval, bQuotationAuthorization, bQuotationPreparation, bJobBriefPrepration, bJobBriefAuthorization, bJobBriefApproval, bAlert As Boolean
Public globalJOBCARDNo, GlobalDepartmentCode As Variant
Public bRenewalAPPROVAL, bRenewalAUTHORIZATION, bLeaseAPPROVAL, bLeaseAUTHORIZATION, bSiteAPPROVAL, bSiteAuthorization, bopenJOBBRIEF, bcloseJobBrief, bsendnoticePREPARATION, bsendnoticeAPPROVAL, bsendnoticeAUTHORIZATION As Boolean
Public bRequisitionAPPROVAL, bRequisitionAUTHORIZATION, bPurchaseOrderAPPROVAL, bPurchaseOrderAUTHORIZATION As Boolean
Public breceivenoticePREPARATION, breceivenoticeAPPROVAL, breceivenoticeAUTHORIZATION, edit As Boolean
Public bBillBoard, bBillBoardFace, CouncilForm, bStreetSign, bSaveCosts, bAllowProcess, bPlotRenewal As Boolean

