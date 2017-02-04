Attribute VB_Name = "ModBooleans"
Option Explicit
Public rsRELATIONSHIP, rsCOSTCENTER  As ADODB.Recordset
Public dtLASTPREMDUEDATE, ThisDate As Date
Public strSQL, strCONTROL, gRequirementCode, globalDepartmentCode As String
Public BoolCate, BoolPolicyDetails, BoolInstallment, BoolProposal, BoolOrders, BoolEmploymentType, BoolAccountTypes, BoolCustomerMessage, BoolSupplierType, BoolCustomerType, BoolCurrency, BoolCompany, BoolVat, BoolChart, Boolemployee, BoolSupplier, Boolcustomer, BoolCust, BoolPay, BoolProduct As Boolean
Public LoadCate, loadPolicyDetails, LoadBeneficiary, LoadPLedger, loadPROPOSAL, loadINSTALLMENT, loadPolicy, LoadEmploymentType, loadRECEIPT, LoadCustomerMessage, LoadSupplierType, LoadCustomerType, LoadCurrency, LoadCompany, LoadVAT, LoadChart, LoadEmployee, LoadCust, LoadCustomer, LoadSupplier, Loadpay, loadPRODUCT As Boolean
Public inquiry, bEmployerUpdate, bCalcDays, bcompleteBENEFICIARY, bsingleBONUS, bloadClaimDeductions, bloadClaimProceeds As Boolean
Public surrender, bRequireAccountNo, bcompleteQUOTATION, bcoyRetentionCalculated As Boolean
Public payvoucher, bsearchRECORD, bloanREFUND, bupdateLOAN, bExitProgram, breverseRECEIPT As Boolean
Public bExitSub, blapseREINSURANCE, bsurrenderREINSURANCE, breinstateREINSURANCE, bdeathREINSURANCE, baddRECORD, beditRECORD, bloadRECORD, bsearchPOLICY, bobtainRIDER, bobtainPlan, bsaveRECORD As Boolean
Public addpen, bstatusUpdate, bCONTINUE, bCHECK, bCreateMedicals, bbankRECEIPT, bsurrenderCLAIM As Boolean
Public bLOCKFORM, bloadINVOICE, BSave, bVAL, bunderwritePROPOSAL, bcompleteTAKEON, bcompleteUNDERWRITING As Boolean
Public bprintMATURTIES, bprintMATSUM, bprintAnnuities, bprintPMaturity, bprintPMaturitySum, bprintAnnuitySum, bprintClaimListing, bprintClaimSummary, bprintClaimConsolidation, bprintMaturityReport As Boolean
Public breceiveDOCUMENTS, bphysicianSTATEMENT, bidentificationSTATEMENT, bPayee, bSearchByJobBriefNo, bapproveINVOICE, bauthorizeINVOICE As Boolean
Public bexitFORM, bclaimantSTATEMENT, bprocessedCLAIMS, bClaimRegApproval, bclaimregAuthorization, bsendDOCUMENTS, bsearchRequisition, bsearchDISCHARGE As Boolean
Public bApproveDischarge, bAuthorizeDischarge, bapproveREQUISITION, bAuthorizeREQUISITION, bApproveCheque, bAuthorizeCheque, bUseClaimNo, bCostingsAuthorization As Boolean
Public bLoanCHECKED, bsaveCLAIMREG, bPrintProductionReport, bPrintProductionListing, bReinstatePolicy As Boolean
Public bApproveREINSTATEMENT, bAuthorizeREINSTATEMENT, bloadSOURCE, bloadBRANCH, bloadUNIT, bPaidup, bApprovePAIDUP, bAuthorizePAIDUP, bMedicalRequisition As Boolean
Public bapproveRECORD, bPremiumREFUND, bSuspenseREFUND, bDepositREFUND, bgenerateFOCUS, bproposal, bPolicy, editRECORD, bauthorizeVOUCHER, bApproveVOUCHER, bCostingsApproval, authorizecheque, bAuthorizationinvoice, bpreprareVOUCHER As Boolean
Public strQRE As Variant
Public rsFind As ADODB.Recordset, Edit, bmakePAYMENT, breversePAYMENT, bissueCHECKS, bCorrectCommission  As Boolean
Public bunderwritePLAN, bunderwriteRIDER, bsaveRIDER, bsavePLAN, bupdateSTATUS, bAmountOff, RatesVoucher, bAmountOn As Boolean
Public bSurrenderAPPROVAL, bSurrenderAUTHORIZATION, bLapseAPPROVAL, blapseAUTHORIZATION As Boolean
Public bEndorseSumAssured, bendorseALL, bEndorsePaymentMethod, bendorseDOC, bendorseProduct, bEndorseTerm, bEndorsePaymentPeriod, bEndorsePAYMENTMODE, bEndorseAge, bEndorseRider As Boolean
