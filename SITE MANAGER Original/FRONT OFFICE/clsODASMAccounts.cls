VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASMAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub calculateVAT()
On Error GoTo err
        With frmODASMAccounts
                .txtJBVATAmount.Text = FormatNumber(CDbl(.cboVATRate.Text) / 100 * CDbl(.txtJBPriceExclusive))
                .txtJBPriceInclusive.Text = FormatNumber(CDbl(.txtJBVATAmount.Text) + CDbl(.txtJBPriceExclusive.Text))
        End With
        
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub loadRECORD()
On Error GoTo err
    With frmODASMAccounts
        
        Set rsCONTROL = New ADODB.Recordset
        strSQL = "SELECT * FROM ODASMJobBrief JB, ODASPAccount AC,  ODASPPaymentMode PM WHERE JB.PaymentMode = PM.PaymentMode and JB.AccountNo = AC.AccountNo and JB.JobBriefNo = '" & .txtJobBriefNo.Text & "' ; "
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
                .txtInvoiceDate.Text = Date
                .txtCurrentPeriod.Text = CurrentPeriod
                .txtDeposit.Text = 0
                .txtPriceInclusive.Text = 0
                .txtPriceExclusive.Text = 0
                .txtTotalCost.Text = 0
                .txtVATAmount.Text = 0
                .txtTotalCost.Text = 0
                .txtAmountQuoted.Text = 0
                .txtJBPriceExclusive.Text = 0
                
                .txtPaymentMethod.Text = rsCONTROL!PaymentMethod & ""
                .txtAccountNo.Text = rsCONTROL!Accountno
                .txtCompanyName.Text = rsCONTROL!CompanyName
                .txtDescriptionOfOrder.Text = rsCONTROL!descriptionOfOrder
                .txtProductCode.Text = rsCONTROL!ProductCode
                .txtDeposit.Text = Format(rsCONTROL!SuspenseAccount) & ""
                .txtTotalCost.Text = FormatNumber(rsCONTROL!TotalOverallCost)
                
                
                
                If IsNull(rsCONTROL!CommencementDate) = True Then
                        .txtCommencementDate.Text = Date
                Else: .txtCommencementDate = rsCONTROL!CommencementDate & ""
                End If
                
                .txtExpiryDate.Text = rsCONTROL!expirydate & ""
                .txtAmountQuoted.Text = FormatNumber(rsCONTROL!PriceExclusive)
                .txtInvoiceDescription.Text = "Invoice For JobBrief No - '" & frmODASMAccounts.txtJobBriefNo & "'"
    
    End With

Exit Sub
err:
    ErrorMessage
End Sub
Public Sub LoadDEFAULT()
On Error GoTo err
    With frmODASMAccounts
        
        Set rsCONTROL = New ADODB.Recordset
        strSQL = "SELECT * FROM ODASPDefault  "
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
        
        .cboVATRate.Text = rsCONTROL!VATRate
        
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub loadINSTALLMENT()
On Error GoTo err
    With frmODASMAccounts
        
        
        Set rsCONTROL = New ADODB.Recordset
        strSQL = "SELECT * FROM ODASMJobBriefInstallment WHERE InvoiceReference = '" & .txtInvoiceReference.Text & "' ; "
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
                
                
                .txtDueDate.Text = rsCONTROL!PaymentDueDate
                .txtJBPriceExclusive.Text = FormatNumber(rsCONTROL!Amount)
                .txtPriceExclusive.Text = FormatNumber(rsCONTROL!Amount)
    End With

Exit Sub
err:
    ErrorMessage
End Sub


Public Sub clearRECORD()
    With frmODASMAccounts
            .txtAmountQuoted.Text = 0
            .txtPriceExclusive.Text = 0
            .txtAmountQuoted.Text = 0
            .txtDeposit.Text = 0
            .txtDescriptionOfOrder.Text = 0
            .txtJobBriefNo.Text = Empty
            .txtPaymentMethod.Text = Empty
            .txtProductCode.Text = Empty
            
    End With
End Sub

Public Sub enableRECORD()
    With frmODASMAccounts
        .txtCommencementDate.Locked = False
    
    
    End With
End Sub

Public Sub calculateTOTALRECEIPTS()
On Error GoTo err
   With frmODASMAccounts
   
            Set rsCONTROL = New ADODB.Recordset
            
            strSQL = "SELECT sum(TransactionAmount) as totals FROM ALISMReceiptDetails Where DocumentNo = '" & .txtJobBriefNo & "';"
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
            
            If IsNull(rsCONTROL!TOTALS) = True Then
                        .txtDeposit = 0
            Else: .txtDeposit.Text = FormatNumber(rsCONTROL!TOTALS)
            
            End If
            
    End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub calculateInvoicesSend()
On Error GoTo err
   With frmODASMAccounts
   
            Set rsCONTROL = New ADODB.Recordset
            
            strSQL = "SELECT sum(PriceExclusive) as totals FROM ODASMInvoiceSENT Where JobBriefNo = '" & .txtJobBriefNo & "';"
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
            
'            If IsNull(rsCONTROL!TOTALS) = True Then
'                        .txtPriceExclusive = 0
'            Else: .txtPriceExclusive.Text = FormatNumber(rsCONTROL!TOTALS)
'
'            End If
'
'
'            If .txtPriceExclusive.Text = Empty Then .txtPriceExclusive.Text = 0
'            .txtPriceExclusive.Text = FormatNumber(CDbl(.txtAmountQuoted.Text) - CDbl(.txtPriceExclusive.Text))

    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub saveINVOICE()
On Error GoTo err
       With frmODASMAccounts
            
                strSQL = "select * from ODASMInvoiceSent Where InvoiceNo = '" & .txtInvoiceNo & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        rsSAVE.AddNew
                        rsSAVE!InvoiceNo = Trim(.txtInvoiceNo.Text)
                        rsSAVE!InvoiceDate = Date
                        rsSAVE!JobBriefNo = Trim(.txtJobBriefNo.Text)
                        rsSAVE!Preparedby = CurrentUserName
                        rsSAVE!dateprepared = Date
                        rsSAVE!CompanyCode = "MAG"
                        rsSAVE!Prepared = "Y"
                        
                        If rsDEFAULT!AutoApproval = "Y" Then
                                rsSAVE!ApprovedBy = CurrentUserName
                                rsSAVE!DateApproved = Date
                                rsSAVE!Approved = "Y"
                                rsSAVE!Status = "INVOICE-APPROVED"
                        Else
                                rsSAVE!Approved = "N"
                                rsSAVE!Status = "INVOICE-CREATED"
                        End If
                        
                        rsSAVE!Despatched = "N"
                        rsSAVE!Approved = "N"
                        rsSAVE!Authorized = "N"
                        rsSAVE!Balance = 0
                End If
                
                rsSAVE!Accountno = .txtAccountNo.Text
                rsSAVE!CurrentPeriod = .txtCurrentPeriod.Text
                rsSAVE!InvoiceDescription = .txtInvoiceDescription.Text
                rsSAVE!PriceInclusive = CDbl(.txtPriceInclusive.Text)
                rsSAVE!VATAmount = CDbl(.txtVATAmount.Text)
                rsSAVE!PriceExclusive = CDbl(.txtPriceExclusive.Text)
                rsSAVE!ReceivedToDate = 0
                rsSAVE!remark = .txtRemark.Text
                rsSAVE!Balance = CDbl(.txtPriceInclusive)
                rsSAVE.Update
                rsSAVE.Requery
                
                Set rsNewRecord = Nothing
        
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub updatePRICEinclusive()
On Error GoTo err

       With frmODASMAccounts
            
                strSQL = "select sum(PriceInclusive) as totals from ODASMInvoiceItemsSent Where InvoiceNo = '" & .txtInvoiceNo & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        .txtPriceInclusive.Text = FormatNumber(.txtJBPriceInclusive.Text)
                ElseIf IsNull(rsSAVE!TOTALS) = True Then
                        .txtPriceInclusive.Text = FormatNumber(.txtJBPriceInclusive.Text)
                Else
                        .txtPriceInclusive.Text = FormatNumber(CDbl(rsSAVE!TOTALS))
                End If
        End With

rsSAVE.Close

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub generateITEMNO()
On Error GoTo err

       With frmODASMAccounts
            
                strSQL = "select * from ODASMInvoiceItemsSent Where InvoiceNo = '" & .txtInvoiceNo & "';"
                Set rsCONTROL = New ADODB.Recordset
                rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsCONTROL.EOF Or rsCONTROL.BOF Then
                        .txtItemNo.Text = Trim(.txtInvoiceNo) + "-1"
                ElseIf IsNull(rsCONTROL!TOTALS) = True Then
                        .txtItemNo.Text = Trim(.txtInvoiceNo) + "-1"
                Else
                        .txtItemNo.Text = Trim(.txtInvoiceNo) + "-" + Trim(Str(rsCONTROL.RecordCount + 1))
                End If
                
        End With

rsCONTROL.Close

Exit Sub

err:
    ErrorMessage
End Sub


Public Sub updateVAT()
On Error GoTo err

       With frmODASMAccounts
            
                strSQL = "select sum(VATAmount) as totals from ODASMInvoiceItemsSent Where InvoiceNo = '" & .txtInvoiceNo & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        .txtVATAmount.Text = FormatNumber(.txtJBVATAmount)
                ElseIf IsNull(rsSAVE!TOTALS) = True Then
                        .txtVATAmount.Text = FormatNumber(.txtJBVATAmount)
                Else
                        .txtVATAmount.Text = FormatNumber(CDbl(rsSAVE!TOTALS))
                End If
        End With

rsSAVE.Close

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub updatePRICEExclusive()
On Error GoTo err

       With frmODASMAccounts
            
                strSQL = "select sum(PriceExclusive) as totals from ODASMInvoiceItemsSent Where InvoiceNo = '" & .txtInvoiceNo & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        .txtPriceExclusive.Text = FormatNumber(.txtJBPriceExclusive)
                ElseIf IsNull(rsSAVE!TOTALS) = True Then
                        .txtPriceExclusive.Text = FormatNumber(.txtJBPriceExclusive)
                Else
                        .txtPriceExclusive.Text = FormatNumber(CDbl(rsSAVE!TOTALS))
                End If
        End With

rsSAVE.Close

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub updateRECORD()
        validateRECORD
        If bsaveRECORD = True Then
                GenerateInvoiceNo
                saveINVOICEITEMS
                updatePRICEExclusive
                updateVAT
                updatePRICEinclusive
                saveINVOICE
                updateJOBBRIEF
                updateJOBBRIEFMAIN
                saveUpdateStatement
                NewRecord = False
                beditRECORD = False
        End If
        showBRIEFINACCOUNT
        showINVOICEitems
Exit Sub
End Sub
Public Sub updateJOBBRIEF()
On Error GoTo err
       With frmODASMAccounts

                strSQL = "select * from ODASMJobBriefInstallment Where InvoiceReference = '" & .txtInvoiceReference & "' ;"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
                    rsSAVE!Invoiced = "Y"
                    rsSAVE!InvoiceDate = Date
                    rsSAVE!InvoiceNo = .txtInvoiceNo.Text
                    rsSAVE.Update
                    rsSAVE.Requery
                Set rsSAVE = Nothing
                strSQL = Empty
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub updateJOBBRIEFMAIN()
On Error GoTo err
    
       With frmODASMAccounts

                strSQL = "select * from ODASMJobBrief Where JobBriefNo = '" & .txtJobBriefNo & "' ;"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
                
                rsSAVE!Invoiced = "Y"
                rsSAVE!InvoiceDate = Date
                rsSAVE!InvoiceNo = .txtInvoiceNo.Text
                
                rsSAVE.Update
                rsSAVE.Requery
                
                
                Set rsSAVE = Nothing
                strSQL = Empty
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub updateStatement()
On Error GoTo err
    
       With frmODASMAccounts

                strSQL = "select * from ODASMCustomerStatement Where InvoiceReference = '" & .txtInvoiceReference & "' ;"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
                
                rsSAVE!Invoiced = "Y"
                rsSAVE!InvoiceDate = Date
                rsSAVE!InvoiceNo = .txtInvoiceNo.Text
                
                rsSAVE.Update
                rsSAVE.Requery
                
                
                Set rsSAVE = Nothing
                strSQL = Empty
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub validateRECORD()
On Error GoTo err

        With frmODASMAccounts
            
            bsaveRECORD = False
            
            If .txtAccountNo.Text <= Empty Then
                    MsgBox "Kindly Load the Proper Accouunt"
                    .txtAccountNo.SetFocus
            
            ElseIf .txtInvoiceDescription.Text <= Empty Then
                    MsgBox "The invoice Description is MANDATORY"
                    .txtInvoiceDescription.SetFocus
                    
            
            ElseIf CDbl(.txtPriceExclusive) < 0 Then
                    MsgBox "The Amount Invoiced Must be Greater Than Zero"
                    .txtPriceExclusive.SetFocus
                
            ElseIf CDbl(.txtAmountQuoted.Text) <= 0 Then
                    MsgBox "The Amount Quoated Must be Greater Than Zero"
                    .txtAmountQuoted.SetFocus
                            
            ElseIf .txtDescriptionOfOrder.Text <= Empty Then
                    MsgBox "The Description of Order is Required"
                    .txtDescriptionOfOrder.SetFocus
            
            ElseIf .txtInvoiceDate.Text <= Empty Then
                    MsgBox "The invoice Date cannot be Blank"
                    .txtInvoiceDate.SetFocus
            
            ElseIf DateDiff("D", .txtInvoiceDate, Date) < 0 Then
                    MsgBox "The Invoice Date cannot be in the Future"
                    .txtInvoiceDate.SetFocus
            
            ElseIf .txtPaymentMethod.Text <= Empty Then
                    MsgBox "The Form of Payment is Required"
                    .txtPaymentMethod.SetFocus
            
            ElseIf CDbl(.txtPriceExclusive.Text) < 0 Then
                    MsgBox "The Price Exclusive of VAT Cannot be Less than Zero"
                    .txtPriceExclusive.SetFocus
            
            ElseIf CDbl(.txtPriceInclusive.Text) < 0 Then
                    MsgBox "The Price Inclusive of VAT is Required"
                    .txtPriceInclusive.SetFocus
            
            ElseIf .txtProductCode.Text <= Empty Then
                    MsgBox "The Product Code Cannot be Blank"
                    .txtProductCode.SetFocus
            Else
                    bsaveRECORD = True
            End If
        End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub saveINVOICEITEMS()
On Error GoTo err
       With frmODASMAccounts
            
                strSQL = "select * from ODASMInvoiceItemsSent Where ItemNo = '" & .txtItemNo & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        rsSAVE.AddNew
                        generateITEMNO
                        rsSAVE!ItemNo = .txtItemNo.Text
                        rsSAVE!InvoiceNo = .txtInvoiceNo.Text
                        rsSAVE!Preparedby = CurrentUserName
                        rsSAVE!dateprepared = Date
                        rsSAVE!Prepared = "Y"
                        rsSAVE!Status = "INVOICE-CREATED"
                End If
                
                rsSAVE!PriceInclusive = CDbl(.txtJBPriceInclusive.Text)
                If .txtNoOfItems.Text = "" Then
                   .txtNoOfItems.Text = "1"
                End If
                rsSAVE!PerItemAmount = CDbl(.txtJBPriceExclusive.Text) / (.txtNoOfItems.Text)
                rsSAVE!VATAmount = CDbl(.txtJBVATAmount.Text)
                rsSAVE!PriceExclusive = CDbl(.txtJBPriceExclusive.Text)
                rsSAVE.Update
                rsSAVE.Requery
                
                Set rsNewRecord = Nothing
        
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub saveUpdateStatement()
On Error GoTo err
       With frmODASMAccounts
            
                strSQL = "select * from ODASMCustomerStatement;" 'Where ItemNo = '" & .txtItemNo & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                'If rsSAVE.EOF Or rsSAVE.BOF Then
                rsSAVE.AddNew
                rsSAVE!Accountno = .txtAccountNo.Text
                rsSAVE!Reference = .txtInvoiceNo.Text
                rsSAVE!Transactionby = CurrentUserName
                rsSAVE!DebitAmount = CDbl(.txtJBPriceInclusive.Text)
                rsSAVE!CreditAmount = "00"
                rsSAVE!TransactionDate = Date
                rsSAVE!Balance = CDbl(.txtAmountQuoted) - (CDbl(.txtJBPriceInclusive.Text) * CDbl(.txtInstallmentNo))
                rsSAVE!Details = .txtInvoiceDescription.Text
                'rsSAVE! = CDbl(.txtJBPriceExclusive.Text)
               ' End If
                rsSAVE.Update
                rsSAVE.Requery
                
                Set rsNewRecord = Nothing
        
        End With
Exit Sub

err:
    ErrorMessage
End Sub


Public Sub GenerateInvoiceNo()
On Error GoTo err

        Set rsCONTROL = New Recordset
        
        strSQL = "SELECT * FROM ODASPLastNumbers WHERE AutoInvoiceNo = 'Y'"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
        With rsCONTROL
                If .BOF Or .EOF = True Then Exit Sub
                        Screen.ActiveForm.txtInvoiceNo.Text = !InvoiceNo & ""
                
                Select Case Len(Trim(frmODASMAccounts.txtInvoiceNo))
                        Case 1: frmODASMAccounts.txtInvoiceNo.Text = Trim(!InvoicePrefix) + "00000" + Trim(frmODASMAccounts.txtInvoiceNo)
                        Case 2: frmODASMAccounts.txtInvoiceNo.Text = Trim(!InvoicePrefix) + "0000" + Trim(frmODASMAccounts.txtInvoiceNo)
                        Case 3: frmODASMAccounts.txtInvoiceNo.Text = Trim(!InvoicePrefix) + "000" + Trim(frmODASMAccounts.txtInvoiceNo)
                        Case 4: frmODASMAccounts.txtInvoiceNo.Text = Trim(!InvoicePrefix) + "00" + Trim(frmODASMAccounts.txtInvoiceNo)
                        Case 5: frmODASMAccounts.txtInvoiceNo.Text = Trim(!InvoicePrefix) + "0" + Trim(frmODASMAccounts.txtInvoiceNo)
                        Case 6: frmODASMAccounts.txtInvoiceNo.Text = Trim(!InvoicePrefix) + Trim(frmODASMAccounts.txtInvoiceNo)
                End Select
                
                !InvoiceNo = !InvoiceNo + 1
                .Update
                .Requery

        End With
        
rsCONTROL.Close
strTRANS = ""

Exit Sub
err:
    ErrorMessage
End Sub

