VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASMInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub loadREQUISTION()
On Error GoTo err
    With frmODASMReceiveinvoice
    
            Dim rsLOAD As ADODB.Recordset
            Set rsLOAD = New ADODB.Recordset
            
            strSQL = "SELECT * FROM ODASMRequisition Where RequisitionNo = '" & .txtRequisitionNo.Text & "' ;"
            rsLOAD.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsLOAD.EOF Or rsLOAD.BOF Then Exit Sub
            
            .txtInvoiceAmountInclusive.Text = FormatNumber(rsLOAD!PriceInclusive)
            .txtInvoiceVATAmount.Text = FormatNumber(rsLOAD!VATAmount)
            .txtInvoiceAmountExclusive.Text = FormatNumber(rsLOAD!PriceExclusive)
            .txtVATRate.Text = 0
            .txtInvoiceDate.Text = Date
            .DTPickerInvoiceDate.Value = Date
            .txtRequisitionDate.Text = rsLOAD!RequisitionDate & ""
            .txtVATAmount.Text = FormatNumber(rsLOAD!VATAmount)
            .txtPriceExclusive.Text = FormatNumber(rsLOAD!PriceExclusive)
            .txtPriceInclusive.Text = FormatNumber(rsLOAD!PriceInclusive)
            
        
    End With
            
rsLOAD.Close
strSQL = ""
         
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub updateREQUISITION()
On Error GoTo err
    With Screen.ActiveForm
    
            Set rsSAVE = New ADODB.Recordset
            
            strSQL = "SELECT * FROM ODASMRequisition Where LPONo = '" & .txtLPONo.Text & "' ;"
            rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
            
            rsSAVE!InvoiceBalance = CDbl(.txtLPOBalance)
            
            If rsSAVE!InvoiceBalance = 0 Then
                    rsSAVE!InvoiceStatus = "INVOICE RECVD"
                    rsSAVE!InvoiceStatusDate = Date
                    rsSAVE!InvoiceReceived = "Y"
            Else: rsSAVE!InvoiceStatus = "INVOICE - PENDING"
                    rsSAVE!InvoiceStatus = "INVOICE - PENDING"
                    rsSAVE!InvoiceReceived = "N"
            End If
            
            rsSAVE.Update
            rsSAVE.Requery
    End With
            
rsSAVE.Close
strSQL = ""
         
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub calculateLPOBalance()
On Error GoTo err
    With frmODASMReceiveinvoice
            
            If .txtPriceInclusive.Text = Empty Then Exit Sub
            Set rsCONTROL = New ADODB.Recordset
            
            strSQL = "SELECT sum(PriceInclusive) as Totals FROM ODASMinvoice Where LPONo = '" & .txtLPONo.Text & "' ;"
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsCONTROL.EOF Or rsCONTROL.BOF Then
                    .txtLPOBalance.Text = FormatNumber(CDbl(.txtPriceInclusive.Text))
            ElseIf IsNull(rsCONTROL!TOTALS) = True Then
                    .txtLPOBalance.Text = FormatNumber(CDbl(.txtPriceInclusive.Text))
            Else
                    .txtLPOBalance.Text = FormatNumber(CDbl(.txtPriceInclusive.Text) - CDbl(rsCONTROL!TOTALS))
            End If
    
            .txtInvoiceAmountInclusive.Text = FormatNumber(.txtLPOBalance)
    End With
            
rsCONTROL.Close
strSQL = ""
         
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub enableRECORD()
On Error GoTo err
    With frmODASMReceiveinvoice
        .txtInvoiceNo.Locked = False
        .txtInvoiceAmountExclusive.Locked = False
        .txtInvoiceAmountInclusive.Locked = False
        .DTPickerInvoiceDate.Enabled = True
        .txtVATRate.Locked = False
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub clearRECORD()
On Error GoTo err
    With frmODASMReceiveinvoice
        .txtInvoiceNo.Text = Empty
        .txtInvoiceAmountExclusive.Text = 0
        .txtInvoiceAmountInclusive.Text = 0
        .DTPickerInvoiceDate.Value = Date
        .txtInvoiceDate.Text = Empty
        .txtLPOBalance.Text = Empty
    End With
Exit Sub

err:
    ErrorMessage
End Sub


Public Sub updateRECORD()
        validateRECORD
        If bsaveRECORD = True Then
                saveINVOICE
                calculateLPOBalance
                updateREQUISITION
                showLPOINVOICES
                NewRecord = False
                beditRECORD = False
        End If

Exit Sub

End Sub
Public Sub validateRECORD()
On Error GoTo err
        With frmODASMReceiveinvoice
            If .txtAccountNo.Text = Empty Then
                    MsgBox "The Supply Code is Required"
                .txtAccountNo.SetFocus
            ElseIf .txtInvoiceNo.Text = Empty Then
                    MsgBox "Invoice number MUST be entered..."
                .txtInvoiceNo.SetFocus
            ElseIf CDbl(.txtLPOBalance.Text) > CDbl(.txtPriceInclusive.Text) Then
                MsgBox "Cannot Over Pay a LPO"
                .txtLPOBalance.SetFocus
            
            ElseIf CDbl(.txtInvoiceAmountExclusive.Text) <= 0 Then
                MsgBox "The Invoice Amount Cannot be Zero"
                .txtInvoiceAmountExclusive.SetFocus
            
            ElseIf CDbl(.txtInvoiceAmountInclusive.Text) <= 0 Then
                MsgBox "The Invoice Amount Cannot be Zero"
                .txtInvoiceAmountInclusive.SetFocus
                
            ElseIf CDbl(.txtInvoiceAmountInclusive.Text) < CDbl(.txtInvoiceAmountExclusive) Then
                MsgBox "The Amount inclusive cannot be Less than the Amount Exclusive"
                .txtInvoiceAmountInclusive.SetFocus
                
            ElseIf CDbl(.txtInvoiceAmountInclusive.Text) > CDbl(.txtPriceInclusive) Then
                MsgBox "The Invoice Amount Cannot be Greater Than The LPO Amount"
                .txtInvoiceAmountInclusive.SetFocus
            
            ElseIf CDbl(.txtVATAmount.Text) < 0 Then
                MsgBox "The VAT Amount Cannot be Zero"
                .txtVATAmount.SetFocus
            
            ElseIf CDbl(.txtVATRate.Text) <= 0 Then
                MsgBox "The VAT Rate Entered is Wrong"
                .txtVATRate.SetFocus
                
            ElseIf .txtInvoiceDate.Text <= Empty Then
                    MsgBox "The Invoice Date Cannot be Left Blank"
                    .txtInvoiceDate.SetFocus
            
            ElseIf DateDiff("D", .txtInvoiceDate, Date) < 0 Then
                    MsgBox "The Invoice Date cannot be in the Future"
                    .txtInvoiceDate.SetFocus
            Else
                 bsaveRECORD = True
            End If
        End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub saveINVOICE()
On Error GoTo err
       With frmODASMReceiveinvoice
            
                strSQL = "select * from ODASMInvoice Where InvoiceNo = '" & .txtInvoiceNo & "' and LPONo = '" & .txtLPONo.Text & "';"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        rsSAVE.AddNew
                        rsSAVE!InvoiceNo = Trim(.txtInvoiceNo.Text)
                        rsSAVE!InvoiceDate = Date
                        rsSAVE!LPONo = .txtLPONo.Text
'                        rsSAVE!JobCardNo = Trim(.txtJobBriefNo.Text)
                        rsSAVE!Preparedby = CurrentUserName
                        rsSAVE!dateprepared = Date
                        rsSAVE!Prepared = "Y"
                        rsSAVE!Approved = "N"
                        rsSAVE!Authorized = "N"
                        rsSAVE!Paid = "N"
                        rsSAVE!Requisitioned = "N"
                        rsSAVE!Status = "INVOICE-RECEIVED"
                End If
                
                rsSAVE!AccountNo = .txtAccountNo.Text
                rsSAVE!CurrentPeriod = .txtCurrentPeriod.Text
                rsSAVE!PriceInclusive = CDbl(.txtInvoiceAmountInclusive.Text)
                rsSAVE!InvoiceBalance = CDbl(rsSAVE!PriceInclusive)
                rsSAVE!VATAmount = CDbl(.txtInvoiceVATAmount.Text)
                rsSAVE!VATRate = CDbl(.txtVATRate.Text)
                rsSAVE!PriceExclusive = CDbl(.txtInvoiceAmountExclusive.Text)
                rsSAVE.Update
                rsSAVE.Requery
                
                Set rsNewRecord = Nothing
        
        End With
Exit Sub

err:
    ErrorMessage
End Sub


Public Sub loadRECORD()
On Error GoTo err
    With frmODASMReceiveinvoice
        
        Set rsCONTROL = New ADODB.Recordset
        strSQL = "SELECT * FROM ODASMLPO L, ODASPAccount AC  WHERE L.AccountNo = AC.AccountNo and L.LPONo = '" & .txtLPONo.Text & "' ; "
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub

                .txtLPOBalance.Text = 0
                .txtCurrentPeriod.Text = CurrentPeriod
                .txtAccountNo.Text = rsCONTROL!AccountNo
                .txtCompanyName.Text = rsCONTROL!CompanyName
                .txtDescriptionOfOrder.Text = rsCONTROL!Remarks
                .txtTotalCost.Text = rsCONTROL!PriceInclusive & ""
                
    
    End With

Exit Sub
err:
    ErrorMessage
End Sub

