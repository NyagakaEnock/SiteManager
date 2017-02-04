Attribute VB_Name = "modCONTROLS"
Public NewSQL As String
Public Sub computeVOUCHERTOTAL()
On Error GoTo err
       With frmODASMVoucher
       
                Set rsSAVE = New ADODB.Recordset
                strSQL = "SELECT sum(AmountPaid) as totals from ODASMVoucherItem where VoucherNo = '" & frmODASMVoucher.txtVoucherNo.Text & "';"
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        .txtVoucherAmount.Text = 0
                ElseIf IsNull(rsSAVE!TOTALS) = True Then
                        .txtVoucherAmount.Text = .txtAmountPaid.Text
                Else:
                        .txtVoucherAmount.Text = FormatNumber(rsSAVE!TOTALS)
                End If
            End With
Exit Sub
err:
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
        rsSAVE.CancelUpdate
        rsSAVE.Requery
    Else
        UpdateErrorMessage
    End If
End Sub

Public Sub updateCostCenter()
On Error GoTo err
            
            Set rsCOSTCENTER = New ADODB.Recordset
            rsCOSTCENTER.Open "SELECT * FROM ODASPCostCentre C, ODASMVoucher V WHERE V.CostCenter = C.CostCentre and V.VoucherNo = '" & frmODASMOperation.txtApplicationNo & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsCOSTCENTER.EOF Or rsCOSTCENTER.BOF Then Exit Sub
            
            If rsCOSTCENTER!Materials = "Y" Then
            ElseIf rsCOSTCENTER!Machinery = "Y" Then
            ElseIf rsCOSTCENTER!Rent = "Y" Then
                    saveINSTALLMENTISSUED

            ElseIf rsCOSTCENTER!Rate = "Y" Then
            ElseIf rsCOSTCENTER!AdminCost = "Y" Then
            ElseIf rsCOSTCENTER!ManPower = "Y" Then
            End If

Exit Sub
err:
    ErrorMessage
End Sub

Public Sub saveINSTALLMENTISSUED()
On Error GoTo err

    With Screen.ActiveForm
    
        Set rsNewRecord = New Recordset
        strCONTROL = "SELECT * from ODASMInstallment I where I.VoucherNo = '" & frmODASMOperation.txtApplicationNo & " ' ; "
        rsNewRecord.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        If rsNewRecord.EOF Or rsNewRecord.BOF Then Exit Sub
    
                
        'Sum the Total Amount Paid from the Installment selected to be
        ' paid by the requisitions
        
        Dim rsINSTALL As ADODB.Recordset, strINSTALL As String, numAmountDue, numBALANCE As Double
        numAmountDue = 0
        
        Set rsINSTALL = New Recordset
        strINSTALL = "SELECT sum(TotalRent) as Totals from ODASMInstallment I where I.VoucherNo = '" & frmODASMOperation.txtApplicationNo & " ' ; "
        rsINSTALL.Open strINSTALL, cnCOMMON, adOpenKeyset, adLockOptimistic

        If rsINSTALL.EOF Or rsINSTALL.BOF Then
                numAmountDue = 0
        ElseIf IsNull(rsINSTALL!TOTALS) = True Then
                numAmountDue = 0
        Else: numAmountDue = CDbl(rsINSTALL!TOTALS)
        End If
                
        numBALANCE = numAmountDue
        
        Do While Not rsNewRecord.EOF

                Set rsSAVE = New Recordset
                strSQL = "SELECT * from ODASMInstallment where Installment = '" & rsNewRecord!Installment & " ' and ContractNo = '" & rsNewRecord!ContractNo & " ' and ContractYear = '" & rsNewRecord!ContractYear & " ' and  PaymentMode = '" & rsNewRecord!PaymentMode & " '; "
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                Else
                        rsSAVE!Status = "CHK-ISSUED"
                        rsSAVE!StatusDate = Date
                        rsSAVE!CurrentPeriod = CurrentPeriod
                        
                        If CDbl(numAmountDue) >= CDbl(rsSAVE!PaymentDue) Or PaymentDue = CDbl(rsSAVE!PaymentDue) Then
                                rsSAVE!PaymentFlag = "Y"
                                rsSAVE!PaymentDue = 0
                                rsSAVE!Balance = 0
                                rsSAVE!AmountPaid = CDbl(rsNewRecord!TotalRent)
                        Else:
                                If CDbl(rsNewRecord!TotalRent) <= numBALANCE Then
                                        rsSAVE!PaymentDue = 0
                                        rsSAVE!PaymentFlag = "Y"
                                        rsSAVE!Balance = 0
                                        rsSAVE!AmountPaid = CDbl(rsNewRecord!TotalRent)
                                        numBALANCE = numBALANCE - CDbl(rsNewRecord!TotalRent)
                                Else
                                        rsSAVE!PaymentDue = rsNewRecord!TotalRent - numBALANCE
                                        rsSAVE!PaymentFlag = "P"
                                        rsSAVE!Balance = rsSAVE!PaymentDue
                                        rsSAVE!AmountPaid = numBALANCE
                                        numBALANCE = 0
                                End If
                        End If
                        
                        rsSAVE!PaymentDate = Date
                                               
                        rsSAVE.Update
                End If
                rsNewRecord.MoveNext
        Loop
    End With
    
    Exit Sub
            
        rsNewRecord.Close
        strSQL = ""
    
   
err:
    UpdateErrorMessage
End Sub

Public Sub LoadDEFAULT()
On Error GoTo err

            strDEFAULT = "select * from ODASPdefault ;"
            Set rsDEFAULT = New ADODB.Recordset
            rsDEFAULT.Open strDEFAULT, cnCOMMON, adOpenKeyset, adLockOptimistic
Exit Sub
err:
    ErrorMessage
End Sub


Public Sub countVOUCHERITEMS()
On Error GoTo err
       With frmODASMVoucher
       
                Set rsSAVE = New ADODB.Recordset
                strSQL = "SELECT count(VoucherNo) as totals from ODASMVoucherItem where VoucherNo = '" & .txtVoucherNo.Text & "';"
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        .txtItems.Text = 0
                ElseIf IsNull(rsSAVE!TOTALS) = True Then
                        .txtItems.Text = 0
                Else
                        .txtItems.Text = FormatNumber(rsSAVE!TOTALS)
                End If
            End With
            

Exit Sub

err:
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
        rsSAVE.CancelUpdate
        rsSAVE.Requery
    Else
        UpdateErrorMessage
    End If
End Sub

Public Sub selectProductCode_GotFocus()
On Error GoTo err
    Set rsCONTROL = New Recordset
    strSQL = "SELECT * FROM ALISPProduct;"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    Screen.ActiveForm.cboProductCode.Clear
    With rsCONTROL
            Do Until .EOF
                    Screen.ActiveForm.cboProductCode.AddItem !ProductCode
                    .MoveNext
            Loop
    End With
rsCONTROL.Close
strSQL = ""
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub selectVATRATE_GotFocus()
On Error GoTo err

    Set rsCONTROL = New Recordset
    
    strSQL = "SELECT * FROM ODASPVAT;"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Screen.ActiveForm.cboVATRate.Clear

    With rsCONTROL
            Do Until .EOF
                    Screen.ActiveForm.cboVATRate.AddItem !VATDescription
                    .MoveNext
            Loop
    
    End With

rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub selectVATRate_LostFocus()
On Error GoTo err

    Set rsCONTROL = New ADODB.Recordset
    
    strSQL = "SELECT * FROM ODASPVAT Where VATDescription = '" & Screen.ActiveForm.cboVATRate.Text & "';"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsCONTROL
            If .BOF Or .EOF Then Exit Sub
                Screen.ActiveForm.cboVATRate.Text = !VATRate
    End With

rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub ClearListview2()
On Error GoTo err

 Dim j, i As Integer
       
        j = Screen.ActiveForm.ListView2.ListItems.Count
            
        For i = 1 To j
                Screen.ActiveForm.ListView2.ListItems(i).Checked = False
        Next i
Exit Sub
err:
ErrorMessage
    
End Sub

Public Sub SaveNewRecord()
On Error GoTo err

With Screen.ActiveForm
    Set rsNewRecord = New ADODB.Recordset
    rsNewRecord.Open NewSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    Set rsNewRecord = Nothing
End With

Exit Sub

err:
    UpdateErrorMessage
End Sub
Public Sub ClearListview1()
On Error GoTo err

 Dim j, i As Integer
       
                j = Screen.ActiveForm.ListView1.ListItems.Count
            
        For i = 1 To j
                Screen.ActiveForm.ListView1.ListItems(i).Checked = False
        Next i
Exit Sub
err:
ErrorMessage
    
End Sub

Public Sub clearALLRECORD()
On Error GoTo err
Dim i
    For Each i In Screen.ActiveForm
        If TypeOf i Is TextBox Then
            i.Text = Empty
        End If
        If TypeOf i Is ComboBox Then
            i.Clear
        End If
            
        If TypeOf i Is DTPicker Then
                i.Value = Date
        End If
        
        If TypeOf i Is CheckBox Then
                i.Value = 0
        End If
        
        If TypeOf i Is OptionButton Then
                i.Value = 0
        End If
        
    Next i

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub enableALLRECORD()
On Error GoTo err

Dim i
    For Each i In Screen.ActiveForm
            If TypeOf i Is TextBox Then
                i.Locked = False
            End If
            
            If TypeOf i Is ComboBox Then
                i.Locked = False
            End If
            
            If TypeOf i Is VScrollBar Then
                i.Enabled = True
            End If
            
            If TypeOf i Is DTPicker Then
                i.Enabled = True
            End If
            
            If TypeOf i Is UpDown Then
                i.Enabled = True
            End If
            
            If TypeOf i Is CheckBox Then
                i.Enabled = True
            End If
            
            If TypeOf i Is OptionButton Then
                i.Enabled = True
            End If
            
            If TypeOf i Is DTPicker Then
                i.Enabled = True
            End If
            
    Next i

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub searchMyRecord()
On Error GoTo err
    bsearchRECORD = True
    disableALLRECORD
    enableButtons
    Screen.ActiveForm.cmdSearch.Enabled = False
    Screen.ActiveForm.cmdEdit.Enabled = False
    Screen.ActiveForm.cmdPrint.Enabled = False

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub EditMyRecord()
On Error GoTo err
        beditRECORD = True
        enableALLRECORD
        disableButtons
        Screen.ActiveForm.cmdEdit.Enabled = False
        Screen.ActiveForm.cmdSearch.Enabled = False
        Screen.ActiveForm.cmdPrint.Enabled = False
        Screen.ActiveForm.cmdAddNew.Enabled = False
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub cancelCMD()
        enableButtons
        clearALLRECORD
        disableALLRECORD
End Sub

Public Sub addCMD()
        clearALLRECORD
        enableALLRECORD
        disableButtons
End Sub

Public Sub disableALLRECORD()
On Error GoTo err

Dim i
    For Each i In Screen.ActiveForm
    
            If TypeOf i Is TextBox Then
                i.Locked = True
            End If
            
            If TypeOf i Is ComboBox Then
                i.Locked = True
            End If
            
            If TypeOf i Is VScrollBar Then
                i.Enabled = False
            End If
            
            If TypeOf i Is DTPicker Then
                i.Enabled = False
            End If
            
            If TypeOf i Is UpDown Then
                i.Enabled = False
            End If
            
            If TypeOf i Is CheckBox Then
                i.Enabled = False
            End If
            
            If TypeOf i Is OptionButton Then
                i.Enabled = False
            End If
            
            If TypeOf i Is DTPicker Then
                i.Enabled = False
            End If
            
    Next i
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub disableButtonsTAB1()
On Error GoTo err

With Screen.ActiveForm
        .cmdUpdateTab1.Enabled = True
        .cmdAddNewTab1.Enabled = False
        .cmdSearchTab1.Enabled = False
        .cmdEditTab1.Enabled = False
        .cmdDeleteTab1.Enabled = False
        .cmdCancelTab1.Enabled = True
        .cmdPrintTab1.Enabled = False
End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub disableButtonsTAB2()
On Error GoTo err

With Screen.ActiveForm
        .cmdUpdateTAB2.Enabled = True
        .cmdAddNewTAB2.Enabled = False
        .cmdsearchTAB2.Enabled = False
        .cmdEditTAB2.Enabled = False
        .cmdDeleteTAB2.Enabled = False
        .cmdCancelTAB2.Enabled = True
End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub disableButtonsTAB3()
On Error GoTo err

With Screen.ActiveForm
        .cmdUpdateTAB3.Enabled = True
        .cmdAddNewTAB3.Enabled = False
        .cmdSearchTAB3.Enabled = False
        .cmdEditTAB3.Enabled = False
        .cmdDeleteTAB3.Enabled = False
        .cmdcancelTAB3.Enabled = True
End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub disableButtons()
On Error GoTo err

With Screen.ActiveForm
        .cmdUpdate.Enabled = True
        .cmdAddNew.Enabled = False
        .cmdDelete.Enabled = False
        .cmdEdit.Enabled = False
        .cmdSearch.Enabled = False
        .cmdCancel.Enabled = True
End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub viewButtons()
On Error GoTo err

    With Screen.ActiveForm
            .cmdUpdate.Enabled = False
            .cmdAddNew.Enabled = False
            .cmdSearch.Enabled = False
            .cmdEdit.Enabled = False
            .cmdDelete.Enabled = False
            .cmdCancel.Enabled = False
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub enableButtonsExtra()
On Error GoTo err

    With Screen.ActiveForm
        .cmdRiders.Enabled = True
        .cmdPlan.Enabled = True
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub disableButtonsExtra()
On Error GoTo err

    With Screen.ActiveForm
        .cmdRiders.Enabled = False
        .cmdPlan.Enabled = False
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub enableButtons()
On Error GoTo err

    With Screen.ActiveForm
            .cmdUpdate.Enabled = False
            .cmdAddNew.Enabled = True
            .cmdSearch.Enabled = True
            .cmdEdit.Enabled = True
            .cmdDelete.Enabled = True
            .cmdCancel.Enabled = True
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub enableSButtons()
On Error GoTo err

    With Screen.ActiveForm
            .cmdUpdate.Enabled = False
            .cmdAddNew.Enabled = True
            .cmdCancel.Enabled = True
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub disableSButtons()
On Error GoTo err

    With Screen.ActiveForm
            .cmdUpdate.Enabled = True
            .cmdAddNew.Enabled = False
            .cmdCancel.Enabled = True
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub enableButtonsTAB1()
On Error GoTo err

    With Screen.ActiveForm
            .cmdUpdateTab1.Enabled = False
            .cmdAddNewTab1.Enabled = True
            .cmdSearchTab1.Enabled = True
            .cmdEditTab1.Enabled = True
            .cmdDeleteTab1.Enabled = True
            .cmdCancelTab1.Enabled = True
            .cmdPrintTab1.Enabled = True
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub enableButtonsTAB2()
On Error GoTo err

    With Screen.ActiveForm
            .cmdUpdateTAB2.Enabled = False
            .cmdAddNewTAB2.Enabled = True
            .cmdsearchTAB2.Enabled = True
            .cmdEditTAB2.Enabled = True
            .cmdDeleteTAB2.Enabled = True
            .cmdCancelTAB2.Enabled = True
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub enableButtonsTAB3()
On Error GoTo err

    With Screen.ActiveForm
            .cmdUpdateTAB3.Enabled = False
            .cmdAddNewTAB3.Enabled = True
            .cmdSearchTAB3.Enabled = True
            .cmdEditTAB3.Enabled = True
            .cmdDeleteTAB3.Enabled = True
            .cmdcancelTAB3.Enabled = True
            .cmdPrintTAB3.Enabled = True
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub ClaimNoGOTFOCUS()
On Error GoTo err
    
    Dim rsCLAIM As ADODB.Recordset
    Set rsCLAIM = New ADODB.Recordset
    
    If bsendDOCUMENTS = True Then
            strSQL = "SELECT * FROM ODASMInvoice Where Approved = 'Y' and (ODASMInvoice.RequirementSENT <> 'Y' );"
    ElseIf breceiveDOCUMENTS = True Then
            strSQL = "SELECT * FROM ODASMInvoice Where ODASMInvoice.RequirementSENT = 'Y' and ODASMInvoice.RequirementReceived <> 'Y';"
    End If
    
    rsCLAIM.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Screen.ActiveForm.ActiveControl.Clear
    
    With rsCLAIM
            Do Until .EOF
                    Screen.ActiveForm.ActiveControl.AddItem !claimno
                    .MoveNext
            Loop
    
    End With

    rsCLAIM.Close
    strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub selectRECEIPTTYPEGOTFOCUS()
On Error GoTo err
    
    Dim rsPAY As ADODB.Recordset, strPAY As String
    Set rsPAY = New Recordset
    
    strPAY = "SELECT * FROM ALISPReceipt;"
    rsPAY.Open strPAY, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Screen.ActiveForm.cboReceiptType.Clear

    With rsPAY
            Do Until .EOF
                    Screen.ActiveForm.cboReceiptType.AddItem !Description
                    .MoveNext
            Loop
    
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub selectRECEIPTTYPELOSTFOCUS()
On Error GoTo err

        Set rsRCPT = New Recordset
        
        rsRCPT.Open "SELECT * FROM ALISPReceipt WHERE Description= '" & Screen.ActiveForm.cboReceiptType.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsRCPT
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboReceiptType.Text = !ReceiptType
                Screen.ActiveForm.txtAccountingPeriod = CurrentPeriod
                Screen.ActiveForm.txtRefundDate.Text = Date
        End With
 Exit Sub

err:
        ErrorMessage

End Sub

Public Sub Form_Unload1(cancel As Integer)
On Error GoTo err
    If bunEXITform = True Then
        cancel = True
        MsgBox "Data entry in progress. Click Refresh to Cancel", vbCritical
    Else
        cancel = False
    End If
Exit Sub

err:
    ErrorMessage
End Sub


Public Sub loadEMPLOYER()
On Error GoTo err

        Dim rsEMPLOYER   As ADODB.Recordset, strEMPLOYER As String
        Set rsEMPLOYER = New Recordset
            
        strEMPLOYER = "SELECT * FROM ODASPAccount WHERE AccountNo = '" & Screen.ActiveForm.txtEmployer & "';"
        rsEMPLOYER.Open strEMPLOYER, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsEMPLOYER
                If .BOF Or .EOF = True Then
                        MsgBox "The Employer Does not Exist in the Database", vbOKOnly
                        Exit Sub
                End If
                
                Screen.ActiveForm.txtEmployer = !CompanyName & ""
                Screen.ActiveForm.txtEmployerCommission = !AccountCommission & ""

        End With

rsEMPLOYER.Close
strEMPLOYER = ""

Exit Sub

err:

If err.Number = 91 Then Exit Sub
If err.Number = 13 Then Resume Next
ErrorMessage

End Sub
Public Sub loadAccountNo()
On Error GoTo err

        Dim rsEMPLOYER   As ADODB.Recordset, strEMPLOYER As String
        Set rsEMPLOYER = New Recordset
            
        strEMPLOYER = "SELECT * FROM ODASPAccount WHERE AccountNo = '" & Screen.ActiveForm.txtEmployer & "';"
        rsEMPLOYER.Open strEMPLOYER, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsEMPLOYER
                If .BOF Or .EOF = True Then
                        MsgBox "The Employer Does not Exist in the Database", vbOKOnly
                        Exit Sub
                End If
                
                Screen.ActiveForm.txtCompanyName = !CompanyName & ""

        End With

rsEMPLOYER.Close
strEMPLOYER = ""

Exit Sub

err:

If err.Number = 91 Then Exit Sub
If err.Number = 13 Then Resume Next
ErrorMessage

End Sub

Public Sub bankNoGotFocus()
On Error GoTo err

          Set rsCONTROL = New ADODB.Recordset
          rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

          Screen.ActiveForm.cboBankNo.Clear

          With rsCONTROL
                  If .EOF Or .BOF Then Exit Sub
                  Do Until .EOF
                      Screen.ActiveForm.cboBankNo.AddItem !Details & ""
                      .MoveNext
                  Loop
          End With
        

rsCONTROL.Close
strSQL = ""
         
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub BankNoLostFocus()
On Error GoTo err
        
        Set rsCONTROL = New ADODB.Recordset
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboBankNo.Text = !BankNo
        End With
        
        
rsCONTROL.Close
strSQL = ""

Exit Sub

err:
        ErrorMessage

End Sub

Public Sub loadBANKDETAILS()
On Error GoTo err
        
        Set rsCONTROL = New ADODB.Recordset
        rsCONTROL.Open "SELECT * FROM ALISPBankAccount WHERE BankNo = '" & Screen.ActiveForm.cboBankNo.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboBankNo.Text = !BankNo
                Screen.ActiveForm.txtBankName.Text = !Details
                Screen.ActiveForm.txtBankAccountNo = !AccountNo
        End With
        
        
rsCONTROL.Close

Exit Sub

err:
        ErrorMessage

End Sub


