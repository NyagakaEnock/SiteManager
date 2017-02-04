Attribute VB_Name = "modCONTROLS"

Public Sub updatePaymentFlag()
On Error GoTo err
With frmUpdatePaymentFlag

        Set rsCONTROL = New Recordset
        strSQL = "SELECT * FROM ODASMInstallment  ;"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
        .ProgressBar1.Value = 0: .ProgressBar1.Visible = True: .ProgressBar1.Min = 0: .ProgressBar1.Max = rsCONTROL.RecordCount
        
        Do While Not rsCONTROL.EOF
        
            Set rsSAVE = New Recordset
            strSAVE = "Select * From ODASMInstallment where InstallmentNo = '" & rsCONTROL!InstallmentNo & "'"
            rsSAVE.Open strSAVE, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsSAVE.EOF Or rsSAVE.BOF Then
            Else
        
                    If rsCONTROL!PaymentDue = 0 And rsCONTROL!AmountPaid = 0 Then
                            rsSAVE!PaymentFlag = "N"
                            rsSAVE.Update
                    ElseIf CDbl(rsCONTROL!PaymentDue) > 0 And CDbl(rsCONTROL!AmountPaid) > 0 And CDbl(rsCONTROL!AmountPaid) < CDbl(rsCONTROL!PaymentDue) Then
                            rsSAVE!PaymentFlag = "N"
                           rsSAVE.Update
                    Else
                         If CDbl(rsCONTROL!AmountPaid) = CDbl(rsCONTROL!PaymentDue) Then
                                     rsSAVE!PaymentFlag = "Y"
                                     rsSAVE.Update
                         End If
                    
                    End If
                    
                    'rsSAVE.Update
                    
                    Set rsSAVE = Nothing
                    strSAVE = Empty
                
            End If
             
             
            DoEvents
            .ProgressBar1.Value = .ProgressBar1.Value + 1
            rsCONTROL.MoveNext
            
        Loop
 
 .ProgressBar1.Visible = False
End With
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub SelectModeGotFocus()
On Error GoTo err

    Set rsCONTROL = New Recordset
    
    strSQL = "SELECT Distinct(PaymentMode) FROM ODASPLandRate where CurrentYear = '" & Screen.ActiveForm.txtCurrentYear.Text & "' and townCode = '" & Screen.ActiveForm.txtTownCode.Text & "' and mediaSize = '" & Screen.ActiveForm.txtMediaSize.Text & "';"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    
    Screen.ActiveForm.cboPaymentMode.Clear

    With rsCONTROL
            Do Until .EOF
                    Screen.ActiveForm.cboPaymentMode.AddItem !PaymentMode
                    .MoveNext
            Loop
    End With
        
rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub updatecurrentperiod()
Dim curYear, curmonth As String
On Error GoTo err

    Set rsCONTROL = New Recordset
    
    strSQL = "SELECT * FROM ODASMInstallment where paymentflag = 'N';"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    

     
    Do Until rsCONTROL.EOF
             curYear = Str(Year(rsCONTROL!PaymentDueDate))
             curmonth = Month(rsCONTROL!PaymentDueDate)
             
             If Val(curmonth) <= 9 Then
                    curmonth = "0" + Trim(curmonth)
             Else
             End If
            
            
            rsCONTROL!CurrentPeriod = Trim((curYear) + "/" + Trim(curmonth))
            rsCONTROL.Update
              
                 
            rsCONTROL.MoveNext
    Loop
        
rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub updateExpiryperiod()
Dim curYear, curmonth As String
On Error GoTo err

    Set rsCONTROL = New Recordset
    
    strSQL = "SELECT * FROM ODASPPlotMast;"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
       
    Do Until rsCONTROL.EOF
             curYear = Str(Year(rsCONTROL!expirydate))
             curmonth = Month(rsCONTROL!expirydate)
             
             If Val(curmonth) <= 9 Then
                    curmonth = "0" + Trim(curmonth)
             Else
             End If
            
            
            rsCONTROL!CurrentPeriod = Trim((curYear) + "/" + Trim(curmonth))
            rsCONTROL.Update
              
                 
            rsCONTROL.MoveNext
    Loop
        
rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub updateCurrentPayment()
Dim curYear, curmonth As String
On Error GoTo err

    Set rsCONTROL = New Recordset
    
    strSQL = "SELECT * FROM ODASMInstallment Where paymentflag= 'Y';"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    

     
    Do Until rsCONTROL.EOF
             curYear = Str(Year(rsCONTROL!PaymentDate))
             curmonth = Month(rsCONTROL!PaymentDate)
             
             If Val(curmonth) <= 9 Then
                    curmonth = "0" + Trim(curmonth)
             Else
             End If
            
            
            rsCONTROL!CurrentPeriod = Trim((curYear) + "/" + Trim(curmonth))
            rsCONTROL.Update
              
                 
            rsCONTROL.MoveNext
    Loop
        
rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub updateTargetPeriod()
Dim curYear, curmonth, curTargetPeriod As String
On Error GoTo err

    Set rsCONTROL = New Recordset
    
    strSQL = "SELECT * FROM ODASPPlotMast;"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    
      Dim StartDate As Date
      Dim EndDate As Date
      Dim Targetdate As Date
      Dim Difference As String
      
     Do Until rsCONTROL.EOF
      StartDate = rsCONTROL!CommencementDate
      EndDate = rsCONTROL!expirydate
      Difference = DateDiff("M", Date, EndDate)
      Targetdate = DateAdd("M", -6, Format(EndDate, "MMMM dd,yyyy"))
     
             curYear = Str(Year(Targetdate))
             curmonth = Month(rsCONTROL!Targetdate)
             
             If Val(curmonth) <= 9 Then
                    curmonth = "0" + Trim(curmonth)
             Else
             End If
            
            'With rsCONTROL
            rsCONTROL!TargetPeriod = Trim((curYear) + "/" + Trim(curmonth))
            rsCONTROL.Update
             'End With
                 
            rsCONTROL.MoveNext
    Loop
  
rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub selectModeLostFocus()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
        
        rsCONTROL.Open "SELECT * FROM ODASPLandRate WHERE PaymentMode = '" & Screen.ActiveForm.cboPaymentMode.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboPaymentMode.Text = !PaymentMode
        End With
        
Exit Sub

rsCONTROL.Close

err:
        ErrorMessage

End Sub

Public Sub loadPAYMENTMODE()
On Error GoTo err
With Screen.ActiveForm

        Set rsCONTROL = New ADODB.Recordset
        rsCONTROL.Open "SELECT * FROM ODASPPaymentMode WHERE PaymentMode = '" & .cboPaymentMode.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF And rsCONTROL.BOF Then Exit Sub
        .cboPaymentMode.Text = rsCONTROL!PaymentMode
End With
        
Exit Sub

rsCONTROL.Close

err:
        ErrorMessage

End Sub

Public Sub ClearListView2()
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
Public Sub LoadAccountType()
On Error GoTo err
        Set rsCONTROL = New ADODB.Recordset
        
        rsCONTROL.Open "SELECT * FROM ODASPAccountType WHERE AccountType = '" & Screen.ActiveForm.cboAccountType.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
  
        With rsCONTROL
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.txtAccountTypeDescription.Text = !AccountTypeDescription
        End With

rsCONTROL.Close
strSQL = ""

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
                i.Enabled = True
            End If
            
            If TypeOf i Is ComboBox Then
                i.Locked = False
                i.Enabled = True
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
    Screen.ActiveForm.cmdSearch.Enabled = False
    Screen.ActiveForm.cmdEdit.Enabled = False
    Screen.ActiveForm.cmdPrint.Enabled = False

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub editMYRECORD()
On Error GoTo err
        beditRECORD = True
        enableALLRECORD
        Screen.ActiveForm.cmdEdit.Enabled = False
        Screen.ActiveForm.cmdSearch.Enabled = False
        Screen.ActiveForm.cmdPrint.Enabled = False
        Screen.ActiveForm.cmdADDNEW.Enabled = False
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub cancelCMD()
        clearALLRECORD
        disableALLRECORD
End Sub
Public Sub addCMD()
        clearALLRECORD
        enableALLRECORD
        
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
Public Sub SelectSidingGotFocus()
On Error GoTo err

    Set rsCONTROL = New Recordset
    
    strSQL = "SELECT * FROM ODASPSiding;"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Screen.ActiveForm.cboSidingCode.Clear

    With rsCONTROL
            Do Until .EOF
                    Screen.ActiveForm.cboSidingCode.AddItem !SidingDescription
                    .MoveNext
            Loop
    End With
        
rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub selectSidingLostFocus()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
        
        rsCONTROL.Open "SELECT * FROM ODASPSiding WHERE SidingDescription = '" & Screen.ActiveForm.cboSidingCode.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboSidingCode.Text = !SidingCode
        End With
        
Exit Sub

rsCONTROL.Close

err:
        ErrorMessage

End Sub
Public Sub SelectTerminationReasonGotFocus()
On Error GoTo err

    Set rsCONTROL = New Recordset
    
    strSQL = "SELECT * FROM ODASPTerminationReasons;"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
   Screen.ActiveForm.cboTerminationCode.Clear

    With rsCONTROL
            Do Until .EOF
                    Screen.ActiveForm.cboTerminationCode.AddItem !TerminationReason
                    .MoveNext
            Loop
    End With
        
rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub selectTerminationReasonLostFocus()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
        
        rsCONTROL.Open "SELECT * FROM ODASPTerminationReasons WHERE TerminationReason = '" & Screen.ActiveForm.cboTerminationCode.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboTerminationCode.Text = !TerminationCode
        End With

rsCONTROL.Close
        
Exit Sub

rsCONTROL.Close

err:
        ErrorMessage

End Sub
Public Sub SelectCostingGotFocus()
On Error GoTo err

    Set rsCONTROL = New Recordset
    
    strSQL = "SELECT * FROM ODASPCOsting;"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Screen.ActiveForm.cboCostingCode.Clear

    With rsCONTROL
            Do Until .EOF
                    Screen.ActiveForm.cboCostingCode.AddItem !CostingDescription
                    .MoveNext
            Loop
    End With
        
rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub selectCostingLostFocus()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
        
        rsCONTROL.Open "SELECT * FROM ODASPCosting WHERE CostingDescription = '" & Screen.ActiveForm.cboCostingCode.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboCostingCode.Text = !CostingCode
        End With
        
Exit Sub

rsCONTROL.Close

err:
        ErrorMessage

End Sub
Public Sub SelectCurrencyGotFocus()
On Error GoTo err

    Set rsCONTROL = New Recordset
    
    strSQL = "SELECT * FROM ODASPCurrency;"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Screen.ActiveForm.cboCurrencyCode.Clear

    With rsCONTROL
            Do Until .EOF
                    Screen.ActiveForm.cboCurrencyCode.AddItem !CurrencyDescription
                    .MoveNext
            Loop
    End With
        
rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub selectCurrencyLostFocus()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
        
        rsCONTROL.Open "SELECT * FROM ODASPCurrency WHERE CurrencyDescription = '" & Screen.ActiveForm.cboCurrencyCode.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboCurrencyCode.Text = !CurrencyCode
                Screen.ActiveForm.txtCurrencySymbol.Text = !CurrencySymbol
        End With
        
Exit Sub

rsCONTROL.Close

err:
        ErrorMessage

End Sub

Public Sub SelectDiscountGotFocus()
On Error GoTo err

    Set rsCONTROL = New Recordset
    
    strSQL = "SELECT * FROM ODASPDiscount;"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Screen.ActiveForm.cboDiscountCode.Clear

    With rsCONTROL
            Do Until .EOF
                    Screen.ActiveForm.cboDiscountCode.AddItem !DiscountDescription
                    .MoveNext
            Loop
    End With
        
rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub SelectAccountingperiod()
On Error GoTo err

    Set rsCONTROL = New Recordset
    
    strSQL = "SELECT * FROM ALISPPeriod;"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Screen.ActiveForm.cbostartperiod.Clear

    With rsCONTROL
            Do Until .EOF
                    Screen.ActiveForm.cbostartperiod.AddItem !AccountingPeriod
                    .MoveNext
            Loop
    End With
        
rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub SelectDescription()
On Error GoTo err

    Set rsCONTROL = New Recordset
    
    strSQL = "SELECT * FROM ALISPPeriod where AccountingPeriod>= '" & Screen.ActiveForm.cbostartperiod.Text & "' ;"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Screen.ActiveForm.cboendperiod.Clear

    With rsCONTROL
            Do Until .EOF
                    Screen.ActiveForm.cboendperiod.AddItem !AccountingPeriod
                    .MoveNext
            Loop
    End With
        
rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub computeTOTALRentThisPeriod()
On Error GoTo err
    
    With frmODASRentPaid
            
            Set rsCONTROL = New Recordset
             strSQL = "SELECT sum(M.AmountPaid) as TOTALS FROM ODASMInstallment M Where M.CurrentPeriod >= '" & frmODASRentPaid.cbostartperiod & "' and (PaymentFlag = 'Y' or PaymentFlag = 'P') AND M.CurrentPeriod <= '" & frmODASRentPaid.cboendperiod & "'   ;"
            'strSQL = "SELECT sum(M.AmountPaid) as TOTALS FROM ODASMInstallment M Where M.CurrentPeriod = '" & .txtCurrentPeriod & "' and (M.PaymentFlag = 'Y' or M.PaymentFlag = 'N') ;"
            rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
            If rsCONTROL.EOF Or rsCONTROL.BOF Then
                    .txtTotalRent = 0
            ElseIf IsNull(rsCONTROL!Totals) = True Then
                    .txtTotalRent = 0
            Else: .txtTotalRent.Text = FormatNumber(rsCONTROL!Totals)
            End If
    End With
        
rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub computeTOTALRentDueThisPeriod()
On Error GoTo err
    
    With frmODASYearsearch
            
            Set rsCONTROL = New Recordset
            strSQL = "SELECT sum(M.Balance) as TOTALS FROM ODASMInstallment M Where M.CurrentPeriod >= '" & frmODASYearsearch.cbostartperiod & "' and PaymentFlag = 'N' AND M.CurrentPeriod <= '" & frmODASYearsearch.cboendperiod & "'   ;"
             rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
            If rsCONTROL.EOF Or rsCONTROL.BOF Then
                    .txtTotalRent = 0
            ElseIf IsNull(rsCONTROL!Totals) = True Then
                    .txtTotalRent = 0
            Else: .txtTotalRent.Text = FormatNumber(rsCONTROL!Totals)
            End If
    End With
        
rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub SelectDescription1()
On Error GoTo err

    Set rsCONTROL = New Recordset
    
    strSQL = "SELECT * FROM ALISPPeriod where AccountingPeriod>= '" & frmODASYearsearch.cbostartperiod.Text & "' ;"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Screen.ActiveForm.cboendperiod.Clear

    With rsCONTROL
            Do Until .EOF
                    Screen.ActiveForm.cboendperiod.AddItem !AccountingPeriod
                    .MoveNext
            Loop
    End With
        
rsCONTROL.Close
strSQL = ""

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub selectDiscountLostFocus()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
        
        rsCONTROL.Open "SELECT * FROM ODASPDiscount WHERE DiscountDescription = '" & Screen.ActiveForm.cboDiscountCode.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboDiscountCode.Text = !DiscountRate
        End With
        
Exit Sub

rsCONTROL.Close

err:
        ErrorMessage

End Sub
Public Sub LoadSidingDescription()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
        
        rsCONTROL.Open "SELECT * FROM ODASPSiding WHERE SidingCode = '" & Screen.ActiveForm.cboSidingCode.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboSidingCode.Text = !SidingCode
        End With

rsCONTROL.Close

Exit Sub

rsCONTROL.Close

err:
    ErrorMessage

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

Public Sub computeVOUCHERTOTAL()
On Error GoTo err
       With frmODASMVoucher
       
                Set rsSAVE = New ADODB.Recordset
                strSQL = "SELECT sum(AmountPaid) as totals from ODASMVoucherItem where VoucherNo = '" & frmODASMVoucher.txtVoucherNo.Text & "';"
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        .txtVoucherAmount.Text = 0
                ElseIf IsNull(rsSAVE!Totals) = True Then
                     If PartialPaid > 0 Then
                     .txtVoucherAmount.Text = CDbl(PartialPaid)
                     Else
                       .txtVoucherAmount.Text = CDbl(.txtAmountPaid.Text)
                     End If
                       '.txtInvoiceBalance = CDbl(.txtInvoiceAmount) - CDbl(.txtAmountPaid.Text)
                Else:
                If PartialPaid = 0 Then
                        .txtVoucherAmount.Text = FormatNumber(rsSAVE!Totals)
                 Else
                       .txtVoucherAmount.Text = PartialPaid
                 End If
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

Public Sub countVOUCHERITEMS()
On Error GoTo err
       With frmODASMVoucher
       
                Set rsSAVE = New ADODB.Recordset
                strSQL = "SELECT count(VoucherNo) as totals from ODASMVoucherItem where VoucherNo = '" & .txtVoucherNo.Text & "';"
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        .txtItems.Text = 0
                ElseIf IsNull(rsSAVE!Totals) = True Then
                        .txtItems.Text = 0
                Else
                        .txtItems.Text = FormatNumber(rsSAVE!Totals)
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

Public Sub enableButtons()
On Error GoTo err

    With Screen.ActiveForm
            .cmdUpdate.Enabled = False
            .cmdADDNEW.Enabled = True
            .cmdSearch.Enabled = True
            .cmdEdit.Enabled = True
            .cmdDelete.Enabled = True
            .cmdCancel.Enabled = True
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub
