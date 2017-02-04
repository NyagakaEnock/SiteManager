Attribute VB_Name = "modODASListView"
Public Sub showJBITEMS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView5.ListItems.Clear
                .ListView5.ColumnHeaders.Clear
                .ListView5.ColumnHeaders.Add , , "Item No", .ListView5.Width / 5
                .ListView5.ColumnHeaders.Add , , "Media Size", .ListView5.Width / 5
                .ListView5.ColumnHeaders.Add , , "Media ", .ListView5.Width / 5
                .ListView5.ColumnHeaders.Add , , "Site", .ListView5.Width / 5
                .ListView5.ColumnHeaders.Add , , "Qty", .ListView5.Width / 5
                .ListView5.ColumnHeaders.Add , , "Price", .ListView5.Width / 5

                .ListView5.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset

                strSQL = "SELECT * FROM ODASMJobBriefItems JB WHERE JB.JobBriefNo =  '" & .txtJobBriefNo & "' ORDER by JB.JobBriefItemNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView5.ListItems.Add(, , CStr(rsLIST!JobBriefItemNo))
                        
                        If Not IsNull(rsLIST!MediaSize) Then
                            MyList.SubItems(1) = CStr(rsLIST!MediaSize)
                        End If

                        If Not IsNull(rsLIST!MediaCode) Then
                                MyList.SubItems(2) = CStr(rsLIST!MediaCode)
                        End If
                        
                        If Not IsNull(rsLIST!SiteNo) Then
                                MyList.SubItems(3) = CStr(rsLIST!SiteNo)
                        End If
                        
                        If Not IsNull(rsLIST!ItemQuantity) Then
                                MyList.SubItems(4) = CStr(rsLIST!ItemQuantity)
                        End If

                        If Not IsNull(rsLIST!NetItemPrice) Then
                                MyList.SubItems(5) = CStr(rsLIST!NetItemPrice)
                        End If

                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showJOBCARDS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView2.ListItems.Clear
                .ListView2.ColumnHeaders.Clear
                
                .ListView2.ColumnHeaders.Add , , "Job Card No", .ListView2.Width / 5 ', lvwColumnCenter
                .ListView2.ColumnHeaders.Add , , "Dept. Code", .ListView2.Width / 7
                .ListView2.ColumnHeaders.Add , , "Department", .ListView2.Width / 5
                .ListView2.ColumnHeaders.Add , , "Price", .ListView2.Width / 5
                .ListView2.ColumnHeaders.Add , , "VAT", .ListView2.Width / 5
                .ListView2.ColumnHeaders.Add , , "Total", .ListView2.Width / 5
                .ListView2.ColumnHeaders.Add , , "Used", .ListView2.Width / 5

                .ListView2.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMjobCard J, ODASPDepartment D Where J.JobCardNo = '" & .txtJobBriefNo.Text & "' and J.DepartmentCode = D.DepartmentCode;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                    Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!JobCardNo))
                    If Not IsNull(rsLIST!DepartmentCode) Then
                        MyList.SubItems(1) = CStr(rsLIST!DepartmentCode)
                    End If
                    If Not IsNull(rsLIST!DepartmentDescription) Then
                        MyList.SubItems(2) = CStr(rsLIST!DepartmentDescription)
                    End If

                    If Not IsNull(rsLIST!PriceExclusive) Then
                        MyList.SubItems(3) = CStr(rsLIST!PriceExclusive)
                    End If
                    
                    If Not IsNull(rsLIST!VatAmount) Then
                        MyList.SubItems(4) = CStr(rsLIST!VatAmount)
                    End If
                    
                    If Not IsNull(rsLIST!TotalCost) Then
                        MyList.SubItems(5) = CStr(rsLIST!TotalCost)
                    End If
                    
                    If Not IsNull(rsLIST!Used) Then
                        MyList.SubItems(6) = CStr(rsLIST!Used)
                    End If

                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub getJobBrief()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "JobBrief No", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "CompanyName", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "AccountNo", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Status", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select ODASMJobBrief.JobBriefNo, ODASPAccount.CompanyName, ODASMJobBrief.Status,ODASPAccount.AccountNo from ODASMJobBrief, ODASPAccount Where ODASMJobBrief.AccountNo = ODASPAccount.AccountNo order by JobBriefNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!JobBriefNo))
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(1) = CStr(rsLIST!CompanyName)
                        End If

                        If Not IsNull(rsLIST!AccountNo) Then
                            MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                        End If
                        
                        If Not IsNull(rsLIST!Status) Then
                            MyList.SubItems(3) = CStr(rsLIST!Status)
                        End If

                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
    If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub listAPPROVALTASKS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "User", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Status", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Operation Date ", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Comment", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Accept", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                rsLIST.Open "SELECT UserCode, Status, OperationDate, Comment, Accept FROM ODASMOperation WHERE ApplicationNo =  '" & Screen.ActiveForm.txtApplicationNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!UserCode))
                        
                        If Not IsNull(rsLIST!Status) Then
                            MyList.SubItems(1) = CStr(rsLIST!Status)
                        End If

                        If Not IsNull(rsLIST!operationDate) Then
                                MyList.SubItems(2) = CStr(rsLIST!operationDate)
                        End If
                        
                        If Not IsNull(rsLIST!Comment) Then
                                MyList.SubItems(3) = CStr(rsLIST!Comment)
                        End If
                        
                        If Not IsNull(rsLIST!Accept) Then
                                MyList.SubItems(4) = CStr(rsLIST!Accept)
                        End If

                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showPAYMENTMETHOD()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Payment Method", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 2

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "select PaymentMethod, PaymentMethodDescription from ODASPPaymentMethod "
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PaymentMethod))
                        
                        If Not IsNull(rsLIST!PaymentMethodDescription) Then
                            MyList.SubItems(1) = CStr(rsLIST!PaymentMethodDescription)
                        End If

                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage

End Sub
Public Sub ListALLInstallments()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Ref", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "InvoiceNo", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Due Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Amount Due", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Invoiced?", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "select * from ODASMJobBriefInstallment where JobbriefNo = '" & .txtJobBriefNo & "' AND (Paid='N' OR Paid IS Null) order by InstallmentNo  "
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!InstallmentNo))
                        
                        If Not IsNull(rsLIST!InvoiceNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!InvoiceNo)
                        End If
                        
                        If Not IsNull(rsLIST!PaymentDueDate) Then
                            MyList.SubItems(2) = Format(rsLIST!PaymentDueDate, "dd/mm/yyyy")
                        End If
                        
                        If Not IsNull(rsLIST!Balance) Then
                            MyList.SubItems(3) = FormatNumber(rsLIST!Balance, 2)
                        End If
                        
                        If Not IsNull(rsLIST!Invoiced) Then
                            MyList.SubItems(4) = CStr(rsLIST!Invoiced)
                        End If

                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage

End Sub

Public Sub showALLInstallments()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Reference", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "InvoiceNo", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Invoice Ref", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Due Date", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Invoiced?", .ListView1.Width / 6

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "select * from ODASMJobBriefInstallment where JobbriefNo = '" & .txtJobBriefNo & "' order by InstallmentNo  "
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!InstallmentNo))
                        
                        If Not IsNull(rsLIST!InvoiceNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!InvoiceNo)
                        End If
                        
                        If Not IsNull(rsLIST!InvoiceReference) Then
                            MyList.SubItems(2) = CStr(rsLIST!InvoiceReference)
                        End If
                        
                        If Not IsNull(rsLIST!PaymentDueDate) Then
                            MyList.SubItems(3) = CStr(rsLIST!PaymentDueDate)
                        End If
                        
                        If Not IsNull(rsLIST!Amount) Then
                            MyList.SubItems(4) = CStr(rsLIST!Amount)
                        End If
                        
                        If Not IsNull(rsLIST!Invoiced) Then
                            MyList.SubItems(5) = CStr(rsLIST!Invoiced)
                        End If

                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage

End Sub

Public Sub showDURATIONMODE()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Duration Mode", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 2

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "select DurationMode, DurationDescription from ODASPDuration "
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!DurationMode))
                        
                        If Not IsNull(rsLIST!DurationDescription) Then
                            MyList.SubItems(1) = CStr(rsLIST!DurationDescription)
                        End If

                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage
End Sub

Public Sub GetCheques()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Cheque No", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Payee Details", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Requisition No", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Cheque Date", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Amount Paid", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Amount Due", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Payment Flag", .ListView1.Width / 8

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ChequeNo))
                        
                        If Not IsNull(rsLIST!vOUCHERnO) Then
                            MyList.SubItems(2) = CStr(rsLIST!vOUCHERnO)
                        End If
                        
                        If Not IsNull(rsLIST!ChequeDate) Then
                                MyList.SubItems(3) = CStr(rsLIST!ChequeDate)
                        End If

                        If Not IsNull(rsLIST!Amount) Then
                                MyList.SubItems(4) = CStr(rsLIST!Amount)
                        End If
                        
                        If Not IsNull(rsLIST!AmountDue) Then
                                MyList.SubItems(5) = CStr(rsLIST!AmountDue)
                        End If
                        
                        If Not IsNull(rsLIST!PaymentFlag) Then
                                MyList.SubItems(6) = CStr(rsLIST!PaymentFlag)
                        End If
                        
                        If Not IsNull(rsLIST!PayeeDetails) Then
                                MyList.SubItems(1) = CStr(rsLIST!PayeeDetails)
                        End If

                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage
End Sub

Public Sub GetChequesISSUEDTHISPERIOD()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Cheque No", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Payee Details", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Account No", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Cheque Date", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Amount ", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Status", .ListView1.Width / 8

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select * from ALISMCheque C Where C.Issued = 'Y' and C.Banked = 'N' and C.CurrentPeriod = '" & CurrentPeriod & "' "
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ChequeNo))
                        
                        If Not IsNull(rsLIST!AccountNo) Then
                            MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                        End If
                        
                        If Not IsNull(rsLIST!ChequeDate) Then
                                MyList.SubItems(3) = CStr(rsLIST!ChequeDate)
                        End If

                        If Not IsNull(rsLIST!ChequeAmount) Then
                                MyList.SubItems(4) = CStr(rsLIST!ChequeAmount)
                        End If
                        
                        If Not IsNull(rsLIST!Status) Then
                                MyList.SubItems(5) = CStr(rsLIST!Status)
                        End If
                        
                        If Not IsNull(rsLIST!PayeeDetails) Then
                                MyList.SubItems(1) = CStr(rsLIST!PayeeDetails)
                        End If

                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage
End Sub

Public Sub GetChequesRELATED()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView2.ListItems.Clear
                .ListView2.ColumnHeaders.Clear
                
                .ListView2.ColumnHeaders.Add , , "Cheque No", .ListView2.Width / 4
                .ListView2.ColumnHeaders.Add , , "AccountNo", .ListView2.Width / 4
                .ListView2.ColumnHeaders.Add , , "Cheque Date", .ListView2.Width / 4
                .ListView2.ColumnHeaders.Add , , "Amount ", .ListView2.Width / 4

                .ListView2.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select * from ALISMCheque C Where C.Issued = 'N' and (C.Banked = 'N' or C.Banked is null) and C.Authorized = 'Y' and C.CurrentPeriod = '" & CurrentPeriod & "' and C.AccountNo = '" & frmODASMCheckIssuance.txtAccountNo.Text & "' "
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                       
                        If rsLIST!ChequeNo = frmODASMCheckIssuance.txtChequeNo.Text Then
                        Else
                                Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!ChequeNo))
                                
                                If Not IsNull(rsLIST!AccountNo) Then
                                    MyList.SubItems(1) = CStr(rsLIST!AccountNo)
                                End If
                                
                                If Not IsNull(rsLIST!ChequeDate) Then
                                        MyList.SubItems(2) = CStr(rsLIST!ChequeDate)
                                End If
        
                                If Not IsNull(rsLIST!ChequeAmount) Then
                                        MyList.SubItems(3) = CStr(rsLIST!ChequeAmount)
                                End If
                        End If

                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage
End Sub

Public Sub GetChequeReference()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Reference", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Total Amount", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Template", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Date Banked", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Bank No", .ListView1.Width / 6

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!Reference))
                        
                        If Not IsNull(rsLIST!TotalAmount) Then
                            MyList.SubItems(1) = CStr(rsLIST!TotalAmount)
                        End If
                        
                        If Not IsNull(rsLIST!templateCode) Then
                                MyList.SubItems(2) = CStr(rsLIST!templateCode)
                        End If

                        If Not IsNull(rsLIST!DateBanked) Then
                                MyList.SubItems(3) = CStr(rsLIST!DateBanked)
                        End If

                        If Not IsNull(rsLIST!BankNo) Then
                                MyList.SubItems(4) = CStr(rsLIST!BankNo)
                        End If
                        
                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage
End Sub

Public Sub GetOtherPayment()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView3.ListItems.Clear
                .ListView3.ColumnHeaders.Clear
                
                .ListView3.ColumnHeaders.Add , , "Cheque No", .ListView3.Width / 6
                .ListView3.ColumnHeaders.Add , , "Payee Details", .ListView3.Width / 4
                .ListView3.ColumnHeaders.Add , , "Amount", .ListView3.Width / 6
                .ListView3.ColumnHeaders.Add , , "Amount Due", .ListView3.Width / 6
                .ListView3.ColumnHeaders.Add , , "Payment Flag", .ListView3.Width / 8
                .ListView3.ColumnHeaders.Add , , "Requisition No", .ListView3.Width / 8

                .ListView3.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select ALISMChequeEntry.ChequeNo, ALISMCheque.PayeeDetails, ALISMChequeEntry.Amount, ALISMChequeEntry.AmountDue, ALISMChequeEntry.PaymentFlag, ALISMChequeEntry.VoucherNo from ALISMChequeEntry, ALISMCheque Where ALISMCheque.ChequeNo = ALISMChequeEntry.ChequeNo and ALISMChequeEntry.chequeNo = '" & Screen.ActiveForm.txtChequeNo.Text & "' and ALISMChequeEntry.PaymentFlag = 'Y';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!ChequeNo))
                        
                        If Not IsNull(rsLIST!Amount) Then
                                MyList.SubItems(2) = CStr(rsLIST!Amount)
                        End If
                        
                        If Not IsNull(rsLIST!AmountDue) Then
                                MyList.SubItems(3) = CStr(rsLIST!AmountDue)
                        End If
                        
                        If Not IsNull(rsLIST!PaymentFlag) Then
                                MyList.SubItems(4) = CStr(rsLIST!PaymentFlag)
                        End If
                        
                        If Not IsNull(rsLIST!PayeeDetails) Then
                                MyList.SubItems(1) = CStr(rsLIST!PayeeDetails)
                        End If
                        
                        If Not IsNull(rsLIST!vOUCHERnO) Then
                                MyList.SubItems(5) = CStr(rsLIST!vOUCHERnO)
                        End If


                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage
End Sub

Public Sub GetInvoicesNotPaid()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Invoice No", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "LPO No", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Invoice Date", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Job Brief No", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "AccountNo", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Supplier", .ListView1.Width / 6

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select *  from ODASMInvoice I, ODASPAccount A Where (I.Requisitioned = 'N' or I.Requisitioned is null) and I.AccountNo = A.AccountNo ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
                        
                        If Not IsNull(rsLIST!LPONo) Then
                                MyList.SubItems(1) = CStr(rsLIST!LPONo)
                        End If

                        If Not IsNull(rsLIST!InvoiceDate) Then
                                MyList.SubItems(2) = CStr(rsLIST!InvoiceDate)
                        End If

                        If Not IsNull(rsLIST!JobCardNo) Then
                                MyList.SubItems(3) = CStr(rsLIST!JobCardNo)
                        End If
                        
                        If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(4) = CStr(rsLIST!AccountNo)
                        End If
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                                MyList.SubItems(5) = CStr(rsLIST!CompanyName)
                        End If

                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage
End Sub
Public Sub GetRentNotPaid()
On Error GoTo err
    
        With Screen.ActiveForm
        
            .ListView1.ListItems.Clear
            .ListView1.ColumnHeaders.Clear
            
            .ListView1.ColumnHeaders.Add , , "Installment ", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Due Date", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "AccountNo", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 6
        
            .ListView1.View = lvwReport
            
            Dim rsLIST As ADODB.Recordset
            Set rsLIST = New ADODB.Recordset
            
            strSQL = "Select * from ODASMInstallment I, ODASPAccount A Where (I.PaymentFlag = 'N' or I.PaymentFlag = 'P') and I.AccountNo = A.AccountNo;"
            rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsLIST.RecordCount <> 0 Then
            
            frmODASMVoucher.ProgressBar1.Visible = True
            frmODASMVoucher.ProgressBar1.Value = 0: frmODASMVoucher.ProgressBar1.Min = 0: frmODASMVoucher.ProgressBar1.Max = rsLIST.RecordCount
            
            Dim MyList As ListItem
                       
            While Not rsLIST.EOF
                    
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
                    
                    If Not IsNull(rsLIST!ContractNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!ContractNo)
                    End If
        
                    If Not IsNull(rsLIST!PaymentDueDate) Then
                            MyList.SubItems(2) = CStr(rsLIST!PaymentDueDate)
                    End If
        
                    If Not IsNull(rsLIST!PaymentDue) Then
                            MyList.SubItems(3) = CStr(rsLIST!PaymentDue)
                    End If
                    
                    If Not IsNull(rsLIST!AccountNo) Then
                            MyList.SubItems(4) = CStr(rsLIST!AccountNo)
                    End If
                    
                    If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(5) = CStr(rsLIST!CompanyName)
                    End If
                    
                    frmODASMVoucher.ProgressBar1.Value = frmODASMVoucher.ProgressBar1.Value + 1
                    rsLIST.MoveNext
            Wend
            frmODASMVoucher.ProgressBar1.Visible = False
            End If
            Set MyList = Nothing
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage
End Sub
Public Sub GetRentNotRequisitioned()
On Error GoTo err
    
        With Screen.ActiveForm
        
            .ListView1.ListItems.Clear
            .ListView1.ColumnHeaders.Clear
            
            .ListView1.ColumnHeaders.Add , , "Installment ", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Due Date", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "AccountNo", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Plot Details", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Installment No", .ListView1.Width / 6
        
            .ListView1.View = lvwReport
            
            Dim rsLIST As ADODB.Recordset
            Set rsLIST = New ADODB.Recordset
            
            strSQL = "Select * from ODASMInstallment I, ODASPAccount A,ODASPPlot P Where P.PlotNo=I.PlotNo AND (I.Requisitioned = 'N' or I.Requisitioned is null or I.Requisitioned = 'Y' ) and A.AccountNo = P.AccountNo  AND (I.PaymentDueDate>='" & Format(.DTPStartDate, "yyyy/MM/dd") & "' AND I.PaymentDueDate<='" & Format(.DTPLastDate, "yyyy/MM/dd") & "') AND I.Balance>0;"
            rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            Debug.Print strSQL
            If rsLIST.RecordCount <> 0 Then
            
            frmODASMVoucher.ProgressBar1.Visible = True
            frmODASMVoucher.ProgressBar1.Value = 0: frmODASMVoucher.ProgressBar1.Min = 0: frmODASMVoucher.ProgressBar1.Max = rsLIST.RecordCount
            Debug.Print strSQL
            Dim MyList As ListItem
                       
            While Not rsLIST.EOF
                    
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
                    
                    If Not IsNull(rsLIST!ContractNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!ContractNo)
                    End If
        
                    If Not IsNull(rsLIST!PaymentDueDate) Then
                            MyList.SubItems(2) = CStr(rsLIST!PaymentDueDate)
                    End If
        
                    If Not IsNull(rsLIST!PaymentDue) Then
                            MyList.SubItems(3) = CStr(rsLIST!PaymentDue)
                    End If
                    
                    If Not IsNull(rsLIST!AccountNo) Then
                            MyList.SubItems(4) = CStr(rsLIST!AccountNo)
                    End If
                    
                    If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(5) = CStr(rsLIST!CompanyName)
                    End If
                    
                    'MyList.SubItems(6) = "[ LR No - " & CStr(rsLIST!LRNo) & " ] " & CStr(rsLIST!PhysicalLocation) & " For the period Starting " & MonthName(Month(DateAdd("yyyy", -1 * rsLIST!ContractLength, rsLIST!PaymentDueDate))) & " " & Year(DateAdd("yyyy", -1 * rsLIST!ContractLength, rsLIST!PaymentDueDate)) & " To " & MonthName(Month(rsLIST!PaymentDueDate)) & " " & Year(rsLIST!PaymentDueDate) & ". Ref. No " & rsLIST!ContractNo
                    MyList.SubItems(6) = "[ LR No - " & CStr(rsLIST!LRNo) & " ] " & CStr(rsLIST!PhysicalLocation) & ". Ref. No " & rsLIST!ContractNo
                                       
                    If Not IsNull(rsLIST!InstallmentNo) Then
                            MyList.SubItems(7) = CStr(rsLIST!InstallmentNo)
                    End If
                    
                    frmODASMVoucher.ProgressBar1.Value = frmODASMVoucher.ProgressBar1.Value + 1
                    rsLIST.MoveNext
            Wend
            frmODASMVoucher.ProgressBar1.Visible = False
            End If
            Set MyList = Nothing
        End With

Exit Sub

err:
'If err.Number = 3265 Then Resume Next
ErrorMessage
End Sub


Public Sub GetRentRequisitioned_ne()
On Error GoTo err
    
        With Screen.ActiveForm
        
            .ListView1.ListItems.Clear
            .ListView1.ColumnHeaders.Clear
            
            .ListView1.ColumnHeaders.Add , , "Installment ", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Due Date", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Requisition Date", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "AccountNo", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Plot Details", .ListView1.Width / 6
        
            .ListView1.View = lvwReport
            
            Dim rsLIST As ADODB.Recordset
            Set rsLIST = New ADODB.Recordset
            
            strSQL = "Select * from ODASMInstallment I, ODASPAccount A,ODASPPlot P Where P.PlotNo=I.ContractNo AND (I.Requisitioned = 'Y') AND (I.ChequeIssued ='N' OR I.ChequeIssued IS NULL) and I.AccountNo = A.AccountNo AND (I.PaymentDueDate>='" & Format(.DTPStartDate, "yyyy/MM/dd") & "' AND I.PaymentDueDate<='" & Format(.DTPLastDate, "yyyy/MM/dd") & "') AND I.PaymentDue>0;"
            rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsLIST.RecordCount <> 0 Then
            
            frmODASMVoucher.ProgressBar1.Visible = True
            frmODASMVoucher.ProgressBar1.Value = 0: frmODASMVoucher.ProgressBar1.Min = 0: frmODASMVoucher.ProgressBar1.Max = rsLIST.RecordCount
            
            Dim MyList As ListItem
                       
            While Not rsLIST.EOF
                    
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
                    
                    If Not IsNull(rsLIST!ContractNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!ContractNo)
                    End If
        
                    If Not IsNull(rsLIST!PaymentDueDate) Then
                            MyList.SubItems(2) = CStr(rsLIST!PaymentDueDate)
                    End If
        
                    If Not IsNull(rsLIST!PaymentDue) Then
                            MyList.SubItems(3) = CStr(rsLIST!PaymentDue)
                    End If
                    
                    If Not IsNull(rsLIST!VoucherDate) Then
                            MyList.SubItems(4) = CStr(rsLIST!VoucherDate)
                    End If
                    
                    
                    If Not IsNull(rsLIST!AccountNo) Then
                            MyList.SubItems(5) = CStr(rsLIST!AccountNo)
                    End If
                    
                    If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(6) = CStr(rsLIST!CompanyName)
                    End If
                    
                    MyList.SubItems(7) = CStr(rsLIST!PhysicalLocation) & " [ LR No - " & CStr(rsLIST!LRNo) & " ]"
                                        
                    frmODASMVoucher.ProgressBar1.Value = frmODASMVoucher.ProgressBar1.Value + 1
                    rsLIST.MoveNext
            Wend
            frmODASMVoucher.ProgressBar1.Visible = False
            End If
            Set MyList = Nothing
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage
End Sub

Public Sub GetVoucherToReverse()
On Error GoTo err
    
        With frmODASMReversePayment
        
            .ListView1.ListItems.Clear
            .ListView1.ColumnHeaders.Clear
            
            .ListView1.ColumnHeaders.Add , , "Installment ", .ListView1.Width / 8
            .ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 7
            .ListView1.ColumnHeaders.Add , , "Installment Year", .ListView1.Width / 9
            .ListView1.ColumnHeaders.Add , , "Rent Payable", .ListView1.Width / 7
            .ListView1.ColumnHeaders.Add , , "AccountNo", .ListView1.Width / 7
            .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 5
            .ListView1.ColumnHeaders.Add , , "Payment Due Date", .ListView1.Width / 8
            .ListView1.ColumnHeaders.Add , , "Amount Paid", .ListView1.Width / 7
            .ListView1.ColumnHeaders.Add , , "Payment Date", .ListView1.Width / 7
        
            .ListView1.View = lvwReport
            
            Dim rsLIST As ADODB.Recordset
            Set rsLIST = New ADODB.Recordset
            
            strSQL = "Select * from ODASMInstallment I,ODASPAccount A Where I.AccountNo=A.AccountNo and I.PaymentFlag = 'Y' and I.ContractNo= '" & frmODASMReversePayment.txtContractNo.Text & "';"
            rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsLIST.RecordCount <> 0 Then
            
            Dim MyList As ListItem
                       
            While Not rsLIST.EOF
                    
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
                    
                    If Not IsNull(rsLIST!ContractNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!ContractNo)
                    End If
        
                    If Not IsNull(rsLIST!Installment) Then
                            MyList.SubItems(2) = CStr(rsLIST!Installment)
                    End If
        
                    If Not IsNull(rsLIST!TotalRent) Then
                            MyList.SubItems(3) = CStr(rsLIST!TotalRent)
                    End If
                    
                    If Not IsNull(rsLIST!AccountNo) Then
                            MyList.SubItems(4) = CStr(rsLIST!AccountNo)
                    End If
                    
                    If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(5) = CStr(rsLIST!CompanyName)
                    End If
                    
                    If Not IsNull(rsLIST!PaymentDueDate) Then
                            MyList.SubItems(6) = CStr(rsLIST!PaymentDueDate)
                    End If
                    
                    If Not IsNull(rsLIST!AmountPaid) Then
                            MyList.SubItems(7) = CStr(rsLIST!AmountPaid)
                    End If
                    
                    If Not IsNull(rsLIST!PaymentDate) Then
                            MyList.SubItems(8) = CStr(rsLIST!PaymentDate)
                    End If
                    
                    rsLIST.MoveNext
            Wend
            End If
            Set MyList = Nothing
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage
End Sub

Public Sub GetRentNotPaidforSpecificAccount()
On Error GoTo err
    
        With frmODASMVoucher
        
            .ListView1.ListItems.Clear
            .ListView1.ColumnHeaders.Clear
            
            .ListView1.ColumnHeaders.Add , , "Installment ", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Due Date", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "AccountNo", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 6
        
            .ListView1.View = lvwReport
            
            Dim rsLIST As ADODB.Recordset
            Set rsLIST = New ADODB.Recordset
            
            strSQL = "Select * from ODASMInstallment I, ODASPAccount A Where (I.PaymentFlag = 'N' or I.PaymentFlag = 'P') and I.AccountNo = A.AccountNo and I.AccountNo = '" & .txtAccountNo & "';"
            rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsLIST.RecordCount <> 0 Then
            
            frmODASMVoucher.ProgressBar1.Visible = True
            frmODASMVoucher.ProgressBar1.Value = 0: frmODASMVoucher.ProgressBar1.Min = 0: frmODASMVoucher.ProgressBar1.Max = rsLIST.RecordCount
            
            Dim MyList As ListItem
                       
            While Not rsLIST.EOF
                    
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
                    
                    If Not IsNull(rsLIST!ContractNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!ContractNo)
                    End If
        
                    If Not IsNull(rsLIST!PaymentDueDate) Then
                            MyList.SubItems(2) = CStr(rsLIST!PaymentDueDate)
                    End If
        
                    If Not IsNull(rsLIST!PaymentDue) Then
                            MyList.SubItems(3) = CStr(rsLIST!PaymentDue)
                    End If
                    
                    If Not IsNull(rsLIST!AccountNo) Then
                            MyList.SubItems(4) = CStr(rsLIST!AccountNo)
                    End If
                    
                    If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(5) = CStr(rsLIST!CompanyName)
                    End If
                    
                    frmODASMVoucher.ProgressBar1.Value = frmODASMVoucher.ProgressBar1.Value + 1
                    rsLIST.MoveNext
            Wend
            frmODASMVoucher.ProgressBar1.Visible = False
            End If
            Set MyList = Nothing
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage
End Sub
Public Sub GetRentRequisitionedforSpecificAccount()
On Error GoTo err
    
        With frmODASMVoucher
        
            .ListView1.ListItems.Clear
            .ListView1.ColumnHeaders.Clear
            
            .ListView1.ColumnHeaders.Add , , "Installment ", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Due Date", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "AccountNo", .ListView1.Width / 6
            .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 6
        
            .ListView1.View = lvwReport
            
            Dim rsLIST As ADODB.Recordset
            Set rsLIST = New ADODB.Recordset
            
            strSQL = "Select * from ODASMInstallment I, ODASPAccount A Where (I.Requisitioned = 'N' or I.Requisitioned IS NULL) and I.AccountNo = A.AccountNo and I.AccountNo = '" & .txtAccountNo & "';"
            rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsLIST.RecordCount <> 0 Then
            
            frmODASMVoucher.ProgressBar1.Visible = True
            frmODASMVoucher.ProgressBar1.Value = 0: frmODASMVoucher.ProgressBar1.Min = 0: frmODASMVoucher.ProgressBar1.Max = rsLIST.RecordCount
            
            Dim MyList As ListItem
                       
            While Not rsLIST.EOF
                    
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
                    
                    If Not IsNull(rsLIST!ContractNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!ContractNo)
                    End If
        
                    If Not IsNull(rsLIST!PaymentDueDate) Then
                            MyList.SubItems(2) = CStr(rsLIST!PaymentDueDate)
                    End If
        
                    If Not IsNull(rsLIST!PaymentDue) Then
                            MyList.SubItems(3) = CStr(rsLIST!PaymentDue)
                    End If
                    
                    If Not IsNull(rsLIST!AccountNo) Then
                            MyList.SubItems(4) = CStr(rsLIST!AccountNo)
                    End If
                    
                    If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(5) = CStr(rsLIST!CompanyName)
                    End If
                    
                    frmODASMVoucher.ProgressBar1.Value = frmODASMVoucher.ProgressBar1.Value + 1
                    rsLIST.MoveNext
            Wend
            frmODASMVoucher.ProgressBar1.Visible = False
            End If
            Set MyList = Nothing
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage
End Sub

Public Sub GetRateNotPaid()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Reference No ", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Site No ", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Start Date", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Amount Due", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "End Date", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Date Due", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Job Brief", .ListView1.Width / 6

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select *  from ODASMCouncilRatedue C Where C.AmountDue > 0 and C.DueDate <= '" & Format(Date, "yyyy/mm/dd") & "'  and C.Requisitioned = 'N' and C.Paid = 'N';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ReferenceNo))
                        
                        If Not IsNull(rsLIST!SiteNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!SiteNo)
                        End If

                        If Not IsNull(rsLIST!StartDate) Then
                                MyList.SubItems(2) = CStr(rsLIST!StartDate)
                        End If

                        If Not IsNull(rsLIST!EndDate) Then
                                MyList.SubItems(4) = CStr(rsLIST!EndDate)
                        End If
                        
                        If Not IsNull(rsLIST!DueDate) Then
                                MyList.SubItems(5) = CStr(rsLIST!DueDate)
                        End If
                        
                        If Not IsNull(rsLIST!AmountDue) Then
                                MyList.SubItems(3) = CStr(rsLIST!AmountDue)
                        End If

                        If Not IsNull(rsLIST!JobBriefItemNo) Then
                                MyList.SubItems(6) = CStr(rsLIST!JobBriefItemNo)
                        End If

                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage
End Sub
Public Sub showALLOTHERREQUISITIONS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Request No", .ListView1.Width / 4 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Request", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Job card", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Requisition No", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Remarks", .ListView1.Width / 4
                                
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMRequisitionItems Where Request='Y' and Authorized = 'Y' and approved = 'Y' and (issued = 'N' or issued is null);"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                       Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ItemNo))

                        If Not IsNull(rsLIST!ProductCode) Then
                            MyList.SubItems(1) = CStr(rsLIST!ProductCode)
                        End If
                        
                        If Not IsNull(rsLIST!JobCardNo) Then
                            MyList.SubItems(2) = CStr(rsLIST!JobCardNo)
                        End If
                                        
                        If Not IsNull(rsLIST!AllowanceTotals) Then
                            MyList.SubItems(3) = CStr(rsLIST!AllowanceTotals)
                        End If
                        
                        If Not IsNull(rsLIST!RequisitionNo) Then
                            MyList.SubItems(4) = CStr(rsLIST!RequisitionNo)
                        End If
                        
                        If Not IsNull(rsLIST!RequestPurpose) Then
                            MyList.SubItems(5) = CStr(rsLIST!RequestPurpose)
                        End If
     
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub
Public Sub GetEachREQUISITIONS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Request No", .ListView1.Width / 4 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Request", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Job card", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Project", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Remarks", .ListView1.Width / 4

                 
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMJobbrief, ODASMRequisition, ODASMRequisitionItems Where ODASMJobbrief.RequisitionNo=ODASMRequisition.RequisitionNo and ODASMRequisition.RequisitionNo = ODASMRequisitionItems.RequisitionNo and ODASMRequisitionItems.Request='Y' and ODASMRequisitionItems.Authorized = 'Y' and ODASMRequisitionItems.approved = 'Y' and (ODASMRequisitionItems.issued = 'N' or ODASMRequisitionItems.issued is null);"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                       Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!RequisitionNo))

                        If Not IsNull(rsLIST!ProductCode) Then
                            MyList.SubItems(1) = CStr(rsLIST!ProductCode)
                        End If
                        
                        If Not IsNull(rsLIST!JobCardNo) Then
                            MyList.SubItems(2) = CStr(rAAsLIST!JobCardNo)
                        End If
                                        
                        If Not IsNull(rsLIST!AllowanceTotals) Then
                            MyList.SubItems(3) = CStr(rsLIST!AllowanceTotals)
                        End If
                        
                        If Not IsNull(rsLIST!RequestPurpose) Then
                            MyList.SubItems(4) = CStr(rsLIST!RequestPurpose)
                        End If
     
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub
Public Sub GetDistinctREQUISITIONS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Request No", .ListView1.Width / 4 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Request", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Job card", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Project", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Remarks", .ListView1.Width / 4

                 
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMJobbrief, ODASMRequisition, ODASMRequisitionItems Where ODASMJobbrief.RequisitionNo=ODASMRequisition.RequisitionNo and ODASMRequisition.RequisitionNo = ODASMRequisitionItems.RequisitionNo and ODASMRequisitionItems.Request='Y' and ODASMRequisitionItems.Authorized = 'Y' and ODASMRequisitionItems.approved = 'Y' and (ODASMRequisitionItems.issued = 'N' or ODASMRequisitionItems.issued is null);"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                       Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!RequisitionNo))

                        If Not IsNull(rsLIST!ProductCode) Then
                            MyList.SubItems(1) = CStr(rsLIST!ProductCode)
                        End If
                        
                        If Not IsNull(rsLIST!JobCardNo) Then
                            MyList.SubItems(2) = CStr(rAAsLIST!JobCardNo)
                        End If
                                        
                        If Not IsNull(rsLIST!AllowanceTotals) Then
                            MyList.SubItems(3) = CStr(rsLIST!AllowanceTotals)
                        End If
                        
                        If Not IsNull(rsLIST!RequestPurpose) Then
                            MyList.SubItems(4) = CStr(rsLIST!RequestPurpose)
                        End If
     
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub
Public Sub showALLLPOS()
On Error GoTo err
    
        With ALISFOManager
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "LPO No", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "LPO Date", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "RequisitionNo", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Supplier", .ListView1.Width / 3.5
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 6

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select *  from ODASMLPO R, ODASPAccount A Where (R.InvoiceRecieved = 'N' or R.InvoiceRecieved is null) and R.AccountNo = A.AccountNo and R.LPOStatus = 'R';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!LPONo))
                        
                        If Not IsNull(rsLIST!LPODate) Then
                                MyList.SubItems(1) = CStr(rsLIST!LPODate)
                        End If

                        If Not IsNull(rsLIST!RequisitionNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!RequisitionNo)
                        End If
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                                MyList.SubItems(3) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!PriceInclusive) Then
                                MyList.SubItems(4) = CStr(rsLIST!PriceInclusive)
                        End If

                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showCOYBankAccounts()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Bank No", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Account No", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Details", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Company Name", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select *  from ALISPBank B, ALISPBankAccount A Where B.BankNo = A.BankNo ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!BankNo))
                        
                        If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!AccountNo)
                        End If

                        If Not IsNull(rsLIST!Details) Then
                                MyList.SubItems(2) = CStr(rsLIST!Details)
                        End If
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                                MyList.SubItems(3) = CStr(rsLIST!CompanyName)
                        End If
                        
                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub GetVoucherPrepared()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Voucher No", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Voucher Date", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Payable to", .ListView1.Width / 2.5
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 6

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select *  from ODASMVoucher,ODASPAccount Where ODASMVoucher.Prepared = 'Y' and ODASMVoucher.AccountNo = ODASPAccount.AccountNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!vOUCHERnO))
                        
                        If Not IsNull(rsLIST!VoucherDate) Then
                                MyList.SubItems(1) = CStr(rsLIST!VoucherDate)
                        End If

                        If Not IsNull(rsLIST!CompanyName) Then
                                MyList.SubItems(2) = CStr(rsLIST!CompanyName)
                        End If

                        If Not IsNull(rsLIST!Amount) Then
                                MyList.SubItems(3) = CStr(rsLIST!Amount)
                        End If
                        
                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub GetVouchers()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView3.ListItems.Clear
                .ListView3.ColumnHeaders.Clear
                
                .ListView3.ColumnHeaders.Add , , "Voucher No", .ListView3.Width / 4
                .ListView3.ColumnHeaders.Add , , "Voucher Date", .ListView3.Width / 4
                .ListView3.ColumnHeaders.Add , , "Account No", .ListView3.Width / 4
                .ListView3.ColumnHeaders.Add , , "Amount", .ListView3.Width / 4

                .ListView3.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select *  from ODASMVoucher Where ChequePrepared = 'N' and Authorized = 'Y' and AccountNo = '" & frmODASMCheck.txtAccountNo.Text & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!vOUCHERnO))
                        
                        If Not IsNull(rsLIST!VoucherDate) Then
                                MyList.SubItems(1) = CStr(rsLIST!VoucherDate)
                        End If

                        If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                        End If

                        If Not IsNull(rsLIST!Amount) Then
                                MyList.SubItems(3) = CStr(rsLIST!Amount)
                        End If
                        
                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub GetVoucherAPPROVED()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Voucher No", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Voucher Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Account No", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select *  from ODASMVoucher Where Approved = 'Y' and Authorized = 'N';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!vOUCHERnO))
                        
                        If Not IsNull(rsLIST!VoucherDate) Then
                                MyList.SubItems(1) = CStr(rsLIST!VoucherDate)
                        End If

                        If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                        End If

                        If Not IsNull(rsLIST!Amount) Then
                                MyList.SubItems(3) = CStr(rsLIST!Amount)
                        End If
                        
                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub
Public Sub GetVoucherAPPROVED2()
On Error GoTo err
    
        With ALISFOManager
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Voucher No", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Voucher Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Account No", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select *  from ODASMVoucher Where Approved = 'Y' and Authorized = 'N';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!vOUCHERnO))
                        
                        If Not IsNull(rsLIST!VoucherDate) Then
                                MyList.SubItems(1) = CStr(rsLIST!VoucherDate)
                        End If

                        If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                        End If

                        If Not IsNull(rsLIST!Amount) Then
                                MyList.SubItems(3) = CStr(rsLIST!Amount)
                        End If
                        
                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub GetPreviousPayment()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView3.ListItems.Clear
                .ListView3.ColumnHeaders.Clear
                
                .ListView3.ColumnHeaders.Add , , "Cheque No", .ListView3.Width / 6
                .ListView3.ColumnHeaders.Add , , "RequsitionNo", .ListView3.Width / 4
        
                .ListView3.ColumnHeaders.Add , , "Amount", .ListView3.Width / 6
                .ListView3.ColumnHeaders.Add , , "Amount Due", .ListView3.Width / 6
                .ListView3.ColumnHeaders.Add , , "Payment Flag", .ListView3.Width / 8

                .ListView3.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select ALISMChequeEntry.ChequeNo, ALISMChequeEntry.VoucherNo, ALISMChequeEntry.Amount, ALISMChequeEntry.AmountDue, ALISMChequeEntry.PaymentFlag from ALISMChequeEntry Where ALISMChequeEntry.ChequeNo = '" & frmALISMCheckReverse.txtChequeNo & "' ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!ChequeNo))
                        
                        If Not IsNull(rsLIST!Amount) Then
                                MyList.SubItems(2) = CStr(rsLIST!Amount)
                        End If
                        
                        If Not IsNull(rsLIST!AmountDue) Then
                                MyList.SubItems(3) = CStr(rsLIST!AmountDue)
                        End If
                        
                        If Not IsNull(rsLIST!PaymentFlag) Then
                                MyList.SubItems(4) = CStr(rsLIST!PaymentFlag)
                        End If
                        
                        If bmakePAYMENT = True Then
                                If Not IsNull(rsLIST!PayeeDetails) Then
                                        MyList.SubItems(1) = CStr(rsLIST!PayeeDetails)
                                End If
                        ElseIf breversePAYMENT = True Then
                                If Not IsNull(rsLIST!vOUCHERnO) Then
                                        MyList.SubItems(1) = CStr(rsLIST!vOUCHERnO)
                                End If
                        End If
                        
                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub GetPreviousPayment1()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView3.ListItems.Clear
                .ListView3.ColumnHeaders.Clear
                
                .ListView3.ColumnHeaders.Add , , "Voucher No", .ListView3.Width / 6
                .ListView3.ColumnHeaders.Add , , "Payee Details", .ListView3.Width / 4
                .ListView3.ColumnHeaders.Add , , "Amount", .ListView3.Width / 6
                .ListView3.ColumnHeaders.Add , , "Amount Due", .ListView3.Width / 6
                .ListView3.ColumnHeaders.Add , , "Payment Flag", .ListView3.Width / 8


                .ListView3.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select ALISMChequeEntry.VoucherNo, ALISMCheque.PayeeDetails, ALISMChequeEntry.Amount, ALISMChequeEntry.AmountDue, ALISMChequeEntry.PaymentFlag from ALISMChequeEntry, ALISMCheque Where ALISMCheque.ChequeNo = ALISMChequeEntry.ChequeNo and ALISMChequeEntry.VoucherNo = '" & Screen.ActiveForm.txtVoucherNo.Text & "' and ALISMChequeEntry.PaymentFlag <> 'C';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                    Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!vOUCHERnO))
                        
                        If Not IsNull(rsLIST!Amount) Then
                                MyList.SubItems(2) = CStr(rsLIST!Amount)
                        End If
                        
                        If Not IsNull(rsLIST!AmountDue) Then
                                MyList.SubItems(3) = CStr(rsLIST!AmountDue)
                        End If
                        
                        If Not IsNull(rsLIST!PaymentFlag) Then
                                MyList.SubItems(4) = CStr(rsLIST!PaymentFlag)
                        End If
                        
                        If bmakePAYMENT = True Then
                                If Not IsNull(rsLIST!PayeeDetails) Then
                                        MyList.SubItems(1) = CStr(rsLIST!PayeeDetails)
                                End If
                        ElseIf breversePAYMENT = True Then
                                If Not IsNull(rsLIST!vOUCHERnO) Then
                                        MyList.SubItems(1) = CStr(rsLIST!vOUCHERnO)
                                End If
                        End If
                        
                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub GetVoucherAUTHORIZED()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Voucher No", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Voucher Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Account No", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select *  from ODASMVoucher Where ChequePrepared = 'N' and Authorized = 'Y';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!vOUCHERnO))
                        
                        If Not IsNull(rsLIST!VoucherDate) Then
                                MyList.SubItems(1) = CStr(rsLIST!VoucherDate)
                        End If

                        If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(2) = CStr(rsLIST!AccountNo)
                        End If

                        If Not IsNull(rsLIST!Amount) Then
                                MyList.SubItems(3) = CStr(rsLIST!Amount)
                        End If
                        
                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub GetInvoicesREQUISITIONED()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView2.ListItems.Clear
                .ListView2.ColumnHeaders.Clear
                
                .ListView2.ColumnHeaders.Add , , "Item No", .ListView2.Width / 4
                .ListView2.ColumnHeaders.Add , , "Invoice No", .ListView2.Width / 4
                .ListView2.ColumnHeaders.Add , , "LPO No", .ListView2.Width / 4
                .ListView2.ColumnHeaders.Add , , "Amount", .ListView2.Width / 4

                .ListView2.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select *  from ODASMVoucherItem Where VoucherNo = '" & .txtVoucherNo & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!VoucherItemNo))
                        
                        If Not IsNull(rsLIST!LPONo) Then
                                MyList.SubItems(2) = CStr(rsLIST!LPONo)
                        End If

                        If Not IsNull(rsLIST!DocumentNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!DocumentNo)
                        End If

                        If Not IsNull(rsLIST!AmountPaid) Then
                                MyList.SubItems(3) = CStr(rsLIST!AmountPaid)
                        End If
                        
                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub getREQUISITIONS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Req No", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Req Date", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Price Excl", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "VAT", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Price Incl", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Paid", .ListView1.Width / 8

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select * from ODASMRequisition R Where (R.InvoiceReceived is null or R.InvoiceReceived = 'N') and R.RequisitionNo = '" & frmODASMReceiveinvoice.txtRequisitionNo.Text & " ';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!RequisitionNo))

                        If Not IsNull(rsLIST!RequisitionDate) Then
                            MyList.SubItems(2) = CStr(rsLIST!RequisitionDate)
                        End If
                        
                        If Not IsNull(rsLIST!PriceExclusive) Then
                            MyList.SubItems(3) = CStr(rsLIST!PriceExclusive)
                        End If
                        
                        If Not IsNull(rsLIST!VatAmount) Then
                            MyList.SubItems(4) = CStr(rsLIST!VatAmount)
                        End If

                        If Not IsNull(rsLIST!PriceInclusive) Then
                            MyList.SubItems(5) = CStr(rsLIST!PriceInclusive)
                        End If
                        
                        If Not IsNull(rsLIST!paid) Then
                            MyList.SubItems(6) = CStr(rsLIST!paid)
                        End If
                
                rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub
Public Sub showBRIEFSNOTINVOICED()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Job Brief No", .ListView1.Width / 7 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Product", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Customer Name", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Description Of Order", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Brief Date", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Balance", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Cost", .ListView1.Width / 7

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMJobBrief J, ODASPAccount C Where J.AccountNo = C.AccountNo AND J.CreditAuthorized = 'Y' and J.TotalPrice is not null AND (J.scheduled = 'N' or J.scheduled is null) order by j.JobBriefNo ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                    
                    If rsLIST!Balance = 0 Then GoTo Continue
                    
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!JobBriefNo))
                        
                        If Not IsNull(rsLIST!ProductCode) Then
                            MyList.SubItems(1) = CStr(rsLIST!ProductCode)
                        End If

                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(2) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!descriptionOfOrder) Then
                            MyList.SubItems(3) = CStr(rsLIST!descriptionOfOrder)
                        End If
                        
                        If Not IsNull(rsLIST!JobBriefDate) Then
                            MyList.SubItems(4) = Format(rsLIST!JobBriefDate, "dd/mm/yyyy")
                        End If
                        
                        If Not IsNull(rsLIST!Balance) Then
                            MyList.SubItems(5) = CStr(rsLIST!Balance)
                        End If
                        
                        If Not IsNull(rsLIST!TotalOverallCost) Then
                            MyList.SubItems(6) = CStr(rsLIST!TotalOverallCost)
                        End If
Continue:
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub
Public Sub showCONTRACTNOTINVOICED()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Job Brief No", .ListView1.Width / 11 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Product", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Customer Name", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Description Of Order", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Brief Date", .ListView1.Width / 12
                .ListView1.ColumnHeaders.Add , , "Balance", .ListView1.Width / 9
                .ListView1.ColumnHeaders.Add , , "Cost", .ListView1.Width / 9
                

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMJobBrief J, ODASPAccount C Where J.AccountNo = C.AccountNo AND J.Contract='Y' AND J.CreditAuthorized = 'Y' and J.TotalPrice is not null AND (J.scheduled = 'N' or J.scheduled is null) order by j.JobBriefNo dESC ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                    
                    If rsLIST!Balance = 0 Then GoTo Continue
                    
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!JobBriefNo))
                        
                        If Not IsNull(rsLIST!ProductCode) Then
                            MyList.SubItems(1) = CStr(rsLIST!ProductCode)
                        End If

                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(2) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!descriptionOfOrder) Then
                            MyList.SubItems(3) = CStr(rsLIST!descriptionOfOrder)
                        End If
                        
                        If Not IsNull(rsLIST!JobBriefDate) Then
                            MyList.SubItems(4) = Format(rsLIST!JobBriefDate, "dd/mm/yyyy")
                        End If
                        
                        If Not IsNull(rsLIST!Balance) Then
                            MyList.SubItems(5) = FormatNumber(rsLIST!Balance, 2)
                        End If
                        
                        If Not IsNull(rsLIST!TotalPrice) Then
                            MyList.SubItems(6) = FormatNumber(rsLIST!TotalPrice, 2)
                        End If
Continue:
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showBRIEFSINVOICED()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Job Brief No", .ListView1.Width / 7 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Product", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Customer Name", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Description Of Order", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Brief Date", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Balance", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Cost", .ListView1.Width / 7

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMJobBrief J, ODASPAccount C Where J.AccountNo = C.AccountNo AND J.CreditAuthorized = 'Y' and J.TotalPrice is not null AND J.scheduled = 'Y' order by j.JobBriefNo ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                    
                    If rsLIST!Balance = 0 Then GoTo Continue
                    
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!JobBriefNo))
                        
                        If Not IsNull(rsLIST!ProductCode) Then
                            MyList.SubItems(1) = CStr(rsLIST!ProductCode)
                        End If

                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(2) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!descriptionOfOrder) Then
                            MyList.SubItems(3) = CStr(rsLIST!descriptionOfOrder)
                        End If
                        
                        If Not IsNull(rsLIST!JobBriefDate) Then
                            MyList.SubItems(4) = Format(rsLIST!JobBriefDate, "dd/mm/yyyy")
                        End If
                        
                        If Not IsNull(rsLIST!Balance) Then
                            MyList.SubItems(5) = CStr(rsLIST!Balance)
                        End If
                        
                        If Not IsNull(rsLIST!TotalOverallCost) Then
                            MyList.SubItems(6) = CStr(rsLIST!TotalOverallCost)
                        End If
Continue:
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showRECEIPTSCHEDULE()
On Error GoTo err
    
        With ALISFOManager
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Reference", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Brief No", .ListView1.Width / 8 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Product", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Customer Name", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Description Of Order", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Brief Date", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Balance", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Inst #", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 8

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMJobBrief J, ODASPAccount C, ODASMJobBriefInstallment I Where I.Invoiced = 'N' and I.JobBriefNo = J.JobBriefNo and J.AccountNo = C.AccountNo and (I.PaymentDueDate >='" & Format(Date, "yyyy/mm/dd") & "')  Order by I.InstallmentNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem

                While Not rsLIST.EOF
                    
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!InvoiceReference))
                            
                            If Not IsNull(rsLIST!JobBriefNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!JobBriefNo)
                            End If
    
                            If Not IsNull(rsLIST!ProductCode) Then
                                MyList.SubItems(2) = CStr(rsLIST!ProductCode)
                            End If
    
                            If Not IsNull(rsLIST!CompanyName) Then
                                MyList.SubItems(3) = CStr(rsLIST!CompanyName)
                            End If
                            
                            If Not IsNull(rsLIST!descriptionOfOrder) Then
                                MyList.SubItems(4) = CStr(rsLIST!descriptionOfOrder)
                            End If
                            
                            If Not IsNull(rsLIST!JobBriefDate) Then
                                MyList.SubItems(5) = CStr(rsLIST!JobBriefDate)
                            End If
                            
                            If Not IsNull(rsLIST!Balance) Then
                                MyList.SubItems(6) = CStr(rsLIST!Balance)
                            End If
                            
                            If Not IsNull(rsLIST!InstallmentNo) Then
                                MyList.SubItems(7) = CStr(rsLIST!InstallmentNo)
                            End If
                            
                            If Not IsNull(rsLIST!Amount) Then
                                MyList.SubItems(8) = CStr(rsLIST!Amount)
                            End If
    
                         rsLIST.MoveNext
                    Wend
                    Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showBRIEFSNOTAUTHORIZED()
On Error GoTo err
    
        With ALISFOManager
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Job Brief No", .ListView1.Width / 7 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Product", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Customer Name", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Description Of Order", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Brief Date", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Balance", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Cost", .ListView1.Width / 7

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMJobBrief J, ODASPAccount C Where (J.TotalPrice is not null ) and J.AccountNo = C.AccountNo and (J.CreditAuthorized = 'N' or J.CreditAuthorized is null) order by j.JobBriefNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!JobBriefNo))
                        
                        If Not IsNull(rsLIST!ProductCode) Then
                            MyList.SubItems(1) = CStr(rsLIST!ProductCode)
                        End If

                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(2) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!descriptionOfOrder) Then
                            MyList.SubItems(3) = CStr(rsLIST!descriptionOfOrder)
                        End If
                        
                        If Not IsNull(rsLIST!JobBriefDate) Then
                            MyList.SubItems(4) = CStr(rsLIST!JobBriefDate)
                        End If
                        
                        If Not IsNull(rsLIST!Balance) Then
                            MyList.SubItems(5) = CStr(rsLIST!Balance)
                        End If
                        
                        If Not IsNull(rsLIST!TotalOverallCost) Then
                            MyList.SubItems(6) = CStr(rsLIST!TotalOverallCost)
                        End If

                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showBRIEFSAUTHORIZED()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Job Brief No", .ListView1.Width / 7 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Product", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Customer Name", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Description Of Order", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Auth Date", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Balance", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Auth #", .ListView1.Width / 7

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMJobBrief J, ODASPAccount C, ODASMCreditAuthorization CR Where CR.JobBriefNo = J.JobBriefNo and J.AccountNo = C.AccountNo and J.CreditAuthorized = 'Y' and CR.CurrentPeriod = '" & CurrentPeriod & "' order by j.JobBriefNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!JobBriefNo))
                        
                        If Not IsNull(rsLIST!ProductCode) Then
                            MyList.SubItems(1) = CStr(rsLIST!ProductCode)
                        End If

                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(2) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!descriptionOfOrder) Then
                            MyList.SubItems(3) = CStr(rsLIST!descriptionOfOrder)
                        End If
                        
                        If Not IsNull(rsLIST!DateAuthorized) Then
                            MyList.SubItems(4) = CStr(rsLIST!DateAuthorized)
                        End If
                        
                        If Not IsNull(rsLIST!Balance) Then
                            MyList.SubItems(5) = CStr(rsLIST!Balance)
                        End If
                        
                        If Not IsNull(rsLIST!AuthorizationNo) Then
                            MyList.SubItems(6) = CStr(rsLIST!AuthorizationNo)
                        End If

                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showBRIEFRECEIPTS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView2.ListItems.Clear
                .ListView2.ColumnHeaders.Clear
                
                .ListView2.ColumnHeaders.Add , , "ReceiptNo", .ListView2.Width / 3 ', lvwColumnCenter
                .ListView2.ColumnHeaders.Add , , "ReceiptDate", .ListView2.Width / 3
                .ListView2.ColumnHeaders.Add , , "Receipt Amount", .ListView2.Width / 3

                .ListView2.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ALISMReceiptDetails Where DocumentNo = '" & .txtJobBriefNo & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!ReceiptNo))
                        

                        If Not IsNull(rsLIST!TransactionDate) Then
                            MyList.SubItems(1) = CStr(rsLIST!TransactionDate)
                        End If
                        
                        If Not IsNull(rsLIST!TransactionAmount) Then
                            MyList.SubItems(2) = CStr(rsLIST!TransactionAmount)
                        End If
 
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showBRIEFINVOICESRECeived()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView3.ListItems.Clear
                .ListView3.ColumnHeaders.Clear
                
                .ListView3.ColumnHeaders.Add , , "Invoice No", .ListView3.Width / 3 ', lvwColumnCenter
                .ListView3.ColumnHeaders.Add , , "Invoice Date", .ListView3.Width / 3
                .ListView3.ColumnHeaders.Add , , "Amount", .ListView3.Width / 3

                .ListView3.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMInvoice Where JobCardNo = '" & .txtJobBriefNo & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
                        
                        If Not IsNull(rsLIST!InvoiceDate) Then
                            MyList.SubItems(1) = CStr(rsLIST!InvoiceDate)
                        End If
                        
                        If Not IsNull(rsLIST!PriceInclusive) Then
                            MyList.SubItems(2) = CStr(rsLIST!PriceInclusive)
                        End If
 
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showINVOICEDETAILS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView3.ListItems.Clear
                .ListView3.ColumnHeaders.Clear
                
                .ListView3.ColumnHeaders.Add , , "Item No", .ListView3.Width / 5 ', lvwColumnCenter
                .ListView3.ColumnHeaders.Add , , "Location", .ListView3.Width / 5 ', lvwColumnCenter
                .ListView3.ColumnHeaders.Add , , "Size", .ListView3.Width / 5
                .ListView3.ColumnHeaders.Add , , "Siding", .ListView3.Width / 5
                .ListView3.ColumnHeaders.Add , , "Amount", .ListView3.Width / 5

                .ListView3.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT JBI.JobBriefItemNo, JBI.PhysicalLocation, JBI.MediaSize, S.SidingDescription, JBI.PriceExclusive FROM ODASMJobBriefItems JBI, ODASPsiding S Where JobBriefNo = '" & .txtJobBriefNo & "' and JBI.SidingCode = S.SidingCode ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!JobBriefItemNo))

                        If Not IsNull(rsLIST!PhysicalLocation) Then
                            MyList.SubItems(1) = CStr(rsLIST!PhysicalLocation)
                        End If
                        
                        If Not IsNull(rsLIST!MediaSize) Then
                            MyList.SubItems(2) = CStr(rsLIST!MediaSize)
                        End If

                        If Not IsNull(rsLIST!SidingDescription) Then
                            MyList.SubItems(3) = CStr(rsLIST!SidingDescription)
                        End If

                        If Not IsNull(rsLIST!PriceExclusive) Then
                            MyList.SubItems(4) = CStr(rsLIST!PriceExclusive)
                        End If
 
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showBRIEFINACCOUNT()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView6.ListItems.Clear
                .ListView6.ColumnHeaders.Clear
                
                .ListView6.ColumnHeaders.Add , , "JobBrief No", .ListView6.Width / 3 ', lvwColumnCenter
                .ListView6.ColumnHeaders.Add , , "Due Date", .ListView6.Width / 3
                .ListView6.ColumnHeaders.Add , , "Amount", .ListView6.Width / 3
                .ListView6.ColumnHeaders.Add , , "Inv Ref", .ListView6.Width / 3

                .ListView6.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMJobBriefInstallment I, ODASMJobBrief J Where J.AccountNo = '" & .txtAccountNo & "' and (I.Invoiced = 'N' or I.Invoiced is null) and J.JobBriefNo = I.jobBriefNo and I.PaymentDueDate <= '" & Format(Date, "yyyy/mm/dd") & "' ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView6.ListItems.Add(, , CStr(rsLIST!JobBriefNo))
                        
                        If Not IsNull(rsLIST!PaymentDueDate) Then
                            MyList.SubItems(1) = CStr(rsLIST!PaymentDueDate)
                        End If
                        
                        If Not IsNull(rsLIST!Amount) Then
                            MyList.SubItems(2) = CStr(rsLIST!Amount)
                        End If
                        
                        If Not IsNull(rsLIST!InvoiceReference) Then
                            MyList.SubItems(3) = CStr(rsLIST!InvoiceReference)
                        End If

 
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showBRIEFINVOICESsenT()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Invoice No", .ListView1.Width / 3 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Invoice Date", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 3

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMInvoiceSENT Where JobBriefNo = '" & .txtJobBriefNo & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
                        
                        If Not IsNull(rsLIST!InvoiceDate) Then
                            MyList.SubItems(1) = CStr(rsLIST!InvoiceDate)
                        End If
                        
                        If Not IsNull(rsLIST!PriceInclusive) Then
                            MyList.SubItems(2) = CStr(rsLIST!PriceInclusive)
                        End If
 
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showINVOICEitems()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView7.ListItems.Clear
                .ListView7.ColumnHeaders.Clear
                
                .ListView7.ColumnHeaders.Add , , "Item No", .ListView7.Width / 4 ', lvwColumnCenter
                .ListView7.ColumnHeaders.Add , , "Invoice No", .ListView7.Width / 4 ', lvwColumnCenter
                .ListView7.ColumnHeaders.Add , , "Invoice Date", .ListView7.Width / 4
                .ListView7.ColumnHeaders.Add , , "Amount", .ListView7.Width / 4

                .ListView7.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMInvoiceItemsSENT IT, ODASMInvoiceSeNT INV   Where IT.InvoiceNo = INV.InvoiceNo and IT.InvoiceNo = '" & .txtInvoiceNo & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView7.ListItems.Add(, , CStr(rsLIST!ItemNo))
                        
                        If Not IsNull(rsLIST!InvoiceNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!InvoiceNo)
                        End If

                        If Not IsNull(rsLIST!InvoiceDate) Then
                            MyList.SubItems(2) = CStr(rsLIST!InvoiceDate)
                        End If
                        
                        If Not IsNull(rsLIST!PriceInclusive) Then
                            MyList.SubItems(3) = CStr(rsLIST!PriceInclusive)
                        End If
 
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showINVOICEeNTRIES()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Item No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Invoice No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Invoice Date", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Exclusivet", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "VAT", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "INclusive", .ListView1.Width / 6

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMInvoiceItemsSENT IT, ODASMInvoiceSeNT INV   Where IT.InvoiceNo = INV.InvoiceNo and IT.InvoiceNo = '" & .txtInvoiceNo & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ItemNo))
                        
                        If Not IsNull(rsLIST!InvoiceNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!InvoiceNo)
                        End If

                        If Not IsNull(rsLIST!InvoiceDate) Then
                            MyList.SubItems(2) = CStr(rsLIST!InvoiceDate)
                        End If
                        
                        If Not IsNull(rsLIST!PriceExclusive) Then
                            MyList.SubItems(3) = CStr(rsLIST!PriceExclusive)
                        End If
                        
                        If Not IsNull(rsLIST!VatAmount) Then
                            MyList.SubItems(4) = CStr(rsLIST!VatAmount)
                        End If
                        
                        If Not IsNull(rsLIST!PriceInclusive) Then
                            MyList.SubItems(5) = CStr(rsLIST!PriceInclusive)
                        End If
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showINVOICESprepared()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Invoice No", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Invoice Date", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Company", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Amount(Excl)", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMInvoiceSeNT I, ODASPAccount A Where  A.AccountNo = I.AccountNo and I.Prepared = 'Y' and I.Approved = 'N' ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
                        
                        If Not IsNull(rsLIST!InvoiceDate) Then
                            MyList.SubItems(1) = CStr(rsLIST!InvoiceDate)
                        End If

                        If Not IsNull(rsLIST!InvoiceDescription) Then
                            MyList.SubItems(2) = CStr(rsLIST!InvoiceDescription)
                        End If
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(3) = CStr(rsLIST!CompanyName)
                        End If

                        If Not IsNull(rsLIST!PriceInclusive) Then
                            MyList.SubItems(4) = CStr(rsLIST!PriceExclusive)
                        End If
 
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showINVOICES()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView3.ListItems.Clear
                .ListView3.ColumnHeaders.Clear
                
                .ListView3.ColumnHeaders.Add , , "Invoice No", .ListView3.Width / 3 ', lvwColumnCenter
                .ListView3.ColumnHeaders.Add , , "Company", .ListView3.Width / 3 ', lvwColumnCenter
                .ListView3.ColumnHeaders.Add , , "Invoice Date", .ListView3.Width / 3
                .ListView3.ColumnHeaders.Add , , "Amount(Incl)", .ListView3.Width / 3

                .ListView3.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMInvoiceSeNT I, ODASPAccount A Where  A.AccountNo = I.AccountNo and I.Prepared = 'Y' and (I.Paid = 'N' or I.Paid is null) and Approved = 'Y';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(1) = CStr(rsLIST!CompanyName)
                        End If

                        If Not IsNull(rsLIST!InvoiceDate) Then
                            MyList.SubItems(2) = CStr(rsLIST!InvoiceDate)
                        End If

                        If Not IsNull(rsLIST!PriceInclusive) Then
                            MyList.SubItems(3) = CStr(rsLIST!PriceInclusive)
                        End If
 
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showJOBBRIEFINVOICES()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView3.ListItems.Clear
                .ListView3.ColumnHeaders.Clear
                
                .ListView3.ColumnHeaders.Add , , "Invoice No", .ListView3.Width / 3 ', lvwColumnCenter
                .ListView3.ColumnHeaders.Add , , "Invoice Date", .ListView3.Width / 3 ', lvwColumnCenter
                .ListView3.ColumnHeaders.Add , , "Company", .ListView3.Width / 3
                .ListView3.ColumnHeaders.Add , , "Amount(Incl)", .ListView3.Width / 3

                .ListView3.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMInvoiceSeNT I, ODASPAccount A Where I.JobBriefNo = '" & Screen.ActiveForm.txtJobBriefNo.Text & "' and A.AccountNo = I.AccountNo  ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
                        
                        If Not IsNull(rsLIST!InvoiceDate) Then
                            MyList.SubItems(1) = CStr(rsLIST!InvoiceDate)
                        End If

                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(2) = CStr(rsLIST!CompanyName)
                        End If

                        If Not IsNull(rsLIST!PriceInclusive) Then
                            MyList.SubItems(3) = CStr(rsLIST!PriceInclusive)
                        End If
 
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showINVOICESApproved()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Invoice No", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Invoice Date", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Company", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Amount(Excl)", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMInvoiceSeNT I, ODASPAccount A Where  A.AccountNo = I.AccountNo and I.Approved = 'Y' and I.Authorized = 'N' ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
                        
                        If Not IsNull(rsLIST!InvoiceDate) Then
                            MyList.SubItems(1) = CStr(rsLIST!InvoiceDate)
                        End If

                        If Not IsNull(rsLIST!InvoiceDescription) Then
                            MyList.SubItems(2) = CStr(rsLIST!InvoiceDescription)
                        End If
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(3) = CStr(rsLIST!CompanyName)
                        End If

                        If Not IsNull(rsLIST!PriceInclusive) Then
                            MyList.SubItems(4) = CStr(rsLIST!PriceExclusive)
                        End If
 
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showINVOICESAuthorized()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Invoice No", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Invoice Date", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Company", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Amount(Excl)", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMInvoiceSeNT I, ODASPAccount A Where  A.AccountNo = I.AccountNo and I.Authorized = 'Y' ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
                        
                        If Not IsNull(rsLIST!InvoiceDate) Then
                            MyList.SubItems(1) = CStr(rsLIST!InvoiceDate)
                        End If

                        If Not IsNull(rsLIST!InvoiceDescription) Then
                            MyList.SubItems(2) = CStr(rsLIST!InvoiceDescription)
                        End If
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(3) = CStr(rsLIST!CompanyName)
                        End If

                        If Not IsNull(rsLIST!PriceInclusive) Then
                            MyList.SubItems(4) = CStr(rsLIST!PriceExclusive)
                        End If
 
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showINVOICESPrinted()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Invoice No", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Invoice Date", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Company", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Amount(Excl)", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMInvoiceSeNT I, ODASPAccount A Where  A.AccountNo = I.AccountNo and I.Printed = 'Y' and (I.Paid = 'N' or I.Paid is null) ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
                        
                        If Not IsNull(rsLIST!InvoiceDate) Then
                            MyList.SubItems(1) = CStr(rsLIST!InvoiceDate)
                        End If

                        If Not IsNull(rsLIST!InvoiceDescription) Then
                            MyList.SubItems(2) = CStr(rsLIST!InvoiceDescription)
                        End If
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(3) = CStr(rsLIST!CompanyName)
                        End If

                        If Not IsNull(rsLIST!PriceInclusive) Then
                            MyList.SubItems(4) = CStr(rsLIST!PriceExclusive)
                        End If
 
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showINVOICESPaid()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Invoice No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Invoice Date", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Company", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Amount(Excl)", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Received TD", .ListView1.Width / 6

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMInvoiceSeNT I, ODASPAccount A Where  A.AccountNo = I.AccountNo and I.Paid = 'Y' and I.currentPeriod = '" & Screen.ActiveForm.txtCurrentPeriod & "' ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
                        
                        If Not IsNull(rsLIST!InvoiceDate) Then
                            MyList.SubItems(1) = CStr(rsLIST!InvoiceDate)
                        End If

                        If Not IsNull(rsLIST!InvoiceDescription) Then
                            MyList.SubItems(2) = CStr(rsLIST!InvoiceDescription)
                        End If
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(3) = CStr(rsLIST!CompanyName)
                        End If

                        If Not IsNull(rsLIST!PriceInclusive) Then
                            MyList.SubItems(4) = CStr(rsLIST!PriceExclusive)
                        End If
                        
                         If Not IsNull(rsLIST!ReceivedToDate) Then
                            MyList.SubItems(5) = CStr(rsLIST!ReceivedToDate)
                        End If

                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showINVOICEitem()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Invoice No", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Invoice Date", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Company", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Amount(Excl)", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMInvoiceSeNT I, ODASPAccount A Where  A.AccountNo = I.AccountNo and I.Authorized = 'N' and I.Approved = 'Y' ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
                        
                        If Not IsNull(rsLIST!InvoiceDate) Then
                            MyList.SubItems(1) = CStr(rsLIST!InvoiceDate)
                        End If

                        If Not IsNull(rsLIST!InvoiceDescription) Then
                            MyList.SubItems(2) = CStr(rsLIST!InvoiceDescription)
                        End If
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(3) = CStr(rsLIST!CompanyName)
                        End If

                        If Not IsNull(rsLIST!PriceInclusive) Then
                            MyList.SubItems(4) = CStr(rsLIST!PriceExclusive)
                        End If
 
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showLPOINVOICES()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView2.ListItems.Clear
                .ListView2.ColumnHeaders.Clear
                
                .ListView2.ColumnHeaders.Add , , "InvoiceNo ", .ListView2.Width / 4 ', lvwColumnCenter
                .ListView2.ColumnHeaders.Add , , "Invoice Date", .ListView2.Width / 4 ', lvwColumnCenter
                .ListView2.ColumnHeaders.Add , , "Amount(Excl)", .ListView2.Width / 4
                .ListView2.ColumnHeaders.Add , , "LPONo", .ListView2.Width / 4
                .ListView2.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMInvoice Where LPONo = '" & .txtLPONo & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
                        
                        If Not IsNull(rsLIST!InvoiceDate) Then
                            MyList.SubItems(1) = CStr(rsLIST!InvoiceDate)
                        End If

                        If Not IsNull(rsLIST!PriceExclusive) Then
                            MyList.SubItems(2) = CStr(rsLIST!PriceExclusive)
                        End If
                        
                        If Not IsNull(rsLIST!LPONo) Then
                            MyList.SubItems(3) = CStr(rsLIST!LPONo)
                        End If

                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub showALLINVOICESRECEIVED()
On Error GoTo err
    
        With ALISFOManager
                .ListView1.FlatScrollBar = True
                
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "InvoiceNo ", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Invoice Date", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Amount(Excl)", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "LPONo", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Received From", .ListView1.Width / 2.5
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMInvoice,ODASPAccount Where ODASMInvoice.AccountNo = ODASPAccount.AccountNo and ODASMInvoice.Prepared = 'Y' and ODASMInvoice.Approved = 'N';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!InvoiceNo))
                        
                        If Not IsNull(rsLIST!InvoiceDate) Then
                            MyList.SubItems(1) = CStr(rsLIST!InvoiceDate)
                        End If

                        If Not IsNull(rsLIST!PriceExclusive) Then
                            MyList.SubItems(2) = CStr(rsLIST!PriceExclusive)
                        End If
                        
                        If Not IsNull(rsLIST!LPONo) Then
                            MyList.SubItems(3) = CStr(rsLIST!LPONo)
                        End If
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(4) = CStr(rsLIST!CompanyName)
                        End If

                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage
End Sub

Public Sub showALLPaymentMODE()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView4.ListItems.Clear
                .ListView4.ColumnHeaders.Clear
                .ListView4.ColumnHeaders.Add , , "Payment Mode", .ListView4.Width / 3
                .ListView4.ColumnHeaders.Add , , "Description", .ListView4.Width / 3
                .ListView4.ColumnHeaders.Add , , "CoverPeriod", .ListView4.Width / 3

                .ListView4.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPaymentMode ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView4.ListItems.Add(, , CStr(rsLIST!PaymentMode))
                        
                        If Not IsNull(rsLIST!PaymentModeDescription) Then
                            MyList.SubItems(1) = CStr(rsLIST!PaymentModeDescription)
                        End If
                        
                        If Not IsNull(rsLIST!CoverPeriod) Then
                            MyList.SubItems(2) = CStr(rsLIST!CoverPeriod)
                        End If

                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage
End Sub

Public Sub showALLPaymentMETHOD()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView5.ListItems.Clear
                .ListView5.ColumnHeaders.Clear
                .ListView5.ColumnHeaders.Add , , "Method", .ListView5.Width / 3
                .ListView5.ColumnHeaders.Add , , "Description", .ListView5.Width / 3
                .ListView5.ColumnHeaders.Add , , "Ac Details?", .ListView5.Width / 3

                .ListView5.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPPaymentMethod ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView5.ListItems.Add(, , CStr(rsLIST!PaymentMethod))
                        
                        If Not IsNull(rsLIST!PaymentMethodDescription) Then
                            MyList.SubItems(1) = CStr(rsLIST!PaymentMethodDescription)
                        End If
                        
                        If Not IsNull(rsLIST!AccountNoRequired) Then
                            MyList.SubItems(2) = CStr(rsLIST!AccountNoRequired)
                        End If

                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage
End Sub
