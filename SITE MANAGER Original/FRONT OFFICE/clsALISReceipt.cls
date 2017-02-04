VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsALISReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub addRECORD()
    clearRECORD
    enableALLRECORD
    disableSButtons
End Sub
Public Sub getMONTH()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Month", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!MonthCode))
                        
                        If Not IsNull(rsLIST!MonthDescription) Then
                            MyList.SubItems(1) = CStr(rsLIST!MonthDescription)
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
                
                strSQL = "Select ODASMJobBrief.JobBriefNo, ODASPAccount.CompanyName, ODASPAccount.AccountNo, ODASMJobBrief.Status from ODASMJobBrief, ODASPAccount Where ODASMJobBrief.AccountNo = ODASPAccount.AccountNo order by JobBriefNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!JobBriefNo))
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(1) = CStr(rsLIST!CompanyName)
                        End If

                        If Not IsNull(rsLIST!Accountno) Then
                            MyList.SubItems(2) = CStr(rsLIST!Accountno)
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
Public Sub getSUSPENSEPOLICY()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Policy No", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "CompanyName", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Received", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Suspense", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Status", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select ODASMJobBrief.JobBriefNo, ODASPAccount.CompanyName, ODASMJobBrief.ReceivedTodate, ODASMJobBrief.SuspenseAccount, ODASMJobBrief.Status from ODASMJobBrief, ODASPAccount Where ODASMJobBrief.AccountNo = ODASPAccount.AccountNo order by JobBriefNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!JobBriefNo))
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(1) = CStr(rsLIST!CompanyName)
                        End If

                        
                        If Not IsNull(rsLIST!ReceivedToDate) Then
                            MyList.SubItems(2) = CStr(rsLIST!ReceivedToDate)
                        End If

                        If Not IsNull(rsLIST!SuspenseAccount) Then
                            MyList.SubItems(3) = CStr(rsLIST!SuspenseAccount)
                        End If
                        
                        If Not IsNull(rsLIST!Status) Then
                            MyList.SubItems(4) = CStr(rsLIST!Status)
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


Public Sub loadRENEWAL()
On Error GoTo err
    
    Set rsCONTROL = New ADODB.Recordset
    strSQL = "Select * from ALISPMonth where monthCode = '" & frmALISMRenewalNotice.txtMonth.Text & "'"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsCONTROL
            If .EOF Or .BOF Then
                    MsgBox "The month Setup NOT in Place", vbOKOnly
                    Exit Sub
            End If
            
            frmALISMRenewalNotice.txtMonthDescription.Text = !MonthDescription
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub loadRENEWALDETAILS()
On Error GoTo err
    
    Set rsCONTROL = New ADODB.Recordset
    
    strSQL = "Select * from ODASMJobBrief where RenewalPeriod = '" & frmALISMRenewalNotice.txtRenewalPeriod.Text & "' "

    'strSQL = "Select * from ODASMJobBrief, ODASPAccount where month(ODASMJobBrief.DueDATE) = '" & frmALISMRenewalNotice.txtMonth.Text & "' and year(ODASMJobBrief.DueDate) =  '" & frmALISMRenewalNotice.txtYear.Text & "';"
    
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Dim MYCOUNT As Double
    
    With rsCONTROL
            
            MYCOUNT = .RecordCount
            
            If .EOF Or .BOF Then
                    MsgBox "The month Setup NOT in Place", vbOKOnly
                    Exit Sub
            End If
            
            frmALISMRenewalNotice.txtMonthDescription.Text = !MonthDescription
    End With

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub RecalcDueDate()
On Error GoTo err
    
        Set rsCONTROL = New ADODB.Recordset
        strSQL = "Select * from ODASMJobBrief where statuscode = 'INFORCE' or statusCode = 'LAPSE';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsCONTROL
                If .BOF Or .EOF Then Exit Sub
                
                With frmALISMRenewalNotice
                        .ProgressBar1.Visible = True
                        .ProgressBar1.Max = rsCONTROL.RecordCount
                        .ProgressBar1.Value = 0
                End With
                
                Do While Not .EOF
                        
                        Set rsSAVE = New ADODB.Recordset
                        strSQL = "Select * from ODASMJobBrief where JobBriefNo = '" & !JobBriefNo & "';"
                        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                        rsSAVE!DueDate = DateAdd("M", !UnitCount, !DateOfCommencement)
                        rsSAVE!renewed = "N"
                        
                        Dim strRenewalMonth, strRenewalYear As String
                        strRenewalMonth = ""
                        strRenewalYear = ""
                        
                        strRenewalMonth = Str(Month(rsSAVE!DueDate))
                        
                        If Len(Trim(strRenewalMonth)) = 1 Then
                                    strRenewalMonth = "0" + Trim(Str(strRenewalMonth))
                                   'strRenewalMonth = Str(strRenewalMonth)
                        End If
                        
                        strRenewalYear = Str(Year(rsSAVE!DueDate))
                        rsSAVE!RenewalPeriod = (strRenewalYear) + "/" + (strRenewalMonth)
                        rsSAVE.Update
                        
                        frmALISMRenewalNotice.ProgressBar1.Value = frmALISMRenewalNotice.ProgressBar1.Value + 1
                        frmALISMRenewalNotice.txtNoOfRecords.Text = frmALISMRenewalNotice.ProgressBar1.Value
                .MoveNext
                Loop
                
                With frmALISMRenewalNotice
                        .ProgressBar1.Visible = False
                End With
        End With

Exit Sub

err:
    ErrorMessage
End Sub


Public Sub getRENEWAL()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Renewal Period", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "No Of Policies", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Total Premiums", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Date Prepared", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Prepared By", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select * from ALISMRenewal Where RenewalPeriod = '" & frmALISMRenewalNotice.txtRenewalPeriod.Text & "'"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!RenewalPeriod))
                        
                        If Not IsNull(rsLIST!Description) Then
                            MyList.SubItems(1) = CStr(rsLIST!Description)
                        End If

                        If Not IsNull(rsLIST!NoOfPolicies) Then
                            MyList.SubItems(2) = CStr(rsLIST!NoOfPolicies)
                        End If
                        
                        If Not IsNull(rsLIST!TotalPremium) Then
                            MyList.SubItems(3) = CStr(rsLIST!TotalPremium)
                        End If

                        If Not IsNull(rsLIST!dateprepared) Then
                            MyList.SubItems(4) = CStr(rsLIST!Preparedby)
                        End If
                        
                        If Not IsNull(rsLIST!Preparedby) Then
                            MyList.SubItems(5) = CStr(rsLIST!Preparedby)
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

Public Sub getRENEWALDETAILS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Policy No", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "CompanyName", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Other Names", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Sum Asssured", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Term", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Product", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Due Date", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "DOC", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Agent Code", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select ODASMJobBrief.JobBriefNo, ODASPAccount.CompanyName, ODASPAccount.OtherNames, ODASMJobBrief.SumAssured, ODASMJobBrief.TermofPolicy, ODASMJobBrief.ProductCode, ODASMJobBrief.ExpectedPremium, ODASMJobBrief.DueDate, ODASMJobBrief.DateofCommencement, ODASMJobBrief.AgentNo from ODASMJobBrief, ODASPAccount Where ODASMJobBrief.AccountNo = ODASPAccount.AccountNo and ODASMJobBrief.PaymentMode <> 'M' order by ODASMJobBrief.JobBriefNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!JobBriefNo))
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(1) = CStr(rsLIST!CompanyName)
                        End If

                        If Not IsNull(rsLIST!OtherNames) Then
                            MyList.SubItems(2) = CStr(rsLIST!OtherNames)
                        End If
                        
                        If Not IsNull(rsLIST!sumassured) Then
                            MyList.SubItems(3) = CStr(rsLIST!sumassured)
                        End If

                        If Not IsNull(rsLIST!TermOfPolicy) Then
                            MyList.SubItems(4) = CStr(rsLIST!TermOfPolicy)
                        End If
                        
                        If Not IsNull(rsLIST!ProductCode) Then
                            MyList.SubItems(5) = CStr(rsLIST!Product)
                        End If
                        
                        If Not IsNull(rsLIST!ExpectedPremium) Then
                            MyList.SubItems(6) = CStr(rsLIST!ExpectedPremium)
                        End If
                        
                        If Not IsNull(rsLIST!DueDate) Then
                            MyList.SubItems(7) = CStr(rsLIST!DueDate)
                        End If

                        If Not IsNull(rsLIST!DateOfCommencement) Then
                            MyList.SubItems(8) = CStr(rsLIST!DateOfCommencement)
                        End If
                        
                        If Not IsNull(rsLIST!Agentno) Then
                            MyList.SubItems(9) = CStr(rsLIST!Agentno)
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

Private Sub clearRECORD()
On Error GoTo err
    With frmALISMReverseReceipt
            .txtRemark.Text = ""
            .cboReversalType.Text = ""
            .txtReversalDate.Text = Date
    End With

Exit Sub

err:
    ErrorMessage
End Sub


Public Sub getRECEIPT()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Receipt No", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Receipt Date", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Period", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Payer", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Method", .ListView1.Width / 10
                .ListView1.ColumnHeaders.Add , , "Cheque No", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Bank No", .ListView1.Width / 7

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select * from ALISMReceiptNew order by receiptno;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ReceiptNo))
                        
                        If Not IsNull(rsLIST!ReceiptDate) Then
                            MyList.SubItems(1) = CStr(rsLIST!ReceiptDate)
                        End If

                        If Not IsNull(rsLIST!ReceiptAmount) Then
                            MyList.SubItems(2) = CStr(rsLIST!ReceiptAmount)
                        End If
                        
                        If Not IsNull(rsLIST!AccountingPeriod) Then
                            MyList.SubItems(3) = CStr(rsLIST!AccountingPeriod)
                        End If
                        
                        If Not IsNull(rsLIST!Payer) Then
                            MyList.SubItems(4) = CStr(rsLIST!Payer)
                        End If

                        If Not IsNull(rsLIST!PaymentMethod) Then
                            MyList.SubItems(5) = CStr(rsLIST!PaymentMethod)
                        End If

                        If Not IsNull(rsLIST!ChequeNo) Then
                            MyList.SubItems(6) = CStr(rsLIST!ChequeNo)
                        End If
                        
                        If Not IsNull(rsLIST!BankNo) Then
                            MyList.SubItems(7) = CStr(rsLIST!BankNo)
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
Public Sub Cancelrecord()
    clearRECORD
    disableALLRECORD
    enableButtons
End Sub

Public Sub getBANK()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Policy No", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "CompanyName", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Other Names", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "BankNo", .ListView1.Width / 8
                .ListView1.ColumnHeaders.Add , , "Account No", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Status Code", .ListView1.Width / 10
                .ListView1.ColumnHeaders.Add , , "Payment Method", .ListView1.Width / 8
              
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!JobBriefNo))
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(1) = CStr(rsLIST!CompanyName)
                        End If

                        If Not IsNull(rsLIST!OtherNames) Then
                            MyList.SubItems(2) = CStr(rsLIST!OtherNames)
                        End If
                        
                        If Not IsNull(rsLIST!BankNo) Then
                            MyList.SubItems(3) = CStr(rsLIST!BankNo)
                        End If
                        
                        If Not IsNull(rsLIST!Accountno) Then
                            MyList.SubItems(4) = CStr(rsLIST!Accountno)
                        End If

                        If Not IsNull(rsLIST!StatusCode) Then
                            MyList.SubItems(5) = CStr(rsLIST!StatusCode)
                        End If

                        If Not IsNull(rsLIST!PaymentMethod) Then
                            MyList.SubItems(6) = CStr(rsLIST!PaymentMethod)
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
Public Sub updateCOMMISSION()
        validateRECORD
        If bsaveRECORD = True Then
            
        End If
End Sub

Public Sub updateRECORD()
        validateRECORD
        If bsaveRECORD = True Then
                cancelRECEIPT
                cancelRECEIPTDETAILS
                voidRECEIPT
                voidRECEIPTDETAILS
                cancelCOMMISSION
                enableSButtons
                disableALLRECORD
        End If
End Sub
Public Sub cancelCOMMISSION()
On Error GoTo err

    Set rsCONTROL = New ADODB.Recordset
    strSQL = "SELECT * FROM ALISMAgentsPay where ReceiptNo = '" & frmALISMReverseReceipt.txtReceiptNo & "' and Status = 'PAID' ;"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

    Set rsSAVE = New ADODB.Recordset
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsCONTROL
    
            If .EOF Or .BOF Then Exit Sub
            
            
            Do While .EOF
            
                !Status = "CNCL"
                .Update
                
                
                rsSAVE.AddNew
                    rsSAVE!Agentno = rsCONTROL!Agentno
                    rsSAVE!ReceiptNo = rsCONTROL!ReceiptNo
                    rsSAVE!ReceiptDate = frmALISMReverseReceipt.txtReversalDate.Text
                    rsSAVE!CurrentPeriod = CurrentPeriod
                    rsSAVE!ProductCode = rsPOLICY!ProductCode
                    rsSAVE!Factor = rsCONTROL!Factor
                    rsSAVE!RateApplied = rsCONTROL!RateApplied
                    rsSAVE!CostCenter = rsCONTROL!CostCenter
                    rsSAVE!Amount = rsCONTROL!Amount * "Y"
                    rsSAVE!TaxRate = rsCONTROL!TaxRate
                    rsSAVE!taxAmount = rsCONTROL!taxAmount * "Y"
                    rsSAVE!JobBriefCount = 0
                    rsSAVE!Status = "VOID"
                    rsSAVE!JobBriefNo = rsCONTROL!JobBriefNo
                rsSAVE.Update
                
                rsCONTROL.MoveNext
                            
            Loop
            
    End With


Exit Sub

err:
    ErrorMessage
End Sub


Public Sub validateRECORD()
On Error GoTo err
    
        With Screen.ActiveForm
            If .txtPayer = "" Then
                MsgBox "The Details of the Person Paying are Necessary and cannot be Left Blank"
                .txtPayer.SetFocus
            
            ElseIf .cboCurrencyCode.Text = "" Then
                MsgBox "The Currency Code is Required"
                .cboCurrencyCode.SetFocus
        
            ElseIf .txtReceiptDate.Text <= "" Then
                MsgBox "The Receipt Date is  required"
                .txtReceiptDate.SetFocus
            
            ElseIf DateDiff("D", Date, .txtReceiptDate.Text) > 0 Then
                MsgBox "Payment cannot be accepted in the future"
                .txtReceiptDate.SetFocus
            
            ElseIf .txtReceiptAmount.Text <= "" Then
                MsgBox "The Receipt Amount is required and There cannot be 0"
                .txtReceiptAmount.SetFocus
            
            ElseIf .txtReceiptNo.Text <= "" Then
                MsgBox "The Receipt No is required and There cannot be 0"
                .txtReceiptNo.SetFocus
        
            ElseIf .cboPaymentMethod = "" Then
                MsgBox "The Payment Method is required", vbOKOnly
                .cboPaymentMethod.SetFocus
            
'            ElseIf .cboReversalType.Text = "" Then
'                MsgBox "The Reversal Type MUST NOT be Left Blank", vbOKOnly
'                .cboReversalType.SetFocus
            
            ElseIf .txtReversalDate.Text <= "" Then
                MsgBox "The Reversal Date is  required"
                .txtReversalDate.SetFocus
            
            ElseIf DateDiff("D", Date, .txtReversalDate.Text) > 0 Then
                MsgBox "The Reversal Date cannot be accepted in the future"
                .txtReversalDate.SetFocus
            
            ElseIf Trim(.txtRemark.Text) <= "" Then
                MsgBox "The Remark MUST not be left Blank", vbOKOnly
                .txtRemark.SetFocus
            
            ElseIf Trim(.txtRemark.Text) <= "                            " Then
                MsgBox "The Remark entered is NOT Substantial", vbOKOnly
                .txtRemark.SetFocus
            Else
                    bsaveRECORD = True
            End If
        End With

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub saveRecord()
On Error GoTo err

    Set rsRCPT = New Recordset
    
    strSQL = "SELECT * FROM ALISMReceiptDetails, ALISPReceipt Where ALISMReceiptDetails.ReceiptNo = '" & rsCONTROL!ReceiptNo & "' and ALISPReceipt.ReceiptType = ALISMReceiptDetails.ReceiptType;"
    rsRCPT.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
 
            If rsRCPT!Deposit = 1 Then
                updatePROPOSALDATA
'            ElseIf rsRCPT!IssuePolicy = 1 Then
'                updateISSUANCE
            ElseIf rsRCPT!Jobbrief = 1 Then
                updateJOBBRIEFDATA
            ElseIf rsRCPT!RevivalFee = 1 Then
                updateREVIVALDATA
                updateREINSTATEMENTFILE
            ElseIf rsRCPT!loan = 1 Then
                updateLOANDATA
            ElseIf rsRCPT!accountsreceivable = 1 Then
                updateACCOUNTSDATA
            End If
    
'rsRCPT = ""

Exit Sub

err:
        ErrorMessage
End Sub

Private Sub updateISSUANCE()
On Error GoTo err

        Set rsSAVE = New Recordset
        strSAVE = "SELECT * FROM ODASMJobBrief WHERE JobBriefNo =  '" & rsCONTROL!DocumentNo & "';"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
        With rsSAVE
            !JobBriefNo = rsCONTROL!DocumentNo
            !DepositReceived = 0
            !IssueDate = Date
            !DateofLastPayment = rsCONTROL!ReceiptDate
            !DueDate = rsCONTROL!DueDate
            !UnitCount = rsCONTROL!UnitCountBeforePayment
            !SuspenseAccount = !SuspenseAccount - rsCONTROL!SuspenseAccount
            !DateofLastPayment = rsCONTROL!ReceiptDate
            !ReceivedToDate = rsCONTROL!ReceivedToDate - rsCONTROL!TransactionAmount
            !UnitCountBeforePayment = 0
            !DueDate = rsCONTROL!DueDate
            !StatusCode = "INFORCE"
            .Update
            .Requery
        End With
        
rsSAVE.Close
strSQL = ""

Exit Sub

err:
        UpdateErrorMessage
        
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
                rsSAVE.CancelUpdate
                rsSAVE.Requery
        End If

End Sub

Public Sub getRECEIPTDETAILS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Trans No", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Type", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Doc No", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Trans Date", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Period", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Status", .ListView1.Width / 6
               
                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!TransactionNo))
                        
                        If Not IsNull(rsLIST!ReceiptType) Then
                            MyList.SubItems(1) = CStr(rsLIST!ReceiptType)
                        End If

                        If Not IsNull(rsLIST!TransactionAmount) Then
                            MyList.SubItems(2) = CStr(rsLIST!TransactionAmount)
                        End If
                        
                        If Not IsNull(rsLIST!DocumentNo) Then
                            MyList.SubItems(3) = CStr(rsLIST!DocumentNo)
                        End If
                        
                        If Not IsNull(rsLIST!TransactionDate) Then
                            MyList.SubItems(4) = CStr(rsLIST!TransactionDate)
                        End If

                        If Not IsNull(rsLIST!AccountingPeriod) Then
                            MyList.SubItems(5) = CStr(rsLIST!AccountingPeriod)
                        End If

                        If Not IsNull(rsLIST!PaymentStatus) Then
                            MyList.SubItems(6) = CStr(rsLIST!PaymentStatus)
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

Public Sub loadRECORD()
On Error GoTo err
        

        Set rsCONTROL = New ADODB.Recordset
        strSQL = "Select * from ALISMReceiptNew Where ReceiptNo like '" & Screen.ActiveForm.txtReceiptNo.Text & "';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsCONTROL
            If .EOF And .BOF Then Exit Sub
                    Screen.ActiveForm.txtReceiptNo.Text = !ReceiptNo
                    Screen.ActiveForm.txtReceiptDate.Text = !ReceiptDate
                    Screen.ActiveForm.txtReceiptAmount.Text = !ReceiptAmount
                    Screen.ActiveForm.txtAccountingPeriod.Text = !AccountingPeriod
                    Screen.ActiveForm.txtPaymentStatus.Text = !PaymentStatus
                    Screen.ActiveForm.txtChequeNo.Text = !ChequeNo & ""
                    Screen.ActiveForm.txtPaymentMethod.Text = !PaymentMethod & ""
                    Screen.ActiveForm.txtBankNo.Text = !BankNo & ""
                    Screen.ActiveForm.txtPayer.Text = !Payer
                    Screen.ActiveForm.txtLocal.Text = !LocalCheque & ""
                    Screen.ActiveForm.cboCurrencyCode.Text = !CurrencyCode & ""
                    Screen.ActiveForm.txtReversalDate.Text = Date
                    Screen.ActiveForm.DTPickerVoidDate.Value = Date
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadPREMIUM()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
        strSQL = "Select * from ODASMJobBrief Where JobBriefNo = '" & Screen.ActiveForm.cboDocumentNo.Text & "';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
        
        frmALISMCorrectReceipt.txtExpectedAmount.Text = rsCONTROL!ExpectedPremium & ""
        frmALISMCorrectReceipt.cboAgentNo.Text = rsCONTROL!Agentno & ""

rsCONTROL.Close
strSQL = Empty

Exit Sub
err:
    ErrorMessage
End Sub
Public Sub loadAGENT()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
        strSQL = "Select * from ALISPAgent Where AgentNo = '" & Screen.ActiveForm.cboAgentNo.Text & "';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
        
        frmALISMCorrectReceipt.txtNames = Trim(rsCONTROL!OtherNames) + " " + Trim(rsCONTROL!CompanyName)
        frmALISMCorrectReceipt.txtAgentStatus.Text = rsCONTROL!Status & ""
        
rsCONTROL.Close
strSQL = Empty

Exit Sub
err:
    ErrorMessage
End Sub


Public Sub loadCOMMISSION()
On Error GoTo err
        

        Set rsCONTROL = New ADODB.Recordset
        strSQL = "Select * from ALISMReceiptNew, ALISMReceiptDetails Where ALISMReceiptNew.ReceiptNo like ALISMReceiptdetails.Receiptno and ALISMReceiptDetails.ReceiptNo like '" & Screen.ActiveForm.txtReceiptNo.Text & "';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsCONTROL
            If .EOF And .BOF Then Exit Sub
                    Screen.ActiveForm.txtReceiptNo.Text = !ReceiptNo
                    Screen.ActiveForm.txtReceiptDate.Text = !ReceiptDate
                    Screen.ActiveForm.txtReceiptAmount.Text = !ReceiptAmount
                    Screen.ActiveForm.txtAccountingPeriod.Text = !AccountingPeriod
                    Screen.ActiveForm.txtPaymentStatus.Text = !PaymentStatus
                    Screen.ActiveForm.txtChequeNo.Text = !ChequeNo & ""
                    Screen.ActiveForm.txtPaymentMethod.Text = !PaymentMethod & ""
                    Screen.ActiveForm.txtBankNo.Text = !BankNo & ""
                    Screen.ActiveForm.txtPayer.Text = !Payer
                    Screen.ActiveForm.txtLocal.Text = !LocalCheque & ""
                    Screen.ActiveForm.cboCurrencyCode.Text = !CurrencyCode & ""
                    frmALISMCorrectReceipt.cboDocumentNo.Text = !DocumentNo & ""
                    frmALISMCorrectReceipt.txtUnitCount.Text = !UnitCount & ""
                    frmALISMCorrectReceipt.txtUnitsPaid.Text = !UnitsPaid & ""
                 End With

rsCONTROL.Close
strSQL = Empty

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub cancelRECEIPT()
On Error GoTo err
        
        Set rsSAVE = New ADODB.Recordset
        strSQL = "Select * from ALISMReceiptNew Where ReceiptNo like '" & frmALISMReverseReceipt.txtReceiptNo.Text & "' and paymentStatus = 'PAID';"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsSAVE
            If .EOF And .BOF Then Exit Sub
            
            !PaymentStatus = "CNCL"
            !ReversalDate = frmALISMReverseReceipt.txtReversalDate.Text
            !remark = frmALISMReverseReceipt.txtRemark.Text
            .Update
            .Requery
        End With
Exit Sub

err:
        UpdateErrorMessage
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
                rsSAVE.CancelUpdate
                rsSAVE.Requery
        End If
End Sub

Private Sub cancelRECEIPTDETAILS()
On Error GoTo err
        
        Set rsSAVE = New ADODB.Recordset
        strSQL = "Select * from ALISMReceiptDetails Where ReceiptNo like '" & frmALISMReverseReceipt.txtReceiptNo.Text & "';"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsSAVE
            If .EOF And .BOF Then Exit Sub
            
            Do While Not .EOF
                
                !PaymentStatus = "CNCL"
                !ReversalDate = frmALISMReverseReceipt.txtReversalDate.Text
                
                .Update
                .MoveNext
            Loop
        End With


Exit Sub

err:
        UpdateErrorMessage
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
                rsSAVE.CancelUpdate
                rsSAVE.Requery
        End If
End Sub

Private Sub voidRECEIPT()
On Error GoTo err
        
        Set rsSAVE = New ADODB.Recordset
        strSQL = "Select * from ALISMReceiptNew Where ReceiptNo like '" & frmALISMReverseReceipt.txtReceiptNo.Text & "' ;"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsSAVE
            If .EOF And .BOF Then
                    .AddNew
                    !ReceiptNo = frmALISMReverseReceipt.txtReceiptNo.Text
            End If
            
            !AccountingPeriod = CurrentPeriod
            !PaymentStatus = "VOID"
            !ReceiptDate = frmALISMReverseReceipt.txtReversalDate.Text
            !ReceiptAmount = CDbl(frmALISMReverseReceipt.txtReceiptAmount.Text) '* "Y"
            !ChequeNo = frmALISMReverseReceipt.txtChequeNo.Text
            !PaymentMethod = frmALISMReverseReceipt.txtPaymentMethod.Text
            !BankNo = frmALISMReverseReceipt.txtBankNo.Text
            !Payer = frmALISMReverseReceipt.txtPayer.Text
            !LocalCheque = frmALISMReverseReceipt.txtLocal.Text
            !CurrencyCode = frmALISMReverseReceipt.cboCurrencyCode.Text
            !ReversalDate = frmALISMReverseReceipt.txtReversalDate.Text
            !remark = frmALISMReverseReceipt.txtRemark.Text
            .Update
        End With
Exit Sub

err:
        UpdateErrorMessage
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
                rsSAVE.CancelUpdate
                rsSAVE.Requery
        End If
End Sub

Private Sub voidRECEIPTDETAILS()
On Error GoTo err

    Set rsSAVE = New ADODB.Recordset
    strSQL = "Select * from ALISMReceiptDetails Where ReceiptNo like '" & frmALISMReverseReceipt.txtReceiptNo & "' ;"
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Set rsCONTROL = New ADODB.Recordset
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

    
    With rsCONTROL
            
            If .BOF Or .EOF Then Exit Sub
            
            Do While Not .EOF
            
'                    rsSAVE.AddNew
'                    rsSAVE!ReceiptNo = rsCONTROL!ReceiptNo
                    rsSAVE!TransactionNo = rsCONTROL!TransactionNo
                    rsSAVE!TransactionDate = Screen.ActiveForm.txtReversalDate
                    rsSAVE!ReceiptType = rsCONTROL!ReceiptType
                    rsSAVE!DocumentNo = rsCONTROL!DocumentNo
                    If IsNull(rsCONTROL!Accountno) = True Then
                            rsSAVE!Accountno = 0
                    Else
                            rsSAVE!Accountno = rsCONTROL!Accountno & ""
                    End If
                    rsSAVE!TransactionAmount = CDbl(rsCONTROL!TransactionAmount) '* "Y"
'                    rsSAVE!AccountingPeriod = CurrentPeriod
                    rsSAVE!PaymentStatus = "VOID"
'                    rsSAVE!UnitsPaid = rsCONTROL!UnitsPaid * "Y"
                    rsSAVE!Accountno = rsCONTROL!Accountno
                    rsSAVE!SuspenseAccount = CDbl(rsCONTROL!SuspenseAccount)
                    rsSAVE!ReceivedToDate = CDbl(rsCONTROL!ReceivedToDate) - CDbl(rsCONTROL!TransactionAmount)
'                    rsSAVE!UnitCount = rsCONTROL!UnitCountBeforePayment
'                    rsSAVE!UnitCountBeforePayment = CDbl(rsCONTROL!UnitCountBeforePayment) - CDbl(rsCONTROL!UnitsPaid)
'
'                    If !UnitCountBeforePayment < 0 Then !UnitCountBeforePayment = 0
                    
'                    rsSAVE!DueDate = DateAdd("M", CDbl(!UnitCount), rsCONTROL!DueDate)
                    rsSAVE!StatusCode = rsCONTROL!StatusCode
                    rsSAVE!TransactionDate = Date
                    
                    .Update
                    .Requery
                    
                    '/Update the mother document with the reversal
                    
                    saveRecord
                    .MoveNext
            Loop
            
    End With
'rsSAVE.Close

Exit Sub

err:
    UpdateErrorMessage
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    End If
End Sub

Private Sub updateJOBBRIEFDATA()
On Error GoTo err

        Set rsSAVE = New Recordset
        If breverseRECEIPT = True Then
            strSQL = "SELECT * FROM ODASMJobBrief WHERE JobBriefNo =  '" & rsCONTROL!DocumentNo & "';"
        Else: strSQL = "SELECT * FROM ODASMJobBrief WHERE JobBriefNo =  '" & rsCONTROL!DocumentNo & "';"
        End If
        
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
        With rsSAVE
            
            If .EOF Or .BOF Then Exit Sub
            
            !JobBriefNo = rsCONTROL!DocumentNo
            !DateofLastPayment = Screen.ActiveForm.txtReceiptDate
'            !UnitCount = rsCONTROL!UnitCountBeforePayment
'            !SuspenseAccount = CDbl(!SuspenseAccount) - CDbl(rsCONTROL!SuspenseAccount)
            !ReceivedToDate = CDbl(rsCONTROL!ReceivedToDate) - CDbl(rsCONTROL!TransactionAmount)
'            !UnitCountBeforePayment = CDbl(rsCONTROL!UnitCountBeforePayment) - CDbl(rsCONTROL!UnitsPaid)
'            !DueDate = DateAdd("M", CDbl(!UnitCount), DueDate)
'            !StatusCode = rsCONTROL!StatusCode
            .Update
            .Requery
        End With
        
rsSAVE.Close

strSQL = ""
Exit Sub

err:
UpdateErrorMessage
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
                rsSAVE.CancelUpdate
                rsSAVE.Requery
        End If
End Sub

Private Sub updateREVIVALDATA()
On Error GoTo err

        Set rsSAVE = New Recordset
        strSQL = "SELECT * FROM ODASMJobBrief WHERE JobBriefNo =  '" & Screen.ActiveForm.cboDocumentNo & "';"
    
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
        With rsSAVE
            !JobBriefNo = Screen.ActiveForm.cboDocumentNo
            !DepositReceived = 0
            !DateofLastPayment = Screen.ActiveForm.txtReceiptDate
            !UnitCount = rsCONTROL!UnitCountBeforePayment
            !SuspenseAccount = CDbl(!SuspenseAccount) - CDbl(rsCONTROL!SuspenseAccount)
            !ReceivedToDate = CDbl(Screen.ActiveForm.txtReceivedTodate) - CDbl(rsCONTROL!TransactionAmount)
            !UnitCountBeforePayment = CDbl(Screen.ActiveForm.txtUnitCountBeforePayment) - CDbl(rsCONTROL!UnitsPaid)
            !DueDate = rsCONTROL!DueDate
            !ReinstatementCharge = !ReinstatementCharge - CDbl(rsCONTROL!TransactionAmount)
            !StatusCode = "LAPSED"
            !StatusDate = Date
            .Update
            .Requery
        End With
        
        rsSAVE.Close

err:
        UpdateErrorMessage
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
                rsSAVE.CancelUpdate
                rsSAVE.Requery
        End If
End Sub
Private Sub updateREINSTATEMENTFILE()
On Error GoTo err

        Dim rsREINSTATEMENT As ADODB.Recordset, strPOL As String
        Set rsREINSTATEMENT = New Recordset
            
        rsREINSTATEMENT.Open "SELECT * FROM ALISMReinstatement WHERE JobBriefNo =  '" & Screen.ActiveForm.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                With rsREINSTATEMENT
                    !Reinstated = "N"
                    !ReceiptDate = Date
                    .Update
                    .Requery
                End With
        
        rsREINSTATEMENT.Close
 
Exit Sub
err:
        UpdateErrorMessage
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
                rsREINSTATEMENT.CancelUpdate
                rsREINSTATEMENT.Requery
        End If

End Sub

Private Sub updateACCOUNTSDATA()
On Error GoTo err

        Dim rsACCOUNT As ADODB.Recordset
        Set rsACCOUNT = New Recordset
            
        rsACCOUNT.Open "SELECT * FROM MEGAMAcTransaction  WHERE AccountNo =  '" & Screen.ActiveForm.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
        With rsACCOUNT
            .AddNew
            !Accountno = Screen.ActiveForm.cboDocumentNo.Text
            !TransactionType = CDbl(rsCONTROL!ReceiptType)
            !TransactionDate = CDbl(rsCONTROL!TransactionDate)
            !InclusiveAmount = InclusiveAmount - CDbl(rsCONTROL!TransactionAmount)
            !Amount = !Amount - CDbl(rsCONTROL!TransactionAmount)
            !VAT = 0
            !VATRate = 0
            !Balance = !InclusiveAmount
            !Status = "PAID"
            !DueDate = Date
            !ReceiptNo = rsCONTROL!ReceiptNo
            !InvoiceNo = rsCONTROL!Account
            !DateofLastPayment = rsCONTROL!ReceiptDate
            !ReceivedToDate = rsCONTROL!ReceivedToDate
            !dateprepared = Date
            !Preparedby = CurrentUserName
            .Update
            .Requery
        End With
        
        rsACCOUNT.Close

Exit Sub


err:
UpdateErrorMessage
If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
        rsSAVE.CancelUpdate
        rsSAVE.Requery
End If

End Sub
Private Sub updatePROPOSALDATA()
On Error GoTo err

        Dim rsPROP As ADODB.Recordset, strPROP As String
        Set rsPROP = New Recordset
            
                rsPROP.Open "SELECT * FROM ALISMProposal WHERE proposalNo =  '" & Screen.ActiveForm.cboDocumentNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                With rsPROP
                    !ProposalNo = Screen.ActiveForm.cboDocumentNo
                    !DepositReceived = !DepositReceived + CDbl(rsCONTROL!TransactionAmount)
                    !DateofLastPayment = rsCONTROL!ReceiptDate
                    !StatusCode = "DepRecvd"
                    .Update
                    .Requery
                End With
        
        rsPROP.Close
        strPROP = ""

Exit Sub

err:
UpdateErrorMessage
If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
        rsSAVE.CancelUpdate
        rsSAVE.Requery
End If
End Sub

Public Sub UpdatePolicyLedger()
On Error GoTo err

With Screen.ActiveForm
    If rsCONTROL!ReceiptType = "BIP" Or rsCONTROL!ReceiptNoDetails <= "" Then Exit Sub
    
    Set rsFindRecord = New ADODB.Recordset
    rsFindRecord.Open "SELECT * FROM ODASPAccount WHERE AccountNo LIKE '" & rsCONTROL!Accountno & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        MySName = Empty: MyOName = Empty
    ElseIf rsFindRecord!CompanyName = "" And rsFindRecord!OtherNames = "" Then
        MySName = Empty: MyOName = Empty
    ElseIf IsNull(rsFindRecord!CompanyName) = True And IsNull(rsFindRecord!OtherNames) = True Then
        MySName = Empty: MyOName = Empty
    Else
        MySName = rsFindRecord!CompanyName & "": MyOName = rsFindRecord!OtherNames & ""
    End If
    
    If rsRCPT!Deposit = "1" Then Exit Sub
    If rsRCPT!accountsreceivable = "1" Then Exit Sub
    
    Set rsSAVE = New ADODB.Recordset

    rsSAVE.Open "SELECT * FROM ODASMJobBriefLedger;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
        rsSAVE.AddNew
            rsSAVE!ReceiptNo = rsCONTROL!ReceiptNo
            rsSAVE!ReceiptDate = rsCONTROL!ReceiptDate
            rsSAVE!ReceiptType = rsCONTROL!ReceiptType
            rsSAVE!DocumentNo = rsCONTROL!DocumentNo
            rsSAVE!Accountno = rsCONTROL!Accountno
            rsSAVE!ReceiptAmount = rsCONTROL!ReceiptAmount
            rsSAVE!AccountingPeriod = CurrentPeriod
            rsSAVE!PaymentStatus = "VOID"
            rsSAVE!UnitsPaid = rsCONTROL!UnitsPaid
            rsSAVE!AccountCode = rsCONTROL!AccountCode
            rsSAVE!DueDate = rsCONTROL!DueDate
            rsSAVE!SuspenseAccount = CDbl(rsCONTROL!SuspenseAccount)
            rsSAVE!PaymentMethod = rsCONTROL!PaymentMethod
            rsSAVE!BankNo = rsCONTROL!BankNo
            rsSAVE!ChequeNo = rsCONTROL!ChequeNo
            rsSAVE!DateofLastPayment = rsCONTROL!DateofLastPayment
            rsSAVE!ReceivedToDate = rsCONTROL!ReceivedToDate
            rsSAVE!UnitCount = rsCONTROL!UnitCount
            rsSAVE!UnitCountBeforePayment = rsCONTROL!UnitCountBeforePayment
            rsSAVE!StatusCode = rsCONTROL!StatusCode
            rsSAVE!TransactionDate = Date
            rsSAVE!TransactionNo = rsCONTROL!TransactionNODetails
            rsSAVE!CompanyName = MySName
            rsSAVE!OtherNames = MyOName
            rsSAVE!Preparedby = CurrentUserName
        rsSAVE.Update
    
    Set rsSAVE = Nothing
    
End With
Exit Sub

err:
UpdateErrorMessage
If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
        rsSAVE.CancelUpdate
        rsSAVE.Requery
End If

End Sub

Private Sub updateLOANDATA()
On Error GoTo err

    Set rsSAVE = New Recordset
    strSQL = "SELECT * FROM ALISMLoanManagement WHERE LoanNo =  '" & Screen.ActiveForm.cboDocumentNo & "';"
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            With rsSAVE
            
                !JobBriefNo = Screen.ActiveForm.cboDocumentNo
                !TotalReceivedTodate = !TotalReceivedTodate - CDbl(rsCONTROL!TransactionAmount)
                !DateofLastPayment = rsCONTROL!ReceiptDate
                !DueDate = rsCONTROL!DueDate
                !suspense = !suspense - rsCONTROL!SuspenseAccount & ""
                !CurrentBalance = !CurrentBalance + !PrincipalAmount
                !InterestReceived = !InterestReceived - (rsCONTROL!ReceiptAmount.Text - !PrincipalAmount)
                !DateofLastPayment = rsCONTROL!DateofLastPayment & ""
                !DueDate = rsCONTROL!DueDate
                !Status = rsCONTROL!StatusCode
                .Update
                .Requery
            End With
    
    rsSAVE.Close
    strSQL = ""
    
Exit Sub
err:
    ErrorMessage
UpdateErrorMessage
If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
        rsSAVE.CancelUpdate
        rsSAVE.Requery
End If

End Sub

