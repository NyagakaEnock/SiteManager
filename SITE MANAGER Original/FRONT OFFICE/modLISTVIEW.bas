Attribute VB_Name = "modLISTVIEW"
Public Sub showBRIEFSInstalments()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Job Brief No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Customer Name", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Description Of Order", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Brief Date", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Deadline Date", .ListView1.Width / 6

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                strSQL = "SELECT DISTINCT (jobbriefNo) FROM ODASMJobBriefInstallment;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                df = rsLIST.RecordCount
                
                'strSQL = "SELECT * FROM ODASMJobBrief, ODASPAccount Where ODASMJobBrief.AccountNo=ODASPAccount.AccountNo ORDER BY ODASMJobBrief.JobBriefNo;"
                'rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set rsCONTROL = New ADODB.Recordset
                        strCONTROL = "SELECT * FROM ODASMJobBrief, ODASPAccount Where ODASMJobBrief.AccountNo=ODASPAccount.AccountNo  AND ODASMJobBrief.JobBriefNo='" & rsLIST!JobBriefNo & "' ORDER BY ODASMJobBrief.JobBriefNo;"
                        rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                        If rsCONTROL.EOF And rsCONTROL.BOF Then
                        Else
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsCONTROL!JobBriefNo))
                        
                            If Not IsNull(rsCONTROL!CompanyName) Then
                                MyList.SubItems(1) = CStr(rsCONTROL!CompanyName)
                            End If
                            
                            If Not IsNull(rsCONTROL!descriptionOfOrder) Then
                                MyList.SubItems(2) = CStr(rsCONTROL!descriptionOfOrder)
                            End If
                            
                            If Not IsNull(rsCONTROL!JobBriefDate) Then
                                MyList.SubItems(3) = CStr(rsCONTROL!JobBriefDate)
                            End If
                            
                            If Not IsNull(rsCONTROL!deadlineDate) Then
                                MyList.SubItems(4) = CStr(rsCONTROL!deadlineDate)
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
Public Sub showALLCLOSEDCOSTINGBRIEFS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Job Brief No", .ListView1.Width / 6 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Customer Name", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Description Of Order", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Brief Date", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Deadline Date", .ListView1.Width / 6

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMJobBrief, ODASPAccount Where ODASMJobBrief.AccountNo=ODASPAccount.AccountNo ORDER BY ODASMJobBrief.JobBriefNo;"
                
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!JobBriefNo))
                        
                       If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(1) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!descriptionOfOrder) Then
                            MyList.SubItems(2) = CStr(rsLIST!descriptionOfOrder)
                        End If
                        
                        If Not IsNull(rsLIST!JobBriefDate) Then
                            MyList.SubItems(3) = CStr(rsLIST!JobBriefDate)
                        End If
                        
                        If Not IsNull(rsLIST!deadlineDate) Then
                            MyList.SubItems(4) = CStr(rsLIST!deadlineDate)
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


Public Sub FindVATRate()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Code", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "VAT Rate [%]", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Starting", .ListView1.Width / 3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Ending", .ListView1.Width / 3 ', lvwColumnCenter
.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ODASPVAT ;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!VATCode))

    If Not IsNull(rsLIST!VATRate) Then
        MyList.SubItems(1) = CStr(rsLIST!VATRate)
    End If
    If Not IsNull(rsLIST!Starting) Then
        MyList.SubItems(2) = CStr(rsLIST!Starting)
    End If
    If Not IsNull(rsLIST!Ending) Then
        MyList.SubItems(3) = CStr(rsLIST!Ending)
    End If
    
    
    rsLIST.MoveNext
    
Wend
.ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
Set MyList = Nothing
End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Public Sub getALLContracts()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "ContractNo", .ListView1.Width / 4 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Plot", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Plot Name", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Expiry Date", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT (L.ExpiryDate) as EDates, L.*, P.*  FROM ODASMLeaseAgreement L, ODASPPlot P where L.Assigned = 'Y' and (L.Terminated = 'N' or L.Terminated is null) AND P.PlotNo = L.PlotNo ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsLIST.RecordCount = 0 Then Exit Sub
                rsLIST.MoveFirst
                While Not rsLIST.EOF
                
                            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
                            
                            If Not IsNull(rsLIST!PlotNo) Then
                                MyList.SubItems(1) = CStr(rsLIST!PlotNo)
                            End If
                            
                            If Not IsNull(rsLIST!PlotName) Then
                                MyList.SubItems(2) = CStr(rsLIST!PlotName)
                            End If
                            
                            If Not IsNull(rsLIST!AccountNo) Then
                                MyList.SubItems(3) = CStr(rsLIST!AccountNo)
                            End If
                            If Not IsNull(rsLIST!EDates) Then
                                MyList.SubItems(4) = CStr(rsLIST!EDates)
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
Public Sub ShowAllClients()
On Error GoTo err
With Screen.ActiveForm
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Client Code", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Company Name", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Address", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Contact Name", .ListView1.Width / 4

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ODASPAccount Where AccountType='CUST' and (status = 'A' or status='1') ORDER BY CompanyName;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!AccountNo))


    If Not IsNull(rsLIST!CompanyName) Then
        MyList.SubItems(1) = CStr(rsLIST!CompanyName)
    End If
     
    If Not IsNull(rsLIST!PostalAddress) Then
        MyList.SubItems(2) = CStr(rsLIST!PostalAddress)
    End If
    
    
    If Not IsNull(rsLIST!ContactPerson) Then
        MyList.SubItems(3) = CStr(rsLIST!ContactPerson)
    End If
    
    rsLIST.MoveNext
    
Wend

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub OpeningBalances()
On Error GoTo err
    With frmODASPCustomerOpeningBal
        
        .ListView1.ListItems.Clear
        .ListView1.ColumnHeaders.Clear
        
        .ListView1.ColumnHeaders.Add , , "Cust No.", .ListView1.Width / 5
        .ListView1.ColumnHeaders.Add , , "Customer Name", .ListView1.Width / 4 ', lvwColumnCenter
        .ListView1.ColumnHeaders.Add , , "Transaction Date", .ListView1.Width / 3 ', lvwColumnCenter
        .ListView1.ColumnHeaders.Add , , "Opening Balance", .ListView1.Width / 3 ', lvwColumnCenter
        .ListView1.View = lvwReport
        
        Dim rsLIST As ADODB.Recordset
        Set rsLIST = New ADODB.Recordset
        
        rsLIST.Open "SELECT * FROM ODASMCustomerStatement S, ODASPAccount A Where S.AccountNo = A.AccountNo and S.Reference = 'BAL C/F' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
        Dim MyList As ListItem
           
        While Not rsLIST.EOF
            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!AccountNo))
            
                If Not IsNull(rsLIST!CompanyName) Then
                    MyList.SubItems(1) = CStr(rsLIST!CompanyName)
                End If
                If Not IsNull(rsLIST!TransactionDate) Then
                    MyList.SubItems(2) = CStr(rsLIST!TransactionDate)
                End If
                If Not IsNull(rsLIST!CreditAmount) Then
                    MyList.SubItems(3) = CStr(rsLIST!CreditAmount)
                End If
                rsLIST.MoveNext
                
            Wend
        .ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
        Set MyList = Nothing
    End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub FindAccountingPeriods()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Period", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Starting", .ListView1.Width / 3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Ending", .ListView1.Width / 3 ', lvwColumnCenter
.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ALISPPeriod ;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!AccountingPeriod))

    If Not IsNull(rsLIST!Description) Then
        MyList.SubItems(1) = CStr(rsLIST!Description)
    End If
    If Not IsNull(rsLIST!StartDate) Then
        MyList.SubItems(2) = CStr(rsLIST!StartDate)
    End If
    If Not IsNull(rsLIST!LastDate) Then
        MyList.SubItems(3) = CStr(rsLIST!LastDate)
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

Public Sub FindCurrencies()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Code", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Symbol", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Base", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ODASPCurrency ;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!CurrencyCode))

    If Not IsNull(rsLIST!CurrencyDescription) Then
        MyList.SubItems(1) = CStr(rsLIST!CurrencyDescription)
    End If
    If Not IsNull(rsLIST!CurrencySymbol) Then
        MyList.SubItems(2) = CStr(rsLIST!CurrencySymbol)
    End If
    If Not IsNull(rsLIST!BaseCurrency) Then
        MyList.SubItems(3) = CStr(rsLIST!BaseCurrency)
    End If
    
    rsLIST.MoveNext
    
Wend
.ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
Set MyList = Nothing
End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub showAUTHORIZEDJOBBRIEF()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Job Brief No", .ListView1.Width / 5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Customer Name", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Description Of Order", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Brief Date", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Deadline Date", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMJobBrief JB, ODASPAccount A Where A.AccountNo = JB.AccountNo and JB.Authorized = 'Y' ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                
                    Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!JobBriefNo))
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(1) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!descriptionOfOrder) Then
                            MyList.SubItems(2) = CStr(rsLIST!descriptionOfOrder)
                        End If
                        
                        If Not IsNull(rsLIST!JobBriefDate) Then
                            MyList.SubItems(3) = CStr(rsLIST!JobBriefDate)
                        End If
                        
                        If Not IsNull(rsLIST!deadlineDate) Then
                            MyList.SubItems(4) = CStr(rsLIST!deadlineDate)
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

Public Sub showALLMEDIA2()
On Error GoTo err

    With Screen.ActiveForm
            .ListView2.ListItems.Clear
            .ListView2.ColumnHeaders.Clear
            
            .ListView2.ColumnHeaders.Add , , "Media", .ListView2.Width / 4
            .ListView2.ColumnHeaders.Add , , "Description", .ListView2.Width / 4
            .ListView2.ColumnHeaders.Add , , "Inventory Item?", .ListView2.Width / 4
            .ListView2.ColumnHeaders.Add , , "Status", .ListView2.Width / 4

            .ListView2.View = lvwReport
            
            Dim rsLIST As ADODB.Recordset
            Set rsLIST = New ADODB.Recordset
            
            strSQL = "SELECT * FROM ODASPMedia"
            rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            Dim MyList As ListItem
            
            If rsLIST.EOF And rsLIST.BOF Then
                .ListView2.View = lvwList
                Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
                Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
            End If
            
            While Not rsLIST.EOF
            
            Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!MediaCode))
            
                If Not IsNull(rsLIST!MediaDescription) Then
                    MyList.SubItems(1) = CStr(rsLIST!MediaDescription)
                End If
                
                If Not IsNull(rsLIST!InventoryItem) Then
                    MyList.SubItems(2) = CStr(rsLIST!InventoryItem)
                End If
                
                If Not IsNull(rsLIST!Status) Then
                    MyList.SubItems(3) = CStr(rsLIST!Status)
                End If

                rsLIST.MoveNext
                
            Wend
        
        Set MyList = Nothing: Set rsLIST = Nothing
        
        End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub showALLMEDIASIZES()
On Error GoTo err

    With Screen.ActiveForm
            .ListView1.ListItems.Clear
            .ListView1.ColumnHeaders.Clear
            
            .ListView1.ColumnHeaders.Add , , "Media Code", .ListView1.Width / 5
            .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 5
            .ListView1.ColumnHeaders.Add , , "Size", .ListView1.Width / 5
            .ListView1.ColumnHeaders.Add , , "Length", .ListView1.Width / 5
            .ListView1.ColumnHeaders.Add , , "Width", .ListView1.Width / 5

            .ListView1.View = lvwReport
            
            Dim rsLIST As ADODB.Recordset
            Set rsLIST = New ADODB.Recordset
            
            strSQL = "SELECT * FROM ODASPMedia, ODASPMediaSize Where ODASPMedia.MediaCode = ODASPMediaSize.MediaCode and ODASPMediaSize.MediaCode = '" & frmODASPLandRate.txtMediaCode.Text & "'"
            rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            Dim MyList As ListItem
            
            If rsLIST.EOF And rsLIST.BOF Then
                .ListView1.View = lvwList
                Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
                Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
            End If
            
            While Not rsLIST.EOF
            
            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!MediaCode))
            
                If Not IsNull(rsLIST!MediaDescription) Then
                    MyList.SubItems(1) = CStr(rsLIST!MediaDescription)
                End If
                
                If Not IsNull(rsLIST!MediaSize) Then
                    MyList.SubItems(2) = CStr(rsLIST!MediaSize)
                End If
                
                If Not IsNull(rsLIST!Length) Then
                    MyList.SubItems(3) = CStr(rsLIST!Length)
                End If
                
                If Not IsNull(rsLIST!Width) Then
                    MyList.SubItems(4) = CStr(rsLIST!Width)
                End If

                rsLIST.MoveNext
                
            Wend
        
        Set MyList = Nothing: Set rsLIST = Nothing
        
        End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub showALLLANDRATES()
On Error GoTo err

    With Screen.ActiveForm
            .ListView3.ListItems.Clear
            .ListView3.ColumnHeaders.Clear
            
            .ListView3.ColumnHeaders.Add , , "Town", .ListView3.Width / 5
            .ListView3.ColumnHeaders.Add , , "Media", .ListView3.Width / 5
            .ListView3.ColumnHeaders.Add , , "Mode", .ListView3.Width / 5
            .ListView3.ColumnHeaders.Add , , "Size", .ListView3.Width / 5
            .ListView3.ColumnHeaders.Add , , "Rate", .ListView3.Width / 5

            .ListView3.View = lvwReport
            
            Dim rsLIST As ADODB.Recordset
            Set rsLIST = New ADODB.Recordset
            
            strSQL = "SELECT * FROM ODASPLandRate Where TownCode =  '" & frmODASPLandRate.txtTownCode.Text & "'"
            rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            Dim MyList As ListItem
            
            If rsLIST.EOF And rsLIST.BOF Then
                .ListView3.View = lvwList
                Set MyList = .ListView3.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
                Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
            End If
            
            While Not rsLIST.EOF
            
            Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!townCode))
            
                If Not IsNull(rsLIST!MediaCode) Then
                    MyList.SubItems(1) = CStr(rsLIST!MediaCode)
                End If
                
                If Not IsNull(rsLIST!MediaSize) Then
                    MyList.SubItems(3) = CStr(rsLIST!MediaSize)
                End If
                
                If Not IsNull(rsLIST!PaymentMode) Then
                    MyList.SubItems(2) = CStr(rsLIST!PaymentMode)
                End If
                
                If Not IsNull(rsLIST!Amount) Then
                    MyList.SubItems(4) = CStr(rsLIST!Amount)
                End If

                rsLIST.MoveNext
                
            Wend
        
        Set MyList = Nothing: Set rsLIST = Nothing
        
        End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub GetApprovedChecks()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Cheque No", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Payee Details", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Date", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Scheduled", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                strSQL = "Select ALISMCheque.ChequeNo, ALISMCheque.PayeeDetails, ALISMCheque.ChequeDate, ALISMCheque.chequeAmount, ALISMCheque.status from ALISMCheque Where  ALISMCheque.Prepared = 'Y' and ALISMCheque.approved = 'N' "
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ChequeNo))
                        
                        If Not IsNull(rsLIST!ChequeDate) Then
                            MyList.SubItems(2) = CStr(rsLIST!ChequeDate)
                        End If

                        If Not IsNull(rsLIST!ChequeAmount) Then
                                MyList.SubItems(3) = CStr(rsLIST!ChequeAmount)
                        End If
                        
                        If Not IsNull(rsLIST!Status) Then
                                MyList.SubItems(4) = CStr(rsLIST!Status)
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

Public Sub GetIssuedChecks()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Cheque No", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Payee Details", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Date", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Scheduled", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select ALISMCheque.ChequeNo, ALISMCheque.PayeeDetails, ALISMCheque.ChequeDate, ALISMCheque.ChequeAmount, ALISMCheque.status from ALISMCheque Where  ALISMCheque.authorized = 'Y' and ALISMCheque.issued = 'N' "
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ChequeNo))
                        
                        If Not IsNull(rsLIST!ChequeDate) Then
                            MyList.SubItems(2) = CStr(rsLIST!ChequeDate)
                        End If

                        If Not IsNull(rsLIST!ChequeAmount) Then
                                MyList.SubItems(3) = CStr(rsLIST!ChequeAmount)
                        End If
                        
                        If Not IsNull(rsLIST!Status) Then
                                MyList.SubItems(4) = CStr(rsLIST!Status)
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

Public Sub GetAuthorizedChecks()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Cheque No", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Payee Details", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Date", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Scheduled", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select ALISMCheque.ChequeNo, ALISMCheque.PayeeDetails, ALISMCheque.ChequeDate, ALISMCheque.ChequeAmount, ALISMCheque.status from ALISMCheque Where  ALISMCheque.Approved = 'Y' and ALISMCheque.authorized = 'N' "
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ChequeNo))
                        
                        If Not IsNull(rsLIST!ChequeDate) Then
                            MyList.SubItems(2) = CStr(rsLIST!ChequeDate)
                        End If

                        If Not IsNull(rsLIST!ChequeAmount) Then
                                MyList.SubItems(3) = CStr(rsLIST!ChequeAmount)
                        End If
                        
                        If Not IsNull(rsLIST!Status) Then
                                MyList.SubItems(4) = CStr(rsLIST!Status)
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

Public Sub GetSchedule()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Bank No", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "AccountNo", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Details", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset

                strSQL = "Select BankNo, AccountNo, Details from ALISPBankAccount "
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

                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub GetScheduledchecks()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Bank No", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "AccountNo", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Details", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset

                strSQL = "Select BankNo, AccountNo, Details from ALISPBankAccount "
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

                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
If err.Number = 3265 Then Resume Next
ErrorMessage
End Sub

Public Sub GetRequisition1()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Req No", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Payee Details", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Reference", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Date", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Job Brief No", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Payment Flag", .ListView1.Width / 8

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select VoucherNo, CompanyName RequisitionDate, Reference, Amount, DocumentNo, PaymentFlag from ODASMVoucher Where Prepared = 'Y' and approved = 'N'"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!vOUCHERnO))
                        
                        If Not IsNull(rsLIST!RequisitionDate) Then
                            MyList.SubItems(3) = CStr(rsLIST!RequisitionDate)
                        End If
                        
                        If Not IsNull(rsLIST!Reference) Then
                                MyList.SubItems(2) = CStr(rsLIST!Reference)
                        End If

                        If Not IsNull(rsLIST!Amount) Then
                                MyList.SubItems(4) = CStr(rsLIST!Amount)
                        End If
                        
                        If Not IsNull(rsLIST!DocumentNo) Then
                                MyList.SubItems(5) = CStr(rsLIST!DocumentNo)
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

Public Sub showALLCHKENTRIES()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView2.ListItems.Clear
                .ListView2.ColumnHeaders.Clear
                .ListView2.ColumnHeaders.Add , , "EntryNo", .ListView2.Width / 7
                .ListView2.ColumnHeaders.Add , , "Cheque No", .ListView2.Width / 7
                .ListView2.ColumnHeaders.Add , , "Voucher No", .ListView2.Width / 7
                .ListView2.ColumnHeaders.Add , , "Status", .ListView2.Width / 7
                .ListView2.ColumnHeaders.Add , , "Amount", .ListView2.Width / 7
                .ListView2.ColumnHeaders.Add , , "Period", .ListView2.Width / 7
                .ListView2.ColumnHeaders.Add , , "Bank ", .ListView2.Width / 7

                .ListView2.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select * from ALISMChequeEntry Where ChequeNo = '" & frmODASMCheck.txtChequeNo.Text & "'"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!ChequeEntryNo))
                        
                        If Not IsNull(rsLIST!ChequeNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!ChequeNo)
                        End If
                        
                        If Not IsNull(rsLIST!vOUCHERnO) Then
                                MyList.SubItems(2) = CStr(rsLIST!vOUCHERnO)
                        End If

                        If Not IsNull(rsLIST!Status) Then
                                MyList.SubItems(3) = CStr(rsLIST!Status)
                        End If
                        
                        If Not IsNull(rsLIST!Amount) Then
                                MyList.SubItems(4) = CStr(rsLIST!Amount)
                        End If
                        
                        If Not IsNull(rsLIST!CurrentPeriod) Then
                                MyList.SubItems(5) = CStr(rsLIST!CurrentPeriod)
                        End If
                        
                        If Not IsNull(rsLIST!BankNo) Then
                                MyList.SubItems(6) = CStr(rsLIST!BankNo)
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

Public Sub showALLCUSTOMERS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView3.ListItems.Clear
                .ListView3.ColumnHeaders.Clear
                .ListView3.ColumnHeaders.Add , , "Account No", .ListView3.Width / 2
                .ListView3.ColumnHeaders.Add , , "Names", .ListView3.Width / 2
      
                .ListView3.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPAccount, ODASPAccountType  Where ODASPAccount.Accounttype = ODASPAccount.AccountType And ODASPAccountType.cUSTOMER = 'y';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!AccountNo))
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                                MyList.SubItems(1) = CStr(rsLIST!CompanyName)
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
Public Sub showALLPaymentMethods()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "PaymentMethod", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 3

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ALISPPaymentMethod ;"
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

Public Sub showALLTOWNS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Town Code", .ListView1.Width / 3
                .ListView1.ColumnHeaders.Add , , "Town", .ListView1.Width / 3

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPTown ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!townCode))
                        
                        If Not IsNull(rsLIST!Town) Then
                            MyList.SubItems(1) = CStr(rsLIST!Town)
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

Public Sub showALLGUARANTOR()
On Error GoTo err
    
        With ALISFOManager
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Type", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Guarantor", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Status", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Req Remark", .ListView1.Width / 4

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASPGuarantor ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!GuarantorType))
                        
                        If Not IsNull(rsLIST!Guarantor) Then
                            MyList.SubItems(1) = CStr(rsLIST!Guarantor)
                        End If
                        
                        If Not IsNull(rsLIST!Status) Then
                            MyList.SubItems(2) = CStr(rsLIST!Status)
                        End If

                        If Not IsNull(rsLIST!RequireREMARK) Then
                            MyList.SubItems(3) = CStr(rsLIST!RequireREMARK)
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

Public Sub ShowALLJOBBRIEFS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView3.ListItems.Clear
                .ListView3.ColumnHeaders.Clear
                .ListView3.ColumnHeaders.Add , , "Job Brief No", .ListView3.Width / 2
                .ListView3.ColumnHeaders.Add , , "NAMES", .ListView3.Width / 2
                .ListView3.ColumnHeaders.Add , , "DOC", .ListView3.Width / 2
                .ListView3.ColumnHeaders.Add , , "Premium", .ListView3.Width / 2

                .ListView3.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT * FROM ODASMJobBrief JB, ODASPAccount A, ODASPAccountType T WHERE JB.AccountNo = A.AccountNo and A.AccountType = T.AccountType and T.Customer = 'Y';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!JobBriefNo))
                        
                        If Not IsNull(rsLIST!CompanyName) Then
                            MyList.SubItems(1) = CStr(rsLIST!CompanyName)
                        End If
                        
                        If Not IsNull(rsLIST!descriptionOfOrder) Then
                            MyList.SubItems(2) = CStr(rsLIST!descriptionOfOrder)
                        End If
                        
                        If Not IsNull(rsLIST!Balance) Then
                            MyList.SubItems(3) = CStr(rsLIST!Balance)
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

Public Sub showRECEIPTITEMS()
On Error GoTo err
    
        With frmODASMReceipt
        
                .ListView2.ListItems.Clear
                .ListView2.ColumnHeaders.Clear
                .ListView2.ColumnHeaders.Add , , "Document No", .ListView2.Width / 3
                .ListView2.ColumnHeaders.Add , , "No", .ListView2.Width / 3
                .ListView2.ColumnHeaders.Add , , "Type", .ListView2.Width / 3
                .ListView2.ColumnHeaders.Add , , "Amount", .ListView2.Width / 3
                .ListView2.ColumnHeaders.Add , , "Date", .ListView2.Width / 3
                .ListView2.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT DocumentNo, TransactionNo, ReceiptType, TransactionAmount, transactionDate FROM ALISMreceiptdetails WHERE ReceiptNo LIKE  '" & .txtReceiptNo.Text & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!DocumentNo))
                        
                        If Not IsNull(rsLIST!TransactionNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!TransactionNo)
                        End If

                        If Not IsNull(rsLIST!ReceiptType) Then
                            MyList.SubItems(2) = CStr(rsLIST!ReceiptType)
                        End If

                        If Not IsNull(rsLIST!TransactionAmount) Then
                            MyList.SubItems(3) = CStr(rsLIST!TransactionAmount)
                        End If
                        
                        If Not IsNull(rsLIST!TransactionDate) Then
                            MyList.SubItems(4) = CStr(rsLIST!TransactionDate)
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

Public Sub showJOBBRIEFRECEIPT()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView2.ListItems.Clear
                .ListView2.ColumnHeaders.Clear
                .ListView2.ColumnHeaders.Add , , "Receipt No", .ListView2.Width / 3
                .ListView2.ColumnHeaders.Add , , "No", .ListView2.Width / 3
                .ListView2.ColumnHeaders.Add , , "Type", .ListView2.Width / 3
                .ListView2.ColumnHeaders.Add , , "Amount", .ListView2.Width / 3
                .ListView2.ColumnHeaders.Add , , "Date", .ListView2.Width / 3
                .ListView2.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT DocumentNo, ReceiptNo, ReceiptType, TransactionAmount, transactionDate FROM ALISMreceiptdetails WHERE ReceiptNo LIKE  '" & Screen.ActiveForm.txtReceiptNo.Text & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!DocumentNo))
                        
                        If Not IsNull(rsLIST!TransactionNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!TransactionNo)
                        End If

                        If Not IsNull(rsLIST!ReceiptType) Then
                            MyList.SubItems(2) = CStr(rsLIST!ReceiptType)
                        End If

                        If Not IsNull(rsLIST!TransactionAmount) Then
                            MyList.SubItems(3) = CStr(rsLIST!TransactionAmount)
                        End If
                        
                        If Not IsNull(rsLIST!TransactionDate) Then
                            MyList.SubItems(4) = CStr(rsLIST!TransactionDate)
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

Public Sub showJOBBRIEFRECEIPTs()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView2.ListItems.Clear
                .ListView2.ColumnHeaders.Clear
                .ListView2.ColumnHeaders.Add , , "Receipt No", .ListView2.Width / 3
                .ListView2.ColumnHeaders.Add , , "No", .ListView2.Width / 3
                .ListView2.ColumnHeaders.Add , , "Type", .ListView2.Width / 3
                .ListView2.ColumnHeaders.Add , , "Amount", .ListView2.Width / 3
                .ListView2.ColumnHeaders.Add , , "Date", .ListView2.Width / 3
                .ListView2.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "SELECT DocumentNo, ReceiptNo, ReceiptType, TransactionAmount, transactionDate FROM ALISMreceiptdetails WHERE JobBriefNo =  '" & Screen.ActiveForm.txtJobBriefNo.Text & "';"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!DocumentNo))
                        
                        If Not IsNull(rsLIST!ReceiptNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!ReceiptNo)
                        End If

                        If Not IsNull(rsLIST!ReceiptType) Then
                            MyList.SubItems(2) = CStr(rsLIST!ReceiptType)
                        End If

                        If Not IsNull(rsLIST!TransactionAmount) Then
                            MyList.SubItems(3) = CStr(rsLIST!TransactionAmount)
                        End If
                        
                        If Not IsNull(rsLIST!TransactionDate) Then
                            MyList.SubItems(4) = CStr(rsLIST!TransactionDate)
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

Public Sub getRECEIPTS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Receipt No", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Trans No", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Type", .ListView1.Width / 4
                .ListView1.ColumnHeaders.Add , , "Trans Date", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Period", .ListView1.Width / 5
                .ListView1.ColumnHeaders.Add , , "Status", .ListView1.Width / 5

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select ALISMReceiptDetails.ReceiptNo, TransactionNo, ReceiptType, TransactionDate, TransactionAmount,  SuspenseAccount, AccountingPeriod, PaymentStatus from ALISMReceiptDetails Where ALISMReceiptDetails.DocumentNo = '" & frmALISMSuspense.cboDocumentNo.Text & "' order by ReceiptNo, TransactionNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ReceiptNo))
                        
                        If Not IsNull(rsLIST!TransactionNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!TransactionNo)
                        End If

                        If Not IsNull(rsLIST!ReceiptType) Then
                            MyList.SubItems(2) = CStr(rsLIST!ReceiptType)
                        End If
                        
                        If Not IsNull(rsLIST!TransactionDate) Then
                            MyList.SubItems(3) = CStr(rsLIST!TransactionDate)
                        End If

                        If Not IsNull(rsLIST!TransactionAmount) Then
                            MyList.SubItems(4) = CStr(rsLIST!TransactionAmount)
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

Public Sub showALLPreviousRECEIPTS()
On Error GoTo err
    
        With frmODASMReceipt
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Receipt No", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Trans No", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Type", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Trans Date", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Period", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Status", .ListView1.Width / 7

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select ALISMReceiptDetails.ReceiptNo, TransactionNo, ReceiptType, TransactionDate, TransactionAmount,  SuspenseAccount, AccountingPeriod, PaymentStatus from ALISMReceiptDetails Where ALISMReceiptDetails.DocumentNo = '" & .cboDocumentNo.Text & "' order by ReceiptNo, TransactionNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ReceiptNo))
                        
                        If Not IsNull(rsLIST!TransactionNo) Then
                            MyList.SubItems(1) = CStr(rsLIST!TransactionNo)
                        End If

                        If Not IsNull(rsLIST!ReceiptType) Then
                            MyList.SubItems(2) = CStr(rsLIST!ReceiptType)
                        End If
                        
                        If Not IsNull(rsLIST!TransactionDate) Then
                            MyList.SubItems(3) = CStr(rsLIST!TransactionDate)
                        End If

                        If Not IsNull(rsLIST!TransactionAmount) Then
                            MyList.SubItems(4) = CStr(rsLIST!TransactionAmount)
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

Public Sub showALLPreviousRECEIPTS1()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Receipt No", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Date", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Type", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Payer", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Period", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Status", .ListView1.Width / 7

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select ALISMReceiptNew.ReceiptNo, ReceiptDate, Remark, Payer, ReceiptAmount,  AccountingPeriod, PaymentStatus from ALISMReceiptNew Where ALISMReceiptNew.DocumentNo = '" & Screen.ActiveForm.cboDocumentNo.Text & "' order by ReceiptNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ReceiptNo))
                        
                        If Not IsNull(rsLIST!ReceiptDate) Then
                            MyList.SubItems(1) = CStr(rsLIST!ReceiptDate)
                        End If

                        If Not IsNull(rsLIST!remark) Then
                            MyList.SubItems(2) = CStr(rsLIST!remark)
                        End If
                        
                        If Not IsNull(rsLIST!Payer) Then
                            MyList.SubItems(3) = CStr(rsLIST!Payer)
                        End If

                        If Not IsNull(rsLIST!ReceiptAmount) Then
                            MyList.SubItems(4) = CStr(rsLIST!ReceiptAmount)
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

Public Sub showALLRECEIPTS()
On Error GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                .ListView1.ColumnHeaders.Add , , "Receipt No", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Date", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Remark", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Payer", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Amount", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Period", .ListView1.Width / 7
                .ListView1.ColumnHeaders.Add , , "Status", .ListView1.Width / 7

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select ALISMReceiptNew.ReceiptNo, ReceiptDate, Remark, Payer, ReceiptAmount,  AccountingPeriod, PaymentStatus from ALISMReceiptNew Where ALISMReceiptNew.DocumentNo = '" & Screen.ActiveForm.cboDocumentNo.Text & "' and AccountingPeriod = '" & ALISFOManager.txtcurrentPeriod.Text & "' order by ReceiptNo;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ReceiptNo))
                        
                        If Not IsNull(rsLIST!ReceiptDate) Then
                            MyList.SubItems(1) = CStr(rsLIST!ReceiptDate)
                        End If

                        If Not IsNull(rsLIST!remark) Then
                            MyList.SubItems(2) = CStr(rsLIST!remark)
                        End If
                        
                        If Not IsNull(rsLIST!Payer) Then
                            MyList.SubItems(3) = CStr(rsLIST!Payer)
                        End If

                        If Not IsNull(rsLIST!ReceiptAmount) Then
                            MyList.SubItems(4) = CStr(rsLIST!ReceiptAmount)
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

Public Sub getALLOPERATIONS()
On Error GoTo err

    With Screen.ActiveForm
            .ListView1.ListItems.Clear
            .ListView1.ColumnHeaders.Clear
            
            .ListView1.ColumnHeaders.Add , , "Operation", .ListView1.Width / 2#
            .ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 2
            
            .ListView1.View = lvwReport
            
            Dim rsLIST As ADODB.Recordset
            Set rsLIST = New ADODB.Recordset
            
            strSQL = "SELECT * FROM ODASPOperationType"
            rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
            Dim MyList As ListItem
            
            If rsLIST.EOF And rsLIST.BOF Then
                .ListView1.View = lvwList
                Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
                Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
            End If
            
            While Not rsLIST.EOF
            
            Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!OperationType))
            
                If Not IsNull(rsLIST!Description) Then
                    MyList.SubItems(1) = CStr(rsLIST!Description)
                End If
                 
                rsLIST.MoveNext
                
            Wend
        
        Set MyList = Nothing: Set rsLIST = Nothing
        
        End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Function FillList(Domain As String, lv As ListView, Optional FindString As String = "") As Boolean
 Screen.MousePointer = vbHourglass
      '==================================================================
      '  Purpose:   to fill a ListView control with data from a table or
      '             query
      '  Arguments: a Domain which is the name of the table or query, and
      '             a ListView control object
      '  Returns:   A Boolean value to indicate if the function was
      '             successful
      '==================================================================

      Dim rs As ADODB.Recordset
      Dim intTotCount As Integer
      Dim intCount1 As Integer, intCount2 As Integer
      Dim colNew As ColumnHeader, NewLine As ListItem

      On Error GoTo Err_Man

        ' Clear the ListView control.
        lv.ListItems.Clear
        lv.ColumnHeaders.Clear
    
        ' Set Variables.
         
        Set rs = New ADODB.Recordset
        cnCOMMON.CursorLocation = adUseClient
        rs.Open Domain, cnCOMMON, adOpenStatic, adLockOptimistic
       
        If Trim(FindString) = "" Then
        Else
                Dim strFilterString
                strFilterString = ""
            
                'Build filter string
                For i = 0 To rs.Fields.Count - 1
    
'                        If rs.Fields(i).Type = 202 Then
                                        strFilterString = strFilterString & "[" & rs.Fields(i).Name & "] like '%" & FindString & "%' " & " OR "
'                        End If
                        
                Next i
                'remove the last part of the string " OR "
                strFilterString = Left(strFilterString, Len(strFilterString) - Len(" OR "))
                
                rs.Filter = strFilterString
        End If
      
        ' Set Column Headers.
        For intCount1 = 0 To rs.Fields.Count - 1
             Set colNew = lv.ColumnHeaders.Add(, , rs(intCount1).Name, 1850)
        Next intCount1
        lv.View = lvwReport    ' Set View property to 'Report'.
    
        If rs.EOF Or rs.BOF Then
                    lv.View = lvwList
                    Set NewLine = lv.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
                    lv.Enabled = False
                    Set rs = Nothing: Set NewLine = Nothing:  Screen.MousePointer = vbDefault: Exit Function
                    
        End If
        lv.Enabled = True
        ' Set Total Records Counter.
        rs.MoveLast
        intTotCount = rs.AbsolutePosition
        rs.MoveFirst
        If intTotCount = -1 Then
                rs.MoveFirst
                intTotCount = 0
                While Not rs.EOF
                        intTotCount = intTotCount + 1
                        DoEvents
                rs.MoveNext
                Wend
        End If
        
        rs.MoveFirst
          
        ' Loop through recordset and add Items to the control.
        For intCount1 = 1 To intTotCount
                    If IsNumeric(rs(0).Value) Then
                        Set NewLine = lv.ListItems.Add(, , CStr(rs(0).Value))
                    Else
                        Set NewLine = lv.ListItems.Add(, , rs(0).Value)
                    End If
                          
                    For intCount2 = 1 To rs.Fields.Count - 1
                            If Not IsNull(rs(intCount2)) Then
                                    NewLine.SubItems(intCount2) = rs(intCount2).Value
                            End If
                    Next intCount2
    
                    rs.MoveNext
                DoEvents
        Next intCount1
        
        cnCOMMON.CursorLocation = adUseServer
        
        If lv.ListItems.Count = 1 Then
            lv.ListItems(1).Checked = True
        End If
        
 Screen.MousePointer = vbDefault
    Exit Function

Err_Man:
         ' Ignore Error 94 which indicates you passed a NULL value.
         If err = 94 Then
            Resume Next
         Else
         ' Otherwise display the error message.
            MsgBox "Error: " & err.Number & Chr(13) & _
               Chr(10) & err.Description
         End If
Screen.MousePointer = vbDefault
      End Function
      
      
Public Sub checkOne(Item, lstView As ListView)
        Dim i, j As Double
        
        If Item.Checked = True Then
                    j = lstView.ListItems.Count
                    
                    If j = 0 Then Exit Sub
                    
                    For i = 1 To j
                                If lstView.ListItems(i) <> Item Then
                                   lstView.ListItems(i).Checked = False
                                End If
                    Next i
        Else
                    Item.Checked = False
        End If
End Sub

Public Sub checkAll(lstView As ListView)
        Dim i, j As Double
        
        j = lstView.ListItems.Count
                    
        If j = 0 Then Exit Sub
                    
        For i = 1 To j
                    lstView.ListItems(i).Checked = True
        Next i
End Sub

Public Sub UnCheckAll(lstView As ListView)
        Dim i, j As Double
        
        j = lstView.ListItems.Count
                    
        If j = 0 Then Exit Sub
                    
        For i = 1 To j
                    lstView.ListItems(i).Checked = False
        Next i
End Sub

Function SortListViewColumn(lv As Object, ColumnHeader)
'Check if the Sortkey is the same a the current one
    If lv.SortKey <> ColumnHeader.Index - 1 Then
        'When a column is clicked set the sortkey
        'to the columnheader index -1
        lv.SortKey = ColumnHeader.Index - 1
        lv.SortOrder = lvwAscending
    Else
        'If the column is already selected then change the
        'sortorder to be the opposite of what is currently
        'being used
        lv.SortOrder = IIf(lv.SortOrder = lvwAscending, _
                                lvwDescending, lvwAscending)
    End If
    
    'Set the sorted property to use the new sortkey
    'and sort the contents
    lv.Sorted = True
End Function

'Procedure used to search in listview
Public Sub search_in_listview(ByRef sListView As ListView, ByVal sFindText As String)
    Dim tmp_listtview As ListItem
    Set tmp_listtview = sListView.FindItem(sFindText, lvwSubItem)
    If Not tmp_listtview Is Nothing Then
        tmp_listtview.EnsureVisible
        tmp_listtview.Selected = True
    End If
End Sub


