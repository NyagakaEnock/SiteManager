VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSysManager"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowSystem()
On Error GoTo err
    Load frmLogin
    frmLogin.Show 1
    Exit Sub
err:
    ErrorMessage
End Sub

Public Sub OpenMains()
On Error GoTo err
    Load ALISFOManager
    ALISFOManager.Show 1
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub NewTreeSetup()
On Error GoTo err
With ALISFOManager
       
    Dim NodB As Node
        Set NodB = .TreeView2.Nodes.Add(, , "C", "Receipt Manager", "C")
        Set NodB = .TreeView2.Nodes.Add("C", tvwChild, "C1", "Receipts", "F")
        Set NodB = .TreeView2.Nodes.Add("C", tvwChild, "C4", "Print Receipts - PeriodWise", "F")
        Set NodB = .TreeView2.Nodes.Add("C", tvwChild, "C5", "Ledger Details", "F")
        Set NodB = .TreeView2.Nodes.Add("C", tvwChild, "C6", "Suspense Payment", "F")
        Set NodB = .TreeView2.Nodes.Add("C", tvwChild, "C8", "Void Payments", "F")
        Set NodB = .TreeView2.Nodes.Add("C", tvwChild, "C7", "Voided Payments", "F")
        Set NodB = .TreeView2.Nodes.Add("C", tvwChild, "C9", "Current Receipts", "F")
        Set NodB = .TreeView2.Nodes.Add("C", tvwChild, "C11", "Bank Receipts", "F")
        
    Dim NodR As Node
        Set NodR = .TreeView2.Nodes.Add(, , "R", "Payments Requisition", "C")
        Set NodR = .TreeView2.Nodes.Add("R", tvwChild, "R2", "Receive Invoices", "F")
        Set NodR = .TreeView2.Nodes.Add("R", tvwChild, "R3", "show Invoices Received Invoices", "F")
        Set NodR = .TreeView2.Nodes.Add("R", tvwChild, "R1", "Prepare Voucher", "F")
            If rsDEFAULT!AutoApproval <> "Y" Then
                    Set NodR = .TreeView2.Nodes.Add("R", tvwChild, "R10", "VoucherApproval", "F")
            End If
        Set NodR = .TreeView2.Nodes.Add("R", tvwChild, "R11", "VoucherAuthorization", "F")
        Set NodR = .TreeView2.Nodes.Add("R", tvwChild, "R7", "Print Voucher", "F")
        Set NodR = .TreeView2.Nodes.Add("R", tvwChild, "K6", "Reverse Payment", "F")
        Set NodR = .TreeView2.Nodes.Add("R", tvwChild, "R8", "Summary of Vouchers ", "F")
        Set NodR = .TreeView2.Nodes.Add("R", tvwChild, "R9", "Authorized Vouchers", "F")
    
    If rsDEFAULT!AutopaymentUpdate <> "Y" Then
        Dim NodK As Node
            Set NodK = .TreeView2.Nodes.Add(, , "K", "Cheque Requisition", "C")
            Set NodK = .TreeView2.Nodes.Add("K", tvwChild, "K2", "Cheque Processing", "F")
            Set NodK = .TreeView2.Nodes.Add("K", tvwChild, "K4", "Approval/Authorization", "F")
                If rsDEFAULT!AutoApproval <> "Y" Then
                        Set NodK = .TreeView2.Nodes.Add("K4", tvwChild, "K41", "ChequeApproval", "F")
                End If
                Set NodK = .TreeView2.Nodes.Add("K4", tvwChild, "K42", "ChequeApproval", "F")
            Set NodK = .TreeView2.Nodes.Add("K", tvwChild, "K3", "Issue Cheque", "F")
            
        End If
    Dim NodT As Node
        Set NodT = .TreeView2.Nodes.Add(, , "T", "Credit Manager", "C")
        
        If rsDEFAULT!AutoCreditAuthorization <> "Y" Then
                Set NodT = .TreeView2.Nodes.Add("T", tvwChild, "T1", "Authorize Credit", "F")
                Set NodT = .TreeView2.Nodes.Add("T", tvwChild, "T2", "Credit Authorized Briefs ", "F")
        End If
        
        Set NodT = .TreeView2.Nodes.Add("T", tvwChild, "T3", "Create Installments", "F")
        Set NodT = .TreeView2.Nodes.Add("T", tvwChild, "T11", "Installment Payment", "F")
        Set NodT = .TreeView2.Nodes.Add("T", tvwChild, "T4", "Prepare Invoices", "F")
        
        Set NodT = .TreeView2.Nodes.Add("T", tvwChild, "T51", "Approve Invoices", "F")
        Set NodT = .TreeView2.Nodes.Add("T", tvwChild, "T5", "Authorization", "F")
        Set NodT = .TreeView2.Nodes.Add("T", tvwChild, "T7", "Print Invoice", "F")
        Set NodT = .TreeView2.Nodes.Add("T", tvwChild, "T8", "List Invoices", "F")
        Set NodT = .TreeView2.Nodes.Add("T", tvwChild, "T9", "Invoices Paid", "F")
        Set NodT = .TreeView2.Nodes.Add("T", tvwChild, "T6", "Invoice Listings", "F")
        Set NodT = .TreeView2.Nodes.Add("T", tvwChild, "T10", "Print Installments", "F")
    
    Dim NodA As Node
        Set NodA = .TreeView2.Nodes.Add(, , "A", "Statements", "C")
        Set NodA = .TreeView2.Nodes.Add("A", tvwChild, "A1", "Print Statement", "F")
        
    Dim NodD As Node
        Set NodD = .TreeView2.Nodes.Add(, , "D", "Rates / Settings / Setup", "C")
        Set NodD = .TreeView2.Nodes.Add("D", tvwChild, "D7", "Land Rates", "F")
        Set NodD = .TreeView2.Nodes.Add("D", tvwChild, "D8", "Create Guarantor", "F")
        Set NodD = .TreeView2.Nodes.Add("D", tvwChild, "D9", "Accounting Periods", "F")
        Set NodD = .TreeView2.Nodes.Add("D", tvwChild, "P1", "MS Office Settings", "F")
    Dim NodP As Node
        Set NodP = .TreeView2.Nodes.Add(, , "P", "Listings / Options", "C")
        Set NodP = .TreeView2.Nodes.Add("P", tvwChild, "P3", "View Currencies", "F")
        Set NodP = .TreeView2.Nodes.Add("P", tvwChild, "P4", "View V.A.T. Rate", "F")
        Set NodP = .TreeView2.Nodes.Add("P", tvwChild, "D1", "Payment Type", "F")
        Set NodP = .TreeView2.Nodes.Add("P", tvwChild, "D2", "Payment Configuration", "F")
        Set NodP = .TreeView2.Nodes.Add("P", tvwChild, "D5", "Reversal Types", "F")
        Set NodP = .TreeView2.Nodes.Add("P", tvwChild, "D6", "Payment Methods", "F")
        Set NodP = .TreeView2.Nodes.Add("P", tvwChild, "P5", "View Country Code [Telecom]", "F")
        Set NodP = .TreeView2.Nodes.Add("P", tvwChild, "P6", "View Local/Area Code [Telecom]", "F")
        Set NodP = .TreeView2.Nodes.Add("P", tvwChild, "P7", "View Default Payment Method", "F")
        Set NodP = .TreeView2.Nodes.Add("P", tvwChild, "V3", "Authorized Currencies", "F")
        Set NodP = .TreeView2.Nodes.Add("P", tvwChild, "V4", "Taxes and Rates", "F")
        Set NodP = .TreeView2.Nodes.Add("P", tvwChild, "D10", "Duration", "F")
        .TreeView2.BorderStyle = vbFixedSingle
        
End With
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub ResizeControls()
On Error GoTo err
    With ALISFOManager
        .TreeView2.Height = .Height - (9000 - 7335)
        .ListView1.Height = .Height - (9000 - 7335)
        
        .ListView1.Width = .Width - (12000 - 9135)
        .ProgressBar1.Width = .Width
    End With
Exit Sub
err:
If err.Number = 380 Then Resume Next
    ErrorMessage
End Sub
Public Sub HelpContents()
On Error GoTo err
With ALISFOManager
    Dim nRet As Integer
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, .Caption
    Else
        '''On Error Resume Next
        nRet = WinHelp(.hwnd, App.HelpFile, 3, 0)
        If err Then
            MsgBox err.Description
        End If
    End If
End With
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub HelpIndex()
On Error GoTo err
With ALISFOManager
    Dim nRet As Integer

    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, .Caption
    Else
        '''On Error Resume Next
        nRet = WinHelp(.hwnd, App.HelpFile, 261, 0)
        If err Then
            MsgBox err.Description
        End If
    End If
End With
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub HelpAbout()
On Error GoTo err
    Load frmAbout
    frmAbout.Show 1
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub GetMainStructure()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Company Code", .ListView1.Width / 7
.ListView1.ColumnHeaders.Add , , "Name of Company", .ListView1.Width / 3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Postal Address", .ListView1.Width / 3.5
.ListView1.ColumnHeaders.Add , , "Country", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Phone No", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Fax/Telex", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "E-Mail Address", .ListView1.Width / 4.5
.ListView1.ColumnHeaders.Add , , "I.T. NO", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "PIN NUMBER", .ListView1.Width / 5.8 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "NHIF NO", .ListView1.Width / 6

.ListView1.View = lvwReport
End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub FindMyCompany()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Company Code", .ListView1.Width / 7
.ListView1.ColumnHeaders.Add , , "Name of Company", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 2 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Postal Address", .ListView1.Width / 2.5
.ListView1.ColumnHeaders.Add , , "Country", .ListView1.Width / 7
.ListView1.ColumnHeaders.Add , , "Phone No", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Fax/Telex", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "E-Mail Address", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "I.T. NO", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "PIN NUMBER", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "NHIF NO", .ListView1.Width / 5.5

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamCompanyMaster ORDER BY CompanyCode;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!CompanyCode))

    If Not IsNull(rsLIST!CompanyName) Then
        MyList.SubItems(1) = CStr(rsLIST!CompanyName)
    End If
    
    If Not IsNull(rsLIST!PhysicalAddress) Then
        MyList.SubItems(2) = CStr(rsLIST!PhysicalAddress)
    End If
    
    If Not IsNull(rsLIST!PostalAddress) And Not IsNull(rsLIST!city) Then
        MyList.SubItems(3) = CStr(Trim(rsLIST!PostalAddress)) & " " & CStr(Trim(rsLIST!city))
    End If
    
    If Not IsNull(rsLIST!country) Then
        MyList.SubItems(4) = CStr(rsLIST!country)
    End If
    
    If Not IsNull(rsLIST!TelephoneNo) Then
        MyList.SubItems(5) = CStr(rsLIST!TelephoneNo)
    End If
    
    If Not IsNull(rsLIST!telexfax) Then
        MyList.SubItems(6) = CStr(Trim(rsLIST!telexfax))
    End If
    
    If Not IsNull(rsLIST!email) Then
        MyList.SubItems(7) = CStr(rsLIST!email)
    End If
    
    If Not IsNull(rsLIST!coyitno) Then
        MyList.SubItems(8) = CStr(rsLIST!coyitno)
    End If
    
    If Not IsNull(rsLIST!coypinno) Then
        MyList.SubItems(9) = CStr(Trim(rsLIST!coypinno))
    End If
    
    If Not IsNull(rsLIST!coynhifno) Then
        MyList.SubItems(10) = CStr(Trim(rsLIST!coynhifno))
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

Public Sub FindCompanyBranch()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Branch Code", .ListView1.Width / 7
.ListView1.ColumnHeaders.Add , , "Name of Branch", .ListView1.Width / 3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Postal Address", .ListView1.Width / 2.5
.ListView1.ColumnHeaders.Add , , "Town/City", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Phone No", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Fax/Telex", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "E-Mail Address", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Contact Name", .ListView1.Width / 3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Staff ID", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Official Title", .ListView1.Width / 3

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamCompanyBranch ORDER BY BranchCode;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!BranchCode))

    If Not IsNull(rsLIST!BranchName) Then
        MyList.SubItems(1) = CStr(rsLIST!BranchName)
    End If
    
    If Not IsNull(rsLIST!phyadd) Then
        MyList.SubItems(2) = CStr(rsLIST!phyadd)
    End If
    
    If Not IsNull(rsLIST!postadd) And Not IsNull(rsLIST!towncity) Then
        MyList.SubItems(3) = CStr(Trim(rsLIST!postadd)) & " " & CStr(Trim(rsLIST!towncity))
    End If
    
    If Not IsNull(rsLIST!towncity) Then
        MyList.SubItems(4) = CStr(rsLIST!towncity)
    End If
    
    If Not IsNull(rsLIST!phoneno) Then
        MyList.SubItems(5) = CStr(rsLIST!phoneno)
    End If
    
    If Not IsNull(rsLIST!faxtelex) Then
        MyList.SubItems(6) = CStr(Trim(rsLIST!faxtelex))
    End If
    
    If Not IsNull(rsLIST!email) Then
        MyList.SubItems(7) = CStr(rsLIST!email)
    End If
    
    If Not IsNull(rsLIST!Contactname) Then
        MyList.SubItems(8) = CStr(rsLIST!Contactname)
    End If
    
    If Not IsNull(rsLIST!StaffIdNo) Then
        MyList.SubItems(9) = CStr(Trim(rsLIST!StaffIdNo))
    End If
    
    If Not IsNull(rsLIST!officialtitle) Then
        MyList.SubItems(10) = CStr(Trim(rsLIST!officialtitle))
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

Public Sub FindMSOffice()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Code", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Applications", .ListView1.Width / 2

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamMSOffice ORDER BY Code;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!CODE))

    If Not IsNull(rsLIST!appname) Then
        MyList.SubItems(1) = CStr(rsLIST!appname)
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

Public Sub FindAreaCode()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Code", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Area Name", .ListView1.Width / 3
.ListView1.ColumnHeaders.Add , , "Area Code", .ListView1.Width / 4

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM Setareacode ORDER BY Codenumber;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!CodeNumber))

    If Not IsNull(rsLIST!areaname) Then
        MyList.SubItems(1) = CStr(rsLIST!areaname)
    End If
    
    If Not IsNull(rsLIST!AreaCode) Then
        MyList.SubItems(2) = CStr(rsLIST!AreaCode)
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

Public Sub FindCountryCode()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Code", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Country Name", .ListView1.Width / 3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Country Code", .ListView1.Width / 4 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM Setcountrycode ORDER BY Codenumber;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df

Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!CodeNumber))

    If Not IsNull(rsLIST!countryname) Then
        MyList.SubItems(1) = CStr(rsLIST!countryname)
    End If
    
    If Not IsNull(rsLIST!CountryCode) Then
        MyList.SubItems(2) = CStr(rsLIST!CountryCode)
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

Public Sub FindLocalCurrency()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Code", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "Currency Name", .ListView1.Width / 3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Curency Symbol", .ListView1.Width / 3 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM SetLocalCurrency ORDER BY Codenumber;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!CodeNumber))

    If Not IsNull(rsLIST!currencyname) Then
        MyList.SubItems(1) = CStr(rsLIST!currencyname)
    End If
    
    If Not IsNull(rsLIST!Currency) Then
        MyList.SubItems(2) = CStr(rsLIST!Currency)
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

Public Sub FindDefaultPayMethod()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Serial No", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Payment Code", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Payment Method", .ListView1.Width / 4 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM SetDefaultPayMethod ORDER BY serialno;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SerialNo))

    If Not IsNull(rsLIST!paycode) Then
        MyList.SubItems(1) = CStr(rsLIST!paycode)
    End If
    
    If Not IsNull(rsLIST!PayMethod) Then
        MyList.SubItems(2) = CStr(rsLIST!PayMethod)
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

Public Sub FindDefaultCurrency()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Code", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Currency Name", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Currency Symbol", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Exchange Rate", .ListView1.Width / 4 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM SetDefaultCurrency ORDER BY codenumber;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!CodeNumber))

    If Not IsNull(rsLIST!currencyname) Then
        MyList.SubItems(1) = CStr(rsLIST!currencyname)
    End If
    
    If Not IsNull(rsLIST!Currency) Then
        MyList.SubItems(2) = CStr(rsLIST!Currency)
    End If
    
    If Not IsNull(rsLIST!exchrate) Then
        MyList.SubItems(3) = CStr(rsLIST!exchrate)
        MyList.SubItems(3) = FormatNumber(rsLIST!exchrate, 5, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    rsLIST.MoveNext
    
Wend
.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
Set MyList = Nothing
End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub FindCompany()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Dept. Code", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Department Name", .ListView1.Width / 2.3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Dept. Head", .ListView1.Width / 4.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Notes/Descriptions", .ListView1.Width / 1.3

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamCompanyDepts ORDER BY DeptCode;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
df = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!deptcode))

    If Not IsNull(rsLIST!deptname) Then
        MyList.SubItems(1) = CStr(rsLIST!deptname)
    End If
    
    If Not IsNull(rsLIST!hod) Then
        MyList.SubItems(2) = CStr(rsLIST!hod)
    End If
    
    If Not IsNull(rsLIST!deptnotes) Then
        MyList.SubItems(3) = CStr(rsLIST!deptnotes)
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

Public Sub FindDepartments()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Dept. Code", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Department Name", .ListView1.Width / 2.3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Head of Department", .ListView1.Width / 4.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Staff ID", .ListView1.Width / 8 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Official Title", .ListView1.Width / 3.1 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Notes/Comments", .ListView1.Width / 1.4

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamCompanyDepts ORDER BY DeptCode;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
df = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!deptcode))

    If Not IsNull(rsLIST!deptname) Then
        MyList.SubItems(1) = CStr(rsLIST!deptname)
    End If
    
    If Not IsNull(rsLIST!hod) Then
        MyList.SubItems(2) = CStr(rsLIST!hod)
    End If
    
    If Not IsNull(rsLIST!hodstaffid) Then
        MyList.SubItems(3) = CStr(rsLIST!hodstaffid)
    End If
    
    If Not IsNull(rsLIST!officialtitle) Then
        MyList.SubItems(4) = CStr(rsLIST!officialtitle)
    End If
    
    If Not IsNull(rsLIST!deptnotes) Then
        MyList.SubItems(5) = CStr(rsLIST!deptnotes)
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

Public Sub FindEmployers()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Emp. Code", .ListView1.Width / 7
.ListView1.ColumnHeaders.Add , , "Name of Employer", .ListView1.Width / 3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Postal Address", .ListView1.Width / 3
.ListView1.ColumnHeaders.Add , , "Country", .ListView1.Width / 7
.ListView1.ColumnHeaders.Add , , "Phone No", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Fax/Telex", .ListView1.Width / 7
.ListView1.ColumnHeaders.Add , , "Mobile No", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contact Person", .ListView1.Width / 3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Official Title", .ListView1.Width / 3

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamEmployer ORDER BY empcode;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
df = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!empcode))

    If Not IsNull(rsLIST!empname) Then
        MyList.SubItems(1) = CStr(rsLIST!empname)
    End If
    
    If Not IsNull(rsLIST!PhysicalAddress) Then
        MyList.SubItems(2) = CStr(rsLIST!PhysicalAddress)
    End If
    
    If Not IsNull(rsLIST!PostalAddress) And Not IsNull(rsLIST!towncity) Then
        MyList.SubItems(3) = CStr(Trim(rsLIST!PostalAddress)) & " " & CStr(Trim(rsLIST!towncity))
    End If
    
    If Not IsNull(rsLIST!country) Then
        MyList.SubItems(4) = CStr(rsLIST!country)
    End If
    
    If Not IsNull(rsLIST!TelephoneNo) Then
        MyList.SubItems(5) = CStr(rsLIST!TelephoneNo)
    End If
    
    If Not IsNull(rsLIST!mobileno) Then
        MyList.SubItems(6) = CStr(rsLIST!mobileno)
    End If
    
    If Not IsNull(rsLIST!faxtelex) Then
        MyList.SubItems(7) = CStr(rsLIST!faxtelex)
    End If
    
    If Not IsNull(rsLIST!contallnames) Then
        MyList.SubItems(8) = CStr(rsLIST!contallnames)
    End If
    
    If Not IsNull(rsLIST!contofficialtitle) Then
        MyList.SubItems(9) = CStr(rsLIST!contofficialtitle)
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

Public Sub FindAccPeriods()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Period Code", .ListView1.Width / 7
.ListView1.ColumnHeaders.Add , , "Month Name", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Quarter", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Sample Period", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Start Date", .ListView1.Width / 5.2
.ListView1.ColumnHeaders.Add , , "End Date", .ListView1.Width / 5.2

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamAccPeriod ORDER BY periodcode;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
df = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PeriodCode))

    If Not IsNull(rsLIST!AccPeriod) Then
        MyList.SubItems(1) = CStr(rsLIST!AccPeriod)
    End If
    
    If Not IsNull(rsLIST!quarterofyear) Then
        MyList.SubItems(2) = CStr(rsLIST!quarterofyear)
    End If
    
    If Not IsNull(rsLIST!fullperiod) Then
        MyList.SubItems(3) = CStr(rsLIST!fullperiod)
    End If
    
    If Not IsNull(rsLIST!StartDate) Then
        MyList.SubItems(4) = CStr(rsLIST!StartDate)
    End If
    
    If Not IsNull(rsLIST!EndDate) Then
        MyList.SubItems(5) = CStr(rsLIST!EndDate)
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

Public Sub FindTaxesRates()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Tax Code", .ListView1.Width / 7
.ListView1.ColumnHeaders.Add , , "Tax Name", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Tax Rate[%]", .ListView1.Width / 5 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamTaxes ORDER BY taxcode;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
df = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!TaxCode))

    If Not IsNull(rsLIST!taxname) Then
        MyList.SubItems(1) = CStr(rsLIST!taxname)
    End If
    
    If Not IsNull(rsLIST!TaxRate) Then
        MyList.SubItems(2) = CStr(rsLIST!TaxRate)
        MyList.SubItems(2) = FormatNumber(rsLIST!TaxRate, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    rsLIST.MoveNext
    
Wend

.ListView1.ColumnHeaders(3).Alignment = lvwColumnRight

Set MyList = Nothing
End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub FindFeeServices()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Code No", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Service / Fee", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 1.6 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamFeesServices ORDER BY codenumber;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
df = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!CodeNumber))

    If Not IsNull(rsLIST!servicefee) Then
        MyList.SubItems(1) = CStr(rsLIST!servicefee)
    End If
    
    If Not IsNull(rsLIST!descriptions) Then
        MyList.SubItems(2) = CStr(rsLIST!descriptions)
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

Public Sub FindPaymentMethods()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Code No", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "Payment Method", .ListView1.Width / 3 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamPayMethods ORDER BY codenumber;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
df = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!CodeNumber))

    If Not IsNull(rsLIST!PayMethod) Then
        MyList.SubItems(1) = CStr(rsLIST!PayMethod)
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

Public Sub FindTitles()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Title ID", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Title of Courtesy", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 3 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamTitles ORDER BY TitleID;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
df = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!titleid))

    If Not IsNull(rsLIST!Title) Then
        MyList.SubItems(1) = CStr(rsLIST!Title)
    End If
    
    If Not IsNull(rsLIST!Description) Then
        MyList.SubItems(2) = CStr(rsLIST!Description)
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

.ListView1.ColumnHeaders.Add , , "Curr. Code", .ListView1.Width / 7
.ListView1.ColumnHeaders.Add , , "Name of Currency", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Symbol", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Exchange Rate", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Country Code", .ListView1.Width / 7

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamCurrencies ORDER BY codenumber;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
df = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!CodeNumber))

    If Not IsNull(rsLIST!desccurrency) Then
        MyList.SubItems(1) = CStr(rsLIST!desccurrency)
    End If
    
    If Not IsNull(rsLIST!Currency) Then
        MyList.SubItems(2) = CStr(rsLIST!Currency)
    End If
    
    If Not IsNull(rsLIST!exchrate) Then
        MyList.SubItems(3) = CStr(rsLIST!exchrate)
        MyList.SubItems(3) = FormatNumber(rsLIST!exchrate, 5, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!CountryCode) Then
        MyList.SubItems(4) = CStr(rsLIST!CountryCode)
    End If

    rsLIST.MoveNext
    
Wend

.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight

Set MyList = Nothing
End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub FindCities()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "City Code", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Name of City", .ListView1.Width / 2.3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Coyntry Code", .ListView1.Width / 4.5 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamCities ORDER BY citycode;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
df = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!citycode))

    If Not IsNull(rsLIST!cityname) Then
        MyList.SubItems(1) = CStr(rsLIST!cityname)
    End If
    
    If Not IsNull(rsLIST!CountryCode) Then
        MyList.SubItems(2) = CStr(rsLIST!CountryCode)
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

Public Sub FindCountries()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Country Code", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "Country Name", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Commercial Capital", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Continent", .ListView1.Width / 4

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamCountries ORDER BY countrycode;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
df = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!CountryCode))

    If Not IsNull(rsLIST!country) Then
        MyList.SubItems(1) = CStr(rsLIST!country)
    End If
    
    If Not IsNull(rsLIST!capitalcity) Then
        MyList.SubItems(2) = CStr(rsLIST!capitalcity)
    End If
    
    If Not IsNull(rsLIST!continent) Then
        MyList.SubItems(3) = CStr(rsLIST!continent)
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

Public Sub FindEmployees()
On Error GoTo err
With ALISFOManager

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Staff ID", .ListView1.Width / 7.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Full Names", .ListView1.Width / 3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Birth Date", .ListView1.Width / 7.5
.ListView1.ColumnHeaders.Add , , "Marital Status", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Gender", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Nationality", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Nat. ID No", .ListView1.Width / 7.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Passport No", .ListView1.Width / 7.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "PIN Number", .ListView1.Width / 7.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Date Hired", .ListView1.Width / 7.5
.ListView1.ColumnHeaders.Add , , "Official Title", .ListView1.Width / 3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Employ. Type", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Grade", .ListView1.Width / 9 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Dept. Code", .ListView1.Width / 8 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 2.3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Postal Address", .ListView1.Width / 3.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Country", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Phone No", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Mobile No", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "E-Mail Address", .ListView1.Width / 4.5 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ParamEmpMaster ORDER BY StaffIDNo;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
df = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!StaffIdNo))

    If Not IsNull(rsLIST!allnames) Then
        MyList.SubItems(1) = CStr(rsLIST!allnames)
    End If
    
    If Not IsNull(rsLIST!DateofBirth) Then
        MyList.SubItems(2) = CStr(rsLIST!DateofBirth)
    End If
    
    If Not IsNull(rsLIST!maritalstatus) Then
        MyList.SubItems(3) = CStr(rsLIST!maritalstatus)
    End If
    
    If Not IsNull(rsLIST!gender) Then
        MyList.SubItems(4) = CStr(rsLIST!gender)
    End If
  
    If Not IsNull(rsLIST!nationality) Then
        MyList.SubItems(5) = CStr(rsLIST!nationality)
    End If
    
    If Not IsNull(rsLIST!natidno) Then
        MyList.SubItems(6) = CStr(rsLIST!natidno)
    End If
    
    If Not IsNull(rsLIST!passportno) Then
        MyList.SubItems(7) = CStr(rsLIST!passportno)
    End If
    
    If Not IsNull(rsLIST!pinnumber) Then
        MyList.SubItems(8) = CStr(rsLIST!pinnumber)
    End If
    
    If Not IsNull(rsLIST!datehired) Then
        MyList.SubItems(9) = CStr(rsLIST!datehired)
    End If
  
    If Not IsNull(rsLIST!officialtitle) Then
        MyList.SubItems(10) = CStr(rsLIST!officialtitle)
    End If
    
    If Not IsNull(rsLIST!employtype) Then
        MyList.SubItems(11) = CStr(rsLIST!employtype)
    End If
    
    If Not IsNull(rsLIST!gradecode) Then
        MyList.SubItems(12) = CStr(rsLIST!gradecode)
    End If
    
    If Not IsNull(rsLIST!deptcode) Then
        MyList.SubItems(13) = CStr(rsLIST!deptcode)
    End If
    
    If Not IsNull(rsLIST!PhysicalAddress) Then
        MyList.SubItems(14) = CStr(rsLIST!PhysicalAddress)
    End If
    
    If Not IsNull(rsLIST!PostalAddress) And Not IsNull(rsLIST!conttowncity) Then
        MyList.SubItems(15) = CStr(Trim(rsLIST!PostalAddress)) & " " & CStr(Trim(rsLIST!conttowncity))
    End If
  
    If Not IsNull(rsLIST!contcountry) Then
        MyList.SubItems(16) = CStr(rsLIST!contcountry)
    End If
    
    If Not IsNull(rsLIST!conttelephone) Then
        MyList.SubItems(17) = CStr(rsLIST!conttelephone)
    End If
    
    If Not IsNull(rsLIST!contmobile) Then
        MyList.SubItems(18) = CStr(rsLIST!contmobile)
    End If
  
    If Not IsNull(rsLIST!contemail) Then
        MyList.SubItems(19) = CStr(rsLIST!contemail)
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

Public Sub FindAllClaimsDrugs()
On Error GoTo err
With ALISFOManager
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Visit No", .ListView1.Width / 7.5
.ListView1.ColumnHeaders.Add , , "Membership No", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Family No", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Reg. Date", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Trans. Date", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Company", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Procedures", .ListView1.Width / 7
.ListView1.ColumnHeaders.Add , , "Visit Date", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Visit Time", .ListView1.Width / 7
.ListView1.ColumnHeaders.Add , , "Clinic", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Drug Cost", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Proc. Cost", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Total Cost", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Claim With", .ListView1.Width / 7
.ListView1.ColumnHeaders.Add , , "Description", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Approval", .ListView1.Width / 7

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ProcessClaims;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim df
df = rsLIST.RecordCount
Dim MyList As ListItem
   
While Not rsLIST.EOF
Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!visitno))

    If Not IsNull(rsLIST!visitno) Then
        MyList.SubItems(1) = CStr(rsLIST!visitno)
    End If
    
    If Not IsNull(rsLIST!membershipno) Then
        MyList.SubItems(2) = CStr(rsLIST!membershipno)
    End If
    
    If Not IsNull(rsLIST!familyno) Then
        MyList.SubItems(3) = CStr(rsLIST!familyno)
    End If
    
    If Not IsNull(rsLIST!registrationdate) Then
        MyList.SubItems(4) = CStr(rsLIST!registrationdate)
    End If
  
    If Not IsNull(rsLIST!TransactionDate) Then
        MyList.SubItems(5) = CStr(rsLIST!TransactionDate)
    End If
    
    If Not IsNull(rsLIST!company) Then
        MyList.SubItems(6) = CStr(rsLIST!company)
    End If
    
    If Not IsNull(rsLIST!procedures) Then
        MyList.SubItems(7) = CStr(rsLIST!procedures)
    End If
    
    If Not IsNull(rsLIST!visitdate) Then
        MyList.SubItems(8) = CStr(rsLIST!visitdate)
    End If
    
    If Not IsNull(rsLIST!visittime) Then
        MyList.SubItems(9) = CStr(rsLIST!visittime)
    End If
    
    If Not IsNull(rsLIST!clinic) Then
        MyList.SubItems(10) = CStr(rsLIST!clinic)
    End If
  
    If Not IsNull(rsLIST!drugcost) Then
        MyList.SubItems(11) = CStr(rsLIST!drugcost)
    End If
    
    If Not IsNull(rsLIST!procedurecost) Then
        MyList.SubItems(12) = CStr(rsLIST!procedurecost)
    End If
    
    If Not IsNull(rsLIST!TotalCost) Then
        MyList.SubItems(13) = CStr(rsLIST!TotalCost)
    End If
    
    If Not IsNull(rsLIST!claimwith) Then
        MyList.SubItems(14) = CStr(rsLIST!claimwith)
    End If
    
    If Not IsNull(rsLIST!Description) Then
        MyList.SubItems(15) = CStr(rsLIST!Description)
    End If
    
    If Not IsNull(rsLIST!Approved) Then
        MyList.SubItems(16) = CStr(rsLIST!Approved)
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

