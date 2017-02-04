VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSearchForRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public rsSEARCH As ADODB.Recordset, MyCurrentRecord As String

Public Sub SearchEmployeeByID()
On Error GoTo err
    Dim QRY As String
    QRY = InputBox("Enter the Staff ID Number to Search for...", "Search Item")
    Set rsSEARCH = New ADODB.Recordset
    If Len(QRY) = 0 Then Exit Sub
    rsSEARCH.Open "SELECT * FROM ParamEmpMaster WHERE staffidno LIKE '" & Trim(QRY) & "%';", cnCOMMON, adOpenKeyset, adLockOptimistic
    With rsSEARCH
    If .EOF And .BOF Then
        MsgBox "No Matching Record(s) Found! Check Your Search Statement for Spelling!", vbInformation + vbOKOnly, "Search"
        GoTo OUTS
    Else
        GoTo DisplayResults
    End If
    End With
    
DisplayResults:
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
        
        Dim MyList As ListItem
           
        While Not rsSEARCH.EOF
        Set MyList = .ListView1.ListItems.Add(, , CStr(rsSEARCH!StaffIdNo))
        
            If Not IsNull(rsSEARCH!allnames) Then
                MyList.SubItems(1) = CStr(rsSEARCH!allnames)
            End If
            
            If Not IsNull(rsSEARCH!DateofBirth) Then
                MyList.SubItems(2) = CStr(rsSEARCH!DateofBirth)
            End If
            
            If Not IsNull(rsSEARCH!maritalstatus) Then
                MyList.SubItems(3) = CStr(rsSEARCH!maritalstatus)
            End If
            
            If Not IsNull(rsSEARCH!gender) Then
                MyList.SubItems(4) = CStr(rsSEARCH!gender)
            End If
          
            If Not IsNull(rsSEARCH!nationality) Then
                MyList.SubItems(5) = CStr(rsSEARCH!nationality)
            End If
            
            If Not IsNull(rsSEARCH!natidno) Then
                MyList.SubItems(6) = CStr(rsSEARCH!natidno)
            End If
            
            If Not IsNull(rsSEARCH!passportno) Then
                MyList.SubItems(7) = CStr(rsSEARCH!passportno)
            End If
            
            If Not IsNull(rsSEARCH!pinnumber) Then
                MyList.SubItems(8) = CStr(rsSEARCH!pinnumber)
            End If
            
            If Not IsNull(rsSEARCH!datehired) Then
                MyList.SubItems(9) = CStr(rsSEARCH!datehired)
            End If
          
            If Not IsNull(rsSEARCH!officialtitle) Then
                MyList.SubItems(10) = CStr(rsSEARCH!officialtitle)
            End If
            
            If Not IsNull(rsSEARCH!employtype) Then
                MyList.SubItems(11) = CStr(rsSEARCH!employtype)
            End If
            
            If Not IsNull(rsSEARCH!gradecode) Then
                MyList.SubItems(12) = CStr(rsSEARCH!gradecode)
            End If
            
            If Not IsNull(rsSEARCH!deptcode) Then
                MyList.SubItems(13) = CStr(rsSEARCH!deptcode)
            End If
            
            If Not IsNull(rsSEARCH!PhysicalAddress) Then
                MyList.SubItems(14) = CStr(rsSEARCH!PhysicalAddress)
            End If
            
            If Not IsNull(rsSEARCH!PostalAddress) And Not IsNull(rsSEARCH!conttowncity) Then
                MyList.SubItems(15) = CStr(Trim(rsSEARCH!PostalAddress)) & " " & CStr(Trim(rsSEARCH!conttowncity))
            End If
          
            If Not IsNull(rsSEARCH!contcountry) Then
                MyList.SubItems(16) = CStr(rsSEARCH!contcountry)
            End If
            
            If Not IsNull(rsSEARCH!conttelephone) Then
                MyList.SubItems(17) = CStr(rsSEARCH!conttelephone)
            End If
            
            If Not IsNull(rsSEARCH!contmobile) Then
                MyList.SubItems(18) = CStr(rsSEARCH!contmobile)
            End If
          
            If Not IsNull(rsSEARCH!contemail) Then
                MyList.SubItems(19) = CStr(rsSEARCH!contemail)
            End If
        
            rsSEARCH.MoveNext
        Wend
        Set MyList = Nothing
    End With
OUTS:
Set rsSEARCH = Nothing
    Exit Sub
err:
    If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub SearchEmployeeByCompanyName()
On Error GoTo err
    Dim QRY As String
    QRY = InputBox("Enter the Staff Member's CompanyName to Search for...", "Search Item")
    Set rsSEARCH = New ADODB.Recordset
    If Len(QRY) = 0 Then Exit Sub
    rsSEARCH.Open "SELECT * FROM ParamEmpMaster WHERE CompanyName LIKE '" & Trim(QRY) & "%';", cnCOMMON, adOpenKeyset, adLockOptimistic
    With rsSEARCH
    If .EOF And .BOF Then
        MsgBox "No Matching Record(s) Found! Check Your Search Statement for Spelling!", vbInformation + vbOKOnly, "Search"
        GoTo OUTS
    Else
        GoTo DisplayResults
    End If
    End With
    
DisplayResults:
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
        
        Dim MyList As ListItem
           
        While Not rsSEARCH.EOF
        Set MyList = .ListView1.ListItems.Add(, , CStr(rsSEARCH!StaffIdNo))
        
            If Not IsNull(rsSEARCH!allnames) Then
                MyList.SubItems(1) = CStr(rsSEARCH!allnames)
            End If
            
            If Not IsNull(rsSEARCH!DateofBirth) Then
                MyList.SubItems(2) = CStr(rsSEARCH!DateofBirth)
            End If
            
            If Not IsNull(rsSEARCH!maritalstatus) Then
                MyList.SubItems(3) = CStr(rsSEARCH!maritalstatus)
            End If
            
            If Not IsNull(rsSEARCH!gender) Then
                MyList.SubItems(4) = CStr(rsSEARCH!gender)
            End If
          
            If Not IsNull(rsSEARCH!nationality) Then
                MyList.SubItems(5) = CStr(rsSEARCH!nationality)
            End If
            
            If Not IsNull(rsSEARCH!natidno) Then
                MyList.SubItems(6) = CStr(rsSEARCH!natidno)
            End If
            
            If Not IsNull(rsSEARCH!passportno) Then
                MyList.SubItems(7) = CStr(rsSEARCH!passportno)
            End If
            
            If Not IsNull(rsSEARCH!pinnumber) Then
                MyList.SubItems(8) = CStr(rsSEARCH!pinnumber)
            End If
            
            If Not IsNull(rsSEARCH!datehired) Then
                MyList.SubItems(9) = CStr(rsSEARCH!datehired)
            End If
          
            If Not IsNull(rsSEARCH!officialtitle) Then
                MyList.SubItems(10) = CStr(rsSEARCH!officialtitle)
            End If
            
            If Not IsNull(rsSEARCH!employtype) Then
                MyList.SubItems(11) = CStr(rsSEARCH!employtype)
            End If
            
            If Not IsNull(rsSEARCH!gradecode) Then
                MyList.SubItems(12) = CStr(rsSEARCH!gradecode)
            End If
            
            If Not IsNull(rsSEARCH!deptcode) Then
                MyList.SubItems(13) = CStr(rsSEARCH!deptcode)
            End If
            
            If Not IsNull(rsSEARCH!PhysicalAddress) Then
                MyList.SubItems(14) = CStr(rsSEARCH!PhysicalAddress)
            End If
            
            If Not IsNull(rsSEARCH!PostalAddress) And Not IsNull(rsSEARCH!conttowncity) Then
                MyList.SubItems(15) = CStr(Trim(rsSEARCH!PostalAddress)) & " " & CStr(Trim(rsSEARCH!conttowncity))
            End If
          
            If Not IsNull(rsSEARCH!contcountry) Then
                MyList.SubItems(16) = CStr(rsSEARCH!contcountry)
            End If
            
            If Not IsNull(rsSEARCH!conttelephone) Then
                MyList.SubItems(17) = CStr(rsSEARCH!conttelephone)
            End If
            
            If Not IsNull(rsSEARCH!contmobile) Then
                MyList.SubItems(18) = CStr(rsSEARCH!contmobile)
            End If
          
            If Not IsNull(rsSEARCH!contemail) Then
                MyList.SubItems(19) = CStr(rsSEARCH!contemail)
            End If
        
            rsSEARCH.MoveNext
        Wend
        Set MyList = Nothing
    End With
OUTS:
Set rsSEARCH = Nothing
    Exit Sub
err:
    If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub SearchCompanyByName()
On Error GoTo err
    Dim QRY As String
    QRY = InputBox("Enter the Company's Name to Search for...", "Search Item")
    Set rsSEARCH = New ADODB.Recordset
    If Len(QRY) = 0 Then Exit Sub
    rsSEARCH.Open "SELECT * FROM ParamCompanyMaster WHERE companyname LIKE '" & Trim(QRY) & "%';", cnCOMMON, adOpenKeyset, adLockOptimistic
    With rsSEARCH
    If .EOF And .BOF Then
        MsgBox "No Matching Record(s) Found! Check Your Search Statement for Spelling!", vbInformation + vbOKOnly, "Search"
        GoTo OUTS
    Else
        GoTo DisplayResults
    End If
    End With
    
DisplayResults:
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
    
    Dim MyList As ListItem
       
    While Not rsSEARCH.EOF
    Set MyList = .ListView1.ListItems.Add(, , CStr(rsSEARCH!CompanyCode))
    
        If Not IsNull(rsSEARCH!CompanyName) Then
            MyList.SubItems(1) = CStr(rsSEARCH!CompanyName)
        End If
        
        If Not IsNull(rsSEARCH!PhysicalAddress) Then
            MyList.SubItems(2) = CStr(rsSEARCH!PhysicalAddress)
        End If
        
        If Not IsNull(rsSEARCH!PostalAddress) And Not IsNull(rsSEARCH!city) Then
            MyList.SubItems(3) = CStr(Trim(rsSEARCH!PostalAddress)) & " " & CStr(Trim(rsSEARCH!city))
        End If
        
        If Not IsNull(rsSEARCH!country) Then
            MyList.SubItems(4) = CStr(rsSEARCH!country)
        End If
        
        If Not IsNull(rsSEARCH!TelephoneNo) Then
            MyList.SubItems(5) = CStr(rsSEARCH!TelephoneNo)
        End If
        
        If Not IsNull(rsSEARCH!telexfax) Then
            MyList.SubItems(6) = CStr(Trim(rsSEARCH!telexfax))
        End If
        
        If Not IsNull(rsSEARCH!email) Then
            MyList.SubItems(7) = CStr(rsSEARCH!email)
        End If
        
        If Not IsNull(rsSEARCH!coyitno) Then
            MyList.SubItems(8) = CStr(rsSEARCH!coyitno)
        End If
        
        If Not IsNull(rsSEARCH!coypinno) Then
            MyList.SubItems(9) = CStr(Trim(rsSEARCH!coypinno))
        End If
        
        If Not IsNull(rsSEARCH!coynhifno) Then
            MyList.SubItems(10) = CStr(Trim(rsSEARCH!coynhifno))
        End If
    
        rsSEARCH.MoveNext
        
    Wend
    Set MyList = Nothing
End With

OUTS:
Set rsSEARCH = Nothing
    Exit Sub
err:
    If err.Number = 3265 Then Resume Next
        ErrorMessage
End Sub

Public Sub SearchCompanyByCode()
On Error GoTo err
    Dim QRY As String
    QRY = InputBox("Enter the Company's Code to Search for...", "Search Item")
    Set rsSEARCH = New ADODB.Recordset
    If Len(QRY) = 0 Then Exit Sub
    rsSEARCH.Open "SELECT * FROM ParamCompanyMaster WHERE companycode LIKE '" & Trim(QRY) & "%';", cnCOMMON, adOpenKeyset, adLockOptimistic
    With rsSEARCH
    If .EOF And .BOF Then
        MsgBox "No Matching Record(s) Found! Check Your Search Statement for Spelling!", vbInformation + vbOKOnly, "Search"
        GoTo OUTS
    Else
        GoTo DisplayResults
    End If
    End With
    
DisplayResults:
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
    
    Dim MyList As ListItem
       
    While Not rsSEARCH.EOF
    Set MyList = .ListView1.ListItems.Add(, , CStr(rsSEARCH!CompanyCode))
    
        If Not IsNull(rsSEARCH!CompanyName) Then
            MyList.SubItems(1) = CStr(rsSEARCH!CompanyName)
        End If
        
        If Not IsNull(rsSEARCH!PhysicalAddress) Then
            MyList.SubItems(2) = CStr(rsSEARCH!PhysicalAddress)
        End If
        
        If Not IsNull(rsSEARCH!PostalAddress) And Not IsNull(rsSEARCH!city) Then
            MyList.SubItems(3) = CStr(Trim(rsSEARCH!PostalAddress)) & " " & CStr(Trim(rsSEARCH!city))
        End If
        
        If Not IsNull(rsSEARCH!country) Then
            MyList.SubItems(4) = CStr(rsSEARCH!country)
        End If
        
        If Not IsNull(rsSEARCH!TelephoneNo) Then
            MyList.SubItems(5) = CStr(rsSEARCH!TelephoneNo)
        End If
        
        If Not IsNull(rsSEARCH!telexfax) Then
            MyList.SubItems(6) = CStr(Trim(rsSEARCH!telexfax))
        End If
        
        If Not IsNull(rsSEARCH!email) Then
            MyList.SubItems(7) = CStr(rsSEARCH!email)
        End If
        
        If Not IsNull(rsSEARCH!coyitno) Then
            MyList.SubItems(8) = CStr(rsSEARCH!coyitno)
        End If
        
        If Not IsNull(rsSEARCH!coypinno) Then
            MyList.SubItems(9) = CStr(Trim(rsSEARCH!coypinno))
        End If
        
        If Not IsNull(rsSEARCH!coynhifno) Then
            MyList.SubItems(10) = CStr(Trim(rsSEARCH!coynhifno))
        End If
    
        rsSEARCH.MoveNext
        
    Wend
    Set MyList = Nothing
End With

OUTS:
Set rsSEARCH = Nothing
    Exit Sub
err:
    If err.Number = 3265 Then Resume Next
        ErrorMessage
End Sub

Public Sub SearchDepartmentByCode()
On Error GoTo err
    Dim QRY As String
    QRY = InputBox("Enter the Department's Code to Search for...", "Search Item")
    Set rsSEARCH = New ADODB.Recordset
    If Len(QRY) = 0 Then Exit Sub
    rsSEARCH.Open "SELECT * FROM ParamCompanyDepts WHERE deptcode LIKE '" & Trim(QRY) & "%';", cnCOMMON, adOpenKeyset, adLockOptimistic
    With rsSEARCH
    If .EOF And .BOF Then
        MsgBox "No Matching Record(s) Found! Check Your Search Statement for Spelling!", vbInformation + vbOKOnly, "Search"
        GoTo OUTS
    Else
        GoTo DisplayResults
    End If
    End With
    
DisplayResults:
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
        
        Dim MyList As ListItem
           
        While Not rsSEARCH.EOF
        Set MyList = .ListView1.ListItems.Add(, , CStr(rsSEARCH!deptcode))
        
            If Not IsNull(rsSEARCH!deptname) Then
                MyList.SubItems(1) = CStr(rsSEARCH!deptname)
            End If
            
            If Not IsNull(rsSEARCH!hod) Then
                MyList.SubItems(2) = CStr(rsSEARCH!hod)
            End If
            
            If Not IsNull(rsSEARCH!hodstaffid) Then
                MyList.SubItems(3) = CStr(rsSEARCH!hodstaffid)
            End If
            
            If Not IsNull(rsSEARCH!officialtitle) Then
                MyList.SubItems(4) = CStr(rsSEARCH!officialtitle)
            End If
            
            If Not IsNull(rsSEARCH!deptnotes) Then
                MyList.SubItems(5) = CStr(rsSEARCH!deptnotes)
            End If
            
            rsSEARCH.MoveNext
            
        Wend
        Set MyList = Nothing
End With

OUTS:
Set rsSEARCH = Nothing
    Exit Sub
err:
    If err.Number = 3265 Then Resume Next
        ErrorMessage
End Sub

Public Sub SearchDepartmentByName()
On Error GoTo err
    Dim QRY As String
    QRY = InputBox("Enter the Department's Name to Search for...", "Search Item")
    Set rsSEARCH = New ADODB.Recordset
    If Len(QRY) = 0 Then Exit Sub
    rsSEARCH.Open "SELECT * FROM ParamCompanyDepts WHERE deptname LIKE '" & Trim(QRY) & "%';", cnCOMMON, adOpenKeyset, adLockOptimistic
    With rsSEARCH
    If .EOF And .BOF Then
        MsgBox "No Matching Record(s) Found! Check Your Search Statement for Spelling!", vbInformation + vbOKOnly, "Search"
        GoTo OUTS
    Else
        GoTo DisplayResults
    End If
    End With
    
DisplayResults:
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
        
        Dim MyList As ListItem
           
        While Not rsSEARCH.EOF
        Set MyList = .ListView1.ListItems.Add(, , CStr(rsSEARCH!deptcode))
        
            If Not IsNull(rsSEARCH!deptname) Then
                MyList.SubItems(1) = CStr(rsSEARCH!deptname)
            End If
            
            If Not IsNull(rsSEARCH!hod) Then
                MyList.SubItems(2) = CStr(rsSEARCH!hod)
            End If
            
            If Not IsNull(rsSEARCH!hodstaffid) Then
                MyList.SubItems(3) = CStr(rsSEARCH!hodstaffid)
            End If
            
            If Not IsNull(rsSEARCH!officialtitle) Then
                MyList.SubItems(4) = CStr(rsSEARCH!officialtitle)
            End If
            
            If Not IsNull(rsSEARCH!deptnotes) Then
                MyList.SubItems(5) = CStr(rsSEARCH!deptnotes)
            End If
            
            rsSEARCH.MoveNext
            
        Wend
        Set MyList = Nothing
End With

OUTS:
Set rsSEARCH = Nothing
    Exit Sub
err:
    If err.Number = 3265 Then Resume Next
        ErrorMessage
End Sub

Public Property Let ThisRecord(ByVal sRecord As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ThisDate = 5
    MyCurrentRecord = sRecord
End Property


Public Property Get ThisRecord() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ThisDate
    ThisRecord = MyCurrentRecord
End Property

Private Sub Class_Terminate()
    Set rsSEARCH = Nothing
End Sub
