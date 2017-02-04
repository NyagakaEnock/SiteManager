VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpPersonal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Function ValidRecord() As Boolean
'On Error GoTo err
Dim strMessage As String
    With frmEmployeesPersonal
    If .txtStaffIDNo.Text = Empty Then
        strMessage = "Required Staff ID Number!"
        .SSTab1.Tab = 0: .txtStaffIDNo.SetFocus
    ElseIf .txtCompanyName.Text = Empty Then
        strMessage = "Required CompanyName!"
        .SSTab1.Tab = 0: .txtCompanyName.SetFocus
    ElseIf .txtOtherNames.Text = Empty Then
        strMessage = "Required Other Names!"
        .SSTab1.Tab = 0: .txtOtherNames.SetFocus
    ElseIf .txtDateofBirth.Text = Empty Then
        strMessage = "Required Date of Birth!"
        .SSTab1.Tab = 0: .txtDateofBirth.SetFocus
    ElseIf .txtDateHired.Text = Empty Then
        strMessage = "Required Date Hired!"
        .SSTab1.Tab = 0: .txtDateHired.SetFocus
    ElseIf .cboGender.Text = Empty Then
        strMessage = "Required Gender!"
        .SSTab1.Tab = 0: .cboGender.SetFocus
    ElseIf .cboNationality.Text = Empty Then
        strMessage = "Required Nationality!"
        .SSTab1.Tab = 0: .cboNationality.SetFocus
    ElseIf .cboDepartment.Text = Empty Then
        strMessage = "Required Department!"
        .SSTab1.Tab = 0: .cboDepartment.SetFocus
    ElseIf .txtPostalAddress.Text = Empty Then
        strMessage = "Required Postal Address!"
        .SSTab1.Tab = 1: .txtPostalAddress.SetFocus
    ElseIf .cboTownCity.Text = Empty Then
        strMessage = "Required Town/City for Address!"
        .SSTab1.Tab = 1: .cboTownCity.SetFocus
    Else
        ValidRecord = True
    End If
    If Not ValidRecord Then
        MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
    End If
    End With
Exit Function
err:
    ErrorMessage
End Function

Public Sub CheckEditRecord()
'On Error GoTo err
If NewRecord Then Exit Sub

With frmEmployeesPersonal
    If .txtStaffIDNo.Text = Empty Then
        MsgBox "There is No Current Record to Edit!", vbCritical + vbOKOnly, "Cannot Edit"
    Else
        editRECORD = True
        .cmdEditRecord.Caption = "SAVE &CHANGES"
        .txtStaffIDNo.SetFocus
        .txtStaffIDNo.Locked = True
    End If
End With

Exit Sub
err:
    ErrorMessage
End Sub

Public Sub ClearMyScreen()
With frmEmployeesPersonal
    .cboCountry = Empty
    .cboDepartment = Empty
    .cboEmployType = Empty
    .cboGender = Empty
    .cboGrade = Empty
    .cboMaritalStatus = Empty
    .cboNationality = Empty
    .cboOfficialTitle = Empty
    .cboStaffIDNo = Empty
    .cboTitle = Empty
    .cboTownCity = Empty
    .txtDateHired = Empty
    .txtDateofBirth = Empty
    .txtDepartment = Empty
    .txtEmail = Empty
    .txtFullNames = Empty
    .txtMobileNo = Empty
    .txtNatIDNo = Empty
    .txtOtherNames = Empty
    .txtPassportNo = Empty
    .txtPhysicalAddress = Empty
    .txtPINNumber = Empty
    .txtPostalAddress = Empty
    .txtStaffIDNo = Empty
    .txtTelephoneNo = Empty
    .txtCompanyName = Empty
    .imgPHOTO.Picture = LoadPicture("")
End With
End Sub

Public Sub LoadEmployees()
With rsFindRecord
If .EOF And .BOF Then Exit Sub
    frmEmployeesPersonal.cboCountry = !contcountry & ""
    frmEmployeesPersonal.cboDepartment = !deptcode & ""
    frmEmployeesPersonal.cboEmployType = !employtype & ""
    frmEmployeesPersonal.cboGender = !gender & ""
    frmEmployeesPersonal.cboGrade = !gradecode & ""
    frmEmployeesPersonal.cboMaritalStatus = !maritalstatus & ""
    frmEmployeesPersonal.cboNationality = !nationality & ""
    frmEmployeesPersonal.cboOfficialTitle = !officialtitle & ""
    frmEmployeesPersonal.cboStaffIDNo = !StaffIdNo & ""
    frmEmployeesPersonal.cboTitle = !titleofcourtesy & ""
    frmEmployeesPersonal.cboTownCity = !conttowncity & ""
    frmEmployeesPersonal.txtDateHired = !datehired & ""
    frmEmployeesPersonal.txtDateofBirth = !DateofBirth & ""
    frmEmployeesPersonal.txtDepartment = !deptcode & ""
    frmEmployeesPersonal.txtEmail = !contemail & ""
    frmEmployeesPersonal.txtFullNames = !allnames & ""
    frmEmployeesPersonal.txtMobileNo = !contmobile & ""
    frmEmployeesPersonal.txtNatIDNo = !natidno & ""
    frmEmployeesPersonal.txtOtherNames = !OtherNames & ""
    frmEmployeesPersonal.txtPassportNo = !passportno & ""
    frmEmployeesPersonal.txtPhysicalAddress = !PhysicalAddress & ""
    frmEmployeesPersonal.txtPINNumber = !pinnumber & ""
    frmEmployeesPersonal.txtPostalAddress = !PostalAddress & ""
    frmEmployeesPersonal.txtStaffIDNo = !StaffIdNo & ""
    frmEmployeesPersonal.txtTelephoneNo = !conttelephone & ""
    frmEmployeesPersonal.txtCompanyName = !CompanyName & ""
    frmEmployeesPersonal.imgPHOTO.Picture = LoadPicture(!photo & "")
End With
End Sub

Public Sub GetNationality()
'On Error GoTo err
If Not NewRecord And Not editRECORD Then Exit Sub
    If frmEmployeesPersonal.cboNationality.ListCount <> 0 Then frmEmployeesPersonal.cboNationality.Refresh: Exit Sub
    
    Dim rsLIST As ADODB.Recordset
    Set rsLIST = cnCOMMON.Execute("SELECT * FROM ParamCountries ORDER BY countrycode;")
    frmEmployeesPersonal.cboNationality.Clear
    
    With rsLIST
    If .EOF And .BOF Then Exit Sub
       .MoveFirst
       Do While Not .EOF
            frmEmployeesPersonal.cboNationality.AddItem !country
            .MoveNext
       Loop
    End With
    
    Set rsLIST = Nothing
    Exit Sub
err:
    ErrorMessage
End Sub

Public Sub FindFullNames()
'On Error GoTo err

With frmEmployeesPersonal
    Set rsFindRecord = New ADODB.Recordset
    
    rsFindRecord.Open "SELECT * FROM ParamEmpMAster WHERE StaffIDNo='" & Trim(.cboStaffIDNo.Text) & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then GoTo OUTS
        .cboStaffIDNo.Text = Trim(rsFindRecord!StaffIdNo)
        .txtFullNames.Text = Trim(rsFindRecord!CompanyName & "") & " " & Trim(rsFindRecord!OtherNames & "")
        .txtPhysicalAddress.SetFocus
OUTS:
    Set rsFindRecord = Nothing
End With

Exit Sub
err:
    ErrorMessage
End Sub

Public Sub FindNationalityCode()
'On Error GoTo err

With frmEmployeesPersonal
    Set rsFindRecord = New ADODB.Recordset
    
    rsFindRecord.Open "SELECT * FROM ParamCountries WHERE Country='" & Trim(.cboNationality.Text) & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then GoTo OUTS
        .cboNationality.Text = Trim(rsFindRecord!CountryCode)
    
OUTS:
    Set rsFindRecord = Nothing
End With

Exit Sub
err:
    ErrorMessage
End Sub

Public Sub FindCountryCode()
'On Error GoTo err

With frmEmployeesPersonal
    Set rsFindRecord = New ADODB.Recordset
    
    rsFindRecord.Open "SELECT * FROM ParamCountries WHERE Country='" & Trim(.cboCountry.Text) & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then GoTo OUTS
        .cboCountry.Text = Trim(rsFindRecord!CountryCode)
    
OUTS:
    Set rsFindRecord = Nothing
End With

Exit Sub
err:
    ErrorMessage
End Sub


Public Sub FindDesignationCode()
'On Error GoTo err

With frmEmployeesPersonal
    Set rsFindRecord = New ADODB.Recordset
    
    rsFindRecord.Open "SELECT * FROM ParamDesignations WHERE Designation='" & Trim(.cboOfficialTitle.Text) & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then GoTo OUTS
        .cboOfficialTitle.Text = Trim(rsFindRecord!descode)
    
OUTS:
    Set rsFindRecord = Nothing
End With

Exit Sub
err:
    ErrorMessage
End Sub

Public Sub FindEmpTypeCode()
'On Error GoTo err

With frmEmployeesPersonal
    Set rsFindRecord = New ADODB.Recordset
    
    rsFindRecord.Open "SELECT * FROM ParamEmpTypes WHERE EmpType='" & Trim(.cboEmployType.Text) & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then GoTo OUTS
        .cboEmployType.Text = Trim(rsFindRecord!typecode)
    
OUTS:
    Set rsFindRecord = Nothing
End With

Exit Sub
err:
    ErrorMessage
End Sub

Public Sub FindDeptCode()
'On Error GoTo err

With frmEmployeesPersonal
    Set rsFindRecord = New ADODB.Recordset
    
    rsFindRecord.Open "SELECT * FROM ParamCompanyDepts WHERE deptname='" & Trim(.cboDepartment.Text) & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then GoTo OUTS
        .cboDepartment.Text = Trim(rsFindRecord!deptcode)
        .txtDepartment.Text = Trim(rsFindRecord!deptname & "")
OUTS:
    Set rsFindRecord = Nothing
End With

Exit Sub
err:
    ErrorMessage
End Sub

Public Sub GetDesignations()
'On Error GoTo err
If Not NewRecord And Not editRECORD Then Exit Sub
    If frmEmployeesPersonal.cboOfficialTitle.ListCount <> 0 Then frmEmployeesPersonal.cboOfficialTitle.Refresh: Exit Sub
    
    Dim rsLIST As ADODB.Recordset
    Set rsLIST = cnCOMMON.Execute("SELECT * FROM ParamDesignations ORDER BY DesCode;")
    frmEmployeesPersonal.cboOfficialTitle.Clear
    
    With rsLIST
    If .EOF And .BOF Then Exit Sub
       .MoveFirst
       Do While Not .EOF
            If Not IsNull(!designation) Then
                frmEmployeesPersonal.cboOfficialTitle.AddItem !designation
            End If
            .MoveNext
       Loop
    End With
    
    Set rsLIST = Nothing
    
    Exit Sub
err:
    ErrorMessage
End Sub

Public Sub GetJobGroups()
'On Error GoTo err
If Not NewRecord And Not editRECORD Then Exit Sub
    If frmEmployeesPersonal.cboGrade.ListCount <> 0 Then frmEmployeesPersonal.cboGrade.Refresh: Exit Sub
    
    Dim rsLIST As ADODB.Recordset
    Set rsLIST = cnCOMMON.Execute("SELECT * FROM ParamJobGroups ORDER BY JobGroup;")
    frmEmployeesPersonal.cboGrade.Clear
    
    With rsLIST
    If .EOF And .BOF Then Exit Sub
       .MoveFirst
       Do While Not .EOF
            If Not IsNull(!JobGroup) Then
                frmEmployeesPersonal.cboGrade.AddItem !JobGroup
            End If
            .MoveNext
       Loop
    End With
    
    Set rsLIST = Nothing
    
    Exit Sub
err:
    ErrorMessage
End Sub

Public Sub GetDepartments()
'On Error GoTo err
If Not NewRecord And Not editRECORD Then Exit Sub
    If frmEmployeesPersonal.cboDepartment.ListCount <> 0 Then frmEmployeesPersonal.cboDepartment.Refresh: Exit Sub
    
    Dim rsLIST As ADODB.Recordset
    Set rsLIST = cnCOMMON.Execute("SELECT * FROM ParamCompanyDepts ORDER BY DeptCode;")
    frmEmployeesPersonal.cboDepartment.Clear
    
    With rsLIST
    If .EOF And .BOF Then Exit Sub
       .MoveFirst
       Do While Not .EOF
            If Not IsNull(!deptname) Then
                frmEmployeesPersonal.cboDepartment.AddItem !deptname
            End If
            .MoveNext
       Loop
    End With
    
    Set rsLIST = Nothing
    
    Exit Sub
err:
    ErrorMessage
End Sub

Public Sub GetStaffIDNo()
'On Error GoTo err
If Not NewRecord And Not editRECORD Then Exit Sub
    If frmEmployeesPersonal.cboStaffIDNo.ListCount <> 0 Then frmEmployeesPersonal.cboStaffIDNo.Refresh: Exit Sub
    
    Dim rsLIST As ADODB.Recordset
    Set rsLIST = cnCOMMON.Execute("SELECT * FROM ParamEmpMaster ORDER BY StaffIDNo;")
    frmEmployeesPersonal.cboStaffIDNo.Clear
    
    With rsLIST
    If .EOF And .BOF Then Exit Sub
       .MoveFirst
       Do While Not .EOF
            If Not IsNull(!emptype) Then
                frmEmployeesPersonal.cboStaffIDNo.AddItem !StaffIdNo
            End If
            .MoveNext
       Loop
    End With
    
    Set rsLIST = Nothing
    
    Exit Sub
err:
    ErrorMessage
End Sub

Public Sub GetEmploymentTypes()
'On Error GoTo err
If Not NewRecord And Not editRECORD Then Exit Sub
    If frmEmployeesPersonal.cboEmployType.ListCount <> 0 Then frmEmployeesPersonal.cboEmployType.Refresh: Exit Sub
    
    Dim rsLIST As ADODB.Recordset
    Set rsLIST = cnCOMMON.Execute("SELECT * FROM ParamEmpTypes ORDER BY TypeCode;")
    frmEmployeesPersonal.cboEmployType.Clear
    
    With rsLIST
    If .EOF And .BOF Then Exit Sub
       .MoveFirst
       Do While Not .EOF
            If Not IsNull(!emptype) Then
                frmEmployeesPersonal.cboEmployType.AddItem !emptype
            End If
            .MoveNext
       Loop
    End With
    
    Set rsLIST = Nothing
    
    Exit Sub
err:
    ErrorMessage
End Sub

Public Sub GetCountryCode()
'On Error GoTo err
If Not NewRecord And Not editRECORD Then Exit Sub
    If frmEmployeesPersonal.cboCountry.ListCount <> 0 Then frmEmployeesPersonal.cboCountry.Refresh: Exit Sub
    
    Dim rsLIST As ADODB.Recordset
    Set rsLIST = cnCOMMON.Execute("SELECT * FROM ParamCountries ORDER BY countrycode;")
    frmEmployeesPersonal.cboCountry.Clear
    
    With rsLIST
    If .EOF And .BOF Then Exit Sub
       .MoveFirst
       Do While Not .EOF
            frmEmployeesPersonal.cboCountry.AddItem !country
            .MoveNext
       Loop
    End With
    
    Set rsLIST = Nothing
    Exit Sub
err:
    ErrorMessage
End Sub

Public Sub GetTitles()
'On Error GoTo err
If Not NewRecord And Not editRECORD Then Exit Sub
    If frmEmployeesPersonal.cboTitle.ListCount <> 0 Then frmEmployeesPersonal.cboTitle.Refresh: Exit Sub
    
    Dim rsLIST As ADODB.Recordset
    Set rsLIST = cnCOMMON.Execute("SELECT * FROM ParamTitles ORDER BY TitleID;")
    frmEmployeesPersonal.cboTitle.Clear
    
    With rsLIST
    If .EOF And .BOF Then Exit Sub
       .MoveFirst
       Do While Not .EOF
            frmEmployeesPersonal.cboTitle.AddItem !Title
            .MoveNext
       Loop
    End With
    
    Set rsLIST = Nothing
    Exit Sub
err:
    ErrorMessage
End Sub

Public Sub GetOfficialTitle()
'On Error GoTo err
If Not NewRecord And Not editRECORD Then Exit Sub
    If frmEmployeesPersonal.cboTitle.ListCount <> 0 Then frmEmployeesPersonal.cboTitle.Refresh: Exit Sub
    
    Dim rsLIST As ADODB.Recordset
    Set rsLIST = cnCOMMON.Execute("SELECT * FROM ParamTitles ORDER BY TitleID;")
    frmEmployeesPersonal.cboTitle.Clear
    
    With rsLIST
    If .EOF And .BOF Then Exit Sub
       .MoveFirst
       Do While Not .EOF
            frmEmployeesPersonal.cboTitle.AddItem !Title
            .MoveNext
       Loop
    End With
    
    Set rsLIST = Nothing
    Exit Sub
err:
    ErrorMessage
End Sub

Public Sub GetMainCity()
'On Error GoTo err
If Not NewRecord And Not editRECORD Then Exit Sub
    If frmEmployeesPersonal.cboTownCity.ListCount <> 0 Then frmEmployeesPersonal.cboTownCity.Refresh: Exit Sub
    
    Dim rsLIST As ADODB.Recordset
    Set rsLIST = cnCOMMON.Execute("SELECT * FROM ParamCities ORDER BY cityname;")
    frmEmployeesPersonal.cboTownCity.Clear
    
    With rsLIST
    If .EOF And .BOF Then Exit Sub
       .MoveFirst
       Do While Not .EOF
            frmEmployeesPersonal.cboTownCity.AddItem !cityname
            .MoveNext
       Loop
    End With
    
    Set rsLIST = Nothing
    Exit Sub
err:
    ErrorMessage
End Sub

Public Sub AddNewRecord()
If editRECORD Then Exit Sub
NewRecord = True
With frmEmployeesPersonal
    .txtStaffIDNo.SetFocus
    .cmdAddNew.Caption = "SAVE &RECORD"
    .txtStaffIDNo.Locked = False
    .cmdAddNew.Default = True
End With
End Sub

Public Sub SaveNewRecord()
'On Error GoTo err
If Not NewRecord Then Exit Sub
If ValidRecord Then
    Set rsNewRecord = New ADODB.Recordset
    rsNewRecord.Open "SELECT * FROM ParamEmpMaster ORDER BY staffidno;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With frmEmployeesPersonal
    rsNewRecord.AddNew
    
        rsNewRecord!contcountry = .cboCountry
        rsNewRecord!deptcode = .cboDepartment
        rsNewRecord!employtype = .cboEmployType
        rsNewRecord!gender = .cboGender
        rsNewRecord!gradecode = .cboGrade
        rsNewRecord!maritalstatus = .cboMaritalStatus
        rsNewRecord!nationality = .cboNationality
        rsNewRecord!officialtitle = .cboOfficialTitle
        rsNewRecord!StaffIdNo = .cboStaffIDNo
        rsNewRecord!titleofcourtesy = .cboTitle
        rsNewRecord!conttowncity = .cboTownCity
        rsNewRecord!datehired = .txtDateHired
        rsNewRecord!DateofBirth = .txtDateofBirth
        rsNewRecord!deptcode = .txtDepartment
        rsNewRecord!contemail = .txtEmail
        rsNewRecord!allnames = .txtFullNames
        rsNewRecord!contmobile = .txtMobileNo
        rsNewRecord!natidno = .txtNatIDNo
        rsNewRecord!OtherNames = .txtOtherNames
        rsNewRecord!passportno = .txtPassportNo
        rsNewRecord!PhysicalAddress = .txtPhysicalAddress
        rsNewRecord!pinnumber = .txtPINNumber
        rsNewRecord!PostalAddress = .txtPostalAddress
        rsNewRecord!StaffIdNo = .txtStaffIDNo
        rsNewRecord!conttelephone = .txtTelephoneNo
        rsNewRecord!CompanyName = .txtCompanyName
        rsNewRecord!photo = .dlgPHOTO.FileName
        rsNewRecord!createdby = CurrentUserName
        rsNewRecord!datecreated = Now
        rsNewRecord!AccPeriod = CurrentPeriod
    
    rsNewRecord.Update
    rsNewRecord.Requery
    
    .cmdAddNew.Caption = "&New"
    NewRecord = False
    
    Set rsNewRecord = Nothing
    
    End With
End If
Exit Sub
err:
    UpdateErrorMessage
End Sub

Public Sub EditMyRecord()
'On Error GoTo err
If editRECORD Then
    If ValidRecord Then

    With frmEmployeesPersonal
    Set rsEditRecord = New ADODB.Recordset
    rsEditRecord.Open "SELECT * FROM ParamEmpMaster WHERE staffidno='" & .txtStaffIDNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsEditRecord.EOF And rsEditRecord.BOF Then GoTo OUTS
        CurrentPic = Trim(!photo & "")
   
        rsEditRecord!contcountry = .cboCountry
        rsEditRecord!deptcode = .cboDepartment
        rsEditRecord!employtype = .cboEmployType
        rsEditRecord!gender = .cboGender
        rsEditRecord!gradecode = .cboGrade
        rsEditRecord!maritalstatus = .cboMaritalStatus
        rsEditRecord!nationality = .cboNationality
        rsEditRecord!officialtitle = .cboOfficialTitle
        rsEditRecord!StaffIdNo = .cboStaffIDNo
        rsEditRecord!titleofcourtesy = .cboTitle
        rsEditRecord!conttowncity = .cboTownCity
        rsEditRecord!datehired = .txtDateHired
        rsEditRecord!DateofBirth = .txtDateofBirth
        rsEditRecord!deptcode = .txtDepartment
        rsEditRecord!contemail = .txtEmail
        rsEditRecord!allnames = .txtFullNames
        rsEditRecord!contmobile = .txtMobileNo
        rsEditRecord!natidno = .txtNatIDNo
        rsEditRecord!OtherNames = .txtOtherNames
        rsEditRecord!passportno = .txtPassportNo
        rsEditRecord!PhysicalAddress = .txtPhysicalAddress
        rsEditRecord!pinnumber = .txtPINNumber
        rsEditRecord!PostalAddress = .txtPostalAddress
        rsEditRecord!StaffIdNo = .txtStaffIDNo
        rsEditRecord!conttelephone = .txtTelephoneNo
        rsEditRecord!CompanyName = .txtCompanyName
        rsEditRecord!photo = .dlgPHOTO.FileName
        rsEditRecord!createdby = CurrentUserName
        rsEditRecord!datecreated = Now
        rsEditRecord!AccPeriod = TransactionPeriod
        
        If .dlgPHOTO.FileName = Empty Then
            rsEditRecord!photo = CurrentPic
        Else
            rsEditRecord!photo = .dlgPHOTO.FileName
        End If
        
        rsEditRecord.Update
        rsEditRecord.Requery
        
        .cmdEditRecord.Caption = "E&dit"
        .txtStaffIDNo.Locked = False
        editRECORD = False
        
    End With
    
    End If
End If
OUTS:
Set rsEditRecord = Nothing
Exit Sub
err:
    UpdateErrorMessage
End Sub

Public Sub RefreshScreen()
If MsgBox("This action terminates all ongoing processes! Are you sure you want to continue? All un-saved date will be lost!", vbQuestion + vbYesNo + vbDefaultButton2, "Screen Refresher") = vbNo Then Exit Sub
    NewRecord = False
    editRECORD = False
    
    Call ClearMyScreen
    
    With frmEmployeesPersonal
        .cmdEditRecord.Caption = "E&dit"
        .cmdAddNew.Caption = "&New"
    End With
End Sub

Public Sub ShowOpenDialog()
If Not NewRecord And Not editRECORD Then Exit Sub
With frmEmployeesPersonal
    .dlgPHOTO.ShowOpen
    If .dlgPHOTO.FileName = "" Then Exit Sub
    .imgPHOTO.Picture = LoadPicture(.dlgPHOTO.FileName)
End With
End Sub
