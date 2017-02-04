VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParamEmployers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub FindCurrentEmployer()
On Error GoTo Err
    Dim rsFindRecord As ADODB.Recordset
    Set rsFindRecord = New ADODB.Recordset
    
    rsFindRecord.Open "SELECT * FROM ParamEmployer WHERE Empcode='" & CurrentRecord & "';", cnALIS, adOpenKeyset, adLockOptimistic
    
    With rsFindRecord
    If .EOF And .BOF Then Exit Sub
        frmParamEmployers.txtContOfficialTitle = !contofficialtitle & ""
        frmParamEmployers.txtContOtherNames = !contothernames & ""
        frmParamEmployers.txtContSurname = !contsurname & ""
        frmParamEmployers.txtEmployerCode = !empcode & ""
        frmParamEmployers.txtEmpName.Text = !empname & ""
        frmParamEmployers.txtPhyAddress.Text = !PhysicalAddress & ""
        frmParamEmployers.txtPostalAddress.Text = !PostalAddress & ""
        frmParamEmployers.cboCountry = !country & ""
        frmParamEmployers.cboTownCity = !towncity & ""
        frmParamEmployers.txtTelephoneNo = !TelephoneNo & ""
        frmParamEmployers.txtMobileNo = !mobileno & ""
        frmParamEmployers.txtEmail = !email & ""
        frmParamEmployers.txtFaxTelex = !faxtelex & ""
    End With
    Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub LoadEmployer()
With rsFindRecord
If .EOF And .BOF Then Exit Sub
    frmParamEmployers.txtContOfficialTitle = !contofficialtitle & ""
    frmParamEmployers.txtContOtherNames = !contothernames & ""
    frmParamEmployers.txtContSurname = !contsurname & ""
    frmParamEmployers.txtEmployerCode = !empcode & ""
    frmParamEmployers.txtEmpName.Text = !empname & ""
    frmParamEmployers.txtPhyAddress.Text = !PhysicalAddress & ""
    frmParamEmployers.txtPostalAddress.Text = !PostalAddress & ""
    frmParamEmployers.cboCountry = !country & ""
    frmParamEmployers.cboTownCity = !towncity & ""
    frmParamEmployers.txtTelephoneNo = !TelephoneNo & ""
    frmParamEmployers.txtMobileNo = !mobileno & ""
    frmParamEmployers.txtEmail = !email & ""
    frmParamEmployers.txtFaxTelex = !faxtelex & ""
End With
End Sub

Public Sub GetEmpCountry()
On Error GoTo Err
    If Not NewRecord And Not EditRecord Then Exit Sub
    
    With frmParamEmployers
    Dim MySQL$
    Dim TypeStr$
    Dim Rst As ADODB.Recordset
    If .cboCountry.ListCount <> 0 Then .cboCountry.Refresh: Exit Sub
    .cboCountry.Clear
    MySQL = "SELECT * FROM PAramCountries;"
    Set Rst = cnALIS.Execute(MySQL)
    Do While Not Rst.EOF
        .cboCountry.AddItem Rst!country
        
        Rst.MoveNext
    Loop
    Set Rst = Nothing
    MySQL = Empty
    
    End With
    Exit Sub
Err:
        ErrorMessage
End Sub

Public Sub ClearMyScreen()
With frmParamEmployers
    .txtContOfficialTitle = Empty
    .txtContOtherNames = Empty
    .txtContSurname = Empty
    .txtEmployerCode = Empty
    .txtEmpName.Text = Empty
    .txtPhyAddress.Text = Empty
    .txtPostalAddress.Text = Empty
    .cboCountry = Empty
    .cboTownCity = Empty
    .txtTelephoneNo = Empty
    .txtMobileNo = Empty
    .txtEmail = Empty
    .txtFaxTelex = Empty
End With
End Sub

Private Function GetEmployerCode() As String
On Error GoTo Err
Dim rsLastID As ADODB.Recordset 'used to retrieve current LastId in the Table
Dim strLastID As String 'SQL statement

Dim strTemp As String 'store current record
Dim iNumPos As Integer 'store position of the first numeral
Dim strPrefix As String 'stores Id Prefix

'Retrieve the last record in the recrdset where order is ascending

'strLastID = "SELECT max(TransTravelDetails.TravelID) as lastid from TransTravelDetails"
strLastID = "SELECT ParamEmployer.EmpCode  from ParamEmployer ORDER BY EmpCode;"
Set rsLastID = New ADODB.Recordset

With rsLastID
'open the recordset
    .Open strLastID, cnALIS, adOpenKeyset, adLockOptimistic
    If .RecordCount < 1 Then 'shows empty recordset
        GetEmployerCode = "MP000001" 'format of desired format of the string value
    Else
       ' If .EOF And .BOF Then .MoveFirst
        .MoveLast
        strTemp = !empcode
        iNumPos = 1
        Dim sChar As String
        Dim iIDLen As Integer
        iIDLen = Len(strTemp)
        sChar = Mid(strTemp, iNumPos, 1)
        While InStr("1234567890", sChar) = 0
            iNumPos = iNumPos + 1
            sChar = Mid(strTemp, iNumPos, 1)
        Wend
        'store the ID prefix eg AP
        strPrefix = Left(strTemp, iNumPos - 1)
        'store the number portion eg and the length with leading Zeros
        strTemp = Right(strTemp, Len(strTemp) + 1 - iNumPos)
        strTemp = Format(Int(strTemp) + 1, String(iIDLen + 1 - iNumPos, "0"))
        GetEmployerCode = strPrefix & strTemp
    End If
End With
Exit Function
Err:
ErrorMessage
End Function

Public Sub GetEmpTownCity()
On Error GoTo Err
    If Not NewRecord And Not EditRecord Then Exit Sub
    With frmParamEmployers
    
    Dim MySQL$
    Dim TypeStr$
    Dim Rst As ADODB.Recordset
    If .cboTownCity.ListCount <> 0 Then .cboTownCity.Refresh: Exit Sub
    .cboTownCity.Clear
    MySQL = "SELECT * FROM PAramCities;"
    Set Rst = cnALIS.Execute(MySQL)
    Do While Not Rst.EOF
        .cboTownCity.AddItem Rst!cityname
        
        Rst.MoveNext
    Loop
    Set Rst = Nothing
    MySQL = Empty
    
    End With
    Exit Sub
Err:
    ErrorMessage
End Sub

Private Function ValidRecord() As Boolean
On Error GoTo Err
Dim strMessage As String
    With frmParamEmployers
        If .txtEmployerCode.Text = Empty Then
            strMessage = "Required Employer Code!"
            .txtEmployerCode.SetFocus
        ElseIf .txtEmpName.Text = Empty Then
            strMessage = "Required Name of Employer!"
            .txtEmpName.SetFocus
        ElseIf .txtPostalAddress.Text = Empty Then
            strMessage = "Required Postal Address of Employer!"
            .txtPostalAddress.SetFocus
        ElseIf .cboTownCity.Text = Empty Then
            strMessage = "Required City of Employer!"
            .cboTownCity.SetFocus
        ElseIf .cboCountry.Text = Empty Then
            strMessage = "Required Country of Employer!"
            .cboCountry.SetFocus
        Else
            ValidRecord = True
        End If
        If Not ValidRecord Then
            MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
        End If
    End With
Exit Function
Err:
ErrorMessage
End Function

Public Sub CheckEditRecord()
On Error GoTo Err
If NewRecord Then Exit Sub

With frmParamEmployers
    If .txtEmployerCode.Text = Empty Then
        MsgBox "There is No Current Record to Edit!", vbCritical + vbOKOnly, "Cannot Edit"
    Else
        EditRecord = True
        .cmdEditRecord.Caption = "SAVE &CHANGES"
        .txtEmployerCode.SetFocus
        .txtEmployerCode.Locked = True
    End If
End With

Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub EditMyRecord()
On Error GoTo Err
If EditRecord Then
    If ValidRecord Then

    With frmParamEmployers
    Set rsEditRecord = New ADODB.Recordset
    rsEditRecord.Open "SELECT * FROM ParamEmployer WHERE EmpCode='" & .txtEmployerCode.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
        If rsEditRecord.EOF And rsEditRecord.BOF Then GoTo OUTS
            
        rsEditRecord!contofficialtitle = .txtContOfficialTitle
        rsEditRecord!contothernames = .txtContOtherNames
        rsEditRecord!contsurname = .txtContSurname
        rsEditRecord!contallnames = Trim(.txtContSurname.Text) & " " & Trim(.txtContOtherNames.Text)
        rsEditRecord!empcode = .txtEmployerCode
        rsEditRecord!empname = .txtEmpName.Text
        rsEditRecord!PhysicalAddress = .txtPhyAddress.Text
        rsEditRecord!PostalAddress = .txtPostalAddress.Text
        rsEditRecord!country = .cboCountry
        rsEditRecord!towncity = .cboTownCity
        rsEditRecord!TelephoneNo = .txtTelephoneNo
        rsEditRecord!mobileno = .txtMobileNo
        rsEditRecord!email = .txtEmail
        rsEditRecord!faxtelex = .txtFaxTelex
        
        rsEditRecord.Update
        rsEditRecord.Requery
        
        .cmdEditRecord.Caption = "E&dit"
        .txtEmployerCode.Locked = False
        EditRecord = False
        
    End With
    
    End If
End If
OUTS:
Set rsEditRecord = Nothing
Exit Sub
Err:
    UpdateErrorMessage
End Sub

Public Sub AddNewRecord()
If EditRecord Then Exit Sub
NewRecord = True
With frmParamEmployers
    .txtEmployerCode.SetFocus
    .cmdAddNew.Caption = "SAVE &RECORD"
    .txtEmployerCode.Locked = False
    .cmdAddNew.Default = True
    If .chkAuto.Value = 1 Then .txtEmployerCode.Text = GetEmployerCode
End With
End Sub

Public Sub SaveNewRecord()
On Error GoTo Err
If Not NewRecord Then Exit Sub
If ValidRecord Then
    Set rsNewRecord = New ADODB.Recordset
    rsNewRecord.Open "SELECT * FROM ParamEmployer ORDER BY EmpCode;", cnALIS, adOpenKeyset, adLockOptimistic
    
    With frmParamEmployers
    rsNewRecord.AddNew
    
        rsNewRecord!contofficialtitle = .txtContOfficialTitle
        rsNewRecord!contothernames = .txtContOtherNames
        rsNewRecord!contsurname = .txtContSurname
        rsNewRecord!contallnames = Trim(.txtContSurname.Text) & " " & Trim(.txtContOtherNames.Text)
        rsNewRecord!empcode = .txtEmployerCode
        rsNewRecord!empname = .txtEmpName.Text
        rsNewRecord!PhysicalAddress = .txtPhyAddress.Text
        rsNewRecord!PostalAddress = .txtPostalAddress.Text
        rsNewRecord!country = .cboCountry
        rsNewRecord!towncity = .cboTownCity
        rsNewRecord!TelephoneNo = .txtTelephoneNo
        rsNewRecord!mobileno = .txtMobileNo
        rsNewRecord!email = .txtEmail
        rsNewRecord!faxtelex = .txtFaxTelex
        
    rsNewRecord.Update
    rsNewRecord.Requery
    
    .cmdAddNew.Caption = "&New"
    NewRecord = False
    
    Set rsNewRecord = Nothing
    
    End With
End If
Exit Sub
Err:
    UpdateErrorMessage
End Sub

Public Sub RefreshScreen()
If MsgBox("This action terminates all ongoing processes! Are you sure you want to continue? All un-saved date will be lost!", vbQuestion + vbYesNo + vbDefaultButton2, "Screen Refresher") = vbNo Then Exit Sub
    NewRecord = False
    EditRecord = False
    
    Call ClearMyScreen
    
    With frmParamEmployers
        .cmdEditRecord.Caption = "E&dit"
        .cmdAddNew.Caption = "&New"
    End With
End Sub


