VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCountries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub FindCurrentCountry()
On Error GoTo Err
    Dim rsFindRecord As ADODB.Recordset
    Set rsFindRecord = New ADODB.Recordset
    
    rsFindRecord.Open "SELECT * FROM ParamCountries WHERE countrycode='" & CurrentRecord & "';", cnALIS, adOpenKeyset, adLockOptimistic
    
    With rsFindRecord
    If .EOF And .BOF Then Exit Sub
        frmParamCountries.txtCapitalCity = !capitalcity & ""
        frmParamCountries.cboContinent = !continent & ""
        frmParamCountries.txtCountryCode = !CountryCode & ""
        frmParamCountries.txtNameCountry = !country & ""
    End With
    Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub LoadCountry()
With rsFindRecord
If .EOF And .BOF Then Exit Sub
    frmParamCountries.txtCapitalCity = !capitalcity & ""
    frmParamCountries.cboContinent = !continent & ""
    frmParamCountries.txtCountryCode = !CountryCode & ""
    frmParamCountries.txtNameCountry = !country & ""
End With
End Sub

Private Function ValidRecord() As Boolean
On Error GoTo Err
Dim strMessage As String
    With frmParamCountries
        If .txtCountryCode.Text = Empty Then
            strMessage = "Required Country Code!"
            .txtCountryCode.SetFocus
        ElseIf .txtNameCountry.Text = Empty Then
            strMessage = "Required Name of Country!"
            .txtNameCountry.SetFocus
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

Public Sub ClearMyScreen()
With frmParamCountries
    .txtCapitalCity = Empty
    .cboContinent = Empty
    .txtCountryCode = Empty
    .txtNameCountry = Empty
End With
End Sub

Public Sub CheckEditRecord()
On Error GoTo Err
If NewRecord Then Exit Sub

With frmParamCountries
    If .txtCountryCode.Text = Empty Then
        MsgBox "There is No Current Record to Edit!", vbCritical + vbOKOnly, "Cannot Edit"
    Else
        EditRecord = True
        .cmdEditRecord.Caption = "SAVE &CHANGES"
        .txtCountryCode.SetFocus
        .txtCountryCode.Locked = True
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

    With frmParamCountries
    Set rsEditRecord = New ADODB.Recordset
    rsEditRecord.Open "SELECT * FROM ParamCountries WHERE CountryCode='" & .txtCountryCode.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
        If rsEditRecord.EOF And rsEditRecord.BOF Then GoTo OUTS
            
        rsEditRecord!capitalcity = .txtCapitalCity
        rsEditRecord!continent = .cboContinent
        rsEditRecord!CountryCode = .txtCountryCode
        rsEditRecord!country = .txtNameCountry
        
        rsEditRecord.Update
        rsEditRecord.Requery
        
        .cmdEditRecord.Caption = "E&dit"
        .txtCountryCode.Locked = False
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
With frmParamCountries
    .txtCountryCode.SetFocus
    .cmdAddNew.Caption = "SAVE &RECORD"
    .txtCountryCode.Locked = False
    .cmdAddNew.Default = True
End With
End Sub

Public Sub SaveNewRecord()
On Error GoTo Err
If Not NewRecord Then Exit Sub
If ValidRecord Then
    Set rsNewRecord = New ADODB.Recordset
    rsNewRecord.Open "SELECT * FROM ParamCountries ORDER BY CountryCode;", cnALIS, adOpenKeyset, adLockOptimistic
    
    With frmParamCountries
    rsNewRecord.AddNew
    
        rsNewRecord!capitalcity = .txtCapitalCity
        rsNewRecord!continent = .cboContinent
        rsNewRecord!CountryCode = .txtCountryCode
        rsNewRecord!country = .txtNameCountry
    
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
    
    With frmParamCountries
        .cmdEditRecord.Caption = "E&dit"
        .cmdAddNew.Caption = "&New"
    End With
End Sub


