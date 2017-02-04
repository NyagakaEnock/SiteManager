VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCompanyBranch"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub CheckEditRecord()
On Error GoTo Err
If NewRecord Then Exit Sub

With frmCompanyBranch
    If .txtBranchCode.Text = Empty Then
        MsgBox "There is No Current Record to Edit!", vbCritical + vbOKOnly, "Cannot Edit"
    Else
        EditRecord = True
        .cmdEditRecord.Caption = "SAVE &CHANGES"
        .txtBranchCode.SetFocus
        .txtBranchCode.Locked = True
    End If
End With

Exit Sub
Err:
    ErrorMessage
End Sub

Private Function ValidRecord() As Boolean
On Error GoTo Err
With frmCompanyBranch
    If .cboCompanyCode.Text = Empty Then
        strMessage = "Required Company Code!"
        .cboCompanyCode.SetFocus
    ElseIf .txtBranchCode.Text = Empty Then
        strMessage = "Required Branch Code !"
        .txtBranchCode.SetFocus
    ElseIf .txtBranchName.Text = Empty Then
        strMessage = "Required Name of Branch!"
        .txtBranchName.SetFocus
    Else
        ValidRecord = True
    End If
    If Not ValidRecord Then
        MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation", "ALISHELP.HLP"
    End If
End With
Exit Function
Err:
    ErrorMessage
End Function

Public Sub EditMyRecord()
On Error GoTo Err
If EditRecord Then
    If ValidRecord Then

    With frmCompanyBranch
    Set rsEditRecord = New ADODB.Recordset
    rsEditRecord.Open "SELECT * FROM ParamCompanyBranch WHERE BranchCode='" & .txtBranchCode.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
        If rsEditRecord.EOF And rsEditRecord.BOF Then GoTo OUTS
            
        rsEditRecord!CompanyCode = .cboCompanyCode.Text
        rsEditRecord!BranchCode = .txtBranchCode.Text
        rsEditRecord!BranchName = .txtBranchName.Text
        rsEditRecord!ContactName = .txtContactName.Text
        rsEditRecord!postadd = .txtPostAdd.Text
        rsEditRecord!phyadd = .txtPhyAdd.Text
        rsEditRecord!officialtitle = .cboContactTitle
        rsEditRecord!towncity = .cboTownCity.Text
        rsEditRecord!staffidno = .txtStaffIDNo.Text
        rsEditRecord!phoneno = .txtPhoneNo.Text
        rsEditRecord!faxtelex = .txtFaxTelex.Text
        rsEditRecord!email = .txtEmail.Text
            
        rsEditRecord.Update
        rsEditRecord.Requery
        
        .cmdEditRecord.Caption = "E&dit"
        .txtBranchCode.Locked = False
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

Public Sub SearchByKeyCode()
On Error GoTo Err
With frmCompanyBranch
    Set rsFindRecord = New ADODB.Recordset
    
    rsFindRecord.Open "SELECT * FROM ParamCompanyBranch WHERE BranchCode='" & .txtBranchCode.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then GoTo OUTS
        .cboCompanyCode.Text = rsFindRecord!CompanyCode & ""
        .txtBranchCode.Text = rsFindRecord!BranchCode & ""
        .txtBranchName.Text = rsFindRecord!BranchName & ""
        .txtContactName.Text = rsFindRecord!ContactName & ""
        .txtPostAdd.Text = rsFindRecord!postadd & ""
        .txtPhyAdd.Text = rsFindRecord!phyadd & ""
        .cboContactTitle = rsFindRecord!officialtitle & ""
        .cboTownCity.Text = rsFindRecord!towncity & ""
        .txtStaffIDNo.Text = rsFindRecord!staffidno & ""
        .txtPhoneNo.Text = rsFindRecord!phoneno & ""
        .txtFaxTelex.Text = rsFindRecord!faxtelex & ""
        .txtEmail.Text = rsFindRecord!email & ""
        .txtCompanyName.Text = GetCompanyName
       
    .cmdEditRecord.SetFocus
    
OUTS:
    Set rsFindRecord = Nothing
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub AddNewRecord()
If EditRecord Then Exit Sub
NewRecord = True
With frmCompanyBranch
    .txtBranchCode.SetFocus
    .cmdAddNew.Caption = "SAVE &RECORD"
    .txtBranchCode.Locked = False
    .cmdAddNew.Default = True
End With
End Sub

Public Sub SaveNewRecord()
On Error GoTo Err
If Not NewRecord Then Exit Sub
If ValidRecord Then
    Set rsNewRecord = New ADODB.Recordset
    rsNewRecord.Open "SELECT * FROM ParamCompanyBranch ORDER BY BranchCode;", cnALIS, adOpenKeyset, adLockOptimistic
    
    With frmCompanyBranch
    rsNewRecord.AddNew
    
        rsNewRecord!CompanyCode = .cboCompanyCode.Text
        rsNewRecord!BranchCode = .txtBranchCode.Text
        rsNewRecord!BranchName = .txtBranchName.Text
        rsNewRecord!ContactName = .txtContactName.Text
        rsNewRecord!postadd = .txtPostAdd.Text
        rsNewRecord!phyadd = .txtPhyAdd.Text
        rsNewRecord!officialtitle = .cboContactTitle
        rsNewRecord!towncity = .cboTownCity.Text
        rsNewRecord!staffidno = .txtStaffIDNo.Text
        rsNewRecord!phoneno = .txtPhoneNo.Text
        rsNewRecord!faxtelex = .txtFaxTelex.Text
        rsNewRecord!email = .txtEmail.Text
    
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

Private Function GetCompanyName() As String
On Error GoTo Err
With frmCompanyBranch
    Set rsFindRecord = New ADODB.Recordset
    rsFindRecord.Open "SELECT * FROM ParamCompanyBranch WHERE BranchCode='" & Trim(.txtBranchCode.Text) & "';", cnALIS, adOpenKeyset, adLockOptimistic
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetCompanyName = Empty: GoTo OUTS
    Else
        GetCompanyName = Trim(rsFindRecord!CompanyName & "")
    End If
    
OUTS:
    Set rsFindRecord = Nothing
End With
Exit Function
Err:
    ErrorMessage
End Function

Public Sub ClearMyScreen()
With frmCompanyBranch
    .cboCompanyCode.Text = Empty
    .txtBranchCode.Text = Empty
    .txtBranchName.Text = Empty
    .txtContactName.Text = Empty
    .txtPostAdd.Text = Empty
    .txtPhyAdd.Text = Empty
    .cboContactTitle = Empty
    .cboTownCity.Text = Empty
    .txtStaffIDNo.Text = Empty
    .txtCompanyName.Text = Empty
    .txtPhoneNo.Text = Empty
    .txtFaxTelex.Text = Empty
    .txtEmail.Text = Empty
End With
End Sub

Public Sub LoadCompanyBranch()
With frmCompanyBranch
If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub
    .cboCompanyCode.Text = rsFindRecord!CompanyCode & ""
    .txtBranchCode.Text = rsFindRecord!BranchCode & ""
    .txtBranchName.Text = rsFindRecord!BranchName & ""
    .txtContactName.Text = rsFindRecord!ContactName & ""
    .txtPostAdd.Text = rsFindRecord!postadd & ""
    .txtPhyAdd.Text = rsFindRecord!phyadd & ""
    .cboContactTitle = rsFindRecord!officialtitle & ""
    .cboTownCity.Text = rsFindRecord!towncity & ""
    .txtStaffIDNo.Text = rsFindRecord!staffidno & ""
    .txtPhoneNo.Text = rsFindRecord!phoneno & ""
    .txtFaxTelex.Text = rsFindRecord!faxtelex & ""
    .txtEmail.Text = rsFindRecord!email & ""
    .txtCompanyName.Text = GetCompanyName
End With
End Sub

Public Sub RefreshScreen()
If MsgBox("This action terminates all ongoing processes! Are you sure you want to continue? All un-saved date will be lost!", vbQuestion + vbYesNo + vbDefaultButton2, "Screen Refresher") = vbNo Then Exit Sub
    NewRecord = False
    EditRecord = False
    
    Call ClearMyScreen
    
    With frmCompanyBranch
        .cmdEditRecord.Caption = "E&dit"
        .cmdAddNew.Caption = "&New"
    End With
End Sub

Public Sub GetCompanyInfo()
On Error GoTo Err
    Set rsFindRecord = New ADODB.Recordset
    rsFindRecord.Open "SELECT * FROM ParamCompanyMaster;", cnALIS, adOpenKeyset, adLockOptimistic
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then GoTo OUTS
    
    With frmCompanyBranch
        .cboCompanyCode.Text = rsFindRecord!CompanyCode & ""
        .txtCompanyName.Text = rsFindRecord!CompanyName & ""
    End With
    
OUTS:
    Set rsFindRecord = Nothing
    
Exit Sub
Err:
    ErrorMessage
End Sub
