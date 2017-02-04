VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAccPeriod"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub RefreshScreen()
If MsgBox("This action terminates all ongoing processes! Are you sure you want to continue? All un-saved date will be lost!", vbQuestion + vbYesNo + vbDefaultButton2, "Screen Refresher") = vbNo Then Exit Sub
    NewRecord = False
    EditRecord = False
    
    Call ClearMyScreen
    
    With frmParamAccPeriods
        .cmdEditRecord.Caption = "E&dit"
        .cmdAddNew.Caption = "&New"
    End With
End Sub

Public Sub ClearMyScreen()
With frmParamAccPeriods
    .txtBEginDate.Text = Empty
    .txtEndDate.Text = Empty
    .txtPeriodCode.Text = Empty
    .txtNamePeriod.Text = Empty
    .cboHalfYear.Text = Empty
    .cboQuarterYear.Text = Empty
End With
End Sub

Public Sub CheckEditRecord()
On Error GoTo Err
If NewRecord Then Exit Sub

With frmParamAccPeriods
    If .txtPeriodCode.Text = Empty Then
        MsgBox "There is No Current Record to Edit!", vbCritical + vbOKOnly, "Cannot Edit"
    Else
        EditRecord = True
        .cmdEditRecord.Caption = "SAVE &CHANGES"
        .txtPeriodCode.SetFocus
        .txtPeriodCode.Locked = True
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

    With frmParamAccPeriods
    Set rsEditRecord = New ADODB.Recordset
    rsEditRecord.Open "SELECT * FROM ParamAccPeriod WHERE PeriodCode='" & .txtPeriodCode.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
        If rsEditRecord.EOF And rsEditRecord.BOF Then GoTo OUTS
            
        rsEditRecord!StartDate = .txtBEginDate
        rsEditRecord!EndDate = .txtEndDate
        rsEditRecord!PeriodCode = .txtPeriodCode
        rsEditRecord!AccPeriod = .txtNamePeriod
        rsEditRecord!halfofyear = .cboHalfYear
        rsEditRecord!quarterofyear = .cboQuarterYear
            
        rsEditRecord.Update
        rsEditRecord.Requery
        
        .cmdEditRecord.Caption = "E&dit"
        .txtPeriodCode.Locked = False
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
With frmParamAccPeriods
    Set rsFindRecord = New ADODB.Recordset
    
    rsFindRecord.Open "SELECT * FROM ParamAccPeriod WHERE PeriodCode='" & .txtPeriodCode.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then GoTo OUTS
       .txtBEginDate = rsFindRecord!StartDate & ""
       .txtEndDate = rsFindRecord!EndDate & ""
       .txtPeriodCode = rsFindRecord!PeriodCode & ""
       .txtNamePeriod = rsFindRecord!AccPeriod & ""
       .cboHalfYear = rsFindRecord!halfofyear & ""
       .cboQuarterYear = rsFindRecord!quarterofyear & ""
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
With frmParamAccPeriods
    .txtPeriodCode.SetFocus
    .cmdAddNew.Caption = "SAVE &RECORD"
    .txtPeriodCode.Locked = False
    .cmdAddNew.Default = True
End With
End Sub

Public Sub SaveNewRecord()
On Error GoTo Err
If Not NewRecord Then Exit Sub
If ValidRecord Then
    Set rsNewRecord = New ADODB.Recordset
    rsNewRecord.Open "SELECT * FROM ParamAccPeriod ORDER BY PeriodCode;", cnALIS, adOpenKeyset, adLockOptimistic
    
    With frmParamAccPeriods
    rsNewRecord.AddNew
        rsNewRecord!StartDate = .txtBEginDate
        rsNewRecord!EndDate = .txtEndDate
        rsNewRecord!PeriodCode = .txtPeriodCode
        rsNewRecord!AccPeriod = .txtNamePeriod
        rsNewRecord!halfofyear = .cboHalfYear
        rsNewRecord!quarterofyear = .cboQuarterYear
        
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

Private Function ValidRecord() As Boolean
On Error GoTo Err
    With frmParamAccPeriods
    If .txtBEginDate.Text = Empty Then
        strMessage = "Required Beginning Date!"
        .txtBEginDate.SetFocus
    ElseIf .txtEndDate.Text = Empty Then
        strMessage = "Required Ending Date!"
        .txtEndDate.SetFocus
    ElseIf .txtPeriodCode.Text = Empty Then
        strMessage = "Required Period Code!"
        .txtPeriodCode.SetFocus
    ElseIf .txtNamePeriod.Text = Empty Then
        strMessage = "Required Name of Period"
        .txtNamePeriod.SetFocus
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

Private Sub UpdateStaffFile()
On Error GoTo Err

With frmParamAccPeriods
    Set rsLineUpdate = New ADODB.Recordset
    
    rsLineUpdate.Open "SELECT * FROM ParamEmpMaster WHERE PeriodCode='" & Trim(.txtPeriodCode.Text) & "';", cnALIS, adOpenKeyset, adLockOptimistic
    
    If rsLineUpdate.EOF And rsLineUpdate.BOF Then GoTo OUTS
        rsLineUpdate!userreg = "Y"
        rsLineUpdate.Update
        rsLineUpdate.Requery
End With
    
OUTS:
    Set rsLineUpdate = Nothing
    
Exit Sub
Err:
    ErrorMessage
End Sub


