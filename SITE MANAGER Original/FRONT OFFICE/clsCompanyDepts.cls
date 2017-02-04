VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCompanyDepts"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub FindCurrentDepts()
On Error GoTo err
    Dim rsFindRecord As ADODB.Recordset
    Set rsFindRecord = New ADODB.Recordset
    
    rsFindRecord.Open "SELECT * FROM ParamCompanyDepts WHERE DeptCode='" & CurrentRecord & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsFindRecord
    If .EOF And .BOF Then Exit Sub
        frmCompanyDepartments.txtDeptCode = !deptcode & ""
        frmCompanyDepartments.txtDeptName = !deptname & ""
        frmCompanyDepartments.cboDeptHead = !hod & ""
        frmCompanyDepartments.txtDeptNotes = !deptnotes & ""
        frmCompanyDepartments.txtOfficialTitle = !officialtitle & ""
        frmCompanyDepartments.txtStaffID = !hodstaffid & ""
    End With
    Exit Sub
err:
    ErrorMessage
End Sub

Private Function ValidRecord() As Boolean
On Error GoTo err
Dim strMessage As String
    With frmCompanyDepartments
        If .txtDeptCode = "" Then
            strMessage = "Required Department Code!"
            .txtDeptCode.SetFocus
        ElseIf .txtDeptName = "" Then
            strMessage = "Required name of Department!"
            .txtDeptName.SetFocus
        ElseIf .cboDeptHead = "" Then
            strMessage = "Required names of the Head of Department!"
            .cboDeptHead.SetFocus
        ElseIf .txtStaffID = Empty Then
            strMessage = "Required HOD Staff ID NO!"
            .cboDeptHead.SetFocus
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

Public Sub RefreshScreen()
If MsgBox("This action terminates all ongoing processes! Are you sure you want to continue? All un-saved date will be lost!", vbQuestion + vbYesNo + vbDefaultButton2, "Screen Refresher") = vbNo Then Exit Sub
    NewRecord = False
    beditRECORD = False
    
    Call ClearMyScreen
    
    With frmCompanyDepartments
        .cmdEditRecord.Caption = "E&dit"
        .cmdAddNew.Caption = "&New"
    End With
End Sub

Public Sub ClearMyScreen()
With frmCompanyDepartments
    .txtDeptCode = Empty
    .txtDeptName = Empty
    .cboDeptHead = Empty
    .txtDeptNotes = Empty
    .txtOfficialTitle = Empty
    .txtStaffID = Empty
End With
End Sub

Public Sub GetDeptCode()
On Error GoTo err
With frmCompanyDepartments
If Not NewRecord And Not beditRECORD Then Exit Sub
    If .cboDeptHead.ListCount <> 0 Then .cboDeptHead.Refresh: Exit Sub
    
    Dim rsLIST As ADODB.Recordset
    Set rsLIST = cnCOMMON.Execute("SELECT * FROM ParamCompanyDepts ORDER BY Allnames;")
    
    With rsLIST
    If .EOF And .BOF Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            frmCompanyDepartments.cboDeptHead.AddItem !allnames
        .MoveNext
        Loop
    End With
    
    Set rsLIST = Nothing
    
End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub CheckEditRecord()
On Error GoTo err
If NewRecord Then Exit Sub

With frmCompanyDepartments
    If .txtDeptCode.Text = Empty Then
        MsgBox "There is No Current Record to Edit!", vbCritical + vbOKOnly, "Cannot Edit"
    Else
        beditRECORD = True
        .cmdEditRecord.Caption = "SAVE &CHANGES"
        .txtDeptCode.SetFocus
        .txtDeptCode.Locked = True
    End If
End With

Exit Sub
err:
    ErrorMessage
End Sub

Public Sub EditMyRecord()
On Error GoTo err
If beditRECORD Then
    If ValidRecord Then

    With frmCompanyDepartments
    Set rsEditRecord = New ADODB.Recordset
    rsEditRecord.Open "SELECT * FROM ParamCompanyDepts WHERE deptcode='" & .txtDeptCode.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsEditRecord.EOF And rsEditRecord.BOF Then GoTo OUTS
            
        rsEditRecord!deptcode = .txtDeptCode
        rsEditRecord!deptname = .txtDeptName
        rsEditRecord!hod = .cboDeptHead
        rsEditRecord!deptnotes = .txtDeptNotes
        rsEditRecord!officialtitle = .txtOfficialTitle
        rsEditRecord!hodstaffid = .txtStaffID
        
        rsEditRecord.Update
        rsEditRecord.Requery
        
        .cmdEditRecord.Caption = "E&dit"
        .txtDeptCode.Locked = False
        beditRECORD = False
        
    End With
    
    End If
End If
OUTS:
Set rsEditRecord = Nothing
Exit Sub
err:
    UpdateErrorMessage
End Sub


Public Sub SearchByKeyCode()
On Error GoTo err
With frmCompanyDepartments
    Set rsFindRecord = New ADODB.Recordset
    
    rsFindRecord.Open "SELECT * FROM ParamCompanyDepts WHERE deptcode='" & .txtDeptCode.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then GoTo OUTS
        
        .txtDeptCode = rsFindRecord!deptcode & ""
        .txtDeptName = rsFindRecord!deptname & ""
        .cboDeptHead = rsFindRecord!hod & ""
        .txtDeptNotes = rsFindRecord!deptnotes & ""
        .txtOfficialTitle = rsFindRecord!officialtitle & ""
        .txtStaffID = rsFindRecord!hodstaffid & ""
        .cmdEditRecord.SetFocus
    
OUTS:
    Set rsFindRecord = Nothing
End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub GetDepartmentHead()
On Error GoTo err
If Not NewRecord And beditRECORD Then Exit Sub
    Dim rsFindRecord As ADODB.Recordset
    Set rsFindRecord = New ADODB.Recordset
    
    rsFindRecord.Open "SELECT * FROM ParamCompanyDepts WHERE allnames='" & frmCompanyDepartments.cboDeptHead.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsFindRecord
    If .EOF And .BOF Then Exit Sub
        frmCompanyDepartments.txtStaffID.Text = !StaffIdNo & ""
        frmCompanyDepartments.txtOfficialTitle.Text = !officialtitle & ""
        frmCompanyDepartments.txtDeptNotes.SetFocus
    End With
    Exit Sub
err:
    ErrorMessage
End Sub

Public Sub AddNewRecord()
If beditRECORD Then Exit Sub
NewRecord = True
With frmCompanyDepartments
    .txtDeptCode.SetFocus
    .cmdAddNew.Caption = "SAVE &RECORD"
    .txtDeptCode.Locked = False
    .cmdAddNew.Default = True
End With
End Sub

Public Sub SaveNewRecord()
On Error GoTo err
If Not NewRecord Then Exit Sub
If ValidRecord Then
    Set rsNewRecord = New ADODB.Recordset
    rsNewRecord.Open "SELECT * FROM ParamCompanyDepts ORDER BY deptcode;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With frmCompanyDepartments
    rsNewRecord.AddNew
    
        rsNewRecord!deptcode = .txtDeptCode
        rsNewRecord!deptname = .txtDeptName
        rsNewRecord!hod = .cboDeptHead
        rsNewRecord!deptnotes = .txtDeptNotes
        rsNewRecord!officialtitle = .txtOfficialTitle
        rsNewRecord!hodstaffid = .txtStaffID
    
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

Public Sub LoadDepartment()
With rsFindRecord
If .EOF And .BOF Then Exit Sub
    frmCompanyDepartments.txtDeptCode = !deptcode & ""
    frmCompanyDepartments.txtDeptName = !deptname & ""
    frmCompanyDepartments.cboDeptHead = !hod & ""
    frmCompanyDepartments.txtDeptNotes = !deptnotes & ""
    frmCompanyDepartments.txtOfficialTitle = !officialtitle & ""
    frmCompanyDepartments.txtStaffID = !hodstaffid & ""
End With
End Sub

