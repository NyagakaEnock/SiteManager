VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParamTitles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub FindCurrentTitle()
On Error GoTo Err
    Dim rsFindRecord As ADODB.Recordset
    Set rsFindRecord = New ADODB.Recordset
    
    rsFindRecord.Open "SELECT * FROM ParamTitles WHERE titleid='" & CurrentRecord & "';", cnALIS, adOpenKeyset, adLockOptimistic
    
    With rsFindRecord
    If .EOF And .BOF Then Exit Sub
        frmParamTitles.txtTitleCode = !titleid & ""
        frmParamTitles.txtDescription = !Title & ""
        frmParamTitles.txtTitleofCourtesy = !Description & ""
    End With
    
    Exit Sub
Err:
    ErrorMessage
End Sub

Private Function ValidRecord() As Boolean
On Error GoTo Err
Dim strMessage As String
    With frmParamTitles
        If .txtTitleCode.Text = Empty Then
            strMessage = "Required City Code!"
            .txtTitleCode.SetFocus
        ElseIf .txtTitleofCourtesy.Text = Empty Then
            strMessage = "Required Name of City!"
            .txtTitleofCourtesy.SetFocus
        ElseIf .txtDescription.Text = Empty Then
            strMessage = "Required Country Code!"
            .txtDescription.SetFocus
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

Private Function GetTitleID() As String
On Error GoTo Err
Dim rsLastID As ADODB.Recordset 'used to retrieve current LastId in the Table
Dim strLastID As String 'SQL statement

Dim strTemp As String 'store current record
Dim iNumPos As Integer 'store position of the first numeral
Dim strPrefix As String 'stores Id Prefix

'Retrieve the last record in the recrdset where order is ascending

'strLastID = "SELECT max(TransTravelDetails.TravelID) as lastid from TransTravelDetails"
strLastID = "SELECT Paramtitles.titleid  from ParamTitles ORDER BY titleid;"
Set rsLastID = New ADODB.Recordset

With rsLastID
'open the recordset
    .Open strLastID, cnALIS, adOpenKeyset, adLockOptimistic
    If .RecordCount < 1 Then 'shows empty recordset
        GetTitleID = "T001" 'format of desired format of the string value
    Else
       ' If .EOF And .BOF Then .MoveFirst
        .MoveLast
        strTemp = !titleid
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
        GetTitleID = strPrefix & strTemp
    End If
End With
Exit Function
Err:
    ErrorMessage
End Function

Public Sub LoadCity()
With rsFindRecord
If .EOF And .BOF Then Exit Sub
    frmParamTitles.txtTitleCode = !titleid & ""
    frmParamTitles.txtDescription = !Title & ""
    frmParamTitles.txtTitleofCourtesy = !Description & ""
End With
End Sub

Public Sub ClearMyScreen()
    frmParamTitles.txtTitleCode = Empty
    frmParamTitles.txtDescription = Empty
    frmParamTitles.txtTitleofCourtesy = Empty
End Sub

Public Sub CheckEditRecord()
On Error GoTo Err
If NewRecord Then Exit Sub

With frmParamTitles
    If .txtTitleCode.Text = Empty Then
        MsgBox "There is No Current Record to Edit!", vbCritical + vbOKOnly, "Cannot Edit"
    Else
        EditRecord = True
        .cmdEditRecord.Caption = "SAVE &CHANGES"
        .txtTitleCode.SetFocus
        .txtTitleCode.Locked = True
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

    With frmParamTitles
    Set rsEditRecord = New ADODB.Recordset
    rsEditRecord.Open "SELECT * FROM ParamTitles WHERE TitleID='" & .txtTitleCode.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
        If rsEditRecord.EOF And rsEditRecord.BOF Then GoTo OUTS
            
        rsEditRecord!titleid = .txtTitleCode
        rsEditRecord!Title = .txtDescription
        rsEditRecord!Description = .txtTitleofCourtesy
        
        rsEditRecord.Update
        rsEditRecord.Requery
        
        .cmdEditRecord.Caption = "E&dit"
        .txtTitleCode.Locked = False
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
With frmParamTitles
    .txtTitleCode.SetFocus
    .cmdAddNew.Caption = "SAVE &RECORD"
    .txtTitleCode.Locked = False
    .cmdAddNew.Default = True
    If .chkAuto.Value = 1 Then .txtTitleCode.Text = GetTitleID
End With
End Sub

Public Sub SaveNewRecord()
On Error GoTo Err
If Not NewRecord Then Exit Sub
If ValidRecord Then
    Set rsNewRecord = New ADODB.Recordset
    rsNewRecord.Open "SELECT * FROM ParamTitles ORDER BY TitleID;", cnALIS, adOpenKeyset, adLockOptimistic
    
    With frmParamTitles
    rsNewRecord.AddNew
    
        rsNewRecord!titleid = .txtTitleCode
        rsNewRecord!Title = .txtDescription
        rsNewRecord!Description = .txtTitleofCourtesy
    
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
    
    With frmParamTitles
        .cmdEditRecord.Caption = "E&dit"
        .cmdAddNew.Caption = "&New"
    End With
End Sub
