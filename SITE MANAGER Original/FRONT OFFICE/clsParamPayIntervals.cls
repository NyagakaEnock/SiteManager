VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParamPayIntervals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub FindCurrentInterval()
On Error GoTo Err
    Dim rsFindRecord As ADODB.Recordset
    Set rsFindRecord = New ADODB.Recordset
    
    rsFindRecord.Open "SELECT * FROM ParamPayIntervals WHERE intervalcode='" & CurrentRecord & "';", cnALIS, adOpenKeyset, adLockOptimistic
    
    With rsFindRecord
    If .EOF And .BOF Then Exit Sub
        frmParamPayIntervals.txtIntervalCode = !intervalcode & ""
        frmParamPayIntervals.txtIntervalType = !intervaltype & ""
    End With
    Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub LoadIntervalTypes()
With rsFindRecord
If .EOF And .BOF Then Exit Sub
    frmParamPayIntervals.txtIntervalCode = !intervalcode & ""
    frmParamPayIntervals.txtIntervalType = !intervaltype & ""
    frmParamPayIntervals.txtDescriptions.Text = !descriptions & ""
End With
End Sub

Public Sub ClearMyScreen()
    frmParamPayIntervals.txtIntervalCode = Empty
    frmParamPayIntervals.txtIntervalType = Empty
    frmParamPayIntervals.txtDescriptions.Text = Empty
End Sub

Public Function GetPaymentInterval() As String
On Error GoTo Err
Dim rsLastID As ADODB.Recordset 'used to retrieve current LastId in the Table
Dim strLastID As String 'SQL statement

Dim strTemp As String 'store current record
Dim iNumPos As Integer 'store position of the first numeral
Dim strPrefix As String 'stores Id Prefix

'Retrieve the last record in the recrdset where order is ascending

'strLastID = "SELECT max(TransTravelDetails.TravelID) as lastid from TransTravelDetails"
strLastID = "SELECT ParamPayIntervals.intervalcode  from ParamPayIntervals ORDER BY intervalcode;"
Set rsLastID = New ADODB.Recordset

With rsLastID
'open the recordset
    .Open strLastID, cnALIS, adOpenKeyset, adLockOptimistic
    If .RecordCount < 1 Then 'shows empty recordset
        GetPaymentInterval = "P001" 'format of desired format of the string value
    Else
       ' If .EOF And .BOF Then .MoveFirst
        .MoveLast
        strTemp = !intervalcode
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
        GetPaymentInterval = strPrefix & strTemp
    End If
End With
Exit Function
Err:
    ErrorMessage
End Function

Private Function ValidRecord() As Boolean
On Error GoTo Err
Dim strMessage As String
    With frmParamPayIntervals
        If .txtIntervalCode.Text = Empty Then
            strMessage = "Required Interval Code!"
            .txtIntervalCode.SetFocus
        ElseIf .txtIntervalType.Text = Empty Then
            strMessage = "Required Interval Type!"
            .txtIntervalType.SetFocus
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

With frmParamPayIntervals
    If .txtIntervalCode.Text = Empty Then
        MsgBox "There is No Current Record to Edit!", vbCritical + vbOKOnly, "Cannot Edit"
    Else
        EditRecord = True
        .cmdEditRecord.Caption = "SAVE &CHANGES"
        .txtIntervalCode.SetFocus
        .txtIntervalCode.Locked = True
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

    With frmParamPayIntervals
    Set rsEditRecord = New ADODB.Recordset
    rsEditRecord.Open "SELECT * FROM ParamPayIntervals WHERE intervalcode='" & .txtIntervalCode.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
        If rsEditRecord.EOF And rsEditRecord.BOF Then GoTo OUTS
            
        rsEditRecord!intervalcode = .txtIntervalCode
        rsEditRecord!intervaltype = .txtIntervalType
        rsEditRecord!descriptions = .txtDescriptions.Text
        
        rsEditRecord.Update
        rsEditRecord.Requery
        
        .cmdEditRecord.Caption = "E&dit"
        .txtIntervalCode.Locked = False
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
With frmParamPayIntervals
    .txtIntervalCode.SetFocus
    .cmdAddNew.Caption = "SAVE &RECORD"
    .txtIntervalCode.Locked = False
    .cmdAddNew.Default = True
    If .chkAuto.Value = 1 Then .txtIntervalCode.Text = GetPaymentInterval
End With
End Sub

Public Sub SaveNewRecord()
On Error GoTo Err
If Not NewRecord Then Exit Sub
If ValidRecord Then
    Set rsNewRecord = New ADODB.Recordset
    rsNewRecord.Open "SELECT * FROM ParamPayIntervals ORDER BY intervalcode;", cnALIS, adOpenKeyset, adLockOptimistic
    
    With frmParamPayIntervals
    rsNewRecord.AddNew
    
        rsNewRecord!intervalcode = .txtIntervalCode.Text
        rsNewRecord!intervaltype = .txtIntervalType.Text
        rsNewRecord!descriptions = .txtDescriptions.Text
        
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
    
    With frmParamPayIntervals
        .cmdEditRecord.Caption = "E&dit"
        .cmdAddNew.Caption = "&New"
    End With
End Sub



