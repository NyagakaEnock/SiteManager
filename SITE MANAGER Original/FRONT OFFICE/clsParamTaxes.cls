VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParamTaxes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub FormatTaxes()
On Error GoTo Err
    With frmParamTaxes
    .txtTaxRate.Text = FormatNumber(.txtTaxRate.Text, 2, vbUseDefault, vbUseDefault, vbTrue)
    End With
Exit Sub
Err:
If Err.Number = 5 Or Err.Number = 13 Then Resume Next
    ErrorMessage
End Sub

Private Function ValidRecord() As Boolean
On Error GoTo Err
Dim strMessage As String
    With frmParamTaxes
        If .txtTaxCode.Text = Empty Then
            strMessage = "Required City Code!"
            .txtTaxCode.SetFocus
        ElseIf .txtNameTax.Text = Empty Then
            strMessage = "Required Name of City!"
            .txtNameTax.SetFocus
        ElseIf .txtTaxRate.Text = Empty Then
            strMessage = "Required Country Code!"
            .txtTaxRate.SetFocus
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

Private Sub FindCurrentTaxes()
On Error GoTo Err
    Dim rsFindRecord As ADODB.Recordset
    Set rsFindRecord = New ADODB.Recordset
    
    rsFindRecord.Open "SELECT * FROM ParamTaxes WHERE taxcode='" & CurrentRecord & "';", cnALIS, adOpenKeyset, adLockOptimistic
    
    With rsFindRecord
    If .EOF And .BOF Then Exit Sub
        frmParamTaxes.txtTaxCode = !TaxCode & ""
        frmParamTaxes.txtTaxRate = !TaxRate & ""
        frmParamTaxes.txtNameTax = !taxname & ""
    End With
    Exit Sub
Err:
    ErrorMessage
End Sub

Public Sub LoadTaxes()
With rsFindRecord
If .EOF And .BOF Then Exit Sub
    frmParamTaxes.txtTaxCode = !TaxCode & ""
    frmParamTaxes.txtTaxRate = !TaxRate & ""
    frmParamTaxes.txtNameTax = !taxname & ""
End With
End Sub

Public Sub ClearMyScreen()
    frmParamTaxes.txtTaxCode = ""
    frmParamTaxes.txtTaxRate = ""
    frmParamTaxes.txtNameTax = ""
End Sub


Private Function GetTaxCode() As String
On Error GoTo Err
Dim rsLastID As ADODB.Recordset 'used to retrieve current LastId in the Table
Dim strLastID As String 'SQL statement

Dim strTemp As String 'store current record
Dim iNumPos As Integer 'store position of the first numeral
Dim strPrefix As String 'stores Id Prefix

'Retrieve the last record in the recrdset where order is ascending

'strLastID = "SELECT max(TransTravelDetails.TravelID) as lastid from TransTravelDetails"
strLastID = "SELECT ParamTaxes.taxcode  from ParamTaxes ORDER BY taxcode;"
Set rsLastID = New ADODB.Recordset

With rsLastID
'open the recordset
    .Open strLastID, cnALIS, adOpenKeyset, adLockOptimistic
    If .RecordCount < 1 Then 'shows empty recordset
        GetTaxCode = "CT000001" 'format of desired format of the string value
    Else
       ' If .EOF And .BOF Then .MoveFirst
        .MoveLast
        strTemp = !TaxCode
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
        GetTaxCode = strPrefix & strTemp
    End If
End With
Exit Function
Err:
ErrorMessage
End Function

Public Sub CheckEditRecord()
On Error GoTo Err
If NewRecord Then Exit Sub

With frmParamTaxes
    If .txtTaxCode.Text = Empty Then
        MsgBox "There is No Current Record to Edit!", vbCritical + vbOKOnly, "Cannot Edit"
    Else
        EditRecord = True
        .cmdEditRecord.Caption = "SAVE &CHANGES"
        .txtTaxCode.SetFocus
        .txtTaxCode.Locked = True
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

    With frmParamTaxes
    Set rsEditRecord = New ADODB.Recordset
    rsEditRecord.Open "SELECT * FROM ParamTaxes WHERE TaxCode='" & .txtTaxCode.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
        If rsEditRecord.EOF And rsEditRecord.BOF Then GoTo OUTS
            
        rsEditRecord!TaxCode = .txtTaxCode.Text
        rsEditRecord!TaxRate = .txtTaxRate.Text
        rsEditRecord!taxname = .txtNameTax.Text
        
        rsEditRecord.Update
        rsEditRecord.Requery
        
        .cmdEditRecord.Caption = "E&dit"
        .txtTaxCode.Locked = False
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
With frmParamTaxes
    .txtTaxCode.SetFocus
    .cmdAddNew.Caption = "SAVE &RECORD"
    .txtTaxCode.Locked = False
    .cmdAddNew.Default = True
End With
End Sub

Public Sub SaveNewRecord()
On Error GoTo Err
If Not NewRecord Then Exit Sub
If ValidRecord Then
    Set rsNewRecord = New ADODB.Recordset
    rsNewRecord.Open "SELECT * FROM ParamTaxes ORDER BY TaxCode;", cnALIS, adOpenKeyset, adLockOptimistic
    
    With frmParamTaxes
    rsNewRecord.AddNew
    
        rsNewRecord!TaxCode = .txtTaxCode.Text
        rsNewRecord!TaxRate = .txtTaxRate.Text
        rsNewRecord!taxname = .txtNameTax.Text
    
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
    
    With frmParamTaxes
        .cmdEditRecord.Caption = "E&dit"
        .cmdAddNew.Caption = "&New"
    End With
End Sub
