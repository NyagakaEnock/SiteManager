VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommonData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ClearVariables()
On Error GoTo err
    aA = Empty
    bB = Empty
    cC = Empty
    dD = Empty
    eE = Empty
    fF = Empty
    gG = Empty
    hH = Empty
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub ClearTheScreen()
On Error GoTo err
For Each i In Screen.ActiveForm
'    If TypeOf i Is DTPicker Then
'        i.Value = Empty
'    End If
    If TypeOf i Is TextBox Then
        i.Text = Empty
    End If
    If TypeOf i Is ComboBox Then
        i.Clear
    End If
    If TypeOf i Is ListView Then
        i.ListItems.Clear
    End If
    If TypeOf i Is Image Then
        i.Picture = LoadPicture("")
    End If
    If TypeOf i Is OptionButton Then
        i.Value = False
    End If
    If TypeOf i Is CheckBox And i.Name <> "chkAuto" Then
        i.Value = 0
    End If
Next i
Exit Sub
err:
    If err.Number = 35787 Then
        Resume Next
    Else
        ErrorMessage
    End If
End Sub

Public Sub ClearTextFields()
For Each i In Screen.ActiveForm
    If TypeOf i Is TextBox Then
        i.Text = Empty
    End If
    If TypeOf i Is ComboBox Then
        i.Clear
    End If
    If TypeOf i Is Image Then
        i.Picture = LoadPicture("")
    End If
    If TypeOf i Is UpDown Then
        i.Value = Empty
    End If
Next i
End Sub

Public Sub ClearDataSheets()
For Each i In Screen.ActiveForm
    If TypeOf i Is ListView Then
        i.ListItems.Clear
    End If
Next i
End Sub

Public Sub AttachAccountingPeriod()
On Error GoTo err
Dim BMonth, LMonth, BYear, LYear, PMonth, PYear, AccPeriod

With Screen.ActiveForm
    BMonth = 1: LMonth = 12
    BYear = (Year(Date) - 2): LYear = Year(Date)
    .cboAccPeriod.Clear
    
    For i = BYear To LYear
    PYear = i
        For j = BMonth To LMonth
            If j = 10 Or j = 11 Or j = 12 Then
                PMonth = j
            Else
                PMonth = 0 & j
            End If
            AccPeriod = PYear & "/" & PMonth
            .cboAccPeriod.AddItem AccPeriod
            .cboAccPeriod.Text = MyCurrentPeriod
        Next j
    Next i
    
    
End With
    Exit Sub
err:
    ErrorMessage
End Sub






Public Sub SaveNewRecord()
On Error GoTo err
With Screen.ActiveForm
    Set rsNewRecord = New ADODB.Recordset
    rsNewRecord.Open NewSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    UpdateSucceeded = True
    Set rsNewRecord = Nothing
End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub EditCurrentRecord()
On Error GoTo err
With Screen.ActiveForm
    Set rsEditRecord = New ADODB.Recordset
    rsEditRecord.Open EditSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    UpdateSucceeded = True
    Set rsEditRecord = Nothing
End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub FindCurrentRecord()
On Error GoTo err
With Screen.ActiveForm
    Set rsFindRecord = New ADODB.Recordset
    rsFindRecord.Open FindSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    Set rsFindRecord = Nothing
End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub DeleteCurrentRecord()
On Error GoTo err
With Screen.ActiveForm
    Set rsDeleteRecord = New ADODB.Recordset
    rsDeleteRecord.Open DeleteSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    Set rsDeleteRecord = Nothing
End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub AttachDropDown()
On Error GoTo err
With Screen.ActiveForm
    
    Set rsDropDown = New ADODB.Recordset
    rsDropDown.Open AttachSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    If TypeOf Screen.ActiveForm.ActiveControl Is ComboBox Then
        If rsDropDown.EOF And rsDropDown.BOF Then Exit Sub
            rsDropDown.MoveFirst
            Do While Not rsDropDown.EOF
                Do Until rsDropDown!selectfield <> APCode And Not rsDropDown.EOF
                    rsDropDown.MoveNext: If rsDropDown.EOF Then rsDropDown.MoveLast: GoTo OUTS
                Loop
            
                If Not IsNull(rsDropDown!selectfield) Then
                    Screen.ActiveForm.ActiveControl.AddItem rsDropDown!selectfield & ""
                End If
                
                APCode = Trim(rsDropDown!selectfield & "")
            rsDropDown.MoveNext
            Loop
    End If
OUTS:
    APCode = Empty
    Set rsDropDown = Nothing
End With
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub AttachInventDropDown()
On Error GoTo err
With Screen.ActiveForm
    
    Set rsDropDown = New ADODB.Recordset
    rsDropDown.Open AttachSQL, cnINVENT, adOpenKeyset, adLockOptimistic
    
    If TypeOf Screen.ActiveForm.ActiveControl Is ComboBox Then
        If rsDropDown.EOF And rsDropDown.BOF Then Exit Sub
            rsDropDown.MoveFirst
            Do While Not rsDropDown.EOF
                Do Until rsDropDown!selectfield <> APCode And Not rsDropDown.EOF
                    rsDropDown.MoveNext: If rsDropDown.EOF Then rsDropDown.MoveLast: GoTo OUTS
                Loop
            
                If Not IsNull(rsDropDown!selectfield) Then
                    Screen.ActiveForm.ActiveControl.AddItem rsDropDown!selectfield & ""
                End If
                
                APCode = Trim(rsDropDown!selectfield & "")
            rsDropDown.MoveNext
            Loop
    End If
OUTS:
    APCode = Empty
    Set rsDropDown = Nothing
End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub AttachCalendarYears()
On Error GoTo err
    If Not TypeOf Screen.ActiveForm.ActiveControl Is ComboBox Then Exit Sub
    With Screen.ActiveForm
    Dim a, b, c, d
        a = Year(Date)
        b = a - 35: d = a - 1
        
        For c = b To a
            .ActiveControl.AddItem c
        Next c
        
        .ActiveControl.Text = d
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Function GetStartDate() As Date
On Error GoTo err

    Set rsFindRecord = cnCOMMON.Execute("SELECT StartDate FROM ParamCompanyMaster ")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetStartDate = Empty: Exit Function
    ElseIf IsNull(rsFindRecord!StartDate) = True Or rsFindRecord!StartDate = "" Then
        GetStartDate = Empty: Exit Function
    Else
        GetStartDate = CDate(rsFindRecord!StartDate)
    End If
          
    Set rsFindRecord = Nothing
    

Exit Function
err:
    ErrorMessage
End Function

