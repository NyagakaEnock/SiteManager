VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASDeliveryMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub AddNewRecord()
If EditRecord Then Exit Sub
    NewRecord = True
    With Screen.ActiveForm
    
        Call ClearTheScreen
        
        .cmdNEW.Caption = "&SAVE RECORD"
        .cmdEDIT.Enabled = False
        .txtDeliveryMethodID.SetFocus
    End With
End Sub

Public Sub ClearTheScreen()
        With Screen.ActiveForm
            .txtDescriptions.Text = Empty
            .txtDeliveryMethod.Text = Empty
            .txtDeliveryMethodID.Text = Empty
           
End With
End Sub

Public Sub SaveNewRecord()
'On Error GoTo Err
With Screen.ActiveForm
    
    Set rsNewRecord = New ADODB.Recordset
    
    MySQL = "INSERT INTO ParamDeliveryMethods (DeliveryMethodid,DeliveryMethod,descriptions,createdby,datecreated,accperiod) VALUES('" & Trim(.txtDeliveryMethodID.Text) & "','" & Trim(.txtDeliveryMethod.Text) & "','" & Trim(.txtDescriptions.Text) & "','" & CurrentUserName & "','" & MyCurrentDate & "','" & MyCurrentPeriod & "')"
    
    rsNewRecord.Open MySQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Set rsNewRecord = Nothing
    
    .cmdNEW.Caption = "&NEW"
    
    NewRecord = False
    
End With

    Exit Sub
    
err:
    UpdateErrorMessage
End Sub

Public Sub EditCurrentRecord()
'On Error GoTo Err
With Screen.ActiveForm
    Set rsEditRecord = New ADODB.Recordset
    
    MySQL = "UPDATE ParamDeliveryMethods SET DeliveryMethod='" & Trim(.txtDeliveryMethod.Text) & "',descriptions='" & Trim(.txtDescriptions.Text) & "' WHERE DeliveryMethodID='" & Trim(.txtDeliveryMethodID.Text) & "';"
    
    rsEditRecord.Open MySQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Set rsEditRecord = Nothing

    .cmdEDIT.Caption = "E&DIT"
        
    EditRecord = False: AllowEdit = False
    
End With
    Exit Sub
err:
    UpdateErrorMessage
End Sub

Public Sub CheckForRecord()
'On Error GoTo Err
    If RecordExists Then
        AllowEdit = True
    Else
        AllowEdit = False
    End If
Exit Sub
err:
    ErrorMessage
End Sub

Private Function RecordExists() As Boolean
'On Error GoTo Err
With Screen.ActiveForm
    If .txtDeliveryMethodID.Text = Empty Then
        strMessage = "There is no Record to Edit Please Enter a Shipping Id!!!"
        .txtDeliveryMethodID.SetFocus
    Else
        RecordExists = True
    End If
    If Not RecordExists Then
        MsgBox strMessage, vbCritical + vbOKOnly, "Error Editing"
    End If
End With
Exit Function
err:
    ErrorMessage
End Function


Public Sub FindRecord()
'On Error GoTo Err

With Screen.ActiveForm
    
    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM ParamDeliveryMethods WHERE DeliveryMethodID='" & Trim(.txtDeliveryMethodID.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
    MsgBox "Sorry there is no record to edit Please enter a Shipping Id", vbCritical + vbOKOnly, "Error Editing"
                Set rsFindRecord = Nothing: Exit Sub
    Else
        .txtDeliveryMethodID.Text = rsFindRecord!DeliveryMethodID & ""
        .txtDeliveryMethod.Text = rsFindRecord!DeliveryMethod
        .txtDescriptions.Text = rsFindRecord!descriptions
    End If
    
    Set rsFindRecord = Nothing
      .cmdEDIT.Caption = "SAVE &CHANGES"
        EditRecord = True
    Edit = True
    
End With

Exit Sub
err:
    ErrorMessage
End Sub
