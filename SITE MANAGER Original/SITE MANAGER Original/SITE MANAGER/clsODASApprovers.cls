VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASApprovers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public rsBENEFIT, RsCode As ADODB.Recordset, strBENEFIT As String, bSaveRECORD As Boolean


Public Sub pclearRECORD()
On Error GoTo err

        With Screen.ActiveForm
            .txtLimitAmount.Text = ""
            .txtPassword.Text = ""
            .txtReenterPassword.Text = ""
            .txtDateAssigned.Text = Date
            .cboStatus.Text = "INFORCE"
        End With
        
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub clearALLRECORD()
        With Screen.ActiveForm
            .txtUserCode.Text = ""
            .cboStatus.Text = ""
            .txtNames.Text = ""
            .txtDateRetired.Text = ""
            .txtPassword.Text = ""
            .txtDateAssigned.Text = ""
            .txtLimitAmount.Text = ""
            .txtReenterPassword.Text = ""
    End With
End Sub

Public Sub DisplayOperationType()
On Error GoTo err

    Set rsCONTROL = New Recordset
    
    strSQL = "SELECT * FROM ODASPOperationType Where OperationType  = '" & Screen.ActiveForm.cboOperationType & "';"
    rsCONTROL.Open "SELECT * FROM ODASPOperationType Where OperationType  = '" & Screen.ActiveForm.cboOperationType & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
    
    Screen.ActiveForm.txtOperationDescription = rsCONTROL!Description
         
Exit Sub

err:
    ErrorMessage
End Sub
Sub penableRECORD()
On Error GoTo err
        With Screen.ActiveForm
                .txtUserCode.Locked = True
                .cboStatus.Locked = True
                .txtNames.Locked = True
                .txtDateAssigned.Locked = True
                .txtLimitAmount.Locked = False
                .txtDateRetired.Locked = True
                .txtPassword.Locked = False
                .cboOperationType.Locked = False
                .txtOperationDescription.Locked = True
        End With
        Exit Sub
err:
ErrorMessage

End Sub

Sub enableRECORD()
On Error GoTo err
        With Screen.ActiveForm
                .txtUserCode.Locked = False
                .cboStatus.Locked = False
                .txtNames.Locked = False
                .txtDateAssigned.Locked = False
                .txtLimitAmount.Locked = False
                .txtDateRetired.Locked = False
                .txtPassword.Locked = False
                .cboOperationType.Locked = False
                .txtOperationDescription.Locked = True
                .txtReenterPassword.Locked = False
        End With
        Exit Sub
err:
ErrorMessage

End Sub
Sub disableRECORD()
       
On Error GoTo err

        With Screen.ActiveForm
                .txtUserCode.Locked = True
                .cboStatus.Locked = True
                .txtNames.Locked = True
                .txtDateAssigned.Locked = True
                .txtLimitAmount.Locked = True
                .txtDateRetired.Locked = True
                .txtPassword.Locked = True
                .cboOperationType.Locked = True
                .txtOperationDescription.Locked = True
                .txtReenterPassword.Locked = True
        End With
        Exit Sub
err:
ErrorMessage

End Sub


Private Sub SaveRECORD()
On Error GoTo err
        
        Set rsSAVE = New ADODB.Recordset
        strSQL = "SELECT * FROM ODASPApprovers where OperationType = '" & Screen.ActiveForm.cboOperationType & "' and StaffId = '" & Screen.ActiveForm.txtUserCode.Text & "';"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsSAVE
                If .EOF Or .BOF Then
                    .AddNew
                    !StaffId = Screen.ActiveForm.txtUserCode
                    !DatePrepared = Date
                    !PreparedBy = CurrentUserName
                End If
                
                !Status = Screen.ActiveForm.cboStatus & ""
                !Password = Screen.ActiveForm.txtPassword
                !DateAssigned = Screen.ActiveForm.txtDateAssigned
                !OperationType = Screen.ActiveForm.cboOperationType
                !LimitAmount = Val(Screen.ActiveForm.txtLimitAmount)
                bSaveRECORD = False
            
                .Update
                .Requery
        End With

Exit Sub

err:
If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
    rsSAVE.CancelUpdate
    rsSAVE.Requery
End If
    UpdateErrorMessage

End Sub

Sub showRECORD()
On Error GoTo err

    With RsCode
            Screen.ActiveForm.txtUserCode = !UserName
            Screen.ActiveForm.cboStatus = !Status
            Screen.ActiveForm.txtNames = !Names
            Screen.ActiveForm.txtDateRetired = !DateRetired
            Screen.ActiveForm.txtPassword = !Password
            Screen.ActiveForm.txtDateAssigned = !DateAssigned
            Screen.ActiveForm.txtLimitAmount = !LimitAmount
            Screen.ActiveForm.cboOperationType = !OperationType
     End With
     
    Call DisplayOperationType
'    Call loadGRID

Exit Sub

err:
        ErrorMessage
End Sub

Public Sub DisableCB()
On Error GoTo err
            Screen.ActiveForm.cmdUpdate.Enabled = True
            Screen.ActiveForm.cmdAdd.Enabled = False
            Screen.ActiveForm.cmdSearch.Enabled = False
            Screen.ActiveForm.cmdEdit.Enabled = False
            Screen.ActiveForm.cmdDelete.Enabled = False
            Screen.ActiveForm.cmdCancel.Enabled = True
            Exit Sub
err:
ErrorMessage

End Sub
Public Sub EnableCB()
On Error GoTo err
            Screen.ActiveForm.cmdUpdate.Enabled = False
            Screen.ActiveForm.cmdAdd.Enabled = True
            Screen.ActiveForm.cmdSearch.Enabled = True
            Screen.ActiveForm.cmdEdit.Enabled = True
            Screen.ActiveForm.cmdDelete.Enabled = True
            Screen.ActiveForm.cmdCancel.Enabled = True
            Exit Sub
err:
ErrorMessage

End Sub


Public Sub statusGOTFOCUS()
On Error GoTo err
            Screen.ActiveForm.cboStatus.Clear
            Screen.ActiveForm.cboStatus.AddItem "ACTIVE"
            Screen.ActiveForm.cboStatus.AddItem "IN ACTIVE"
    
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub statusKEYPRESS(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Public Sub statusLOSTFOCUS()
        If Screen.ActiveForm.cboStatus.Text = "ACTIVE" Then
            Screen.ActiveForm.cboStatus.Text = "A"
            Else: Screen.ActiveForm.cboStatus.Text = "I"
        End If
End Sub


Public Sub cboLoanType_KeyPress(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub

Public Sub paddRECORD()
On Error GoTo err
        penableRECORD
        pclearRECORD
'        disableButtons
'        Call loadGRID

        Exit Sub

err:
        ErrorMessage
End Sub


Public Sub AddRECORD()
        enableRECORD
       clearALLRECORD
'        disableButtons
        baddRECORD = True
        beditRECORD = False
        bsearchRECORD = False

End Sub

Public Sub cancelRECORD()
'        enableButtons
       clearALLRECORD
        disableRECORD
        baddRECORD = False
        beditRECORD = False
        bsearchRECORD = False

End Sub


Public Sub DeleteRecord()
On Error GoTo err

If Screen.ActiveForm.txtUserCode.Text = "" Then
            MsgBox "There is no current record to delete", vbInformation, "Delete Information"
ElseIf Screen.ActiveForm.txtNames.Text = "" Then
            MsgBox "There is no current record", vbInformation, "Delete Information"
Else
        If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            With rsBENEFIT
                
                If .EOF And .BOF Then Exit Sub
                .Delete
                .Requery
               clearALLRECORD
            End With
        End If
        '/* End if Msg Box
        
End If
        '/* If txt = ""
        
Exit Sub

err:
    ErrorMessage

End Sub

Public Sub EditRecord()
    enableRECORD
    changeRECORD
End Sub
Private Sub changeRECORD()
On Error GoTo err

Dim strQRE As Variant
Dim rsFIND As ADODB.Recordset, edit As Boolean

        Set rsFIND = New ADODB.Recordset

        Select Case Screen.ActiveForm.cmdEdit.Caption
                Case "&Edit"
    
                        strQRE = InputBox("Enter User Code to search.", "Search Value")
    
                        rsFIND.Open "SELECT * FROM ODASPApprovers WHERE UserName = '" & strQRE & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

                        With rsFIND
                                If .EOF And .BOF Then
                                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                                Else
                                       clearALLRECORD
                                        Screen.ActiveForm.txtUserCode = !UserName
                                        Screen.ActiveForm.cboStatus = !Status
                                        Screen.ActiveForm.txtNames = !Names
                                        Screen.ActiveForm.txtDateRetired = !DateRetired
                                        Screen.ActiveForm.txtPassword = !Password
                                        Screen.ActiveForm.txtDateAssigned = !DateAssigned
                                        Screen.ActiveForm.txtLimitAmount = !LimitAmount
                                        Screen.ActiveForm.cboOperationType = !OperationType

                                        edit = True
                                End If
                        End With
                                    Call DisplayOperationType
                        If edit Then
                                Screen.ActiveForm.cmdEdit.Caption = "Save &Changes"
                        End If
    
                Case "Save &Changes"
                        Dim rsFinder As ADODB.Recordset
                        Set rsFinder = New ADODB.Recordset

                        rsFinder.Open "SELECT * FROM ODASPApproverss WHERE UserName = '" & Screen.ActiveForm.txtUserCode.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                        With rsFinder
                                !UserName = Screen.ActiveForm.txtUserCode
                                !Status = Screen.ActiveForm.cboStatus & ""
                                !Names = Screen.ActiveForm.txtNames
                                !DateRetired = Screen.ActiveForm.txtDateRetired
                                !Password = Screen.ActiveForm.txtPassword
                                !DateAssigned = Val(Screen.ActiveForm.txtDateAssigned)
                                !OperationType = Screen.ActiveForm.cboOperationType
                                !LimitAmount = Val(Screen.ActiveForm.txtLimitAmount)
                                .Update
                                .Requery
                          
                        Screen.ActiveForm.cmdEdit.Caption = "&Edit"
                        End With
                Case Else
        
            Exit Sub

       End Select

Exit Sub

err:
       ErrorMessage
End Sub

Public Sub ValidateRECORD()
On Error GoTo err
        With Screen.ActiveForm
                
                If .txtUserCode.Text = "" Then
                        MsgBox "User Code is Required"
                        .txtUserCode.SetFocus
                 
                 ElseIf .txtPassword.Text = "" Then
                        MsgBox "The Password is  required"
                        .txtPassword.SetFocus
                
                ElseIf .txtNames.Text = "" Then
                        MsgBox "The Names is  required"
                        .txtNames.SetFocus
                
                ElseIf .txtLimitAmount.Text <= "" Then
                        MsgBox "The Limit Must be Entered"
                        .txtLimitAmount.SetFocus
                
                ElseIf .cboStatus.Text <= "" Then
                        MsgBox "The Status cannot be Left Blank"
                        .cboStatus.SetFocus
                
                ElseIf .txtDateAssigned.Text <= "" Then
                        MsgBox "The Date Assigned Cannot be Lef Blank"
                        .txtDateAssigned.SetFocus
                
                ElseIf .txtPassword.Text <> .txtReenterPassword.Text Then
                        MsgBox "The Passwords entered are not Similar", vbOKOnly
                        .txtPassword.SetFocus
                Else
                        bSaveRECORD = True
                End If
        End With
  
Exit Sub
err:
ErrorMessage

End Sub


Public Sub updateRECORD()
        ValidateRECORD
        
        If bSaveRECORD = True Then
            SaveRECORD
            If bSaveRECORD = False Then
'                enableButtons
                disableALLRECORD
                GetUserCode
                GetAPPROVERS
                baddRECORD = False
                beditRECORD = False
                bsearchRECORD = False

            End If
        End If
        
End Sub

Public Sub SearchRecord()
On Error GoTo err


        Set rsCONTROL = New ADODB.Recordset
        
        strSQL = "SELECT * FROM ODASPApprovers, ADminUserRegister WHERE ODASPApprovers.StaffId = AdminUserRegister.StaffIdNo and ODASPApprovers.OperationType = '" & Screen.ActiveForm.cboOperationType & "' and StaffId = '" & Screen.ActiveForm.txtUserCode & "' ;"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsCONTROL
                If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                Else
                            Screen.ActiveForm.txtUserCode = !UserName
                            Screen.ActiveForm.cboStatus = !Status
                            Screen.ActiveForm.txtNames = !AllNames
                            Screen.ActiveForm.txtDateRetired = !DateRetired & ""
                            Screen.ActiveForm.txtPassword = !Password & ""
                            Screen.ActiveForm.txtDateAssigned = !DateAssigned & ""
                            Screen.ActiveForm.txtLimitAmount = !LimitAmount & ""
                            Screen.ActiveForm.cboOperationType = !OperationType & ""
                End If
        
            End With
            Call DisplayOperationType
            disableALLRECORD
        
        Exit Sub

err:
            ErrorMessage

End Sub


