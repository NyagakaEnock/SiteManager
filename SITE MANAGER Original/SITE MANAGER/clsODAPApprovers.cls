VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODAPApprovers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public rsBENEFIT, RsCode As ADODB.Recordset, strBENEFIT As String

Private Sub loadGRID()
On Error GoTo err

        Dim rsAuthorizerGRID As ADODB.Recordset
        Set rsAuthorizerGRID = New Recordset

        rsAuthorizerGRID.Open "SELECT UserCode, Names, Status, LimitAmount, DateAssigned, DateRetired, operationType FROM ODASPApprovers where operationtype = '" & Screen.ActiveForm.cboOperationType & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
        Set Screen.ActiveForm.AuthorizerGrid.DataSource = rsAuthorizerGRID
    
Exit Sub

err:
    ErrorMessage
End Sub
Sub pclearRECORD()
On Error GoTo err

        With Screen.ActiveForm
            .cboOperationType.Text = ""
            .txtLimitAmount.Text = ""
            .txtOperationDescription = ""
            .txtPassword.Text = ""
        End With
        
Exit Sub

err:
    ErrorMessage
End Sub

Sub clearRECORD()
        With Screen.ActiveForm
            .txtUserCode.Text = ""
            .cboStatus.Text = ""
            .txtNames.Text = ""
            .txtDateRetired.Text = ""
            .txtPassword.Text = ""
            .txtDateAssigned.Text = ""
            .cboOperationType.Text = ""
            .txtLimitAmount.Text = ""
            .txtOperationDescription = ""
    End With
End Sub

Public Sub DisplayOperationType()
On Error GoTo err

    Set rsCONTROL = New ADODB.Recordset

    strSQL = "SELECT * FROM ODASPOperationType Where OperationType  = '" & Screen.ActiveForm.cboOperationType & "';"
    rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    If rsCONTROL.BOF Or rsCONTROL.EOF Then Exit Sub

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
        End With
        Exit Sub
err:
ErrorMessage

End Sub
Private Sub disableRECORD()
        
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
        End With
        Exit Sub
err:
ErrorMessage

End Sub


Private Sub SaveRECORD()
On Error GoTo err
        
        Set rsSAVE = New ADODB.Recordset
        strSQL = "SELECT * FROM ODASPApprovers Where UserCode = '" & frmODASPApprovers.txtUserCode.Text & "';"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsSAVE
                If .BOF Or .EOF Then
                        .AddNew
                        !UserCode = Screen.ActiveForm.txtUserCode
                        !DatePrepared = Date
                        !PreparedBy = CurrentUserName
                End If
                
                !Status = Screen.ActiveForm.cboStatus & ""
                !Names = Screen.ActiveForm.txtNames
                !DateRetired = Screen.ActiveForm.txtDateRetired & ""
                !Password = Screen.ActiveForm.txtPassword
                !DateAssigned = Screen.ActiveForm.txtDateAssigned
                !OperationType = Screen.ActiveForm.cboOperationType
                
                !LimitAmount = Val(Screen.ActiveForm.txtLimitAmount)
            
            .Update
            .Requery
        End With

Exit Sub

err:
If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
    RsCode.CancelUpdate
    RsCode.Requery
End If
    UpdateErrorMessage

End Sub

Sub showRECORD()
On Error GoTo err

    With RsCode
            Screen.ActiveForm.txtUserCode = !UserCode
            Screen.ActiveForm.cboStatus = !Status
            Screen.ActiveForm.txtNames = !Names
            Screen.ActiveForm.txtDateRetired = !DateRetired
            Screen.ActiveForm.txtPassword = !Password
            Screen.ActiveForm.txtDateAssigned = !DateAssigned
            Screen.ActiveForm.txtLimitAmount = !LimitAmount
            Screen.ActiveForm.cboOperationType = !OperationType
     End With
     
    Call DisplayOperationType
    Call loadGRID

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

Public Sub operationtypeGOTFOCUS()
On Error GoTo err
    
    Dim rsAGT As ADODB.Recordset, strAGT As String
    Set rsAGT = New Recordset
    
    strAGT = "SELECT * FROM ODASPOperationType;"
    rsAGT.Open strAGT, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Screen.ActiveForm.cboOperationType.Clear

    With rsAGT
            Do Until .EOF
                    Screen.ActiveForm.cboOperationType.AddItem !Description
                    .MoveNext
            Loop
    
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub operationtypeKEYPRESS(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub

Public Sub operationtypeLOSTFOCUS()
On Error GoTo err

        Dim rsPTLF As ADODB.Recordset, strPTLF As String
        Set rsPTLF = New Recordset
        
        rsPTLF.Open "SELECT * FROM ODASPOperationType WHERE Description= '" & Screen.ActiveForm.cboOperationType.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsPTLF
                If .EOF And .BOF Then Exit Sub
                Screen.ActiveForm.cboOperationType.Text = !OperationType
                Screen.ActiveForm.txtOperationDescription.Text = !Description
        End With

loadGRID

Exit Sub

err:
        ErrorMessage
End Sub
Public Sub paddRECORD()
On Error GoTo err
        penableRECORD
        pclearRECORD
        DisableCB
        Call loadGRID

        Exit Sub

err:
        ErrorMessage
End Sub


Public Sub addRECORD()
On Error GoTo err
        enableRECORD
        clearRECORD
        DisableCB
        Screen.ActiveForm.txtUserCode.SetFocus
        Call loadGRID

        Exit Sub

err:
        ErrorMessage
End Sub

Public Sub Cancelrecord()
On Error GoTo err
        EnableCB
        clearRECORD
        disableRECORD
        Exit Sub

err:
        ErrorMessage
End Sub


Public Sub deleteRECORD()
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
                clearRECORD
            End With
        End If
        '/* End if Msg Box
        
End If
        '/* If txt = ""
        
Exit Sub

err:
    ErrorMessage

End Sub

Public Sub beditRECORD()
On Error GoTo err

Dim strQRE As Variant
Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset

        Select Case Screen.ActiveForm.cmdEdit.Caption
                Case "&Edit"
                        enableRECORD

                        strQRE = InputBox("Enter User Code to search.", "Search Value")
    
                        rsFind.Open "SELECT * FROM ODASPApprovers WHERE UserCode = '" & strQRE & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

                        With rsFind
                                If .EOF And .BOF Then
                                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                                Else
                                        clearRECORD
                                        Screen.ActiveForm.txtUserCode = !UserCode
                                        Screen.ActiveForm.cboStatus = !Status
                                        Screen.ActiveForm.txtNames = !Names
                                        Screen.ActiveForm.txtDateRetired = !DateRetired
                                        Screen.ActiveForm.txtPassword = !Password
                                        Screen.ActiveForm.txtDateAssigned = !DateAssigned
                                        Screen.ActiveForm.txtLimitAmount = !LimitAmount
                                        Screen.ActiveForm.cboOperationType = !OperationType

                                        Edit = True
                                End If
                        End With
                                    Call DisplayOperationType
                        If Edit Then
                                Screen.ActiveForm.cmdEdit.Caption = "Save &Changes"
                        End If
    
                Case "Save &Changes"
                        Dim rsFinder As ADODB.Recordset
                        Set rsFinder = New ADODB.Recordset

                        rsFinder.Open "SELECT * FROM ODASPApproverss WHERE UserCode = '" & Screen.ActiveForm.txtUserCode.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                        With rsFinder
                                !UserCode = Screen.ActiveForm.txtUserCode
                                !Status = Screen.ActiveForm.cboStatus & ""
                                !Names = Screen.ActiveForm.txtNames
                                !DateRetired = Screen.ActiveForm.txtDateRetired
                                !Password = Screen.ActiveForm.txtPassword
                                !DateAssigned = Val(Screen.ActiveForm.txtDateAssigned)
                                !OperationType = Screen.ActiveForm.cboOperationType
                                !LimitAmount = Val(Screen.ActiveForm.txtLimitAmount)
                                
                                Dim rsOPERATIONTYPE As ADODB.Recordset
                                Set rsOPERATIONTYPE = New Recordset
                            
                                rsOPERATIONTYPE.Open "SELECT * FROM ODASPOperationType Where OperationType  = '" & Screen.ActiveForm.cboOperationType & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                                        
                                If rsOPERATIONTYPE!Approved = "1" Then
                                            !Approved = rsOPERATIONTYPE!Approved
                                    ElseIf rsOPERATIONTYPE!Checked = "1" Then
                                            !Checked = rsOPERATIONTYPE!Checked
                                    ElseIf rsOPERATIONTYPE!Authorized = "1" Then
                                            !Authorized = rsOPERATIONTYPE!Authorized
                                End If
                        
                                .Update
                          
                        Screen.ActiveForm.cmdEdit.Caption = "&Edit"
                        End With
                Case Else
        
            Exit Sub

       End Select

Exit Sub

err:
       ErrorMessage
End Sub

Public Sub browseRECORD(Index As Integer)
On Error GoTo err
           
       Set RsCode = New ADODB.Recordset
       RsCode.Open "SELECT * FROM ODASPApprovers;", cnCOMMON, adOpenKeyset, adLockOptimistic

        
        Screen.ActiveForm.cmdUpdate.Enabled = False
        
        With RsCode
       
                If .EOF And .BOF Then Exit Sub
    
                    Select Case Index
                        Case 0
                          .MoveFirst
                        Case 1
                            .MovePrevious
                            If .BOF Then .MoveFirst
                        Case 2
                            .MoveNext
                            If .EOF Then .MoveLast
                        Case 3
                            .MoveLast
                    End Select
        End With
        showRECORD
        Call loadGRID

Exit Sub

err:
    ErrorMessage
End Sub

Public Sub ValidateRECORD()

On Error GoTo err

        If Screen.ActiveForm.txtUserCode.Text = "" Then
                MsgBox "User Code is Required"
                Screen.ActiveForm.txtUserCode.SetFocus
         
         ElseIf Screen.ActiveForm.txtPassword.Text = "" Then
                MsgBox "The Password is  required"
                Screen.ActiveForm.txtPassword.SetFocus
        
        ElseIf Screen.ActiveForm.txtNames.Text = "" Then
                MsgBox "The Names is  required"
                Screen.ActiveForm.txtNames.SetFocus
        
        ElseIf Screen.ActiveForm.txtLimitAmount.Text <= "" Then
                MsgBox "The Limit Must be Entered"
                Screen.ActiveForm.txtLimitAmount.SetFocus
        
        ElseIf Screen.ActiveForm.cboStatus.Text <= "" Then
                MsgBox "The Status cannot be Left Blank"
                Screen.ActiveForm.cboStatus.SetFocus
        
        ElseIf Screen.ActiveForm.txtDateAssigned.Text <= "" Then
                MsgBox "The Date Assigned Cannot be Lef Blank"
                Screen.ActiveForm.txtDateAssigned.SetFocus
        Else
                bSaveRECORD = True
        End If
  
Exit Sub
err:
ErrorMessage

End Sub


Public Sub updateRECORD()
On Error GoTo err

        ValidateRECORD
        
        If bSaveRECORD = True Then
            SaveRECORD
        End If
        
        enableButtons
        disableALLRECORD
        loadGRID

'Exit Sub

err:
        'ErrorMessage
End Sub

Public Sub searchRECORD()
On Error GoTo err

        Dim strQRE As Variant
        Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset
        strQRE = InputBox("Enter User Code to search.", "Search Value")
        
        rsFind.Open "SELECT * FROM ODASPApprovers WHERE UserCode = '" & strQRE & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsFind
                If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                Else
                            clearRECORD
                            Screen.ActiveForm.txtUserCode = !UserCode
                            Screen.ActiveForm.cboStatus = !Status
                            Screen.ActiveForm.txtNames = !Names
                            Screen.ActiveForm.txtDateRetired = !DateRetired
                            Screen.ActiveForm.txtPassword = !Password
                            Screen.ActiveForm.txtDateAssigned = !DateAssigned
                            Screen.ActiveForm.txtLimitAmount = !LimitAmount
                            Screen.ActiveForm.cboOperationType = !OperationType

                            
                            Edit = True
                End If
        
            End With
                            Call DisplayOperationType
                            Call disableRECORD
        Exit Sub

err:
            ErrorMessage

End Sub

Public Sub operationAPPROVED()
On Error GoTo err:
 'MsgBox "this is where you changed"
        frmALISMLoanOperation.Show 1

 Exit Sub

err:
        ErrorMessage
End Sub
Public Sub operationAUTHORIZED()
On Error GoTo err:
        
        frmALISMLoanOperation.Show 1

Exit Sub

err:
        ErrorMessage
End Sub


Private Sub Class_Initialize1()
On Error GoTo err
    
    Set cnCOMMON = New ADODB.Connection
    cnCOMMON.ConnectionString = "PROVIDER=MSDASQL;DSN=DSNALIS;UID=;PWD=;"
    cnCOMMON.Open
    'Begin a transaction
    cnCOMMON.BeginTrans
    
    Dim total As Integer
    
    Set rsBENEFIT = New ADODB.Recordset
    rsBENEFIT.Open "SELECT * FROM ODASPApprovers;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    total = rsBENEFIT.RecordCount
    
    Call EnableCB
    
    'Commit the transaction
    cnCOMMON.CommitTrans

Exit Sub
err:
'Incase of an error the connection object shoud call its rollback object
'if an error occured, roll back the changes
'cnCOMMON.RollbackTrans
MsgBox "An error occured changing the records.", vbExclamation
ErrorMessage
End Sub

Private Sub Class_Terminate1()
On Error GoTo err
    'Close the instance of a recordset
    rsBENEFIT.Close
    
    'Destroy the instance of the recordset
    Set rsBENEFIT = Nothing
Exit Sub

err:
    ErrorMessage
End Sub

