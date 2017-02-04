VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODASCreditAuthorization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub calculateDEPOSIT()
On Error GoTo err
    With frmODASMCreditAuthorization
        .txtAmount.Text = FormatNumber(CDbl(.txtPercentAuthorized) * CDbl(.txtBalance.Text) / 100)
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub checkCASHSTATUS()
On Error GoTo err
    With frmODASMCreditAuthorization
    
        Set rsCONTROL = New Recordset
        
        rsCONTROL.Open "SELECT * FROM ODASPGuarantor WHERE GuarantorType = '" & .cboGuarantorType.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
                .txtPercentAuthorized.Text = 0
                .txtAmount.Text = 0
                .txtPercentAuthorized.Locked = True
                .txtAmount.Locked = True
                .txtPercentAuthorized.BackColor = &HFFFFC0
                .txtAmount.BackColor = &HFFFFC0
                .txtRemarkRequired.Text = rsCONTROL!RequireREMARK
                
                .txtPercentAuthorized.Text = rsCONTROL!PercentAUTHORIZED & ""
                .txtAmount.Text = FormatNumber(CDbl(.txtPercentAuthorized) * CDbl(.txtPriceInclusive.Text) / 100)

                If CDbl(.txtAmount.Text) >= CDbl(.txtBalance.Text) Then
                        .txtStatus.Text = "AUTHORIZED"
                ElseIf CDbl(.txtAmount.Text) = 0 Then
                         .txtStatus.Text = "AUTHORIZED"

                Else
                        .txtStatus.Text = "NOT-AUTHORIZED"
                End If
    End With
        
rsCONTROL.Close
strPAYMETHOD = ""

Exit Sub

err:
    ErrorMessage

End Sub
Public Sub processAUTHORIZATION()
On Error GoTo err
        
        
        With frmODASMCreditAuthorization

        End With
Exit Sub
err:
    ErrorMessage
End Sub

Public Sub GenerateAuthorizationNo()
On Error GoTo err

        Set rsCONTROL = New Recordset
        
        strSQL = "SELECT * FROM ODASPLastNumbers WHERE AutoAuthorizationNo = 'Y'"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
        With rsCONTROL
                If .BOF Or .EOF = True Then Exit Sub
                        Screen.ActiveForm.txtAuthorizationNo.Text = !AuthorizationNo & ""
                
                Select Case Len(Trim(frmODASMCreditAuthorization.txtAuthorizationNo))
                        Case 1: frmODASMCreditAuthorization.txtAuthorizationNo.Text = Trim(!AuthorizationPrefix) + "00000" + Trim(frmODASMCreditAuthorization.txtAuthorizationNo)
                        Case 2: frmODASMCreditAuthorization.txtAuthorizationNo.Text = Trim(!AuthorizationPrefix) + "0000" + Trim(frmODASMCreditAuthorization.txtAuthorizationNo)
                        Case 3: frmODASMCreditAuthorization.txtAuthorizationNo.Text = Trim(!AuthorizationPrefix) + "000" + Trim(frmODASMCreditAuthorization.txtAuthorizationNo)
                        Case 4: frmODASMCreditAuthorization.txtAuthorizationNo.Text = Trim(!AuthorizationPrefix) + "00" + Trim(frmODASMCreditAuthorization.txtAuthorizationNo)
                        Case 5: frmODASMCreditAuthorization.txtAuthorizationNo.Text = Trim(!AuthorizationPrefix) + "0" + Trim(frmODASMCreditAuthorization.txtAuthorizationNo)
                        Case 6: frmODASMCreditAuthorization.txtAuthorizationNo.Text = Trim(!AuthorizationPrefix) + Trim(frmODASMCreditAuthorization.txtAuthorizationNo)
                End Select
                
                !AuthorizationNo = !AuthorizationNo + 1
                .Update
                .Requery

        End With
        
rsCONTROL.Close
strTRANS = ""

Exit Sub
err:
    ErrorMessage
End Sub



Public Sub updateRECORD()
        validateRECORD
        If bsaveRECORD = True Then
                GenerateAuthorizationNo
                saveRecord
                updateJOBBRIEF
                disableALLRECORD
                NewRecord = False
                beditRECORD = False
                Screen.ActiveForm.Toolbar1.Buttons(2).Caption = "New &Record ": Screen.ActiveForm.Toolbar1.Buttons(2).Image = 2
                Screen.ActiveForm.Toolbar1.Buttons(4).Caption = "FINISH"
        End If
        
Exit Sub

End Sub
Public Sub updateJOBBRIEF()
On Error GoTo err
    
       With frmODASMCreditAuthorization

                strSQL = "select * from ODASMJobBrief Where JobBriefNo = '" & .txtJobBriefNo & "' ;"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
                
                rsSAVE!CreditAuthorized = "Y"
                
                rsSAVE.Update
                rsSAVE.Requery
                
                
                Set rsSAVE = Nothing
                strSQL = Empty
    End With
Exit Sub

err:
    ErrorMessage
End Sub


Public Sub saveRecord()
On Error GoTo err
    
       With frmODASMCreditAuthorization

                strSQL = "select * from ODASMCreditAuthorization Where AuthorizationNo = '" & .txtAuthorizationNo & "' ;"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                If rsSAVE.EOF Or rsSAVE.BOF Then
                        rsSAVE.AddNew
                        rsSAVE!AuthorizationNo = .txtAuthorizationNo.Text
                        rsSAVE!Preparedby = CurrentUserName
                        rsSAVE!dateprepared = Date
                        rsSAVE!AuthorizedBy = CurrentUserName
                        rsSAVE!DateAuthorized = Date
                        rsSAVE!CurrentPeriod = .txtCurrentPeriod
                End If
                
                rsSAVE!JobBriefNo = .txtJobBriefNo.Text
                rsSAVE!GuarantorType = .cboGuarantorType.Text
                rsSAVE!remark = Trim(.txtRemark.Text)
                rsSAVE!Amount = CDbl(.txtAmount.Text)
                rsSAVE!Accountno = .txtAccountNo.Text
                rsSAVE!DepositPercent = CDbl(.txtPercentAuthorized)
                rsSAVE!Balance = CDbl(.txtBalance.Text)
                rsSAVE!Status = Trim(.txtStatus.Text)
                rsSAVE.Update
                rsSAVE.Requery
                
                Set rsSAVE = Nothing
                strSQL = Empty
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub enableRECORD()
On Error GoTo err
        With frmODASMCreditAuthorization
            .txtPercentAuthorized.Locked = False
            .txtRemark.Locked = False
            .cboGuarantorType.Locked = False
        End With
Exit Sub

err:
    ErrorMessage
End Sub

Public Sub validateRECORD()
On Error GoTo err
        
        With frmODASMCreditAuthorization
            
            bsaveRECORD = False
            
            If .txtAccountNo.Text = Empty Then
                    MsgBox "The Account number is Required"
                    .txtAccountNo.SetFocus
            
            ElseIf .txtRemarkRequired.Text = "Y" And Trim(.txtRemark.Text) = Empty Then
                    MsgBox "The Remark CANNOT Be Left Blank, Explanation Required"
                    .txtRemark.SetFocus
            
            ElseIf .txtAuthorizedBy.Text = Empty Then
                    MsgBox "The Authorized By is Required"
                    .txtAuthorizedBy.SetFocus
            
            ElseIf CDbl(.txtBalance.Text) <= 0 Then
                    MsgBox "Job Brief Fully paid hence No need for Approval"
                    .txtBalance.SetFocus
                    
            ElseIf .cboGuarantorType.Text = Empty Then
                    MsgBox "The Payment Method MUST be entered"
                    .cboGuarantorType.SetFocus
                    
            ElseIf .txtAuthorizationDate.Text > Date Then
                    MsgBox "The Authorization Date Cannot be in the Future"
                    .txtAuthorizationDate.SetFocus
            Else
                    bsaveRECORD = True
            End If
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadRECORD()
On Error GoTo err
    With frmODASMCreditAuthorization
        
        Dim rsCOY As ADODB.Recordset, strCOY As String
        Set rsCOY = New ADODB.Recordset
        strCOY = "SELECT * FROM ODASPGuarantor Where Status = 'A'; "
        rsCOY.Open strCOY, cnCOMMON, adOpenKeyset, adLockOptimistic

        If rsCOY.EOF Or rsCOY.BOF Then Exit Sub
        
        Set rsCONTROL = New ADODB.Recordset
        strSQL = "SELECT * FROM ODASMJobBrief JB, ODASPAccount AC  WHERE JB.AccountNo = AC.AccountNo and JB.JobBriefNo = '" & .txtJobBriefNo.Text & "' ; "
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub

                .txtCurrentPeriod.Text = CurrentPeriod
                .txtAccountNo.Text = rsCONTROL!Accountno
                .txtCompanyName.Text = rsCONTROL!CompanyName
                .txtDescriptionOfOrder.Text = rsCONTROL!descriptionOfOrder
                .txtProductCode.Text = rsCONTROL!ProductCode
                .txtJobBriefDate.Text = CDate(rsCONTROL!JobBriefDate)
                
                If IsNull(rsCONTROL!CommencementDate) = True Then
                       .txtCommencementDate.Text = Date
                Else: .txtCommencementDate.Text = CDate(rsCONTROL!CommencementDate) & ""
                End If
                
                If IsNull(rsCONTROL!TotalPrice) = True Then
                      .txtPriceInclusive.Text = 0
                Else: .txtPriceInclusive.Text = FormatNumber(rsCONTROL!TotalPrice) & ""
                End If
                
                If IsNull(rsCONTROL!Balance) = True Then
                        .txtBalance.Text = 0
                Else: .txtBalance.Text = FormatNumber(rsCONTROL!Balance) & ""
                End If
                
                .txtAuthorizationDate.Text = Date
                .cboGuarantorType.Clear
                .txtAuthorizedBy.Text = CurrentUserName
                
                .txtPercentAuthorized.Text = rsCOY!PercentAUTHORIZED & ""
                
                .txtAmount.Text = FormatNumber(CDbl(.txtPercentAuthorized) * CDbl(.txtPriceInclusive.Text) / 100)
                .txtStatus.Text = "NOT-AUTHORIZATION"
    End With

Exit Sub
err:
    ErrorMessage
End Sub


