VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmODAMApproval 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Approval/Authorization"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   8010
   Icon            =   "frmODAMApproval.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   8010
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9735
      Begin VB.Frame Frame6 
         Height          =   1215
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   6495
         Begin VB.TextBox txtComment 
            BackColor       =   &H00FFFFC0&
            Height          =   735
            Left            =   960
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   240
            Width           =   5415
         End
         Begin VB.Label lblNarration 
            Caption         =   "Comment"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   300
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1935
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   7815
         Begin VB.CheckBox chkAccept 
            Caption         =   "Accept?"
            Height          =   375
            Left            =   3240
            TabIndex        =   1
            Top             =   1020
            Width           =   975
         End
         Begin VB.TextBox cboUserCode 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1200
            TabIndex        =   26
            Top             =   540
            Width           =   1695
         End
         Begin VB.TextBox txtOperationDescription 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1200
            TabIndex        =   22
            Top             =   1440
            Width           =   6135
         End
         Begin VB.TextBox txtPassword 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   5280
            PasswordChar    =   "*"
            TabIndex        =   2
            Top             =   1005
            Width           =   2055
         End
         Begin VB.TextBox txtOperationDate 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1200
            TabIndex        =   9
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtStatus 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4440
            TabIndex        =   10
            Top             =   525
            Width           =   2895
         End
         Begin VB.TextBox txtOperationType 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4440
            TabIndex        =   8
            Top             =   120
            Width           =   2895
         End
         Begin VB.TextBox txtApplicationNo 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1200
            TabIndex        =   7
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Password"
            Height          =   255
            Left            =   4440
            TabIndex        =   21
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Date"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "Status"
            Height          =   255
            Left            =   3240
            TabIndex        =   19
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lbluserCode 
            Caption         =   "User "
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lbPolicyNo 
            AutoSize        =   -1  'True
            Caption         =   "ApplicationNo"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   990
         End
         Begin VB.Label lbNames 
            Caption         =   "Type"
            Height          =   255
            Left            =   3240
            TabIndex        =   14
            Top             =   225
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   6720
         TabIndex        =   12
         Top             =   2040
         Width           =   1215
         Begin VB.CommandButton cmdDelete 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            Picture         =   "frmODAMApproval.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1860
            Width           =   975
         End
         Begin VB.CommandButton cmdSearch 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            Picture         =   "frmODAMApproval.frx":0544
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   990
            Width           =   975
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   1425
            Width           =   975
         End
         Begin VB.CommandButton cmdUpdate 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODAMApproval.frx":0646
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton cmdCancel 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            Picture         =   "frmODAMApproval.frx":0748
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   2295
            Width           =   975
         End
         Begin VB.CommandButton cmdADDNEW 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODAMApproval.frx":084A
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1815
         Left            =   120
         TabIndex        =   11
         Top             =   3240
         Width           =   6495
         Begin MSComctlLib.ListView ListView1 
            Height          =   1575
            Left            =   120
            TabIndex        =   27
            Top             =   120
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   2778
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
   End
End
Attribute VB_Name = "frmODAMApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCode As ADODB.Recordset, strcode As String
Dim rsPAYMENTTYPE As ADODB.Recordset

Sub ClearControls()
    With frmALISMLoanOperation
        .txtComment.Text = ""
       .chkAccept.Value = 0
        .txtPassword.Text = ""
    End With
End Sub

Private Sub EnableControls()
On Error GoTo err

    With frmALISMLoanOperation
        .txtApplicationNo.Locked = True
        .cboUserCode.Locked = True
        .txtComment.Locked = False
        .txtOperationType.Locked = True
        .txtStatus.Locked = True
        .txtOperationDate.Locked = True
        .chkAccept.Enabled = True
        .txtPassword.Locked = False
        .txtOperationDescription.Locked = True
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub DisableControls()
On Error GoTo err
    With frmALISMLoanOperation
        .txtApplicationNo.Locked = True
        .cboUserCode.Locked = True
        .txtComment.Locked = True
        .txtOperationType.Locked = True
        .txtOperationDate.Locked = True
        .txtStatus.Locked = True
        .chkAccept.Enabled = False
        .txtPassword.Locked = True
        .txtOperationDescription.Locked = True
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub SaveRECORD()
On Error GoTo err

    Dim rsSAVE As ADODB.Recordset
    Set rsSAVE = New ADODB.Recordset
    
    strSQL = ""
    strSQL = "SELECT * from ALISMLoanOperation where ApplicationNo = '" & GlobalApplicationNo & "' AND operationType = '" & GlobalOperationType & "';"
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsSAVE
            .AddNew
            !ApplicationNo = frmALISMLoanOperation.txtApplicationNo
            !UserCode = frmALISMLoanOperation.cboUserCode.Text
            !OperationType = frmALISMLoanOperation.txtOperationType
            !operationDate = frmALISMLoanOperation.txtOperationDate & ""
            !Status = frmALISMLoanOperation.txtStatus.Text
            !Comment = frmALISMLoanOperation.txtComment.Text
            
            If frmALISMLoanOperation.chkAccept.Value = 1 Then
                !Accept = "Y"
            Else: !Accept = "N"
            End If
            
            .Update
            .Requery
   End With
Exit Sub

err:
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
            UpdateErrorMessage
    End If

End Sub
Private Sub updateLOAN()
On Error GoTo err

    Dim rsSAVE As ADODB.Recordset, strAPPLICATION As String
    Set rsSAVE = New Recordset

    strAPPLICATION = "SELECT * FROM ALISMLoanApplication where applicationno = '" & frmALISMLoanOperation.txtApplicationNo & "';"
    rsSAVE.Open strAPPLICATION, cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsSAVE
            If .BOF Or .EOF = True Then Exit Sub
            !Status = frmALISMLoanOperation.txtStatus
            
            If rsPAYMENTTYPE!Checked = "1" Then
                    !PreparedBy = CurrentUserName
                    !DatePrepared = Date
                    If frmALISMLoanOperation.chkAccept.Value = 1 Then
                        !Checked = "Y"
                    Else: !Checked = "N"
                    End If

            ElseIf rsPAYMENTTYPE!Approved = "1" Then
                    !ApprovedBy = CurrentUserName
                    !DateApproved = Date
                    If frmALISMLoanOperation.chkAccept.Value = 1 Then
                        !Approved = "Y"
                    Else: !Approved = "N"
                    End If

            ElseIf rsPAYMENTTYPE!Authorized = "1" Then
                    !AuthorizedBy = CurrentUserName
                    !DateAuthorized = Date
                    If frmALISMLoanOperation.chkAccept.Value = 1 Then
                        !Authorized = "Y"
                    Else: !Authorized = "N"
                    End If
            End If
                                
            .Update
            End With
Exit Sub

err:
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
            UpdateErrorMessage
    End If

End Sub
Private Sub updateDISCHARGE()
On Error GoTo err

        Dim rsDischarge As ADODB.Recordset, strDischarge As String
        Set rsDischarge = New ADODB.Recordset
    
        strDischarge = "SELECT * FROM ALISMClaimtOtal where claimno = '" & frmALISMLoanOperation.txtApplicationNo & "';"
        rsDischarge.Open strDischarge, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsDischarge
                !Status = Screen.ActiveForm.txtStatus
                
                If rsPAYMENTTYPE!DischargeApproval = "1" Then
                        !ApprovedBy = CurrentUserName
                        !DateApproved = Date
                        
                        If frmALISMLoanOperation.chkAccept.Value = 1 Then
                            !Approved = "Y"
                        Else: !Approved = "N"
                        End If
                        
                ElseIf rsPAYMENTTYPE!DischargeAuthorization = "1" Then
                        !AuthorizedBy = CurrentUserName
                        !DateAuthorized = Date
                        If frmALISMLoanOperation.chkAccept.Value = 1 Then
                                !Authorized = "Y"
                        Else: !Authorized = "N"
                        End If
                End If
    
                .Update
                .Requery
        End With

Exit Sub
err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsDischarge.CancelUpdate
            rsDischarge.Requery
    Else
            UpdateErrorMessage
    End If

End Sub
Private Sub updateREGISTRATION()
On Error GoTo err

        Dim rsDischarge As ADODB.Recordset, strDischarge As String
        Set rsDischarge = New ADODB.Recordset
    
        strDischarge = "SELECT * FROM ALISMClaimRegistration where claimno = '" & frmALISMLoanOperation.txtApplicationNo & "';"
        rsDischarge.Open strDischarge, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsDischarge

                If rsPAYMENTTYPE!RegistrationApproval = "1" Then
                        !ApprovedBy = CurrentUserName
                        !DateApproved = Date

                        If frmALISMLoanOperation.chkAccept.Value = 1 Then
                            !Approved = "Y"
                        Else: !Approved = "N"
                        End If
                ElseIf rsPAYMENTTYPE!RegistrationAuthorization = "1" Then
                        !AuthorizedBy = CurrentUserName
                        !DateAuthorized = Date

                        If frmALISMLoanOperation.chkAccept.Value = 1 Then
                            !Authorized = "Y"
                        Else: !Authorized = "N"
                        End If
                End If
    
                .Update
                .Requery
        End With

Exit Sub
err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsDischarge.CancelUpdate
            rsDischarge.Requery
    Else
            UpdateErrorMessage
    End If

End Sub
Private Sub updateREINSTATEMENT()
On Error GoTo err

        Dim rsREINSTATEMENT As ADODB.Recordset, strDischarge As String
        Set rsREINSTATEMENT = New ADODB.Recordset
    
        strDischarge = "SELECT * FROM ALISMReinstatement where Policyno = '" & frmALISMLoanOperation.txtApplicationNo & "';"
        rsREINSTATEMENT.Open strDischarge, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsREINSTATEMENT

                If rsPAYMENTTYPE!ReinstatementApproval = "1" Then
                        !ApprovedBy = CurrentUserName
                        !DateApproved = Date

                        If frmALISMLoanOperation.chkAccept.Value = 1 Then
                            !Approved = "Y"
                        Else: !Approved = "N"
                        End If
                ElseIf rsPAYMENTTYPE!ReinstatementAuthorization = "1" Then
                        !AuthorizedBy = CurrentUserName
                        !DateAuthorized = Date

                        If frmALISMLoanOperation.chkAccept.Value = 1 Then
                            !Authorized = "Y"
                        Else: !Authorized = "N"
                        End If
                End If
    
                .Update
                .Requery
        End With

Exit Sub
err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsREINSTATEMENT.CancelUpdate
            rsREINSTATEMENT.Requery
    Else
            UpdateErrorMessage
    End If

End Sub
Private Sub updateLAPSE()
On Error GoTo err

        Set rsSAVE = New ADODB.Recordset
    
        strSQL = "SELECT * FROM ALISMLapses where Policyno = '" & frmALISMLoanOperation.txtApplicationNo & "';"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsSAVE

                If rsPAYMENTTYPE!LapseApproval = "1" Then
                        !ApprovedBy = CurrentUserName
                        !DateApproved = Date
                        
                        If frmALISMLoanOperation.chkAccept.Value = 1 Then
                            !Approved = "Y"
                            updateSTATUS
                        Else: !Approved = "N"
                        End If
                
                ElseIf rsPAYMENTTYPE!LapseAuthorization = "1" Then
                        !AuthorizedBy = CurrentUserName
                        !DateAuthorized = Date

                        If frmALISMLoanOperation.chkAccept.Value = 1 Then
                            !Authorized = "Y"
                        Else: !Authorized = "N"
                        End If
                End If
    
                .Update
                .Requery
        End With

Exit Sub
err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
            UpdateErrorMessage
    End If

End Sub
Private Sub updateSURRENDER()
On Error GoTo err

        Set rsSAVE = New ADODB.Recordset
    
        strSQL = "SELECT * FROM ALISMSurrender where Policyno = '" & frmALISMLoanOperation.txtApplicationNo & "';"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsSAVE

                If rsPAYMENTTYPE!SurrenderApproval = "1" Then
                        !ApprovedBy = CurrentUserName
                        !DateApproved = Date
                        
                        If frmALISMLoanOperation.chkAccept.Value = 1 Then
                            !Approved = "Y"
                            updateSTATUS
                        Else: !Approved = "N"
                        End If
                
                ElseIf rsPAYMENTTYPE!SurrenderAuthorization = "1" Then
                        !AuthorizedBy = CurrentUserName
                        !DateAuthorized = Date

                        If frmALISMLoanOperation.chkAccept.Value = 1 Then
                            !Authorized = "Y"
                        Else: !Authorized = "N"
                        End If
                End If
    
                .Update
                .Requery
        End With

Exit Sub
err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
            UpdateErrorMessage
    End If

End Sub


Private Sub updatePAIDUP()
On Error GoTo err

        Dim rsREINSTATEMENT As ADODB.Recordset, strDischarge As String
        Set rsREINSTATEMENT = New ADODB.Recordset
    
        strDischarge = "SELECT * FROM ALISMPaidup where Policyno = '" & frmALISMLoanOperation.txtApplicationNo & "';"
        rsREINSTATEMENT.Open strDischarge, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsREINSTATEMENT

                If rsPAYMENTTYPE!paidupApproval = "1" Then
                        !ApprovedBy = CurrentUserName
                        !DateApproved = Date

                        If frmALISMLoanOperation.chkAccept.Value = 1 Then
                            !Approved = "Y"
                        Else: !Approved = "N"
                        End If
                ElseIf rsPAYMENTTYPE!paidupAuthorization = "1" Then
                        !AuthorizedBy = CurrentUserName
                        !DateAuthorized = Date

                        If frmALISMLoanOperation.chkAccept.Value = 1 Then
                            !Authorized = "Y"
                        Else: !Authorized = "N"
                        End If
                End If
    
                .Update
                .Requery
        End With

Exit Sub
err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsREINSTATEMENT.CancelUpdate
            rsREINSTATEMENT.Requery
    Else
            UpdateErrorMessage
    End If

End Sub
Private Sub updatePROPOSAL()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
    
        strSQL = "SELECT * FROM ALISMProposal where Proposalno = '" & frmALISMLoanOperation.txtApplicationNo & "';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsCONTROL

                If rsPAYMENTTYPE!ProposalApproval = "1" Then
                        !Approved = "Y"
                        !Authorized = "N"
                        !ApprovedBy = CurrentUserName
                        !DateApproved = Date
                        !StatusCode = "COMPLETE"
                        
                        If frmALISMLoanOperation.chkAccept.Value = 1 Then
                            !Approved = "Y"
                        Else: !Approved = "N"
                        End If
                ElseIf rsPAYMENTTYPE!ProposalAuthorization = "1" Then
                        !AuthorizedBy = CurrentUserName
                        !DateAuthorized = Date

                        If frmALISMLoanOperation.chkAccept.Value = 1 Then
                            !Authorized = "Y"
                        Else: !Authorized = "N"
                        End If
                End If
    
                .Update
                .Requery
        End With

Exit Sub
err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsCONTROL.CancelUpdate
            rsCONTROL.Requery
    Else
            UpdateErrorMessage
    End If

End Sub

Private Sub updatePOLICY()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
    
        strSQL = "SELECT * FROM ALISMPolicy where Policyno = '" & frmALISMLoanOperation.txtApplicationNo & "';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsCONTROL

                If rsPAYMENTTYPE!PolicyApproval = "1" Then
                        !ApprovedBy = CurrentUserName
                        !DateApproved = Date
                        !StatusCode = "COMPLETE"
                        
                        If frmALISMLoanOperation.chkAccept.Value = 1 Then
                            !Approved = "Y"
                        Else: !Approved = "N"
                        End If
                ElseIf rsPAYMENTTYPE!PolicyAuthorization = "1" Then
                        !AuthorizedBy = CurrentUserName
                        !DateAuthorized = Date

                        If frmALISMLoanOperation.chkAccept.Value = 1 Then
                            !Authorized = "Y"
                        Else: !Authorized = "N"
                        End If
                End If
    
                .Update
                .Requery
        End With

Exit Sub
err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsCONTROL.CancelUpdate
            rsCONTROL.Requery
    Else
            UpdateErrorMessage
    End If

End Sub
Private Sub updateSTATUS()
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
    
        strSQL = "SELECT * FROM ALISMPolicy where Policyno = '" & frmALISMLoanOperation.txtApplicationNo & "';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsCONTROL
                !StatusDate = Date

                If rsPAYMENTTYPE!SurrenderApproval = "1" Then
                        !StatusCode = "SURRENDERED"
                ElseIf rsPAYMENTTYPE!LapseAuthorization = "1" Then
                        !StatusCode = "LAPSED"
                End If
    
                .Update
                .Requery
        End With

Exit Sub
err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsCONTROL.CancelUpdate
            rsCONTROL.Requery
    Else
            UpdateErrorMessage
    End If

End Sub

Private Sub updateRECORD()

        Set rsPAYMENTTYPE = New ADODB.Recordset
        
        rsPAYMENTTYPE.Open "SELECT * FROM ODASPOperationType WHERE OperationType = '" & frmALISMLoanOperation.txtOperationType.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsPAYMENTTYPE
        
                If .EOF And .BOF Then Exit Sub
                     
                If rsPAYMENTTYPE!Approved = "1" Or rsPAYMENTTYPE!Checked = "1" Or rsPAYMENTTYPE!Authorized = "1" Then
                       updateLOAN
                ElseIf rsPAYMENTTYPE!PaymentApproval = "1" Or rsPAYMENTTYPE!PaymentPreparation = "1" Or rsPAYMENTTYPE!PaymentAuthorization = "1" Then
                        updatePaymentRequisition
                ElseIf rsPAYMENTTYPE!DischargeApproval = "1" Or rsPAYMENTTYPE!DischargePreparation = "1" Or rsPAYMENTTYPE!DischargeAuthorization = "1" Then
                        updateDISCHARGE
                ElseIf rsPAYMENTTYPE!ChequeApproval = "1" Or rsPAYMENTTYPE!ChequePreparation = "1" Or rsPAYMENTTYPE!ChequeAuthorization = "1" Or rsPAYMENTTYPE!ChequeIssuance = "1" Then
                        updateCHEQUE
                ElseIf rsPAYMENTTYPE!RegistrationApproval = "1" Or rsPAYMENTTYPE!RegistrationAuthorization = "1" Then
                        updateREGISTRATION
                ElseIf rsPAYMENTTYPE!ReinstatementApproval = "1" Or rsPAYMENTTYPE!ReinstatementAuthorization = "1" Then
                        updateREINSTATEMENT
                ElseIf rsPAYMENTTYPE!paidupApproval = "1" Or rsPAYMENTTYPE!paidupAuthorization = "1" Then
                        updatePAIDUP
                ElseIf rsPAYMENTTYPE!ProposalApproval = "1" Or rsPAYMENTTYPE!ProposalAuthorization = "1" Then
                        updatePROPOSAL
                ElseIf rsPAYMENTTYPE!PolicyApproval = "1" Or rsPAYMENTTYPE!PolicyAuthorization = "1" Then
                        updatePOLICY
                ElseIf rsPAYMENTTYPE!SurrenderApproval = "1" Or rsPAYMENTTYPE!SurrenderAuthorization = "1" Then
                        updateSURRENDER
                ElseIf rsPAYMENTTYPE!LapseApproval = "1" Or rsPAYMENTTYPE!LapseAuthorization = "1" Then
                        updateLAPSE

                End If
        End With
        '/ Payment Type

End Sub
Private Sub updatePaymentRequisition()
On Error GoTo err
          
          Dim rsPAYMENT As ADODB.Recordset, strPAYMENT As String
          Set rsPAYMENT = New Recordset
    
          strPAYMENT = "SELECT * FROM ALISMPaymentRequisition where Requisitionno = '" & frmALISMLoanOperation.txtApplicationNo & "';"
          rsPAYMENT.Open strPAYMENT, cnCOMMON, adOpenKeyset, adLockOptimistic

          With rsPAYMENT
                  !Status = Screen.ActiveForm.txtStatus
                   
                  If rsPAYMENTTYPE!PaymentApproval = "1" Then
                            !ApprovedBy = CurrentUserName
                            !DateApproved = Date
                            
                            If frmALISMLoanOperation.chkAccept.Value = 1 Then
                                  !Approved = "Y"
                            Else
                                  !Approved = "N"
                            End If
                  ElseIf rsPAYMENTTYPE!PaymentAuthorization = "1" Then
                            !AuthorizedBy = CurrentUserName
                            !DateAuthorized = Date
                            If frmALISMLoanOperation.chkAccept.Value = 1 Then
                                  !Authorized = "Y"
                            Else: !Authorized = "N"
                            End If
                  End If

                  .Update
          End With
Exit Sub

err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsPAYMENT.CancelUpdate
            rsPAYMENT.Requery
    Else
            UpdateErrorMessage
    End If

End Sub

Private Sub updateCHEQUE()
On Error GoTo err
      Dim rscheque As ADODB.Recordset, strCheque As String
      Set rscheque = New Recordset

      strCheque = "SELECT * FROM ALISMCheque where Chequeno = '" & frmALISMLoanOperation.txtApplicationNo & "';"
      rscheque.Open strCheque, cnCOMMON, adOpenKeyset, adLockOptimistic

      With rscheque
              !Status = txtStatus
              
              If rsPAYMENTTYPE!ChequeApproval = "1" Then
                      !ApprovedBy = CurrentUserName
                      !DateApproved = Date
                      
                      If frmALISMLoanOperation.chkAccept.Value = 1 Then
                            !Approved = "Y"
                      Else: !Approved = "N"
                      End If
                  
              ElseIf rsPAYMENTTYPE!ChequeAuthorization = "1" Then
                      !AuthorizedBy = CurrentUserName
                      !DateAuthorized = Date
                      
                      If frmALISMLoanOperation.chkAccept.Value = 1 Then
                            !Authorized = "Y"
                      Else: !Authorized = "N"
                      End If
              
              ElseIf rsPAYMENTTYPE!ChequeIssuance = "1" Then
                      !IssuedBy = CurrentUserName
                      !DateIssued = Date
                      If frmALISMLoanOperation.chkAccept.Value = 1 Then
                            !Issued = "Y"
                      Else: !Issued = "N"
                      End If
              End If

              .Update
      End With

Exit Sub

err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rscheque.CancelUpdate
            rscheque.Requery
    Else
            UpdateErrorMessage
    End If

End Sub
Private Sub ShowCode()

On Error GoTo err
    Set RsCode = New Recordset
    strcode = "SELECT * from ALISMLoanOperation where ApplicationNo = '" & GlobalApplicationNo & "' AND operationType = '" & GlobalOperationType & "';"

    RsCode.Open strcode, cnCOMMON, adOpenKeyset, adLockOptimistic
  
    With RsCode
            Screen.ActiveForm.txtApplicationNo.Text = !ApplicationNo
            Screen.ActiveForm.txtOperationType.Text = !OperationType
            Screen.ActiveForm.txtOperationDate = !operationDate
            Screen.ActiveForm.txtStatus = !Status
            Screen.ActiveForm.txtComment = !Comment
            Screen.ActiveForm.txtAccept = !Accept
            cboUserCode = !UserCode
            Screen.ActiveForm.txtOperationDescription = GlobalOperationDescription
    End With

Exit Sub

err:
    UpdateErrorMessage
End Sub

Private Sub DisableCommandButtons()
    cmdADDNEW.Enabled = False
    cmdUpdate.Enabled = False
    cmdCancel.Enabled = True
    cmdEdit.Enabled = False
End Sub
Private Sub EnableCommandButtons()
On Error GoTo err
    cmdADDNEW.Enabled = True
    cmdUpdate.Enabled = False
    cmdCancel.Enabled = True
    cmdEdit.Enabled = True
    Exit Sub
err:
ErrorMessage

End Sub

Private Sub cmdAddNew_Click()
        ClearControls
        EnableControls
        disableButtons
        
        With frmALISMLoanOperation
                .txtPassword.BackColor = &HFFC0C0
                .txtComment.BackColor = &HFFC0C0
                .cmdUpdate.Enabled = False
                .chkAccept.Value = 1
        End With
End Sub
Private Sub checkSTATUS()
On Error GoTo err

        Dim rsCHECK As ADODB.Recordset, strCHECK As String
        Set rsCHECK = New ADODB.Recordset
        
        strCHECK = "SELECT * from ALISMLoanOperation where ApplicationNo = '" & GlobalApplicationNo & "' AND operationType = '" & GlobalOperationType & "';"
        rsCHECK.Open strCHECK, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCHECK
                
                If .EOF Or .BOF Then Exit Sub
                
                If !Accept = "Y" Then
                        DisableControls
                        frmALISMLoanOperation.cmdUpdate = False
                        frmALISMLoanOperation.cmdADDNEW = False
                End If
        End With

rsCHECK.Close

Exit Sub

err:
    ErrorMessage

End Sub
Private Sub cmdCancel_Click()
        EnableCommandButtons
        ClearControls
        DisableControls
End Sub


Private Sub ValidateDATA()

On Error GoTo err

    bSaveRECORD = False
    
    With frmALISMLoanOperation
    
              If .txtApplicationNo.Text = "" Then
                      MsgBox "ApplicationNo is Required"
                      .txtApplicationNo.SetFocus
              
              ElseIf .cboUserCode.Text = "" Then
                      MsgBox "The UserCode is  required"
                      .cboUserCode.SetFocus
              
              ElseIf .txtComment.Text = "" And .chkAccept.Value = 0 Then
                      MsgBox "The comment is  required"
                      .txtComment.SetFocus
            
              ElseIf .txtOperationType.Text = "" Then
                      MsgBox "The Operation Type is  required"
                      .txtOperationType.SetFocus
              
              ElseIf .chkAccept.Value = 0 Then
                      MsgBox "There is no Purpose of undertaking this process without checking the Approval Check Box", vbOKOnly
                      .chkAccept.SetFocus
              
              ElseIf Trim(.txtPassword.Text) <= "" Then
                      MsgBox "Passwords are Required", vbOKOnly
                      txtPassword.SetFocus
              
              Else
                         bSaveRECORD = True
              End If
    End With
        
                    
Exit Sub

err:
    UpdateErrorMessage
            
End Sub


Private Sub cmdEdit_Click()
On Error GoTo err

        Dim strQRE As Variant
        Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset

        Select Case cmdEdit.Caption
                Case "&Edit"
                        EnableControls

                        strQRE = InputBox("Enter Application No  to search.", "Search Value")
    
                        rsFind.Open "SELECT * FROM ALISMLoanOperation WHERE ApplicationNo = '" & frmALISMLoanOperation.txtApplicationNo & "' and  operationtype = '" & frmALISMLoanOperation.txtOperationType & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic

                        With rsFind
                                If .EOF And .BOF Then
                                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                                Else
                                        Screen.ActiveForm.txtApplicationNo.Text = !ApplicationNo
                                        Screen.ActiveForm.txtOperationType.Text = !OperationType
                                        Screen.ActiveForm.txtOperationDate = !operationDate
                                        Screen.ActiveForm.txtStatus = !Status
                                        Screen.ActiveForm.txtComment = !Comment
                                        
                                        If !Accept = "Y" Then
                                                Screen.ActiveForm.txtAccept = !Accept
                                        Else
                                                Screen.ActiveForm.txtAccept = !Accept
                                        End If
                                        
                                        Screen.ActiveForm.cboUserCode = !UserCode
                                        Screen.ActiveForm.txtOperationDescription = GlobalOperationDescription
                                        Edit = True
                                End If
                        End With
                        
                        
                        If Edit Then
                                cmdEdit.Caption = "Save &Changes"
                        End If
    
                Case "Save &Changes"
                        Dim rsFinder As ADODB.Recordset
                        Set rsFinder = New ADODB.Recordset

                        rsFinder.Open "SELECT * FROM ALISMLoanOperation WHERE ApplicationNo = '" & frmALISMLoanOperation.txtApplicationNo & "' and  operationtype = '" & frmALISMLoanOperation.txtOperationType & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                        With rsFinder
                            !ApplicationNo = Screen.ActiveForm.txtApplicationNo
                            !UserCode = Screen.ActiveForm.cboUserCode
                            !OperationType = Screen.ActiveForm.txtOperationType
                            !operationDate = Screen.ActiveForm.txtOperationDate & ""
                            !Status = Screen.ActiveForm.txtStatus
                            !Comment = Screen.ActiveForm.txtComment
                            
                            If Screen.ActiveForm.chkAccept.Value = 1 Then
                                !Accept = "Y"
                            Else
                                !Accept = "N"
                            End If
                            .Update
                            .Requery
                            Edit = False
                    End With
                
                    ClearControls
                    cmdEdit.Caption = "&Edit"
            Case Else
        
            Exit Sub

        End Select

Exit Sub

err:

    If err.Number = 40009 Then
            MsgBox "Record requested does not exist in the Database! Check your Entries.", vbInformation, "Searching."
                rsFind.Requery

            If rsFind.BOF Then Exit Sub
                rsFind.MoveFirst

    ElseIf err.Number = 3021 Then
            MsgBox "Requested record not found! Refresh the database and try the search again...or Check your entries.", vbInformation, "Searching."
                rsFind.Requery

            If rsFind.BOF Then Exit Sub
                rsFind.MoveFirst
    Else
                ErrorMessage
End If

End Sub


Private Sub cmdUpdate_Click()
        ValidateDATA
        
        If bSaveRECORD = True Then
                SaveRECORD
                updateRECORD
                bSaveRECORD = False
                listAPPROVALTASKS
                enableButtons
                DisableControls
        End If
End Sub
Private Sub saveCLAIM()
On Error GoTo err
    
    Dim rsCLAIM As ADODB.Recordset, strCLAIM As String
    
    Set rsCLAIM = New Recordset
    
    strCLAIM = "SELECT * from ALISMClaimStatus where claimNo = '" & GlobalClaimNo & " '; "
    rsCLAIM.Open strCLAIM, cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsCLAIM
            If .EOF Or .BOF Then
                    MsgBox "Record cannot be Processed", vbOKOnly
                    Exit Sub
            End If
            
            !claimstatus = frmALISMLoanOperation.txtStatus
            !ClaimSequence = 5
            !Amount = frmALISMPaymentRequisition.txtAmount
            !StatusDate = Date
            .Update
            .Requery
    End With
Exit Sub

rsCLAIM.Close
strCLAIM = ""

err:
    UpdateErrorMessage
End Sub

Private Sub Form_Activate()
    DisableControls
    enableButtons
    loadRECORD
    checkSTATUS
    listAPPROVALTASKS
End Sub

Private Sub Form_Load()
    OpenConnection
End Sub

      
  
Private Sub loadRECORD()
On Error GoTo err:

        With frmALISMLoanOperation
                .txtApplicationNo.Text = GlobalApplicationNo
                .txtOperationType.Text = GlobalOperationType
                .txtOperationDescription.Text = GlobalOperationDescription
                .txtOperationDate.Text = Date
                .txtStatus.Text = GlobalOperationDescription
                .chkAccept.Value = 1
                .cboUserCode = CurrentUserName

        End With


Exit Sub

err:
        ErrorMessage
End Sub


Private Sub txtPassword_lostFocus()
On Error GoTo err
        If Trim(frmALISMLoanOperation.txtPassword.Text) = "" Then Exit Sub
        
        Dim rsPASSWORD As ADODB.Recordset
        Set rsPASSWORD = New Recordset
        rsPASSWORD.Open "SELECT * FROM ALISPLoanApprover WHERE Names = '" & cboUserCode & "' and operationType = '" & frmALISMLoanOperation.txtOperationType & " ';", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsPASSWORD
                If .EOF Or .BOF Then
                        MsgBox "This User is Not allowed to Check the Transaction", vbOKOnly
                        Exit Sub
                Else
                        bSaveRECORD = False
                        
                        If Trim(Screen.ActiveForm.txtPassword.Text) = "" Or Trim(!Password) = "" Then
                                MsgBox "Passwords are Required", vbOKOnly
                            
                        ElseIf Trim(Screen.ActiveForm.txtPassword.Text) <> Trim(!Password) Then
                                MsgBox "The Password Entered is invalid", vbOKOnly
                                Screen.ActiveForm.txtPassword.SetFocus
                                
                        ElseIf Trim(Screen.ActiveForm.txtPassword.Text) = Trim(!Password) Then
                            If Trim(chkAccept.Value) = 0 And Screen.ActiveForm.txtComment.Text <= "" Then
                                
                                MsgBox " The Reason for Rejecting the Application is very Important"
                                Screen.ActiveForm.txtComment.SetFocus
                                Exit Sub
                            Else
                                bSaveRECORD = True
                                Screen.ActiveForm.cmdUpdate.Enabled = True
                            End If
                            
                        End If
                End If
        
        
        End With

Exit Sub


err:
        ErrorMessage
End Sub




