VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmODASMOperation 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Approval/Authorization"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   8010
   Icon            =   "frmODASMOperation.frx":0000
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
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   240
            Width           =   5175
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
            Width           =   6495
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
            Width           =   2415
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
            Width           =   3255
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
            Width           =   3255
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
            Picture         =   "frmODASMOperation.frx":0442
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
            Picture         =   "frmODASMOperation.frx":0544
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
            Picture         =   "frmODASMOperation.frx":0A76
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
            Picture         =   "frmODASMOperation.frx":0B78
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
            Picture         =   "frmODASMOperation.frx":0C7A
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1695
         Left            =   120
         TabIndex        =   11
         Top             =   3240
         Width           =   6495
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   30
            Left            =   240
            TabIndex        =   28
            Top             =   1800
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   53
            _Version        =   393216
            Appearance      =   1
         End
         Begin MSComctlLib.ProgressBar ProgressBar2 
            DragMode        =   1  'Automatic
            Height          =   375
            Left            =   0
            TabIndex        =   29
            Top             =   1320
            Visible         =   0   'False
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1455
            Left            =   120
            TabIndex        =   27
            Top             =   120
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   2566
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
Attribute VB_Name = "frmODASMOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCode As ADODB.Recordset, strcode As String

Sub ClearControls()
    With frmODASMOperation
        .txtComment.Text = ""
        .chkAccept.Value = 0
        .txtPassword.Text = ""
    End With
End Sub

Private Sub EnableControls()
On Error GoTo err

    With frmODASMOperation
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

Private Sub disableControls()
On Error GoTo err
    With frmODASMOperation
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

Private Sub saveRecord()
On Error GoTo err

    Set rsSAVE = New ADODB.Recordset
    
    strSQL = ""
    strSQL = "SELECT * from ODASMOperation where ApplicationNo = '" & GlobalApplicationNo & "' AND operationType = '" & GlobalOperationType & "';"
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsSAVE
            If .BOF Or .EOF Then
                    .AddNew
                    !ApplicationNo = frmODASMOperation.txtApplicationNo
                    !DatePrepared = Date
                    !PreparedBY = CurrentUserName
                    !OperationType = frmODASMOperation.txtOperationType
            End If
            
            !UserCode = frmODASMOperation.cboUserCode.Text
            !operationDate = frmODASMOperation.txtOperationDate & ""
            !Status = frmODASMOperation.txtStatus.Text
            !Comment = frmODASMOperation.txtComment.Text
            
            If frmODASMOperation.chkAccept.Value = 1 Then
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

    strAPPLICATION = "SELECT * FROM ODASMApplication where applicationno = '" & frmODASMOperation.txtApplicationNo & "';"
    rsSAVE.Open strAPPLICATION, cnCOMMON, adOpenKeyset, adLockOptimistic

    With rsSAVE
            If .BOF Or .EOF = True Then Exit Sub
            !Status = frmODASMOperation.txtStatus
            
            If rsCONTROL!Checked = "1" Then
                    !PreparedBY = CurrentUserName
                    !DatePrepared = Date
                    If frmODASMOperation.chkAccept.Value = 1 Then
                        !Checked = "Y"
                    Else: !Checked = "N"
                    End If

            ElseIf rsCONTROL!Approved = "1" Then
                    !ApprovedBy = CurrentUserName
                    !DateApproved = Date
                    If frmODASMOperation.chkAccept.Value = 1 Then
                        !Approved = "Y"
                    Else: !Approved = "N"
                    End If

            ElseIf rsCONTROL!Authorized = "1" Then
                    !AuthorizedBy = CurrentUserName
                    !DateAuthorized = Date
                    If frmODASMOperation.chkAccept.Value = 1 Then
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
    
        strDischarge = "SELECT * FROM ALISMClaimtOtal where claimno = '" & frmODASMOperation.txtApplicationNo & "';"
        rsDischarge.Open strDischarge, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsDischarge
                !Status = Screen.ActiveForm.txtStatus
                
                If rsCONTROL!DischargeApproval = "1" Then
                        !ApprovedBy = CurrentUserName
                        !DateApproved = Date
                        
                        If frmODASMOperation.chkAccept.Value = 1 Then
                            !Approved = "Y"
                        Else: !Approved = "N"
                        End If
                        
                ElseIf rsCONTROL!DischargeAuthorization = "1" Then
                        !AuthorizedBy = CurrentUserName
                        !DateAuthorized = Date
                        If frmODASMOperation.chkAccept.Value = 1 Then
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
    
        strDischarge = "SELECT * FROM ALISMClaimRegistration where claimno = '" & frmODASMOperation.txtApplicationNo & "';"
        rsDischarge.Open strDischarge, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsDischarge

                If rsCONTROL!RegistrationApproval = "1" Then
                        !ApprovedBy = CurrentUserName
                        !DateApproved = Date

                        If frmODASMOperation.chkAccept.Value = 1 Then
                            !Approved = "Y"
                        Else: !Approved = "N"
                        End If
                ElseIf rsCONTROL!RegistrationAuthorization = "1" Then
                        !AuthorizedBy = CurrentUserName
                        !DateAuthorized = Date

                        If frmODASMOperation.chkAccept.Value = 1 Then
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
    
        strDischarge = "SELECT * FROM ALISMReinstatement where Policyno = '" & frmODASMOperation.txtApplicationNo & "';"
        rsREINSTATEMENT.Open strDischarge, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsREINSTATEMENT

                If rsCONTROL!ReinstatementApproval = "1" Then
                        !ApprovedBy = CurrentUserName
                        !DateApproved = Date

                        If frmODASMOperation.chkAccept.Value = 1 Then
                            !Approved = "Y"
                        Else: !Approved = "N"
                        End If
                ElseIf rsCONTROL!ReinstatementAuthorization = "1" Then
                        !AuthorizedBy = CurrentUserName
                        !DateAuthorized = Date

                        If frmODASMOperation.chkAccept.Value = 1 Then
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
    
        strSQL = "SELECT * FROM ALISMLapses where Policyno = '" & frmODASMOperation.txtApplicationNo & "';"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsSAVE

                If rsCONTROL!LapseApproval = "1" Then
                        !ApprovedBy = CurrentUserName
                        !DateApproved = Date
                        
                        If frmODASMOperation.chkAccept.Value = 1 Then
                            !Approved = "Y"
                            updateSTATUS
                        Else: !Approved = "N"
                        End If
                
                ElseIf rsCONTROL!LapseAuthorization = "1" Then
                        !AuthorizedBy = CurrentUserName
                        !DateAuthorized = Date

                        If frmODASMOperation.chkAccept.Value = 1 Then
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
    
        strSQL = "SELECT * FROM ALISMSurrender where Policyno = '" & frmODASMOperation.txtApplicationNo & "';"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsSAVE

                If rsCONTROL!SurrenderApproval = "1" Then
                        !ApprovedBy = CurrentUserName
                        !DateApproved = Date
                        
                        If frmODASMOperation.chkAccept.Value = 1 Then
                            !Approved = "Y"
                            updateSTATUS
                        Else: !Approved = "N"
                        End If
                
                ElseIf rsCONTROL!SurrenderAuthorization = "1" Then
                        !AuthorizedBy = CurrentUserName
                        !DateAuthorized = Date

                        If frmODASMOperation.chkAccept.Value = 1 Then
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
    
        strDischarge = "SELECT * FROM ALISMPaidup where Policyno = '" & frmODASMOperation.txtApplicationNo & "';"
        rsREINSTATEMENT.Open strDischarge, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsREINSTATEMENT

                If rsCONTROL!paidupApproval = "1" Then
                        !ApprovedBy = CurrentUserName
                        !DateApproved = Date

                        If frmODASMOperation.chkAccept.Value = 1 Then
                            !Approved = "Y"
                        Else: !Approved = "N"
                        End If
                ElseIf rsCONTROL!paidupAuthorization = "1" Then
                        !AuthorizedBy = CurrentUserName
                        !DateAuthorized = Date

                        If frmODASMOperation.chkAccept.Value = 1 Then
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
    
        strSQL = "SELECT * FROM ALISMProposal where Proposalno = '" & frmODASMOperation.txtApplicationNo & "';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsCONTROL

                If rsCONTROL!ProposalApproval = "1" Then
                        !Approved = "Y"
                        !Authorized = "N"
                        !ApprovedBy = CurrentUserName
                        !DateApproved = Date
                        !StatusCode = "COMPLETE"
                        
                        If frmODASMOperation.chkAccept.Value = 1 Then
                            !Approved = "Y"
                        Else: !Approved = "N"
                        End If
                ElseIf rsCONTROL!ProposalAuthorization = "1" Then
                        !AuthorizedBy = CurrentUserName
                        !DateAuthorized = Date

                        If frmODASMOperation.chkAccept.Value = 1 Then
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
    
        strSQL = "SELECT * FROM ALISMPolicy where Policyno = '" & frmODASMOperation.txtApplicationNo & "';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsCONTROL

                If rsCONTROL!PolicyApproval = "1" Then
                        !ApprovedBy = CurrentUserName
                        !DateApproved = Date
                        !StatusCode = "COMPLETE"
                        
                        If frmODASMOperation.chkAccept.Value = 1 Then
                            !Approved = "Y"
                        Else: !Approved = "N"
                        End If
                ElseIf rsCONTROL!PolicyAuthorization = "1" Then
                        !AuthorizedBy = CurrentUserName
                        !DateAuthorized = Date

                        If frmODASMOperation.chkAccept.Value = 1 Then
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
    
        strSQL = "SELECT * FROM ALISMPolicy where Policyno = '" & frmODASMOperation.txtApplicationNo & "';"
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
        With rsCONTROL
                !StatusDate = Date

                If rsCONTROL!SurrenderApproval = "1" Then
                        !StatusCode = "SURRENDERED"
                ElseIf rsCONTROL!LapseAuthorization = "1" Then
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
On Error GoTo err

        Set rsCONTROL = New ADODB.Recordset
        
        rsCONTROL.Open "SELECT * FROM ODASPOperationType WHERE OperationType = '" & frmODASMOperation.txtOperationType.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCONTROL
        
                If .EOF And .BOF Then Exit Sub
                     
                If rsCONTROL!QuotationApproval = "1" Or rsCONTROL!QuotationPreparation = "1" Or rsCONTROL!QuotationAuthorization = "1" Then
                        updateQUOTATION
                ElseIf rsCONTROL!JobBriefApproval = "1" Or rsCONTROL!JobBriefAuthorization = "1" Then
                        updateJobBrief
                ElseIf rsCONTROL!RenewalAPPROVAL = "1" Or rsCONTROL!RenewalAuthorization = "1" Then
                        updateJobBriefItem
                ElseIf rsCONTROL!siteApproval = "1" Or rsCONTROL!siteAuthorization = "1" Then
                        updateSITE
                        updateLEASEAGREEMENT
                        updatePLOTMAST
                        updatePLOTSITE
                ElseIf rsCONTROL!OpenJobCard = "1" Or rsCONTROL!CloseJobCard = "1" Then
                        updateJOBCARD
                ElseIf rsCONTROL!PurchaseOrderApproval = "1" Or rsCONTROL!PurchaseOrderAuthorization = "1" Then
                        updatePURCHASEORDER
                ElseIf rsCONTROL!SendNoticeApproval = "1" Or rsCONTROL!SendNoticeAuthorization = "1" Then
                        updateSENDNOTICE
                ElseIf rsCONTROL!ReceiveNoticeApproval = "1" Or rsCONTROL!ReceiveNoticeAuthorization = "1" Then
                        updateRECEIVENOTICE

                End If
        End With
        '/ Payment Type

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub updateQUOTATION()
On Error GoTo err
          
          Set rsSAVE = New Recordset
    
          strSQL = "SELECT * FROM AdvertQuotation where QuotationNo = '" & frmODASMOperation.txtApplicationNo & "';"
          rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

          With rsSAVE
                   
                  If rsCONTROL!QuotationApproval = "1" Then
                            !ApprovedBy = CurrentUserName
                            !DateApproved = Date
                            
                            If frmODASMOperation.chkAccept.Value = 1 Then
                                  !Approved = "Y"
                            Else
                                  !Approved = "N"
                            End If
                  ElseIf rsCONTROL!QuotationAuthorization = "1" Then
                            !AuthorizedBy = CurrentUserName
                            !DateAuthorized = Date
                            If frmODASMOperation.chkAccept.Value = 1 Then
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

Private Sub updateJOBCARD()
On Error GoTo err
          
          Set rsSAVE = New Recordset
    
          strSQL = "SELECT * FROM ODASMJobCard where JobCardNo = '" & CurrentRecord & "' and DepartmentCode = '" & GlobalDepartmentCode & "';"
          rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

          With rsSAVE
                   
                  If rsCONTROL!OpenJobCard = "1" Then
                            !openedBy = CurrentUserName
                            !DateOpened = Date
                            
                               If frmODASMOperation.chkAccept.Value = 1 Then
                                

                                  !opened = "Y"
                                  !Closed = "N"
                                  !Status = "OPENED"
                                  
                            Else
                                  !opened = "N"
                                  !Closed = "N"
                            End If
                            
                  ElseIf rsCONTROL!CloseJobCard = "1" Then
                            
                            !ClosedBy = CurrentUserName
                            !DateClosed = Date
                            If frmODASMOperation.chkAccept.Value = 1 Then
                                  !Closed = "Y"
                                  !Status = "CLOSED"
                            Else: !Closed = "N"
                            End If
                  End If

                  .Update
                  .Requery
          End With
    
    GlobalDepartmentCode = ""
    
Exit Sub

err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
            UpdateErrorMessage
    End If

End Sub
Private Sub updatePURCHASEORDER()
On Error GoTo err
          
          Set rsSAVE = New Recordset
    
          strSQL = "SELECT * FROM ODASMPurchaseOrder where OrderNo = '" & CurrentRecord & "';"
          rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

          With rsSAVE
                   
                  If rsCONTROL!PurchaseOrderApproval = "1" Then
                            !ApprovedBy = CurrentUserName
                            !DateApproved = Date
                            
                            If frmODASMOperation.chkAccept.Value = 1 Then
                                  !Approved = "Y"
                            Else
                                  !Approved = "N"
                            End If
                            
                  ElseIf rsCONTROL!PurchaseOrderAuthorization = "1" Then
                            
                            !AuthorizedBy = CurrentUserName
                            !DateAuthorized = Date
                            If frmODASMOperation.chkAccept.Value = 1 Then
                                  !Authorized = "Y"
                            Else: !Authorized = "N"
                            End If
                  End If

                  .Update
                  .Requery
          End With
    
    GlobalDepartmentCode = ""
    
Exit Sub

err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
            UpdateErrorMessage
    End If

End Sub

Private Sub updateSENDNOTICE()
On Error GoTo err
          
          Set rsSAVE = New Recordset
    
          strSQL = "SELECT * FROM ODASMLeaseAgreement where ContractNo = '" & CurrentRecord & "';"
          rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

          With rsSAVE
                   
                  If rsCONTROL!SendNoticeApproval = "1" Then
                            !NoticeApprovedBy = CurrentUserName
                            !NoticeApprovalDate = Date
                            
                            If frmODASMOperation.chkAccept.Value = 1 Then
                                  !NoticeApproved = "Y"
                            Else
                                  !NoticeApproved = "N"
                            End If
                            
                  ElseIf rsCONTROL!SendNoticeAuthorization = "1" Then
                            
                            !NoticeAuthorizedBy = CurrentUserName
                            !NoticeAuthorizationDate = Date
                            
                            If frmODASMOperation.chkAccept.Value = 1 Then
                                  !NoticeAuthorized = "Y"
                            Else: !NoticeAuthorized = "N"
                            End If
                  End If

                  .Update
                  .Requery
          End With
    
    GlobalDepartmentCode = ""
    
Exit Sub

err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
            UpdateErrorMessage
    End If

End Sub
Private Sub updateRECEIVENOTICE()
On Error GoTo err
          
          Set rsSAVE = New Recordset
    
          strSQL = "SELECT * FROM ODASMLeaseAgreement where ContractNo = '" & CurrentRecord & "';"
          rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

          With rsSAVE
                   
                  If rsCONTROL!ReceiveNoticeApproval = "1" Then
                            !NoticeApprovedBy = CurrentUserName
                            !NoticeApprovalDate = Date
                            
                            If frmODASMOperation.chkAccept.Value = 1 Then
                                  !NoticeApproved = "Y"
                            Else
                                  !NoticeApproved = "N"
                            End If
                            
                  ElseIf rsCONTROL!ReceiveNoticeAuthorization = "1" Then
                            
                            !NoticeAuthorizedBy = CurrentUserName
                            !NoticeAuthorizationDate = Date
                            
                            If frmODASMOperation.chkAccept.Value = 1 Then
                                  !NoticeAuthorization = "Y"
                            Else: !NoticeAuthorization = "N"
                            End If
                  End If

                  .Update
                  .Requery
          End With
    
    GlobalDepartmentCode = ""
    
Exit Sub

err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
            UpdateErrorMessage
    End If

End Sub
Private Sub updateSITE()
On Error GoTo err
          
          Set rsSAVE = New Recordset
        
          strSQL = "SELECT * FROM ODASPPlotMast where ContractNo = '" & frmODASMOperation.txtApplicationNo & "';"
          
          rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

          With rsSAVE
                   
                  If rsCONTROL!siteApproval = "1" Then
                            !ApprovedBy = CurrentUserName
                            !DateApproved = Date
                            
                            If frmODASMOperation.chkAccept.Value = 1 Then
                                  !Approved = "Y"
                                  !ContractStatus = "APPROVED"
                                  !ContractStatusDate = Date
                            Else
                                  !Approved = "N"
                                  !ContractStatus = "NOT-APPROVED"
                                  !ContractStatusDate = Date

                            End If
                  ElseIf rsCONTROL!siteAuthorization = "1" Then
                            !AuthorizedBy = CurrentUserName
                            !DateAuthorized = Date
                            
                            If frmODASMOperation.chkAccept.Value = 1 Then
                                    !Authorized = "Y"
                                    !Status = "SITE-AVAILABLE"
                                    !RentDueDate = CDbl(!CommencementDate)
                                    !RentDue = CDbl(!AnnualRent)
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

Private Sub updatePLOTSITE()
On Error GoTo err
        frmODASMOperation.ProgressBar2.Visible = True
        
          Dim rsPLOT As ADODB.Recordset, strPLOT As String
          Set rsPLOT = New ADODB.Recordset
        
          strPLOT = "SELECT * FROM ODASPPlotMast where ContractNo = '" & frmODASMOperation.txtApplicationNo & "' ;"
          rsPLOT.Open strPLOT, cnCOMMON, adOpenKeyset, adLockOptimistic
                   
          If rsPLOT.EOF And rsPLOT.BOF Then Exit Sub
          
            
            rsPLOT.MoveFirst
            
            Do While Not rsPLOT.EOF
                  Set rsSAVE = New Recordset
                  strSQL = "SELECT * FROM ODASPPlotSite where MastNo='" & rsPLOT!MastNo & "';"
                  rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
                  With rsSAVE
                  
                    .MoveFirst
                          Do While Not .EOF
                              If rsCONTROL!siteApproval = "1" Then
                                        !ApprovedBy = CurrentUserName
                                        !DateApproved = Date
                                        
                                        If frmODASMOperation.chkAccept.Value = 1 Then
                                              !Approved = "Y"
                                        Else
                                              !Approved = "N"
                                        End If
                              ElseIf rsCONTROL!siteAuthorization = "1" Then
                                        !AuthorizedBy = CurrentUserName
                                        !DateAuthorized = Date
                                        
                                        If frmODASMOperation.chkAccept.Value = 1 Then
                                                !Authorized = "Y"
                                                !Status = "SITE-AVAILABLE"
                                                
                                                
                                                Dim SDate, EDate As Date
                                                        SDate = rsPLOT!CommencementDate
                                                        EDate = DateAdd("yyyy", rsPLOT!LeaseDuration, rsPLOT!CommencementDate)
                                                
                                                        frmODASMOperation.ProgressBar2.Max = DateDiff("d", SDate, EDate) + 1
                                                        frmODASMOperation.ProgressBar2.Min = 0
                                                        
                                                        Do While SDate <= EDate
                                                            Set rsSiteSchedule = New ADODB.Recordset
                                                            rsSiteSchedule.Open "SELECT * FROM ODASMSiteSchedule WHERE SiteNo = '" & rsSAVE!SiteNo & "' and ScheduleDate = '" & Format(SDate, "MMMM dd,yyyy") & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                                                            If rsSiteSchedule.RecordCount = 0 Then
                                                                rsSiteSchedule.AddNew
                                                                rsSiteSchedule!SiteNo = rsSAVE!SiteNo
                                                                rsSiteSchedule!ScheduleDate = SDate
                                                                rsSiteSchedule!Reserved = "N"
                                                                rsSiteSchedule!Allocated = "N"
                                                            End If
                                                            rsSiteSchedule.Update
                                                            SDate = DateAdd("d", 1, SDate)
                                                           frmODASMOperation.ProgressBar2.Value = frmODASMOperation.ProgressBar2.Value + 1
                                                        Loop
                                        Else: !Authorized = "N"
                                        
                                        End If
                              End If

                        .Update
                        frmODASMOperation.ProgressBar2.Value = 0
                       .MoveNext
                    Loop
                    End With
                rsPLOT.MoveNext
              Loop
            frmODASMOperation.ProgressBar2.Visible = False
       Exit Sub

err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
            UpdateErrorMessage
    End If

End Sub
Private Sub updatePLOTMAST()
On Error GoTo err
          
          Dim rsPLOT As ADODB.Recordset, strPLOT As String
          Set rsPLOT = New ADODB.Recordset
        
          strPLOT = "SELECT * FROM ODASPPlotMast where ODASPPlotMast.ContractNo = '" & frmODASMOperation.txtApplicationNo & "' ;"
          rsPLOT.Open strPLOT, cnCOMMON, adOpenKeyset, adLockOptimistic
                   
          If rsPLOT.EOF Or rsPLOT.BOF Then Exit Sub
          
          rsPLOT.MoveFirst
          Do While Not rsPLOT.EOF

                  Set rsSAVE = New Recordset
                  strSQL = "SELECT * FROM ODASPPlotMast where MastNo = '" & rsPLOT!MastNo & "' ;"
                  rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
'
                  With rsPLOT
                           
                              If rsCONTROL!siteApproval = "1" Then
                                        !ApprovedBy = CurrentUserName
                                        !DateApproved = Date
                                        
                                        If frmODASMOperation.chkAccept.Value = 1 Then
                                              !Approved = "Y"
                                        Else
                                              !Approved = "N"
                                        End If
                              ElseIf rsCONTROL!siteAuthorization = "1" Then
                                        !AuthorizedBy = CurrentUserName
                                        !DateAuthorized = Date
                                        
                                        If frmODASMOperation.chkAccept.Value = 1 Then
                                                !Authorized = "Y"
                                                !Status = "SITE-AVAILABLE"
                                                !RentDueDate = CDbl(!CommencementDate)
                                                !RentDue = CDbl(!AnnualRent)
                                        Else: !Authorized = "N"
                                        
                                        End If
                              End If

                        .Update
                    End With
                    
                    rsPLOT.MoveNext
          Loop
Exit Sub

err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
            UpdateErrorMessage
    End If

End Sub
Private Sub updateLEASEAGREEMENT()
On Error GoTo err
          
          Set rsSAVE = New Recordset
        
          strSQL = "SELECT * FROM ODASMLeaseAgreement where ContractNo = '" & frmODASMOperation.txtApplicationNo & "';"
          
          rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

          With rsSAVE
                   
                  If rsCONTROL!siteApproval = "1" Then
                            !ApprovedBy = CurrentUserName
                            !DateApproved = Date
                            !WitnessCoy = CurrentUserName
                            If frmODASMOperation.chkAccept.Value = 1 Then
                                  !Approved = "Y"
                            Else
                                  !Approved = "N"

                            End If
                  ElseIf rsCONTROL!siteAuthorization = "1" Then
                            !AuthorizedBy = CurrentUserName
                            !SignedBy = CurrentUserName
                            !DateAuthorized = Date
                            !Status = ""
                            If frmODASMOperation.chkAccept.Value = 1 Then
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
Private Sub updateJobBriefItem()
On Error GoTo err

      Set rsSAVE = New Recordset

      strSQL = "SELECT * FROM ODASMJobBriefItems where JobBriefItemNo = '" & frmODASMOperation.txtApplicationNo & "';"
      rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

      With rsSAVE
              
              If rsCONTROL!RenewalAPPROVAL = "1" Then
                      !NoticeApproved = "Y"
                      !NoticeApprovedBy = CurrentUserName
                      !Status = "NOTICE-APPROVED"
              ElseIf rsCONTROL!RenewalAuthorization = "1" Then
                      !NoticeAuthorizedBy = CurrentUserName
                      !NoticeAuthorized = "Y"
                      !Status = "NOTICE-AUTHORIZED"
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
Private Sub updateJobBrief()
On Error GoTo err

      Set rsSAVE = New Recordset

      strSQL = "SELECT * FROM ODASMJobBrief where JobBriefNo = '" & frmODASMOperation.txtApplicationNo & "';"
      rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

      With rsSAVE
              
              If rsCONTROL!JobBriefApproval = "1" Then
                      !ApprovedBy = CurrentUserName
                      !DateApproved = Date
                      
                      If frmODASMOperation.chkAccept.Value = 1 Then
                            !Approved = "Y"
                      Else: !Approved = "N"
                      End If
                  
              ElseIf rsCONTROL!JobBriefAuthorization = "1" Then
                      !AuthorizedBy = CurrentUserName
                      !DateAuthorized = Date
                      
                      If frmODASMOperation.chkAccept.Value = 1 Then
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
Private Sub ShowCode()

On Error GoTo err
    Set RsCode = New Recordset
    strcode = "SELECT * from ODASMOperation where ApplicationNo = '" & GlobalApplicationNo & "' AND operationType = '" & GlobalOperationType & "';"

    RsCode.Open strcode, cnCOMMON, adOpenKeyset, adLockOptimistic
  
    With RsCode
            Screen.ActiveForm.txtApplicationNo.Text = !ApplicationNo
            Screen.ActiveForm.txtOperationType.Text = !OperationType
            Screen.ActiveForm.txtOperationDate = !operationDate
            Screen.ActiveForm.txtStatus = !Status
            Screen.ActiveForm.txtComment = !Comment
            Screen.ActiveForm.txtAccept = !Accept
            cboUserCode = !UserName
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
'        disableButtons
        
        With frmODASMOperation
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
        
        strCHECK = "SELECT * from ODASMOperation where ApplicationNo = '" & GlobalApplicationNo & "' AND operationType = '" & GlobalOperationType & "';"
        rsCHECK.Open strCHECK, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsCHECK
                
                If .EOF Or .BOF Then Exit Sub
                
                If !Accept = "Y" Then
                        disableControls
                        frmODASMOperation.cmdUpdate = False
                        frmODASMOperation.cmdADDNEW = False
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
        disableControls
End Sub


Private Sub ValidateData()

On Error GoTo err

    bSaveRECORD = False
    
    With frmODASMOperation
    
              If .txtApplicationNo.Text = "" Then
                      MsgBox "ApplicationNo is Required"
                      .txtApplicationNo.SetFocus
              
              ElseIf .cboUserCode.Text = "" Then
                      MsgBox "The UserName is  required"
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
        Dim rsFIND As ADODB.Recordset, edit As Boolean

        Set rsFIND = New ADODB.Recordset

        Select Case cmdEdit.Caption
                Case "&Edit"
                        EnableControls

                        strQRE = InputBox("Enter Application No  to search.", "Search Value")
    
                        rsFIND.Open "SELECT * FROM ODASMOperation WHERE ApplicationNo = '" & frmODASMOperation.txtApplicationNo & "' and  operationtype = '" & frmODASMOperation.txtOperationType & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic

                        With rsFIND
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
                                        
                                        Screen.ActiveForm.cboUserCode = !UserName
                                        Screen.ActiveForm.txtOperationDescription = GlobalOperationDescription
                                        edit = True
                                End If
                        End With
                        
                        
                        If edit Then
                                cmdEdit.Caption = "Save &Changes"
                        End If
    
                Case "Save &Changes"
                        Dim rsFinder As ADODB.Recordset
                        Set rsFinder = New ADODB.Recordset

                        rsFinder.Open "SELECT * FROM ODASMOperation WHERE ApplicationNo = '" & frmODASMOperation.txtApplicationNo & "' and  operationtype = '" & frmODASMOperation.txtOperationType & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                        With rsFinder
                            !ApplicationNo = Screen.ActiveForm.txtApplicationNo
                            !UserName = Screen.ActiveForm.cboUserCode
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
                            edit = False
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
                rsFIND.Requery

            If rsFIND.BOF Then Exit Sub
                rsFIND.MoveFirst

    ElseIf err.Number = 3021 Then
            MsgBox "Requested record not found! Refresh the database and try the search again...or Check your entries.", vbInformation, "Searching."
                rsFIND.Requery

            If rsFIND.BOF Then Exit Sub
                rsFIND.MoveFirst
    Else
                ErrorMessage
End If

End Sub
Private Sub cmdUpdate_Click()
        ValidateData
        CheckRightsStatus
        If bSaveRECORD = True Then
                saveRecord
                updateRECORD
                bSaveRECORD = False
                listAPPROVALTASKS
                disableControls
        End If
End Sub

Private Sub Form_Activate()
    disableControls
    loadRECORD
    checkSTATUS
    listAPPROVALTASKS
End Sub
  
Private Sub loadRECORD()
On Error GoTo err:

        With frmODASMOperation
                '.txtApplicationNo.Text = GlobalApplicationNo
                .txtApplicationNo.Text = CurrentRecord
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
Private Sub CheckRightsStatus()
On Error GoTo err
        bSaveRECORD = False
        If Trim(frmODASMOperation.txtPassword.Text) = "" Then Exit Sub
        
        Screen.ActiveForm.txtPassword.Text = GetFullEncryption
        Dim rsPASSWORD As ADODB.Recordset
        Set rsPASSWORD = New Recordset
        
        Set rsFindRecord = New ADODB.Recordset
        rsFindRecord.Open "Select * From UserMaster Where userName = '" & cboUserCode & "' ", cnSECURE, adOpenKeyset, adLockOptimistic
        
        If rsFindRecord.EOF And rsFindRecord.BOF Then
            MsgBox "This User is Not allowed to Check the Transaction", vbOKOnly
            Screen.ActiveForm.txtPassword.SetFocus
            Screen.ActiveForm.txtPassword.Text = ""
            Screen.ActiveForm.cmdUpdate.Enabled = False
            Exit Sub
        Else
        
        strSQL = "SELECT (A.Password) as APW,A.* FROM ODASPApprovers A WHERE  A.StaffId = '" & rsFindRecord!StaffId & "' and A.operationType = '" & frmODASMOperation.txtOperationType & " ';"
        rsPASSWORD.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsPASSWORD
                If .EOF Or .BOF Then
                        MsgBox "This User is Not allowed to Check the Transaction", vbOKOnly
                        Screen.ActiveForm.txtPassword.SetFocus
                        Screen.ActiveForm.txtPassword.Text = ""
                        Screen.ActiveForm.cmdUpdate.Enabled = False
                        Exit Sub
                Else
                        bSaveRECORD = False
                        If Trim(Screen.ActiveForm.txtPassword.Text) = "" Or Trim(!APW) = "" Then
                                MsgBox "Passwords are Required", vbOKOnly
                                Screen.ActiveForm.txtPassword.SetFocus
                                
                        ElseIf Trim(Screen.ActiveForm.txtPassword.Text) <> Trim(!APW) Then
                                MsgBox "The Password Entered is invalid", vbOKOnly
                                Me.txtPassword.Text = ""
                                Me.txtPassword.SetFocus
                                
                        ElseIf Trim(Screen.ActiveForm.txtPassword.Text) = Trim(!APW) Then
                            If Trim(chkAccept.Value) = 0 And Screen.ActiveForm.txtComment.Text <= "" Then
                                
                                MsgBox " The Reason for Rejecting the Application is very Important"
                                Screen.ActiveForm.txtComment.SetFocus
                                Screen.ActiveForm.cmdUpdate.Enabled = False
                                Exit Sub
                        Else
                                bSaveRECORD = True
'                                Screen.ActiveForm.cmdUpdate.Enabled = True
'                                Screen.ActiveForm.cmdUpdate.SetFocus
                        End If
                            
                        End If
                End If
        
        End With
        End If
Screen.ActiveForm.cmdUpdate.SetFocus
Exit Sub


err:
        ErrorMessage

End Sub


Private Sub txtPassword_lostFocus()
On Error GoTo err

    Screen.ActiveForm.cmdUpdate.Enabled = True
Exit Sub

err:
        ErrorMessage
End Sub




