VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmODASPApprovalType 
   Caption         =   "Operation Type"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   9225
   Icon            =   "frmODASPApprovalType.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmODASPApprovalType.frx":0442
   ScaleHeight     =   6600
   ScaleWidth      =   9225
   Begin VB.Frame Frame12 
      Height          =   6495
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   9015
      Begin VB.TextBox txtOperationType 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtOperationDescription 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   360
         Width           =   6135
      End
      Begin VB.Frame Frame4 
         Caption         =   "Operations"
         Height          =   2415
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   8775
         Begin VB.OptionButton optJobCardApproval 
            Height          =   255
            Left            =   3600
            TabIndex        =   25
            Top             =   1680
            Width           =   255
         End
         Begin VB.OptionButton optJobCardAuthorization 
            Height          =   255
            Left            =   6120
            TabIndex        =   24
            Top             =   1680
            Width           =   255
         End
         Begin VB.OptionButton optJobCardPreparation 
            Height          =   255
            Left            =   1560
            TabIndex        =   23
            Top             =   1680
            Width           =   255
         End
         Begin VB.OptionButton optJobBriefApproval 
            Height          =   255
            Left            =   3600
            TabIndex        =   22
            Top             =   1320
            Width           =   255
         End
         Begin VB.OptionButton optJobBriefAuthorization 
            Height          =   255
            Left            =   6120
            TabIndex        =   21
            Top             =   1320
            Width           =   255
         End
         Begin VB.OptionButton optJobBriefPreparation 
            Height          =   255
            Left            =   1560
            TabIndex        =   20
            Top             =   1320
            Width           =   255
         End
         Begin VB.OptionButton optQuotationChecked 
            Height          =   255
            Left            =   1560
            TabIndex        =   19
            Top             =   960
            Width           =   255
         End
         Begin VB.OptionButton optQuotationAuthorization 
            Height          =   255
            Left            =   6120
            TabIndex        =   18
            Top             =   960
            Width           =   255
         End
         Begin VB.OptionButton optQuotationApproval 
            Height          =   255
            Left            =   3600
            TabIndex        =   17
            Top             =   960
            Width           =   255
         End
         Begin VB.OptionButton optInvoiceApproval 
            Height          =   255
            Left            =   3600
            TabIndex        =   16
            Top             =   2040
            Width           =   255
         End
         Begin VB.OptionButton optAuthorization 
            Height          =   255
            Left            =   6120
            TabIndex        =   15
            Top             =   2040
            Width           =   255
         End
         Begin VB.OptionButton optInvoicePreparation 
            Height          =   255
            Left            =   1560
            TabIndex        =   14
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "Job Card"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Quotation"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Job Brief"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Preparation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   29
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Approval"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3240
            TabIndex        =   28
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Authorization"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5880
            TabIndex        =   27
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Invoice"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   2040
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   3015
         Left            =   120
         TabIndex        =   11
         Top             =   3360
         Width           =   7575
         Begin MSComctlLib.ListView ListView1 
            Height          =   2655
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   4683
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   7800
         TabIndex        =   4
         Top             =   3360
         Width           =   1095
         Begin VB.CommandButton cmdAddNew 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPApprovalType.frx":0784
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdUpdate 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPApprovalType.frx":0886
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   990
            Width           =   855
         End
         Begin VB.CommandButton cmdSearch 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPApprovalType.frx":0988
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1365
            Width           =   855
         End
         Begin VB.CommandButton cmdDelete 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPApprovalType.frx":0A8A
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1740
            Width           =   855
         End
         Begin VB.CommandButton cmdCancel 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPApprovalType.frx":0B8C
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   2115
            Width           =   855
         End
         Begin VB.CommandButton cmdPrint 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Picture         =   "frmODASPApprovalType.frx":0C8E
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   2520
            Width           =   855
         End
      End
      Begin VB.Label lblRelationshipCode 
         Caption         =   "Operation Type"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   435
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmODASPApprovalType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSAVE As ADODB.Recordset, strRELN As String
Sub loadGRID()

        Set OperationTypeGrid.DataSource = rsSAVE
End Sub

Sub clearRELN()
        With frmODASPOperationType
            .txtOperationType.Text = ""
            .txtDescription.Text = ""
            .optChecked.Value = 0
            .optApproved.Value = 0
            .optAuthorized.Value = 0
            .optDischargeApproval.Value = 0
            .optDischargeAuthorization.Value = 0
            .optDischargePreparation.Value = 0
            .optPaymentApproval.Value = 0
            .optPaymentAuthorization.Value = 0
            .optPaymentPreparation.Value = 0
            .optchequeApproval.Value = 0
            .optChequeAuthorization.Value = 0
            .optChequeIssuance.Value = 0
            .optChequePreparation.Value = 0
            .optClaimRegApproval.Value = 0
            .optClaimRegAuthorization.Value = 0
            .optClaimRegPreparation.Value = 0
            .optReinstatementApproval.Value = 0
            .optReinstatementAuthorization.Value = 0
            .optReinstatementPreparation.Value = 0
        End With
End Sub

Sub enableRELN()
        With frmODASPOperationType
            .txtOperationType.Locked = False
            .txtDescription.Locked = False
            .optChecked.Enabled = True
            .optApproved.Enabled = True
            .optAuthorized.Enabled = True
            .optDischargeApproval.Enabled = True
            .optDischargeAuthorization.Enabled = True
            .optDischargePreparation.Enabled = True
            .optPaymentApproval.Enabled = True
            .optPaymentAuthorization.Enabled = True
            .optPaymentPreparation.Enabled = True
            .optchequeApproval.Enabled = True
            .optChequeAuthorization.Enabled = True
            .optPaymentPreparation.Enabled = True
            .optchequeApproval.Enabled = True
            .optChequeAuthorization.Enabled = True
            .optChequeIssuance.Enabled = True
            .optChequePreparation.Enabled = True
            .optClaimRegApproval.Enabled = True
            .optClaimRegAuthorization.Enabled = True
            .optClaimRegPreparation.Enabled = True
            .optReinstatementApproval.Enabled = True
            .optReinstatementAuthorization.Enabled = True
            .optReinstatementPreparation.Enabled = True


        End With
End Sub

Sub disableRELN()
        With frmODASPOperationType
                .txtOperationType.Locked = True
                .txtDescription.Locked = True
                .optChecked.Enabled = False
                .optApproved.Enabled = False
                .optAuthorized.Enabled = False
                .optDischargeApproval.Enabled = False
                .optDischargeAuthorization.Enabled = False
                .optDischargePreparation.Enabled = False
                .optPaymentApproval.Enabled = False
                .optPaymentAuthorization.Enabled = False
                .optPaymentPreparation.Enabled = False
                .optchequeApproval.Enabled = False
                .optChequeAuthorization.Enabled = False
                .optChequeIssuance.Enabled = False
                .optChequePreparation.Enabled = False
                .optClaimRegApproval.Enabled = False
                .optClaimRegAuthorization.Enabled = False
                .optClaimRegPreparation.Enabled = False
                
                .optReinstatementApproval.Enabled = False
                .optReinstatementAuthorization.Enabled = False
                .optReinstatementPreparation.Enabled = False

        End With
End Sub

Sub showRELN()
    With rsSAVE
        frmODASPOperationType.txtOperationType = !OperationType
        frmODASPOperationType.txtDescription = !Description
                            
        If !MedicalApproval = True Then
                frmODASPOperationType.optMedicalApproval.Value = True
        Else: frmODASPOperationType.optMedicalApproval.Value = False
        End If
        
        If !MedicalPreparation = True Then
                frmODASPOperationType.optMedicalPreparation.Value = True
        Else: frmODASPOperationType.optMedicalPreparation.Value = False
        End If
        
        If !MedicalAuthorization = True Then
                frmODASPOperationType.optMedicalAuthorization.Value = True
        Else: frmODASPOperationType.optMedicalAuthorization.Value = False
        End If
    
        
        If !ProposalApproval = True Then
                frmODASPOperationType.optProposalApproval.Value = True
        Else: frmODASPOperationType.optProposalApproval.Value = False
        End If
        
        If !ProposalPreparation = True Then
                frmODASPOperationType.optProposalPreparation.Value = True
        Else: frmODASPOperationType.optProposalPreparation.Value = False
        End If
        
        If !ProposalAuthorization = True Then
                frmODASPOperationType.optProposalAuthorization.Value = True
        Else: frmODASPOperationType.optProposalAuthorization.Value = False
        End If
    
        
        
        If !PolicyApproval = True Then
                frmODASPOperationType.optPolicyApproval.Value = True
        Else: frmODASPOperationType.optPolicyApproval.Value = False
        End If
        
        If !PolicyPreparation = True Then
                frmODASPOperationType.optPolicyPreparation.Value = True
        Else: frmODASPOperationType.optPolicyPreparation.Value = False
        End If
        
        If !PolicyAuthorization = True Then
                frmODASPOperationType.optPolicyAuthorization.Value = True
        Else: frmODASPOperationType.optPolicyAuthorization.Value = False
        End If
    
        
        
        
        If !paidupApproval = True Then
                frmODASPOperationType.optPaidupApproval.Value = True
        Else: frmODASPOperationType.optPaidupApproval.Value = False
        End If
        
        If !paidupPreparation = True Then
                frmODASPOperationType.optPaidupPreparation.Value = True
        Else: frmODASPOperationType.optPaidupPreparation.Value = False
        End If
        
        If !paidupAuthorization = True Then
                frmODASPOperationType.optPaidupAuthorization.Value = True
        Else: frmODASPOperationType.optPaidupAuthorization.Value = False
        End If
        
        If !ReinstatementApproval = True Then
                frmODASPOperationType.optReinstatementApproval.Value = True
        Else: frmODASPOperationType.optReinstatementApproval.Value = False
        End If
        
        If !ReinstatementPreparation = True Then
                frmODASPOperationType.optReinstatementPreparation.Value = True
        Else: frmODASPOperationType.optReinstatementPreparation.Value = False
        End If
        
        If !ReinstatementAuthorization = True Then
                frmODASPOperationType.optReinstatementAuthorization.Value = True
        Else: frmODASPOperationType.optReinstatementAuthorization.Value = False
        End If

        
        
        If !RegistrationPreparation = True Then
              frmODASPOperationType.optClaimRegPreparation.Value = 1
        Else: frmODASPOperationType.optClaimRegPreparation.Value = 0
        End If
            
        If !RegistrationApproval = True Then
              frmODASPOperationType.optClaimRegApproval.Value = 1
        Else: frmODASPOperationType.optClaimRegApproval.Value = 0
        End If
        
        If !RegistrationAuthorization = True Then
              frmODASPOperationType.optClaimRegAuthorization.Value = 1
        Else: frmODASPOperationType.optChequeAuthorization.Value = 0
        End If
       
        If !Checked = True Then
              frmODASPOperationType.optChecked.Value = 1
        Else: frmODASPOperationType.optChecked.Value = 0
        End If
            
        If !Approved = True Then
              frmODASPOperationType.optApproved.Value = 1
        Else: frmODASPOperationType.optApproved.Value = 0
        End If
        
        If !Authorized = True Then
              frmODASPOperationType.optAuthorized.Value = 1
        Else: frmODASPOperationType.optAuthorized.Value = 0
        End If
            
        If !DischargeApproval = True Then
                frmODASPOperationType.optDischargeApproval.Value = 1
        Else: frmODASPOperationType.optDischargeApproval.Value = 0
        End If
        
        If !DischargePreparation = True Then
                frmODASPOperationType.optDischargePreparation.Value = 1
        Else: frmODASPOperationType.optDischargePreparation.Value = 0
        End If
        
        If !DischargeAuthorization = True Then
                frmODASPOperationType.optDischargeAuthorization.Value = 1
        Else: frmODASPOperationType.optDischargeAuthorization.Value = 0
        End If

        If !PaymentApproval = True Then
                frmODASPOperationType.optPaymentApproval.Value = 1
        Else: frmODASPOperationType.optPaymentApproval.Value = 0
        End If
        
        If !PaymentPreparation = True Then
                frmODASPOperationType.optPaymentPreparation.Value = 1
        Else: frmODASPOperationType.optPaymentPreparation.Value = 0
        End If
        
        If !PaymentAuthorization = True Then
                frmODASPOperationType.optPaymentAuthorization.Value = 1
        Else: frmODASPOperationType.optPaymentAuthorization.Value = 0
        End If
        
        If !ChequeApproval = True Then
                frmODASPOperationType.optchequeApproval.Value = 1
        Else: frmODASPOperationType.optchequeApproval.Value = 0
        End If
        
        If !ChequePreparation = True Then
                frmODASPOperationType.optChequePreparation.Value = 1
        Else: frmODASPOperationType.optChequePreparation.Value = 0
        End If
        
        If !ChequeAuthorization = True Then
                frmODASPOperationType.optChequeAuthorization.Value = 1
        Else: frmODASPOperationType.optChequeAuthorization.Value = 0
        End If
        
        If !ChequeIssuance = True Then
                frmODASPOperationType.optChequeIssuance.Value = 1
        Else: frmODASPOperationType.optChequeIssuance.Value = 0
        End If


    End With
End Sub

Private Sub DisableCButtons()
        With frmODASPOperationType
            .cmdUpdate.Enabled = True
            .cmdAdd.Enabled = False
            .cmdSearch.Enabled = False
            .cmdEdit.Enabled = False
            .cmdDelete.Enabled = False
            .cmdCancel.Enabled = True
        End With
End Sub

Private Sub enableCButtons()
    With frmODASPOperationType
            .cmdUpdate.Enabled = False
            .cmdAdd.Enabled = True
            .cmdSearch.Enabled = True
            .cmdEdit.Enabled = True
            .cmdDelete.Enabled = True
            .cmdCancel.Enabled = True
    End With
End Sub

Private Sub cmdAddNew_Click()
        clearALLRECORD
        enableALLRECORD
        disableButtons
End Sub


Private Sub cmdCancel_Click()
        enableButtons
        clearRELN
        disableRELN
End Sub


Private Sub cmdDelete_Click()
On Error GoTo err

If txtOperationType.Text = "" Then
            MsgBox "There is no current record to delete", vbInformation, "Delete Information"
Else
        If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            With rsSAVE
                
                If .EOF And .BOF Then Exit Sub
                .Delete
                .Requery
                clearRELN
            End With
    End If
        '/* End if Msg Box
        
End If
        '/* If txt = ""
        
Exit Sub

err:
    ErrorMessage

End Sub

Private Sub cmdEdit_Click()
        editMYRECORD
End Sub

Private Sub ValidateRECORD()
On Error GoTo err

        bSaveRECORD = False
        
        If Screen.ActiveForm.txtOperationType.Text = "" Then
                MsgBox "The Operation Type MUST be Entered"
                Screen.ActiveForm.txtOperationType.SetFocus
        ElseIf Screen.ActiveForm.txtOperationType.Text <= "" Then
                MsgBox "The Description of the Operation cannot be Left Blank"
                txtOperationType.SetFocus
        Else
                bSaveRECORD = True
        End If
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub SaveRECORD()
On Error GoTo err
    Set rsSAVE = New ADODB.Recordset
    
    strSQL = "Select * from ODASPOperationType Where OperationType = '" & frmODAPApprovalType.txtOperationType.Text & "'"
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
   With rsSAVE
        If .BOF Or .EOF Then
                .AddNew
                !OperationType = frmODASPOperationType.txtOperationType
                !PreparedBy = CurrentUserName
                !DatePrepared = Date
        End If
        
        !Description = frmODASPOperationType.txtDescription
        
        If frmODASPOperationType.optQuotationPreparation = True Then
                !QuotationPreparation = 1
            Else: !QuotationPreparation = 0
        End If
        
        If frmODASPOperationType.optQuotationApproval = True Then
                !QuotationApproval = 1
        Else: !QuotationApproval = 0
        End If
        
        If frmODASPOperationType.optQuotationAuthorization = True Then
                !QuotationAuthorization = 1
        Else: !QuotationAuthorization = 0
        End If

        If frmODASPOperationType.optJobBriefPreparation = True Then
                !JobBriefPreparation = 1
            Else: !JobBriefPreparation = 0
        End If
        
        If frmODASPOperationType.optJobBriefApproval = True Then
                !JobBriefApproval = 1
        Else: !JobBriefApproval = 0
        End If
        
        If frmODASPOperationType.optJobBriefAuthorization = True Then
                !JobBriefAuthorization = 1
        Else: !JobBriefAuthorization = 0
        End If

        If frmODASPOperationType.optJobCardPreparation = True Then
                !JobCardPreparation = 1
            Else: !JobCardPreparation = 0
        End If
        
        If frmODASPOperationType.optJobCardApproval = True Then
                !JobCardApproval = 1
        Else: !JobCardApproval = 0
        End If
        
        If frmODASPOperationType.optJobCardAuthorization = True Then
                !JobCardAuthorization = 1
        Else: !JobCardAuthorization = 0
        End If

        If frmODASPOperationType.optInvoicePreparation = True Then
                !InvoicePreparation = 1
            Else: !InvoicePreparation = 0
        End If
        
        If frmODASPOperationType.optInvoiceApproval = True Then
                !InvoiceApproval = 1
        Else: !InvoiceApproval = 0
        End If
        
        If frmODASPOperationType.optInvoiceAuthorization = True Then
                !InvoiceAuthorization = 1
        Else: !InvoiceAuthorization = 0
        End If

        bSaveRECORD = False
        
         .Update
         .Requery
  End With
Exit Sub

err:
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            MsgBox "Update Cancelled! You cannot save this record because some required fields are blank or have invalid data!", vbCritical, "Cancel Update"
            rsSAVE.CancelUpdate
            rsSAVE.Requery
    Else
        UpdateErrorMessage
    End If

End Sub


Private Sub cmdUpdate_Click()
        bSaveRECORD = True
        ValidateRECORD
        If bSaveRECORD = True Then
            SaveRECORD
                If bSaveRECORD = False Then
                    enableButtons
                    disableALLRECORD
                End If
        End If

        
End Sub

Private Sub cmdSearch_Click()
        searchMyRecord
End Sub

Private Sub Form_Activate()
    disableALLRECORD
    enableButtons
    getALLOPERATIONS
End Sub

Private Sub Form_Load()

    OpenODBCConnection
      
End Sub



