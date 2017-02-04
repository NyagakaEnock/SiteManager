VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmALISPPaymentType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Types"
   ClientHeight    =   6645
   ClientLeft      =   150
   ClientTop       =   1155
   ClientWidth     =   11325
   Icon            =   "frmALISPPaymentType.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   11325
   Begin VB.Frame Frame12 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11055
      Begin VB.Frame fraCButtons 
         Height          =   3375
         Index           =   6
         Left            =   9720
         TabIndex        =   21
         Top             =   2880
         Width           =   1215
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
            Height          =   495
            Left            =   120
            Picture         =   "frmALISPPaymentType.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   2715
            Width           =   975
         End
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
            Height          =   495
            Left            =   120
            Picture         =   "frmALISPPaymentType.frx":0544
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   2220
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
            Height          =   495
            Left            =   120
            Picture         =   "frmALISPPaymentType.frx":0646
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdEdit 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Picture         =   "frmALISPPaymentType.frx":0748
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   1725
            Width           =   975
         End
         Begin VB.CommandButton cmdAddNew 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Picture         =   "frmALISPPaymentType.frx":084A
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   240
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
            Height          =   495
            Left            =   120
            Picture         =   "frmALISPPaymentType.frx":094C
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   1200
            Width           =   975
         End
      End
      Begin VB.Frame frabrowse 
         Height          =   735
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   5520
         Width           =   9495
         Begin VB.CommandButton cmdFirstCode 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   0
            Left            =   1200
            Picture         =   "frmALISPPaymentType.frx":0A4E
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdFirstCode 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   1
            Left            =   2640
            Picture         =   "frmALISPPaymentType.frx":0E90
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdFirstCode 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   2
            Left            =   4080
            Picture         =   "frmALISPPaymentType.frx":12D2
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdFirstCode 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   3
            Left            =   5400
            Picture         =   "frmALISPPaymentType.frx":1714
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame14 
         Height          =   2655
         Left            =   120
         TabIndex        =   13
         Top             =   2880
         Width           =   9495
         Begin MSDataGridLib.DataGrid DataGrid 
            Height          =   2295
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4048
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777152
            HeadLines       =   1
            RowHeight       =   19
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame13 
         Height          =   1815
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   10815
         Begin VB.ComboBox cboCreditAC 
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
            Height          =   360
            Left            =   5760
            TabIndex        =   4
            Top             =   1200
            Width           =   2415
         End
         Begin VB.ComboBox cboDebitAC 
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
            Height          =   360
            Left            =   1920
            TabIndex        =   3
            Top             =   1200
            Width           =   2535
         End
         Begin VB.TextBox txtDescription 
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
            Left            =   1920
            TabIndex        =   2
            Top             =   697
            Width           =   6255
         End
         Begin VB.TextBox txtPaymentType 
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
            Height          =   360
            Left            =   1920
            TabIndex        =   1
            Top             =   165
            Width           =   1215
         End
         Begin VB.Label lblCreditAC 
            Caption         =   "Credit AC"
            Height          =   375
            Left            =   4680
            TabIndex        =   20
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Debit AC"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Description"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   780
            Width           =   855
         End
         Begin VB.Label lblPaymentType 
            Caption         =   "Receipt Type"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Options"
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   10815
         Begin VB.OptionButton optOthers 
            Caption         =   "Others"
            Height          =   195
            Left            =   8760
            TabIndex        =   30
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton optPolicyLoan 
            Caption         =   "Policy Loan"
            Height          =   255
            Left            =   8760
            TabIndex        =   29
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optMedical 
            Caption         =   "Medical"
            Height          =   255
            Left            =   4380
            TabIndex        =   9
            Top             =   600
            Width           =   2175
         End
         Begin VB.OptionButton optAgent 
            Caption         =   "Agent"
            Height          =   255
            Left            =   4380
            TabIndex        =   6
            Top             =   240
            Width           =   1935
         End
         Begin VB.OptionButton optReinsurer 
            Caption         =   "Reinsurer"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton optClaim 
            Caption         =   "Claim"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "frmALISPPaymentType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCode As ADODB.Recordset


Sub SavePayment()
'On Error GoTo err
    
    Set rsSAVE = New ADODB.Recordset
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With RsCode
'        !CostCenter = frmODASPCostCentre.txtCostCenter
'        !Description = frmODASPCostCentre.txtDescription
'        !CreditAC = frmODASPCostCentre.cboCreditAC
'        !DebitAC = frmODASPCostCentre.cboDebitAC
'        !Agent = frmODASPCostCentre.optAgent
'        !Claim = frmODASPCostCentre.optClaim
'        !Medical = frmODASPCostCentre.optMedical
'        !Others = frmODASPCostCentre.optOthers
'        !PolicyLoan = frmODASPCostCentre.optPolicyLoan
'        !Reinsurer = frmODASPCostCentre.optReinsurer
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub

Sub ShowPayment()
'''On Error GoTo err

    With RsCode
'        !CostCenter = frmODASPCostCentre.txtCostCenter
'        !Description = frmODASPCostCentre.txtDescription
'        !CreditAC = frmODASPCostCentre.cboCreditAC
'        !DebitAC = frmODASPCostCentre.cboDebitAC
'        !Agent = frmODASPCostCentre.optAgent
'        !Claim = frmODASPCostCentre.optClaim
'        !Medical = frmODASPCostCentre.optMedical
'        !Other = frmODASPCostCentre.optOthers
'        !PolicyLoan = frmODASPCostCentre.optPolicyLoan
'        !Reinsurer = frmODASPCostCentre.optReinsurer

    End With

Exit Sub

err:
    ErrorMessage
End Sub




Private Sub cmdAddNew_Click()
        enableALLRECORD
        clearALLRECORD
        disableButtons
End Sub


Private Sub cmdCancel_Click()
        enableButtons
        clearALLRECORD
        disableALLRECORD
End Sub


Private Sub cmdDelete_Click()
'''On Error GoTo err

If txtCostCenter.Text = "" Then
            MsgBox "There is no current record to delete", vbInformation, "Delete Information"
ElseIf txtDescription.Text = "" Then
            MsgBox "There is no current record", vbInformation, "Delete Information"
Else
        If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            With RsCode
                
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

Private Sub cmdEdit_Click()
'''On Error GoTo err

Dim strQRE As Variant
Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset

        Select Case cmdedit.Caption
                Case "&Edit"
                        enableALLRECORD

                        strQRE = InputBox("Enter Payment Type to search.", "Search Value")
    
                        rsFind.Open "SELECT * FROM ALISPPayment WHERE CostCenter LIKE '" & strQRE & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

'                        With rsFind
'                                If .EOF And .BOF Then
'                                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
'                                Else
'                                                txtCostCenter = !CostCenter
'                                                txtDescription = !Description
'                                                cboCreditAC = !CreditAC
'                                                cboDebitAC = !DebitAC
'
'                                                If !JobBrief = True Then
'                                                       ' chkPremium.Value = 1
'                                                    'Else: chkPremium.Value = 0
'                                                End If
'
'                                                If !CalculateCommission = True Then
'                                                       ' chkCalculateCommission.Value = 1
'                                                    'Else: chkCalculateCommission.Value = 0
'                                                End If
'
'                                                If !loan = True Then
'                                                       ' chkLoan.Value = 1
'                                                'Else: chkLoan.Value = 0
'                                                End If
'
'                                                If !interest = True Then
'                                                        'chkInterest.Value = 1
'                                                        'Else: chkInterest.Value = 0
'                                                End If
'
'                                                If !Deposit = True Then
'                                                           chkDeposit.Value = 1
'                                                     Else: chkDeposit.Value = 0
'                                                End If
'
'                                                If !RevivalFee = True Then
'                                                        chkRevivalFee.Value = 1
'                                                Else: chkRevivalFee.Value = 0
'                                                End If
'
'                                                If !Ghost = True Then
'                                                        chkGhost.Value = 1
'                                                        Else: chkGhost.Value = 0
'                                                End If
'
'                                                If !Miscellaneous = True Then
'                                                        chkMiscellaneous.Value = 1
'                                                        Else: chkMiscellaneous.Value = 0
'                                                End If
'
'                                                If !Breakdown = True Then
'                                                        chkBreakdown.Value = 1
'                                                    Else: chkBreakdown.Value = 0
'                                                End If
'
'                                                If !Account = True Then
'                                                        chkEmployer.Value = 1
'                                                    Else: chkEmployer.Value = 0
'                                                End If
'
'                                                If !Client = True Then
'                                                        chkClient.Value = 1
'                                                    Else: chkClient.Value = 0
'                                                End If
'                                        Edit = True
'                                End If
'                        End With
'
'                        If Edit Then
'                                cmdEdit.Caption = "Save &Changes"
'                        End If
'
'                Case "Save &Changes"
'                        Dim rsFinder As ADODB.Recordset
'                        Set rsFinder = New ADODB.Recordset
'
'                        rsFinder.Open "SELECT * FROM ALISPPayment WHERE CostCenter LIKE '" & txtCostCenter.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
'
'                        With rsFinder
'                            !CostCenter = txtCostCenter
'                            !Description = txtDescription
'                            !CalculateCommission = chkCalculateCommission
'                            !CreditAC = cboCreditAC
'                            !DebitAC = cboDebitAC
'                            !JobBrief = chkPremium
'                            !loan = chkLoan
'                            !interest = chkInterest
'                            !Deposit = chkDeposit
'                            !RevivalFee = chkRevivalFee
'                            !Ghost = chkGhost
'                            !Miscellaneous = chkMiscellaneous
'                            !Breakdown = chkBreakdown
'                            !Account = chkEmployer
'                            !Client = chkEmployer
'                            .Update
'                            .Requery
'                            Edit = False
'                    End With
'
'                    clearPAYMENT
'                    cmdEdit.Caption = "&Edit"
'            Case Else
'
'            Exit Sub
'
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


Private Sub cmdFirstCode_Click(Index As Integer)
'''On Error GoTo err

        cmdUpdate.Enabled = False

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

                    ShowPayment
                    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ValidatePayment()
'''On Error GoTo err
        
        bsaveRECORD = False
        
        If txtCostCenter.Text = "" Then
                MsgBox "Payment Type is Required"
                txtCostCenter.SetFocus
        ElseIf txtDescription.Text = "" Then
                MsgBox "The Description is  required"
                txtCostCenter.SetFocus
        Else
                bsaveRECORD = True
        End If

Exit Sub

err:
        ErrorMessage
End Sub


Private Sub cmdUpdate_Click()
        ValidatePayment
        enableButtons
        disableALLRECORD
End Sub


Private Sub Form_Activate()

    Set RsCode = New Recordset
            strSQL = "SELECT * from ODASPCostCentre;"

    RsCode.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

    disableALLRECORD
    enableButtons
    


End Sub

Private Sub Form_Load()

    Call OpenConnection
      

End Sub




