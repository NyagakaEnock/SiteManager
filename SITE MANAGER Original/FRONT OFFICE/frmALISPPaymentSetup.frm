VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmALISPPaymentSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Code"
   ClientHeight    =   7500
   ClientLeft      =   150
   ClientTop       =   1155
   ClientWidth     =   11325
   Icon            =   "frmALISPPaymentSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   11325
   Begin VB.Frame Frame12 
      Height          =   7455
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   11055
      Begin VB.Frame Frame16 
         Caption         =   "Request for Details"
         Height          =   1695
         Left            =   120
         TabIndex        =   28
         Top             =   3240
         Width           =   9735
         Begin VB.CheckBox chkAllowLetters 
            Caption         =   "Allow Letters ?"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CheckBox chkRefund 
            Caption         =   "Premium Refund ?"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   1080
            Width           =   2295
         End
         Begin VB.CheckBox chkClaimCauses 
            Caption         =   "Require Causes ?"
            Height          =   255
            Left            =   2880
            TabIndex        =   39
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CheckBox chkNotification 
            Caption         =   "Require Notification ?"
            Height          =   255
            Left            =   2880
            TabIndex        =   38
            Top             =   840
            Width           =   2175
         End
         Begin VB.CheckBox chkInstallment 
            Caption         =   "Involve Installments ?"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   840
            Width           =   2055
         End
         Begin VB.CheckBox chkClaimantDetails 
            Caption         =   "Request Claimant Details"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   270
            Width           =   2295
         End
         Begin VB.CheckBox chkCorrespondentDetails 
            Caption         =   "Request Correspondent Details"
            Height          =   315
            Left            =   2880
            TabIndex        =   35
            Top             =   240
            Width           =   2535
         End
         Begin VB.CheckBox chkAlterStatus 
            Caption         =   "Alter Status"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   570
            Width           =   1335
         End
         Begin VB.CheckBox chkPayLifeAssured 
            Caption         =   "Pay Life Assured"
            Height          =   255
            Left            =   2880
            TabIndex        =   33
            Top             =   570
            Width           =   2295
         End
         Begin VB.CheckBox chkPersonalAccident 
            Caption         =   "Personal Accident"
            Height          =   255
            Left            =   6360
            TabIndex        =   32
            Top             =   570
            Width           =   1695
         End
         Begin VB.CheckBox chkSurrender 
            Caption         =   "Surrender"
            Height          =   315
            Left            =   6360
            TabIndex        =   31
            Top             =   240
            Width           =   1695
         End
         Begin VB.CheckBox chkDeath 
            Caption         =   "Death?"
            Height          =   255
            Left            =   6360
            TabIndex        =   30
            Top             =   840
            Width           =   1695
         End
         Begin VB.CheckBox chkPaidup 
            Caption         =   "Paid up?"
            Height          =   255
            Left            =   6360
            TabIndex        =   29
            Top             =   1080
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Payment Type"
         Height          =   3135
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   5055
         Begin MSComctlLib.ListView ListView1 
            Height          =   2775
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   4895
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame fraCButtons 
         Height          =   3495
         Index           =   6
         Left            =   9960
         TabIndex        =   18
         Top             =   3240
         Width           =   975
         Begin VB.CommandButton cmdPrint 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   120
            Picture         =   "frmALISPPaymentSetup.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   2820
            Width           =   735
         End
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   120
            Picture         =   "frmALISPPaymentSetup.frx":0544
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   2370
            Width           =   735
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
            Height          =   450
            Left            =   120
            Picture         =   "frmALISPPaymentSetup.frx":0646
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   1920
            Width           =   735
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
            Height          =   450
            Left            =   120
            Picture         =   "frmALISPPaymentSetup.frx":0748
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   570
            Width           =   735
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
            Height          =   450
            Left            =   120
            Picture         =   "frmALISPPaymentSetup.frx":084A
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   1470
            Width           =   735
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
            Height          =   450
            Left            =   120
            Picture         =   "frmALISPPaymentSetup.frx":094C
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   120
            Width           =   735
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
            Height          =   450
            Left            =   120
            Picture         =   "frmALISPPaymentSetup.frx":0A4E
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1020
            Width           =   735
         End
      End
      Begin VB.Frame frabrowse 
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   6720
         Width           =   10815
         Begin VB.CommandButton cmdFirstCode 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   0
            Left            =   720
            Picture         =   "frmALISPPaymentSetup.frx":0B50
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdFirstCode 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   1
            Left            =   2160
            Picture         =   "frmALISPPaymentSetup.frx":0F92
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdFirstCode 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   2
            Left            =   3600
            Picture         =   "frmALISPPaymentSetup.frx":13D4
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdFirstCode 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   3
            Left            =   4920
            Picture         =   "frmALISPPaymentSetup.frx":1816
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Frame Frame14 
         Height          =   1815
         Left            =   120
         TabIndex        =   10
         Top             =   4920
         Width           =   9735
         Begin MSComctlLib.ListView ListView2 
            Height          =   1575
            Left            =   120
            TabIndex        =   46
            Top             =   120
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   2778
            View            =   3
            MultiSelect     =   -1  'True
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame13 
         Height          =   3135
         Left            =   5280
         TabIndex        =   7
         Top             =   120
         Width           =   5655
         Begin VB.ComboBox cboNewStatus 
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
            TabIndex        =   44
            Top             =   2520
            Width           =   3495
         End
         Begin VB.ComboBox cboRiderCode 
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
            TabIndex        =   42
            Top             =   1398
            Width           =   3495
         End
         Begin VB.TextBox txtPaymentType 
            BackColor       =   &H00FFFFC0&
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
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   240
            Width           =   3495
         End
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
            Left            =   1920
            TabIndex        =   4
            Top             =   2160
            Width           =   3495
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
            Top             =   1779
            Width           =   3495
         End
         Begin VB.TextBox txtClaimCodeDescription 
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
            Top             =   1002
            Width           =   3495
         End
         Begin VB.TextBox txtClaimCode 
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
            Top             =   600
            Width           =   3495
         End
         Begin VB.Label Label5 
            Caption         =   "New Status"
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   2580
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Rider Code"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   1451
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Payment Type"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   293
            Width           =   1215
         End
         Begin VB.Label lblCreditAC 
            Caption         =   "Credit AC"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   2220
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Debit AC"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   1830
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Description"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   1062
            Width           =   855
         End
         Begin VB.Label lblPaymentType 
            Caption         =   "Payment Code"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   674
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmALISPPaymentSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCode As ADODB.Recordset


Private Sub SavePayment()
''oN ERROR GoTo err
    
    Set rsSAVE = New ADODB.Recordset
    strSQL = "Select * from ALISPPaymentSetup Where CostCenter = '" & frmALISPPaymentSetup.txtCostCenter.Text & "' and PaymentCode = '" & frmALISPPaymentSetup.txtCostCenter.Text & "';"
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With RsCode
'        !CostCenter = frmALISPPaymentSetup.txtCostCenter
'        !Description = frmALISPPaymentSetup.txtPaymentCodeDescription
'        !CreditAC = frmALISPPaymentSetup.cboCreditAC
'        !DebitAC = frmALISPPaymentSetup.cboDebitAC
'        !Agent = frmALISPPaymentSetup.optAgent
'        !Claim = frmALISPPaymentSetup.optClaim
'        !Medical = frmALISPPaymentSetup.optMedical
'        !Others = frmALISPPaymentSetup.optOthers
'        !PolicyLoan = frmALISPPaymentSetup.optPolicyLoan
'        !Reinsurer = frmALISPPaymentSetup.optReinsurer
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ShowPayment()
''oN ERROR GoTo err

    With RsCode
        
        frmALISPPaymentSetup.txtCostCenter.Text = !CostCenter & ""
        frmALISPPaymentSetup.cboCreditAC = !CreditAC & ""
        frmALISPPaymentSetup.cboDebitAC = !DebitAC & ""
        frmALISPPaymentSetup.txtPaymentCode = !PaymentCode
        frmALISPPaymentSetup.txtPaymentCodeDescription = !PaymentCodeDescription
        frmALISPPaymentSetup.cboNewStatus.Text = !NewStatus & ""
        
        frmALISPPaymentSetup.chkAlterStatus = !AlterStatus
        frmALISPPaymentSetup.chkClaimantDetails = !claimantdetails
        frmALISPPaymentSetup.chkCorrespondentDetails = !CorrespondentDetails
        frmALISPPaymentSetup.chkInstallment = !Installment
        frmALISPPaymentSetup.chkPayLifeAssured = !PayLifeAssured
        frmALISPPaymentSetup.chkNotification = !Notification
        frmALISPPaymentSetup.chkAllowLetters = !AllowLetters
        frmALISPPaymentSetup.chkRefund = !Refund
        frmALISPPaymentSetup.chkClaimCauses = !claimcauses
        frmALISPPaymentSetup.chkSurrender = !surrender
        frmALISPPaymentSetup.chkPersonalAccident = !PersonalAccident
        frmALISPPaymentSetup.chkDeath = !Death



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
''oN ERROR GoTo err
With frmALISPPaymentSetup
        If .txtCostCenter.Text = "" Then
                    MsgBox "There is no current record to delete", vbInformation, "Delete Information"
        ElseIf .txtPaymentCodeDescription.Text = "" Then
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
End With

Exit Sub

err:
    ErrorMessage

End Sub

Private Sub cmdEdit_Click()
''oN ERROR GoTo err

Dim strQRE As Variant
Dim rsFinder As ADODB.Recordset, Edit As Boolean

        Set rsFinder = New ADODB.Recordset

        Select Case frmALISPPaymentSetup.cmdEdit.Caption
                Case "&Edit"
                        enableALLRECORD

                        strQRE = InputBox("Enter Claim Code to search.", "Search Value")
    
                        rsFinder.Open "SELECT * FROM ODASPPaymentCode WHERE PaymentCode = '" & strQRE & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic

                        With rsFinder
                                If .EOF And .BOF Then
                                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                                Else
                                        frmALISPPaymentSetup.txtPaymentCode = !PaymentCode
                                        frmALISPPaymentSetup.txtCostCenter = !CostCenter & ""
                                        frmALISPPaymentSetup.cboNewStatus.Text = !NewStatus & ""
                                        frmALISPPaymentSetup.txtPaymentCodeDescription = !PaymentCodeDescription
                                    
                                        frmALISPPaymentSetup.chkAlterStatus = !AlterStatus
                                        frmALISPPaymentSetup.chkClaimantDetails = !claimantdetails
                                        frmALISPPaymentSetup.chkCorrespondentDetails = !CorrespondentDetails
                                        frmALISPPaymentSetup.chkInstallment = !Installment
                                        frmALISPPaymentSetup.chkPayLifeAssured = !PayLifeAssured
                                        frmALISPPaymentSetup.chkNotification = !Notification
                                        frmALISPPaymentSetup.chkAllowLetters = !AllowLetters
                                        frmALISPPaymentSetup.chkRefund = !Refund
                                        frmALISPPaymentSetup.chkClaimCauses = !claimcauses
                                        frmALISPPaymentSetup.chkSurrender = !surrender
                                        frmALISPPaymentSetup.chkPersonalAccident = !PersonalAccident
                                        frmALISPPaymentSetup.chkDeath = !Death
                                        frmALISPPaymentSetup.chkPaidup = !Paidup

                                        Edit = True
                                End If
                        End With
        
                        If Edit Then
                                frmALISPPaymentSetup.cmdEdit.Caption = "Save &Changes"
                        End If
    
                Case "Save &Changes"
                        'Dim rsFINDER As ADODB.Recordset
                        Set rsFinder = New ADODB.Recordset

                        rsFinder.Open "SELECT * FROM ODASPPaymentCode WHERE PaymentCode = '" & frmALISPPaymentSetup.txtPaymentCode.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                    
                        With rsFinder
                        
                            !PaymentCode = frmALISPPaymentSetup.txtPaymentCode
                            !NewStatus = frmALISPPaymentSetup.cboNewStatus
                            !CostCenter = frmALISPPaymentSetup.txtCostCenter.Text
                            !PaymentCodeDescription = frmALISPPaymentSetup.txtPaymentCodeDescription
                            !AlterStatus = frmALISPPaymentSetup.chkAlterStatus
                            !claimantdetails = frmALISPPaymentSetup.chkClaimantDetails
                            !CorrespondentDetails = frmALISPPaymentSetup.chkCorrespondentDetails
                            !Installment = frmALISPPaymentSetup.chkInstallment
                            !PayLifeAssured = frmALISPPaymentSetup.chkPayLifeAssured
                            !AllowLetters = frmALISPPaymentSetup.chkAllowLetters
                            !Refund = frmALISPPaymentSetup.chkRefund
                            !claimcauses = frmALISPPaymentSetup.chkClaimCauses
                            !surrender = frmALISPPaymentSetup.chkSurrender
                            !PersonalAccident = frmALISPPaymentSetup.chkPersonalAccident
                            !Death = frmALISPPaymentSetup.chkDeath
                            !Paidup = frmALISPPaymentSetup.chkPaidup

                            .Update
                            .Requery
                            Edit = False
                    End With
                
                    frmALISPPaymentSetup.cmdEdit.Caption = "&Edit"
            Case Else
        
            Exit Sub

        End Select
GetInvoicesNotPaid

Exit Sub

err:

    If err.Number = 40009 Then
            MsgBox "Record requested does not exist in the Database! Check your Entries.", vbInformation, "Searching."
                rsFinder.Requery

            If rsFinder.BOF Then Exit Sub
                rsFinder.MoveFirst

    ElseIf err.Number = 3021 Then
            MsgBox "Requested record not found! Refresh the database and try the search again...or Check your entries.", vbInformation, "Searching."
                rsFinder.Requery

            If rsFinder.BOF Then Exit Sub
                rsFinder.MoveFirst
    Else
                UpdateErrorMessage
End If
End Sub




Private Sub cmdFirstCode_Click(Index As Integer)

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
End Sub

Private Sub ValidatePayment()
''oN ERROR GoTo err
        
        bsaveRECORD = False
        With frmALISPPaymentSetup
                If .txtCostCenter.Text = "" Then
                        MsgBox "Payment Type is Required"
                        .txtCostCenter.SetFocus
                ElseIf .txtPaymentCodeDescription.Text = "" Then
                        MsgBox "The Description is  required"
                        .txtPaymentCodeDescription.SetFocus
                Else
                        bsaveRECORD = True
                End If
        End With
Exit Sub

err:
        ErrorMessage
End Sub


Private Sub cmdUpdate_Click()
        ValidateProduct
        If bsaveRECORD = True Then
            savePRODUCT
            
            If bsaveRECORD = False Then
                enableButtons
                disableALLRECORD
            End If
        End If
        
        GetInvoicesNotPaid

End Sub
Private Sub savePRODUCT()
''oN ERROR GoTo err
        
        Set rsSAVE = New ADODB.Recordset
        strSQL = "Select * from ODASPPaymentCode where PaymentCode = '" & frmALISPPaymentSetup.txtPaymentCode & "' and CostCenter = '" & frmALISPPaymentSetup.txtCostCenter & "';"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
        With rsSAVE
                    If .EOF Or .BOF Then
                            .AddNew
                            !PaymentCode = frmALISPPaymentSetup.txtPaymentCode
                            !CostCenter = frmALISPPaymentSetup.txtCostCenter.Text
                    End If
                    
                    !PaymentCodeDescription = frmALISPPaymentSetup.txtPaymentCodeDescription
                    !AlterStatus = frmALISPPaymentSetup.chkAlterStatus
                    !claimantdetails = frmALISPPaymentSetup.chkClaimantDetails
                    !CorrespondentDetails = frmALISPPaymentSetup.chkCorrespondentDetails
                    !Installment = frmALISPPaymentSetup.chkInstallment
                    !Notification = frmALISPPaymentSetup.chkNotification
                    !PayLifeAssured = frmALISPPaymentSetup.chkPayLifeAssured
                    !AllowLetters = frmALISPPaymentSetup.chkAllowLetters
                    !Refund = frmALISPPaymentSetup.chkRefund
                    !claimcauses = frmALISPPaymentSetup.chkClaimCauses
                    !surrender = frmALISPPaymentSetup.chkSurrender
                    !PersonalAccident = frmALISPPaymentSetup.chkPersonalAccident
                    !Death = frmALISPPaymentSetup.chkDeath
                    !Paidup = frmALISPPaymentSetup.chkPaidup.Value
                    bsaveRECORD = False
                .Update
                .Requery
        End With
  
rsSAVE.Close
strSQL = ""
Exit Sub

err:
    
    UpdateErrorMessage
End Sub
Private Sub ValidateProduct()
''oN ERROR GoTo err
        bsaveRECORD = True
        With frmALISPPaymentSetup
                If .txtCostCenter.Text = "" Then
                        MsgBox "The Payment Type MUST be Entered Before further Processing", vbOKCancel
                        .txtCostCenter.SetFocus
                
                ElseIf .txtPaymentCode.Text = "" Then
                        MsgBox "The Payment Code is Required"
                        .txtPaymentCode.SetFocus
                
                ElseIf .txtPaymentCodeDescription.Text = "" Then
                        MsgBox "The Payment Code description is Required", vbOKOnly
                        .txtPaymentCodeDescription.SetFocus
                
                ElseIf .cboNewStatus.Text = "" Then
                        MsgBox "The New Status is required"
                        .cboNewStatus.SetFocus
        
                Else
                        bsaveRECORD = True
                End If
        End With
Exit Sub

err:
    ErrorMessage

End Sub

Private Sub Form_Activate()

    Set RsCode = New Recordset
            strSQL = "SELECT * from ODASPPaymentCode;"

    RsCode.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

    disableALLRECORD
    enableButtons
    GetInvoicesNotPaid

    GetPayment

End Sub

Private Sub Form_Load()

    Call OpenConnection
      

End Sub
Public Sub GetInvoicesNotPaid()
''oN ERROR GoTo err
    
        With Screen.ActiveForm
        
                .ListView2.ListItems.Clear
                .ListView2.ColumnHeaders.Clear
                
                .ListView2.ColumnHeaders.Add , , "Claim Code", .ListView2.Width / 5
                .ListView2.ColumnHeaders.Add , , "Payment Type", .ListView2.Width / 5
                .ListView2.ColumnHeaders.Add , , "Description", .ListView2.Width / 5
                .ListView2.ColumnHeaders.Add , , "New Status", .ListView2.Width / 5
                .ListView2.ColumnHeaders.Add , , "Debit AC", .ListView2.Width / 5
                .ListView2.ColumnHeaders.Add , , "Credit AC", .ListView2.Width / 5

                .ListView2.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select PaymentCode,CostCenter, PaymentCodeDescription, NewStatus, DebitAc, CreditAc  from ODASPPaymentCode ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!PaymentCode))
                        
                        If Not IsNull(rsLIST!PaymentCodeDescription) Then
                                MyList.SubItems(1) = CStr(rsLIST!PaymentCodeDescription)
                        End If

                        If Not IsNull(rsLIST!PaymentCodeDescription) Then
                                MyList.SubItems(2) = CStr(rsLIST!PaymentCodeDescription)
                        End If
                        
                        If Not IsNull(rsLIST!NewStatus) Then
                                MyList.SubItems(3) = CStr(rsLIST!NewStatus)
                        
                        If Not IsNull(rsLIST!NewStatus) Then
                                MyList.SubItems(4) = CStr(rsLIST!NewStatus)
                        End If

                        If Not IsNull(rsLIST!CreditAC) Then
                                MyList.SubItems(5) = CStr(rsLIST!CreditAC)
                        End If

                        End If

                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub GetPayment()
''oN ERROR GoTo err
    
        With Screen.ActiveForm
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Payment Type", .ListView1.Width / 2
                .ListView1.ColumnHeaders.Add , , "Payment Description", .ListView1.Width / 2

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                
                strSQL = "Select CostCenter, Description from ODASPCostCentre ;"
                rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                df = rsLIST.RecordCount
                Dim MyList As ListItem
                           
                While Not rsLIST.EOF
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!CostCenter))
                        
                        If Not IsNull(rsLIST!Description) Then
                                MyList.SubItems(1) = CStr(rsLIST!Description)
                        End If

                        rsLIST.MoveNext
                Wend
                Set MyList = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub


Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
        With frmALISPPaymentSetup
            .txtCostCenter.Text = Item.Text
            
        End With


End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
''oN ERROR GoTo err
        
        With frmALISPPaymentSetup
            .txtCostCenter.Text = Item.Text
            GetInvoicesNotPaid
            
        End With

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
''oN ERROR GoTo err
        
        With frmALISPPaymentSetup
            .txtPaymentCode.Text = Item.Text
            
            Set RsCode = New Recordset
                    strSQL = "SELECT * from ODASPPaymentCode where PaymentCode = '" & Item.Text & "';"

            RsCode.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

            ShowPayment
        End With
Exit Sub

err:
    ErrorMessage
End Sub
