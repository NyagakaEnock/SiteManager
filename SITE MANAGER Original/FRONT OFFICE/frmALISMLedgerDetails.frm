VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmALISMLedgerDetails 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Ledger Details"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   11220
   Icon            =   "frmALISMLedgerDetails.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   11220
   Begin VB.Frame Frame14 
      Height          =   7215
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   11055
      Begin VB.Frame Frame3 
         Caption         =   "Receipts"
         Height          =   1695
         Left            =   7680
         TabIndex        =   46
         Top             =   5400
         Width           =   3255
         Begin VB.CommandButton cmdLedger 
            Appearance      =   0  'Flat
            Caption         =   "&Old Ledger"
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
            Left            =   240
            TabIndex        =   51
            Top             =   720
            Width           =   2655
         End
         Begin VB.CommandButton cmdCancel 
            Appearance      =   0  'Flat
            Caption         =   "&Cancel"
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
            Left            =   240
            TabIndex        =   48
            Top             =   1200
            Width           =   2655
         End
         Begin VB.CommandButton cmdprintledgerdetails 
            Appearance      =   0  'Flat
            Caption         =   "&New Ledger"
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
            Left            =   240
            TabIndex        =   47
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2535
         Left            =   7680
         TabIndex        =   25
         Top             =   120
         Width           =   3255
         Begin VB.TextBox txtPremiumDue 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   10
            Top             =   1620
            Width           =   1695
         End
         Begin VB.TextBox txtPlanPremium 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   11
            Top             =   2025
            Width           =   1695
         End
         Begin VB.TextBox txtStatusCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   8
            Top             =   810
            Width           =   1695
         End
         Begin VB.TextBox txtDueDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   7
            Top             =   405
            Width           =   1695
         End
         Begin VB.TextBox txtUnitCountBeforePayment 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   9
            Top             =   1215
            Width           =   1695
         End
         Begin VB.Label Label11 
            Caption         =   "Prem Due"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   1755
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Plan Prem"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   2115
            Width           =   735
         End
         Begin VB.Label lblDueDate 
            Caption         =   "Due Date:"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblStatusCode 
            Caption         =   "Status Code"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   885
            Width           =   975
         End
         Begin VB.Label lblPremiumDetails 
            Caption         =   "Premium Details"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   27
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label lblUnitCountBeforePayment 
            Caption         =   "Prem Prior:"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1335
            Width           =   855
         End
      End
      Begin VB.Frame Frame15 
         Height          =   2535
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   7455
         Begin VB.TextBox cboJobBriefNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   240
            Width           =   3015
         End
         Begin VB.TextBox txtPremiumcount 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   2160
            Width           =   1815
         End
         Begin VB.TextBox txtDateofCommencement 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5520
            TabIndex        =   41
            Top             =   1680
            Width           =   1815
         End
         Begin VB.TextBox txtMaturityDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtTermofPolicy 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox txtlastpaydate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtdob 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   1680
            Width           =   3015
         End
         Begin VB.TextBox txtsex 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   0
            Top             =   1200
            Width           =   3015
         End
         Begin VB.TextBox txtagent 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   2160
            Width           =   3015
         End
         Begin VB.TextBox txtNames 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   720
            Width           =   3015
         End
         Begin VB.Label lblCurrentPeriod 
            Caption         =   "Prem Count"
            Height          =   210
            Left            =   4440
            TabIndex        =   44
            Top             =   2130
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "DOC"
            Height          =   255
            Left            =   4440
            TabIndex        =   42
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label Label13 
            Caption         =   "Maturity Date"
            Height          =   255
            Left            =   4440
            TabIndex        =   40
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Term"
            Height          =   255
            Left            =   4440
            TabIndex        =   39
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label Label8 
            Caption         =   "Last pay Date"
            Height          =   255
            Left            =   4440
            TabIndex        =   36
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "DoB"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Sex"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1155
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Agent"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label lblDocumentNo 
            Caption         =   "Brief No"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Names"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame Frame16 
         Height          =   1575
         Left            =   120
         TabIndex        =   20
         Top             =   5400
         Width           =   7455
         Begin VB.TextBox cboAccountNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5520
            TabIndex        =   53
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtPaymentMode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4440
            TabIndex        =   52
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtexpectedpremium 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtpsuspense 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1080
            Width           =   2895
         End
         Begin VB.ComboBox cboPaymentMethod 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1320
            TabIndex        =   14
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txtDateOfLastPayment 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   645
            Width           =   1695
         End
         Begin VB.TextBox txtReceivedTodate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   180
            Width           =   2895
         End
         Begin VB.Label Label14 
            Caption         =   "Suspense"
            Height          =   330
            Left            =   3120
            TabIndex        =   45
            Top             =   1155
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Modal Premium"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   195
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Pay Method"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblDateOfLastPayment 
            Caption         =   "Last Pay Date"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblReceiptAmount 
            Caption         =   "Received to Date"
            Height          =   255
            Left            =   3120
            TabIndex        =   22
            Top             =   300
            Width           =   1335
         End
         Begin VB.Label lblPaymentStatus 
            Caption         =   "Payment Mode"
            Height          =   210
            Left            =   3120
            TabIndex        =   21
            Top             =   690
            Width           =   1215
         End
      End
      Begin VB.Frame Frame17 
         Height          =   2775
         Left            =   120
         TabIndex        =   19
         Top             =   2640
         Width           =   10815
         Begin MSComctlLib.ListView ListView1 
            Height          =   2415
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   4260
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
   End
End
Attribute VB_Name = "frmALISMLedgerDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLedgerdetails As clsLedgerDetails
Public rsDEP As ADODB.Recordset, strDEP As String

Private Sub cboJobBriefNo_LostFocus()
On Error GoTo err
    If cboJobBriefNo.Text = "" Then Exit Sub
        Set rsLedgerdetails = New clsLedgerDetails
        rsLedgerdetails.loadPolicy
    Exit Sub
err: ErrorMessage
End Sub

Private Sub cmdCancel_Click()
        enableALLRECORD
End Sub

Private Sub loadRECORD()
On Error GoTo err
        
        Set rsCONTROL = New Recordset
        rsCONTROL.Open "SELECT * FROM ODASMJobBrief, ODASPAccount where ODASMJobBrief.JobBriefNo = '" & CurrentRecord & "' and ODASPAccount.AccountNo LIKE ODASMJobBrief.AccountNo ;", cnCOMMON, adOpenKeyset, adLockOptimistic

         With rsCONTROL
                If .BOF Or .EOF Then Exit Sub
                
                frmALISMLedgerDetails.cboJobBriefNo.Text = !JobBriefNo & ""
                frmALISMLedgerDetails.txtReceivedTodate.Text = !ReceivedToDate & ""
                frmALISMLedgerDetails.txtpsuspense.Text = !SuspenseAccount & ""
                frmALISMLedgerDetails.txtDateOfLastPayment.Text = !DateofLastPayment & ""
                frmALISMLedgerDetails.txtReceivedTodate.Text = !ReceivedToDate & ""
                frmALISMLedgerDetails.txtUnitCountBeforePayment.Text = !UnitCountBeforePayment & ""
                frmALISMLedgerDetails.txtDueDate.Text = !DueDate & ""
                frmALISMLedgerDetails.txtStatusCode.Text = !StatusCode & ""
                frmALISMLedgerDetails.txtexpectedpremium.Text = !ExpectedPremium & ""
                frmALISMLedgerDetails.txtDateOfCommencement.Text = !DateOfCommencement & ""
                frmALISMLedgerDetails.txtPlanPremium.Text = !PlanPremium & ""
                frmALISMLedgerDetails.txtMaturityDate.Text = !MaturityDate & ""
                frmALISMLedgerDetails.txtlastpaydate.Text = !DateofLastPayment & ""
                frmALISMLedgerDetails.txtPaymentMode.Text = !PaymentMode & ""
                frmALISMLedgerDetails.txtTermOfPolicy.Text = !TermOfPolicy & ""
                frmALISMLedgerDetails.cboPaymentMethod.Text = !PaymentMethod & ""
                frmALISMLedgerDetails.txtPremiumDue.Text = !NoofPremiumsdue & ""
                frmALISMLedgerDetails.txtDateOfLastPayment.Text = !DateofLastPayment & ""
                frmALISMLedgerDetails.txtUnitCountBeforePayment.Text = !UnitCountBeforePayment & ""
                frmALISMLedgerDetails.txtDueDate.Text = !DueDate & ""
                frmALISMLedgerDetails.txtPremiumcount.Text = !UnitCount & ""

        End With

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cmdLedger_Click()
    Screen.ActiveForm.cmdLedger.Enabled = False
    Load frmLedgerDetailsOld
    frmLedgerDetailsOld.Show 1, Me
    Screen.ActiveForm.cmdLedger.Enabled = True
End Sub

Private Sub cmdprintledgerdetails_Click()
    Screen.ActiveForm.cmdprintledgerdetails.Enabled = False
    Load frmLedgerDetails
    frmLedgerDetails.Show 1, Me
    Screen.ActiveForm.cmdprintledgerdetails.Enabled = True
End Sub

Private Sub Form_Activate()
        disableALLRECORD
        
        Set rsLedgerdetails = New clsLedgerDetails
        rsLedgerdetails.loadPolicy
        strSQL = "SELECT  ODASMJobBriefLedger.txtReceiptNo,ODASMJobBriefLedger.receiptAmount, ODASMJobBriefLedger.Receiptdate,ODASMJobBriefLedger.unitcount,ODASMJobBriefLedger.Receivedtodate,ODASMJobBriefLedger.receiptType,ODASMJobBriefLedger.statuscode,ODASMJobBriefLedger.suspenseAccount FROM ODASMJobBriefLedger WHERE DocumentNo =  '" & frmALISMLedgerDetails.cboJobBriefNo.Text & "';"
        rsLedgerdetails.getLedgerDetails
        Set rsLedgerdetails = Nothing
End Sub

