VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmALISMLedgerDetails 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Ledger Details"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   11220
   Icon            =   "frmALISFOLedgerDetails.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   11220
   Begin VB.Frame Frame14 
      Height          =   7215
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   11055
      Begin VB.Frame frabrowse 
         Height          =   735
         Index           =   3
         Left            =   120
         TabIndex        =   53
         Top             =   6360
         Width           =   7455
         Begin VB.CommandButton cmdFirstCode 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   3
            Left            =   5160
            Picture         =   "frmALISFOLedgerDetails.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdFirstCode 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   2
            Left            =   3840
            Picture         =   "frmALISFOLedgerDetails.frx":0884
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdFirstCode 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   1
            Left            =   2400
            Picture         =   "frmALISFOLedgerDetails.frx":0CC6
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdFirstCode 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   0
            Left            =   960
            Picture         =   "frmALISFOLedgerDetails.frx":1108
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame18 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   7680
         TabIndex        =   34
         Top             =   3360
         Width           =   3255
         Begin VB.Frame Frame3 
            Caption         =   "Receipts"
            Height          =   3375
            Left            =   120
            TabIndex        =   35
            Top             =   120
            Width           =   3015
            Begin VB.CommandButton cmdprintledgerdetails 
               Appearance      =   0  'Flat
               Caption         =   "&Print Ledger Details"
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
               TabIndex        =   1
               Top             =   600
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
               TabIndex        =   36
               Top             =   1785
               Width           =   2655
            End
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3255
         Left            =   7680
         TabIndex        =   29
         Top             =   120
         Width           =   3255
         Begin VB.TextBox txtPremiumDue 
            Appearance      =   0  'Flat
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
            Height          =   405
            Left            =   1200
            TabIndex        =   12
            Top             =   2280
            Width           =   1695
         End
         Begin VB.TextBox txtPlanPremium 
            Appearance      =   0  'Flat
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
            Height          =   405
            Left            =   1200
            TabIndex        =   13
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox txtStatusCode 
            Appearance      =   0  'Flat
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
            Height          =   405
            Left            =   1200
            TabIndex        =   10
            Top             =   1170
            Width           =   1695
         End
         Begin VB.TextBox txtDueDate 
            Appearance      =   0  'Flat
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
            Height          =   405
            Left            =   1200
            TabIndex        =   9
            Top             =   645
            Width           =   1695
         End
         Begin VB.TextBox txtUnitCountBeforePayment 
            Appearance      =   0  'Flat
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
            Height          =   405
            Left            =   1200
            TabIndex        =   11
            Top             =   1740
            Width           =   1695
         End
         Begin VB.Label Label11 
            Caption         =   "Prem Due"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   2355
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Plan Prem"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   2835
            Width           =   735
         End
         Begin VB.Label lblDueDate 
            Caption         =   "Due Date:"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblStatusCode 
            Caption         =   "Status Code"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1245
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
            TabIndex        =   31
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label lblUnitCountBeforePayment 
            Caption         =   "Prem Prior:"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1815
            Width           =   855
         End
      End
      Begin VB.Frame Frame15 
         Height          =   2535
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   7455
         Begin VB.TextBox txtPremiumcount 
            Appearance      =   0  'Flat
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
            Height          =   405
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   2040
            Width           =   1815
         End
         Begin VB.TextBox txtDateofCommencement 
            Appearance      =   0  'Flat
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
            Height          =   405
            Left            =   5520
            TabIndex        =   48
            Top             =   1560
            Width           =   1815
         End
         Begin VB.TextBox txtMaturityDate 
            Appearance      =   0  'Flat
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
            Height          =   405
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtTermofPolicy 
            Appearance      =   0  'Flat
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
            Height          =   405
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtlastpaydate 
            Appearance      =   0  'Flat
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
            Height          =   375
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   120
            Width           =   1815
         End
         Begin VB.TextBox txtdob 
            Appearance      =   0  'Flat
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
            Height          =   375
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1560
            Width           =   3015
         End
         Begin VB.TextBox txtsex 
            Appearance      =   0  'Flat
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
            Height          =   375
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox txtagent 
            Appearance      =   0  'Flat
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
            Height          =   375
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   2040
            Width           =   3015
         End
         Begin VB.ComboBox cboPolicyNo 
            Appearance      =   0  'Flat
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
            Left            =   1320
            TabIndex        =   0
            Top             =   120
            Width           =   3015
         End
         Begin VB.TextBox txtNames 
            Appearance      =   0  'Flat
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
            Height          =   375
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   600
            Width           =   3015
         End
         Begin VB.Label lblCurrentPeriod 
            Caption         =   "Prem Count"
            Height          =   210
            Left            =   4440
            TabIndex        =   51
            Top             =   2130
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "DOC"
            Height          =   255
            Left            =   4440
            TabIndex        =   49
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label Label13 
            Caption         =   "Maturity Date"
            Height          =   255
            Left            =   4440
            TabIndex        =   47
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Term"
            Height          =   255
            Left            =   4440
            TabIndex        =   46
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label Label8 
            Caption         =   "Last pay Date"
            Height          =   255
            Left            =   4440
            TabIndex        =   43
            Top             =   180
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "DoB"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Sex"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   1155
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Agent"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label lblDocumentNo 
            Caption         =   "Policy  No"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   180
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Names"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame Frame16 
         Height          =   1575
         Left            =   120
         TabIndex        =   24
         Top             =   4920
         Width           =   7455
         Begin VB.TextBox txtexpectedpremium 
            Appearance      =   0  'Flat
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
            Height          =   375
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtPaymentMode 
            Appearance      =   0  'Flat
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
            Height          =   405
            Left            =   4440
            TabIndex        =   19
            Top             =   600
            Width           =   2895
         End
         Begin VB.TextBox txtpsuspense 
            Appearance      =   0  'Flat
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
            Height          =   405
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   1080
            Width           =   2895
         End
         Begin VB.ComboBox cboPaymentMethod 
            Appearance      =   0  'Flat
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
            Left            =   1320
            TabIndex        =   16
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txtDateOfLastPayment 
            Appearance      =   0  'Flat
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
            Height          =   375
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   645
            Width           =   1695
         End
         Begin VB.TextBox txtReceivedTodate 
            Appearance      =   0  'Flat
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
            Height          =   375
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   180
            Width           =   2895
         End
         Begin VB.Label Label14 
            Caption         =   "Suspense"
            Height          =   330
            Left            =   3120
            TabIndex        =   52
            Top             =   1155
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Modal Premium"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   195
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Pay Method"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label lblDateOfLastPayment 
            Caption         =   "Last Pay Date"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   705
            Width           =   1095
         End
         Begin VB.Label lblReceiptAmount 
            Caption         =   "Received to Date"
            Height          =   255
            Left            =   3120
            TabIndex        =   26
            Top             =   300
            Width           =   1335
         End
         Begin VB.Label lblPaymentStatus 
            Caption         =   "Payment Mode"
            Height          =   210
            Left            =   3120
            TabIndex        =   25
            Top             =   690
            Width           =   1215
         End
      End
      Begin VB.Frame Frame17 
         Height          =   2295
         Left            =   120
         TabIndex        =   23
         Top             =   2640
         Width           =   7455
         Begin MSDataGridLib.DataGrid ReceiptGrid 
            Height          =   1935
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   3413
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777152
            HeadLines       =   1
            RowHeight       =   15
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
   End
End
Attribute VB_Name = "frmALISMLedgerDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLedgerdetails As clsLedgerDetails
Dim rsRCPT As ADODB.Recordset, strSQL As String
Public rsDEP As ADODB.Recordset, strDEP As String


Sub LoadReceiptGrid()

    On Error GoTo err:
            Set ReceiptGrid.DataSource = Nothing
            Dim rsGRID As ADODB.Recordset, StrGRID As String
            Set rsGRID = New Recordset
               rsGRID.Open "SELECT  ALISMPolicyLedger.ReceiptNo,ALISMPolicyLedger.Receiptdate,ALISMPolicyLedger.unitcount,ALISMPolicyLedger.Receivedtodate,ALISMPolicyLedger.unitcountbeforepayment,ALISMPolicyLedger.statuscode,ALISMPolicyLedger.suspenseAccount FROM ALISMPolicyLedger WHERE DocumentNo =  '" & frmALISMLedgerDetails.cboPolicyNo.Text & "';", cnALIS, adOpenKeyset, adLockOptimistic
               'If Not rsGRID.EOF And Not rsGRID.BOF Then
               Set ReceiptGrid.DataSource = rsGRID
               'Else
               
               'End If
            Exit Sub
            
err:     ErrorMessage
End Sub

Sub ClearReceipt()
        clearRECORD
        
End Sub

Private Sub DisableControls()
On Error GoTo err
        Set rsLedgerdetails = New clsLedgerDetails

        Dim Bval As Boolean
        rsLedgerdetails.disableDATAENTRY

Exit Sub

err:
    ErrorMessage
End Sub


Sub enableCBRECEIPT()
    frmALISMLedgerDetails.cmdCancel.Enabled = True
End Sub

Sub DisableCBRECEIPT()
    frmALISMLedgerDetails.cmdCancel.Enabled = True
End Sub

Private Sub cboPolicyNo_LostFocus()
On Error GoTo err
        If cboPolicyNo.Text = "" Then Exit Sub
            Set rsLedgerdetails = New clsLedgerDetails
            rsLedgerdetails.loadPolicy
            rsLedgerdetails.LoadNAMES
            LoadReceiptGrid
        Exit Sub
err: ErrorMessage
End Sub




Private Sub cmdCancel_Click()
   On Error GoTo err
        enableCBRECEIPT
        ClearReceipt
        DisableControls
        frmALISMLedgerDetails.cboPolicyNo.SetFocus
        Exit Sub
err:
        ErrorMessage
    
End Sub

Private Sub cmdFirstCode_Click(Index As Integer)
    
    On Error GoTo Myerr
             
                With rsDEP
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
        
        clearRECORD
        'disableDATAENTRY
        'loadRECORD
         With rsDEP
                                 
            
             frmALISMLedgerDetails.cboPolicyNo.Text = !PolicyNo & ""
            frmALISMLedgerDetails.txtReceivedTodate.Text = !ReceivedTodate & ""
            frmALISMLedgerDetails.txtpsuspense.Text = !SuspenseAccount & ""
            frmALISMLedgerDetails.txtDateOfLastPayment.Text = !DateOfLastPayment & ""
            frmALISMLedgerDetails.txtReceivedTodate.Text = !ReceivedTodate & ""
            frmALISMLedgerDetails.txtUnitCountBeforePayment.Text = !UnitCountBeforePayment & ""
            frmALISMLedgerDetails.txtDueDate.Text = !DueDate & ""
            frmALISMLedgerDetails.txtStatusCode.Text = !StatusCode & ""
            frmALISMLedgerDetails.txtexpectedpremium.Text = !expectedpremium & ""
            frmALISMLedgerDetails.txtDateofCommencement.Text = !DateofCommencement & ""
            frmALISMLedgerDetails.txtPlanPremium.Text = !PlanPremium & ""
            frmALISMLedgerDetails.txtMaturityDate.Text = !MaturityDate & ""
            frmALISMLedgerDetails.txtlastpaydate.Text = !DateOfLastPayment & ""
            frmALISMLedgerDetails.txtPaymentMode.Text = !paymentMode & ""
            frmALISMLedgerDetails.txtTermofPolicy.Text = !TermOfPolicy & ""
            frmALISMLedgerDetails.cboPaymentMethod.Text = !PaymentMethod & ""
            frmALISMLedgerDetails.txtPremiumDue.Text = !NoofPremiumsdue & ""
            frmALISMLedgerDetails.txtDateOfLastPayment.Text = !DateOfLastPayment & ""
            frmALISMLedgerDetails.txtUnitCountBeforePayment.Text = !UnitCountBeforePayment & ""
            frmALISMLedgerDetails.txtDueDate.Text = !DueDate & ""
            frmALISMLedgerDetails.txtPremiumcount.Text = !UnitCount & ""
        End With
        rsLedgerdetails.LoadNAMES
        frmALISMLedgerDetails.LoadReceiptGrid
Exit Sub

Myerr:
    ErrorMessage
End Sub

Private Sub cmdprintledgerdetails_Click()

On Error GoTo hiserror
    
    Load frmLedgerDetails
    frmLedgerDetails.Show 1, Me
    
Exit Sub

hiserror: ErrorMessage
End Sub

Private Sub Form_Load()
On Error GoTo err

        OpenConnection
        Call DisableControls
        
         
         Set rsDEP = New Recordset
        rsDEP.Open "SELECT * FROM ALISMPolicy where policyNo = '" & CurrentRecord & "';", cnALIS, adOpenKeyset, adLockOptimistic
    Exit Sub

err:
        ErrorMessage
End Sub


Public Sub clearRECORD()

On Error GoTo err:

        With frmALISMLedgerDetails
            
           Set ReceiptGrid.DataSource = Nothing

            .txtDateOfLastPayment.Text = ""
            .txtReceivedTodate.Text = ""
            .txtNames.Text = ""
            .txtUnitCountBeforePayment.Text = ""
            .txtDueDate.Text = ""
            .txtStatusCode.Text = ""
            .txtReceivedTodate.Text = ""
            .cboPolicyNo.Text = ""
            .txtReceivedTodate.Text = ""
            .txtpsuspense.Text = ""
            .txtDateOfLastPayment.Text = ""
            .txtReceivedTodate.Text = ""
            .txtUnitCountBeforePayment.Text = ""
            .txtDueDate.Text = ""
            .txtStatusCode.Text = ""
            .txtexpectedpremium.Text = ""
            .txtDateofCommencement.Text = ""
            .txtPlanPremium.Text = ""
            .txtMaturityDate.Text = ""
            .txtlastpaydate.Text = ""
            .txtPaymentMode.Text = ""
            .txtTermofPolicy.Text = ""
            .cboPaymentMethod.Text = ""
            .txtPremiumDue.Text = ""
            .txtDateOfLastPayment.Text = ""
            .txtUnitCountBeforePayment.Text = ""
            .txtDueDate.Text = ""
            .txtPremiumcount.Text = ""
            .txtagent = ""
            .txtNames = ""
            .txtsex = ""
            .txtdob = ""
    End With

Exit Sub

err:
    ErrorMessage
End Sub




Private Sub Form_Resize()
 
On Error GoTo err
    
Exit Sub
err:
ErrorMessage
End Sub
