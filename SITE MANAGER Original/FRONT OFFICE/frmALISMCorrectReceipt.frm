VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmALISMCorrectReceipt 
   Caption         =   "Post Commission"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Receipt Entries"
      Height          =   2535
      Left            =   120
      TabIndex        =   31
      Top             =   3240
      Width           =   7935
      Begin MSComctlLib.ListView ListView1 
         Height          =   2175
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   3836
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
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
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   8160
      TabIndex        =   27
      Top             =   3240
      Width           =   1215
      Begin VB.CommandButton cmdSearch 
         Appearance      =   0  'Flat
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
         Picture         =   "frmALISMCorrectReceipt.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
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
         TabIndex        =   29
         Top             =   3120
         Width           =   2295
      End
      Begin VB.CommandButton cmdAddNew 
         Appearance      =   0  'Flat
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
         Picture         =   "frmALISMCorrectReceipt.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Appearance      =   0  'Flat
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
         Picture         =   "frmALISMCorrectReceipt.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
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
         Picture         =   "frmALISMCorrectReceipt.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1200
         Width           =   975
      End
   End
   Begin VB.Frame Frame15 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   9255
      Begin VB.TextBox txtAgentStatus 
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
         Height          =   315
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   1680
         Width           =   1575
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
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1680
         Width           =   2895
      End
      Begin VB.ComboBox cboAgentNo 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1560
         TabIndex        =   46
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtCommissionRecords 
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
         Height          =   315
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox txtUnitsPaid 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox cboDocumentNo 
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
         Height          =   315
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   255
         Width           =   1575
      End
      Begin VB.TextBox cboBankNo 
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
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1320
         Width           =   6135
      End
      Begin VB.TextBox cboCurrencyCode 
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
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox cboPaymentMethod 
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
         Height          =   315
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox txtUnitCount 
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
         Height          =   315
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtExpectedAmount 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtReceiptDate 
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
         Height          =   315
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   255
         Width           =   1215
      End
      Begin VB.TextBox txtReceiptNo 
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
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   255
         Width           =   1815
      End
      Begin VB.TextBox txtPayer 
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
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox txtChequeNo 
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
         Height          =   315
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtAccountingPeriod 
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
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtReceiptAmount 
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
         Height          =   315
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtLocal 
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
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox txtPaymentStatus 
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
         Height          =   315
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtTransactionNo 
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
         Height          =   315
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtBankNo 
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
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtPaymentMethod 
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
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Agent Status"
         Height          =   255
         Left            =   6120
         TabIndex        =   49
         Top             =   1710
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Agent No"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   1710
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Comm Recs"
         Height          =   255
         Left            =   6120
         TabIndex        =   44
         Top             =   2790
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Units Paid"
         Height          =   255
         Left            =   6120
         TabIndex        =   42
         Top             =   2430
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Policy No"
         Height          =   255
         Left            =   6120
         TabIndex        =   39
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Unit count"
         Height          =   255
         Left            =   6120
         TabIndex        =   35
         Top             =   2070
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Expected Premium"
         Height          =   255
         Left            =   6120
         TabIndex        =   30
         Top             =   990
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Currency Code "
         Height          =   210
         Left            =   240
         TabIndex        =   26
         Top             =   2445
         Width           =   1215
      End
      Begin VB.Label lblReceiptDate 
         Caption         =   " Receipt Date"
         Height          =   255
         Left            =   3600
         TabIndex        =   25
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label27 
         Caption         =   "Receipt No"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Received From"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Pay Method"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   990
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Cheque No"
         Height          =   255
         Left            =   6120
         TabIndex        =   21
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Bank"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1365
         Width           =   495
      End
      Begin VB.Label lblCurrentPeriod 
         Caption         =   "Period"
         Height          =   210
         Left            =   240
         TabIndex        =   19
         Top             =   2085
         Width           =   495
      End
      Begin VB.Label lblReceiptAmount 
         Caption         =   "Amount"
         Height          =   255
         Left            =   3600
         TabIndex        =   18
         Top             =   2070
         Width           =   615
      End
      Begin VB.Label lblReferenceNo 
         Caption         =   "Local?"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2790
         Width           =   735
      End
      Begin VB.Label lblPaymentStatus 
         Caption         =   "Status"
         Height          =   210
         Left            =   3600
         TabIndex        =   16
         Top             =   2445
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Entries"
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   2790
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmALISMCorrectReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsREVERSE As clsALISReceipt
Public rsreceipt As clsReceipting

Private Sub cboAgentNo_gotFocus()
'    selectAgentNo_Gotfocus
End Sub

Private Sub cboAgentNo_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboAgentNo_LostFocus()
'    selectAgentNo_LostFocus
End Sub

Private Sub cmdAddNew_Click()
        bCorrectCommission = True
        Set rsREVERSE = New clsALISReceipt
        rsREVERSE.addRECORD
        Set reverse = Nothing
End Sub
Private Sub cboReversalType_GotFocus()
'        selectReversalGotFocus
End Sub

Private Sub cboReversalType_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub cboReversalType_LostFocus()
'        selectReversalLostFocus
End Sub

Private Sub cmdCancel_Click()
        Set rsREVERSE = New clsALISReceipt
        rsREVERSE.Cancelrecord
        Set reverse = Nothing
        bCorrectCommission = False
End Sub

Private Sub locateRECORD()
''''On Error GoTo Myerr

        Dim strQRE As Variant
        Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset
            strQRE = InputBox("Enter Receipt No to search.", "Search Value")
            rsFind.Open "SELECT * FROM ALISMReceiptNew WHERE ReceiptNo = '" & strQRE & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsFind
            If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
            Else:
                    Screen.ActiveForm.txtReceiptNo.Text = !ReceiptNo
                    Screen.ActiveForm.txtReceiptDate.Text = !ReceiptDate
                    Screen.ActiveForm.txtReceiptAmount.Text = !ReceiptAmount
                    Screen.ActiveForm.txtAccountingPeriod.Text = !AccountingPeriod
                    Screen.ActiveForm.txtPaymentStatus.Text = !PaymentStatus
                    Screen.ActiveForm.txtChequeNo.Text = !ChequeNo & ""
                    Screen.ActiveForm.txtPaymentMethod.Text = !PaymentMethod & ""
                    Screen.ActiveForm.txtBankNo.Text = !BankNo & ""
                    Screen.ActiveForm.txtPayer.Text = !Payer
                    Screen.ActiveForm.txtLocal.Text = !LocalCheque & ""
                    Screen.ActiveForm.cboCurrencyCode.Text = !CurrencyCode & ""
                    Screen.ActiveForm.txtTotalAmount.Text = !TotalAmount & ""
                    Screen.ActiveForm.txtBalance.Text = (!ReceiptAmount - !TotalAmount) & ""
                    Screen.ActiveForm.txtReceiptNo.Text = !ReceiptNo
                    Screen.ActiveForm.txtReceiptAmountDetails.Text = !ReceiptAmount
                    Screen.ActiveForm.txtTransactionNo.Text = !TransactionNo & ""
                    Screen.ActiveForm.txtTransactionNo.Text = !TransactionNo & ""
                    Screen.ActiveForm.txtBalance.Text = !ReceiptAmount
            End If
        End With

Exit Sub

Myerr:
    ErrorMessage
End Sub

Private Sub cmdUpdate_Click()
        If bsaveRECORD = True Then
                rsCOMMISSION.CalculateCommission
                disableALLRECORD
                enableSButtons
        End If
        bCorrectCommission = False
End Sub

Private Sub DTPickerVoidDate_Change()
On Error GoTo err
        Screen.ActiveForm.txtVoidDate.Text = Screen.ActiveForm.DTPickerVoidDate.Value
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Form_Activate()
    disableALLRECORD
    enableSButtons
    Set rsREVERSE = New clsALISReceipt
    rsREVERSE.loadCOMMISSION
    rsREVERSE.loadPREMIUM
    rsREVERSE.loadAGENT
    strSQL = "Select TransactionNo, ReceiptType, TransactionAmount,DocumentNo, TransactionDate, AccountingPeriod, PaymentStatus from ALISMReceiptDetails Where PaymentStatus = 'PAID' and receiptNo = '" & Screen.ActiveForm.txtReceiptNo.Text & "';"
    Set rsREVERSE = Nothing
    
    Set rsreceipt = New clsReceipting
    rsreceipt.loadBANK
    rsreceipt.loadPAYMENTMETHOD
    Set rsreceipt = Nothing
End Sub

Private Sub Form_Unload(cancel As Integer)
    'getCOMMRECEIPT
End Sub
