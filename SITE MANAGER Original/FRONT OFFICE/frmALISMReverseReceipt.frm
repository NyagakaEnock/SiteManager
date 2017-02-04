VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmALISMReverseReceipt 
   Caption         =   "Reverse Receipt"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Receipt Entries"
      Height          =   2175
      Left            =   120
      TabIndex        =   32
      Top             =   4440
      Width           =   7935
      Begin MSComctlLib.ListView ListView1 
         Height          =   1815
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   3201
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
      Height          =   2175
      Left            =   8160
      TabIndex        =   28
      Top             =   4440
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
         Picture         =   "frmALISMReverseReceipt.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   34
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
         TabIndex        =   30
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
         Picture         =   "frmALISMReverseReceipt.frx":0102
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
         Picture         =   "frmALISMReverseReceipt.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmALISMReverseReceipt.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1200
         Width           =   975
      End
   End
   Begin VB.Frame Frame15 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   9255
      Begin VB.TextBox txtReferenceNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   3120
         Width           =   1575
      End
      Begin VB.ComboBox cboReversalType 
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
         Left            =   1560
         TabIndex        =   40
         Top             =   3547
         Width           =   3015
      End
      Begin VB.TextBox cboBankNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox cboCurrencyCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox cboPaymentMethod 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1185
         Width           =   2295
      End
      Begin VB.TextBox txtReversalDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   3540
         Width           =   1695
      End
      Begin VB.TextBox txtRemark 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   3960
         Width           =   7215
      End
      Begin VB.TextBox txtReceiptDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   360
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   262
         Width           =   1935
      End
      Begin VB.TextBox txtReceiptNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   262
         Width           =   3015
      End
      Begin VB.TextBox txtPayer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   720
         Width           =   7215
      End
      Begin VB.TextBox txtChequeNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtAccountingPeriod 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox txtReceiptAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtLocal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   3120
         Width           =   3015
      End
      Begin VB.TextBox txtPaymentStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox txtTransactionNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtBankNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtPaymentMethod 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1185
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPickerVoidDate 
         Height          =   315
         Left            =   8520
         TabIndex        =   4
         Top             =   3540
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   54919169
         CurrentDate     =   37953
      End
      Begin VB.Label Label9 
         Caption         =   "Reversal Type"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Void Date"
         Height          =   255
         Left            =   5760
         TabIndex        =   36
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Remark"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   4020
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Currency Code "
         Height          =   210
         Left            =   240
         TabIndex        =   27
         Top             =   2715
         Width           =   1215
      End
      Begin VB.Label lblReceiptDate 
         Caption         =   " Receipt Date"
         Height          =   255
         Left            =   5760
         TabIndex        =   26
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label Label27 
         Caption         =   "Receipt No"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Received From"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Pay Method"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1245
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Cheque No"
         Height          =   255
         Left            =   5760
         TabIndex        =   22
         Top             =   1245
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Bank"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1725
         Width           =   495
      End
      Begin VB.Label lblCurrentPeriod 
         Caption         =   "Period"
         Height          =   210
         Left            =   240
         TabIndex        =   20
         Top             =   2250
         Width           =   495
      End
      Begin VB.Label lblReceiptAmount 
         Caption         =   "Amount"
         Height          =   255
         Left            =   5760
         TabIndex        =   19
         Top             =   2220
         Width           =   615
      End
      Begin VB.Label lblReferenceNo 
         Caption         =   "Local?"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   3180
         Width           =   735
      End
      Begin VB.Label lblPaymentStatus 
         Caption         =   "Status"
         Height          =   210
         Left            =   5760
         TabIndex        =   17
         Top             =   2715
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Entries"
         Height          =   255
         Left            =   5760
         TabIndex        =   16
         Top             =   3180
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmALISMReverseReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsREVERSE As clsALISReceipt
Public rsreceipt As clsReceipting

Private Sub cmdAddNew_Click()
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
        Set rsREVERSE = New clsALISReceipt
        rsREVERSE.updateRECORD
        Set reverse = Nothing
End Sub

Private Sub DTPickerVoidDate_Change()
On Error GoTo err
        Me.txtReversalDate.Text = Me.DTPickerVoidDate.Value
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Form_Activate()
    disableALLRECORD
    enableSButtons
    Set rsREVERSE = New clsALISReceipt
    rsREVERSE.loadRECORD
    strSQL = "Select TransactionNo, ReceiptType, TransactionAmount,DocumentNo, TransactionDate, AccountingPeriod, PaymentStatus from ALISMReceiptDetails Where PaymentStatus = 'PAID' and receiptNo = '" & Screen.ActiveForm.txtReceiptNo.Text & "';"
    rsREVERSE.getRECEIPTDETAILS
    Set rsREVERSE = Nothing
    
    Set rsreceipt = New clsReceipting
    rsreceipt.loadBANK
    rsreceipt.loadPAYMENTMETHOD
    Set rsreceipt = Nothing
End Sub

