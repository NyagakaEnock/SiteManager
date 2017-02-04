VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmALISMCheck1 
   Caption         =   "cheque Processing"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7095
      Left            =   0
      TabIndex        =   7
      Top             =   -120
      Width           =   9975
      Begin VB.Frame Frame4 
         Caption         =   "Included Requisitions"
         Height          =   2175
         Left            =   120
         TabIndex        =   34
         Top             =   4800
         Width           =   8655
         Begin MSComctlLib.ListView ListView2 
            Height          =   1695
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   2990
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
      Begin VB.Frame Frame5 
         Height          =   3855
         Left            =   8880
         TabIndex        =   28
         Top             =   3120
         Width           =   975
         Begin VB.CommandButton Command1 
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
            TabIndex        =   43
            Top             =   3270
            Width           =   735
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
            Height          =   450
            Left            =   120
            Picture         =   "frmALISMCheck1.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   33
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
            Picture         =   "frmALISMCheck1.frx":0102
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   570
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
            Picture         =   "frmALISMCheck1.frx":0204
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
            Picture         =   "frmALISMCheck1.frx":0306
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   1020
            Width           =   735
         End
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
            Picture         =   "frmALISMCheck1.frx":0408
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   2370
            Width           =   735
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
            Height          =   450
            Left            =   120
            TabIndex        =   30
            Top             =   2820
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
            Picture         =   "frmALISMCheck1.frx":050A
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   1470
            Width           =   735
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Similar Requisitions"
         Height          =   1695
         Left            =   120
         TabIndex        =   27
         Top             =   3120
         Width           =   8655
         Begin MSComctlLib.ListView ListView3 
            Height          =   1335
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   2355
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
            BackColor       =   16761024
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
      Begin VB.Frame Frame3 
         Height          =   3015
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   9735
         Begin VB.TextBox txtNoOfEntries 
            Alignment       =   1  'Right Justify
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
            Left            =   3960
            TabIndex        =   41
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox txtTotalPaid 
            Alignment       =   1  'Right Justify
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
            Left            =   1560
            TabIndex        =   39
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox txtStatus 
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
            Left            =   1560
            TabIndex        =   37
            Top             =   1584
            Width           =   3735
         End
         Begin VB.TextBox cboRequisitionNo 
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
            Left            =   6960
            TabIndex        =   36
            Top             =   678
            Width           =   2175
         End
         Begin VB.TextBox txtChequeNo 
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
            TabIndex        =   1
            Top             =   225
            Width           =   3735
         End
         Begin VB.TextBox txtChequeDate 
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
            Left            =   6960
            TabIndex        =   14
            Top             =   225
            Width           =   1935
         End
         Begin VB.ComboBox cboBankNo 
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
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   678
            Width           =   3735
         End
         Begin VB.TextBox txtPayeeDetails 
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
            TabIndex        =   3
            Top             =   1131
            Width           =   3735
         End
         Begin VB.TextBox txtDocumentNo 
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
            Left            =   6960
            TabIndex        =   13
            Top             =   1584
            Width           =   2175
         End
         Begin VB.TextBox txtRequisitionDate 
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
            Left            =   6960
            TabIndex        =   12
            Top             =   1131
            Width           =   2175
         End
         Begin VB.TextBox txtChequeAmount 
            Alignment       =   1  'Right Justify
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
            Left            =   6960
            TabIndex        =   11
            Top             =   2520
            Width           =   2175
         End
         Begin VB.TextBox txtAmountPaid 
            Alignment       =   1  'Right Justify
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
            Left            =   3960
            TabIndex        =   4
            Top             =   2520
            Width           =   1335
         End
         Begin VB.TextBox txtPaymentFlag 
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
            Left            =   6960
            TabIndex        =   10
            Top             =   2040
            Width           =   2175
         End
         Begin VB.TextBox txtAmountDue 
            Alignment       =   1  'Right Justify
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
            Left            =   1560
            TabIndex        =   9
            Top             =   2520
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DTPickerChequeDate 
            Height          =   375
            Left            =   8880
            TabIndex        =   15
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Format          =   19791873
            CurrentDate     =   37945
         End
         Begin VB.Label Label11 
            Caption         =   "Entries"
            Height          =   255
            Left            =   3120
            TabIndex        =   42
            Top             =   2093
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Total Paid"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   2093
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Status"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   1637
            Width           =   975
         End
         Begin VB.Label Label15 
            Caption         =   "Cheque No"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label21 
            Caption         =   "Chk Date"
            Height          =   255
            Left            =   5640
            TabIndex        =   25
            Top             =   285
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "Bank No"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   731
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Payee "
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   1184
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Requisition No"
            Height          =   255
            Left            =   5640
            TabIndex        =   22
            Top             =   731
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Document #"
            Height          =   255
            Left            =   5640
            TabIndex        =   21
            Top             =   1637
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Req Date"
            Height          =   255
            Left            =   5640
            TabIndex        =   20
            Top             =   1184
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Chk Amount"
            Height          =   255
            Left            =   5640
            TabIndex        =   19
            Top             =   2573
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Amt Paid"
            Height          =   255
            Left            =   3120
            TabIndex        =   18
            Top             =   2573
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "Payment Flag"
            Height          =   255
            Left            =   5640
            TabIndex        =   17
            Top             =   2093
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Amount Due"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   2573
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmALISMCheck1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsREQUISITION As clsALISCheque

Private Sub cbobankNo_GotFocus()
        strSQL = "SELECT * FROM ALISPBankAccount"
        bankNoGotFocus
End Sub

Private Sub cboBankNo_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub cbobankNo_LostFocus()
        strSQL = "SELECT * FROM ALISPBankAccount WHERE Details = '" & cbobankNo.Text & "'"
        BankNoLostFocus
End Sub

Private Sub cmdAddNew_Click()
        bmakePAYMENT = True
        breversePAYMENT = False
        Set rsREQUISITION = New clsALISCheque
        rsREQUISITION.addRECORD
        Set rsREQUISITION = Nothing
      
End Sub


Private Sub cmdCancel_Click()
        Set rsREQUISITION = New clsALISCheque
        rsREQUISITION.Cancelrecord
        Set rsREQUISITION = Nothing
        bmakePAYMENT = False
        breversePAYMENT = False

End Sub




Private Sub cmdUpdate_Click()
        Set rsREQUISITION = New clsALISCheque
        bsaveRECORD = False
        rsREQUISITION.UpdateALLRECORDS
        Set rsREQUISITION = Nothing

End Sub

Private Sub Command1_Click()
        Set rsREQUISITION = New clsALISCheque
        rsREQUISITION.contructDATA
        Set rsREQUISITION = Nothing

End Sub

Private Sub DTPickerChequeDate_Change()
        Set rsREQUISITION = New clsALISCheque
        rsREQUISITION.ChangeDATE
        Set rsREQUISITION = Nothing

End Sub

Private Sub Form_Activate()
        
        Set rsREQUISITION = New clsALISCheque
                disableALLRECORD

        If bApproveCheque = True Or bAuthorizeCheque = True Then
                rsREQUISITION.loadRECORD
                viewButtons
        Else
                bmakePAYMENT = True
                breverpayment = False
                rsREQUISITION.loadRECORD
                rsREQUISITION.GetPreviousPayment
                enableButtons
        End If
        
        Set rsREQUISITION = Nothing

End Sub

Private Sub Form_Load()
        OpenConnection
End Sub


Private Sub Form_Unload(Cancel As Integer)
        bmakePAYMENT = False
        breversePAYMENT = False
End Sub


Private Sub txtAmountPaid_LostFocus()
        Set rsREQUISITION = New clsALISCheque
        rsREQUISITION.checkSTATUS
        Set rsREQUISITION = Nothing
End Sub
