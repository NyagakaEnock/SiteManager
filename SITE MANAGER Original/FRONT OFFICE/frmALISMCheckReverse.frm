VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmALISMCheckReverse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reverse Cheque"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6855
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   9975
      Begin VB.Frame Frame5 
         Height          =   3015
         Left            =   8760
         TabIndex        =   23
         Top             =   3840
         Width           =   1095
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
            Picture         =   "frmALISMCheckReverse.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   2040
            Width           =   855
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
            Picture         =   "frmALISMCheckReverse.frx":0102
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   690
            Width           =   855
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
            Picture         =   "frmALISMCheckReverse.frx":0204
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   240
            Width           =   855
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
            Picture         =   "frmALISMCheckReverse.frx":0306
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1155
            Width           =   855
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
            Picture         =   "frmALISMCheckReverse.frx":0408
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   2520
            Width           =   855
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
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   3570
            Width           =   855
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
            Picture         =   "frmALISMCheckReverse.frx":050A
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   1605
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Similar Requisitions"
         Height          =   2895
         Left            =   120
         TabIndex        =   22
         Top             =   3840
         Width           =   8535
         Begin MSComctlLib.ListView ListView3 
            Height          =   2535
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   4471
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
            BackColor       =   16777152
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
         Height          =   3735
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   9735
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
            Left            =   6960
            TabIndex        =   39
            Top             =   2040
            Width           =   2175
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
            Sorted          =   -1  'True
            TabIndex        =   37
            Top             =   2040
            Width           =   3735
         End
         Begin VB.TextBox txtComment 
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
            Height          =   600
            Left            =   1560
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Top             =   3000
            Width           =   6975
         End
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
            Left            =   6960
            TabIndex        =   33
            Top             =   1143
            Width           =   2175
         End
         Begin VB.TextBox cboRequisitionNo 
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
            TabIndex        =   31
            Top             =   684
            Width           =   2175
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
            TabIndex        =   29
            Top             =   1611
            Width           =   3735
         End
         Begin VB.TextBox txtChequeNo 
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
            TabIndex        =   12
            Top             =   225
            Width           =   1935
         End
         Begin VB.ComboBox cboBankNo 
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
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   687
            Width           =   3735
         End
         Begin VB.TextBox txtPayeeDetails 
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
            TabIndex        =   3
            Top             =   1149
            Width           =   3735
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
            TabIndex        =   4
            Top             =   1602
            Width           =   2175
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
            Left            =   4080
            TabIndex        =   10
            Top             =   2535
            Width           =   1215
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
            Top             =   2535
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DTPickerChequeDate 
            Height          =   375
            Left            =   8880
            TabIndex        =   13
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Format          =   56492033
            CurrentDate     =   37945
         End
         Begin VB.Label Label2 
            Caption         =   "TotalPaid"
            Height          =   255
            Left            =   5640
            TabIndex        =   40
            Top             =   2115
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Reversal Type"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   2126
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "Comment"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   3180
            Width           =   735
         End
         Begin VB.Label Label11 
            Caption         =   "Entries"
            Height          =   255
            Left            =   5640
            TabIndex        =   34
            Top             =   1196
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Requisition No"
            Height          =   255
            Left            =   5640
            TabIndex        =   32
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label9 
            Caption         =   "Status"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1664
            Width           =   975
         End
         Begin VB.Label Label15 
            Caption         =   "Cheque No"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   278
            Width           =   1215
         End
         Begin VB.Label Label21 
            Caption         =   "Chk Date"
            Height          =   255
            Left            =   5640
            TabIndex        =   20
            Top             =   278
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "Bank No"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   740
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Payee "
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Chk Amount"
            Height          =   255
            Left            =   5640
            TabIndex        =   17
            Top             =   2573
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Amount Paid"
            Height          =   255
            Left            =   5640
            TabIndex        =   16
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Payment Flag"
            Height          =   255
            Left            =   3000
            TabIndex        =   15
            Top             =   2595
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Amount Due"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   2580
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmALISMCheckReverse"
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
        strSQL = "SELECT * FROM ALISPBankAccount WHERE Details = '" & cboBankNo.Text & "'"
        BankNoLostFocus
End Sub

Private Sub cboReversalType_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub cmdAddNew_Click()
        bmakePAYMENT = False
        breversePAYMENT = True
        Set rsREQUISITION = New clsALISCheque
        rsREQUISITION.addRECORD
        Set rsREQUISITION = Nothing

End Sub

Private Sub cmdCancel_Click()
        Set rsREQUISITION = New clsALISCheque
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

Private Sub DTPickerChequeDate_Change()
        Set rsREQUISITION = New clsALISCheque
        rsREQUISITION.ChangeDATE
        Set rsREQUISITION = Nothing

End Sub

Private Sub Form_Activate()
        breversePAYMENT = True
        bmakePAYMENT = False
        Set rsREQUISITION = New clsALISCheque
        Set rsREQUISITION = Nothing
        disableALLRECORD
        enableButtons
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
