VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmALISMBankReceipt 
   Caption         =   "Bank Receipts"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   10035
   Icon            =   "frmALISMBankReceipts.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   10035
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      Begin VB.Frame Frame5 
         Caption         =   "Receipts Banked"
         Height          =   2535
         Left            =   120
         TabIndex        =   16
         Top             =   4200
         Width           =   9855
         Begin VB.TextBox txtNoofEntries 
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
            Height          =   375
            Left            =   960
            TabIndex        =   25
            Top             =   2040
            Width           =   2295
         End
         Begin VB.TextBox txtTotalAmount 
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
            Height          =   375
            Left            =   4680
            TabIndex        =   23
            Top             =   2040
            Width           =   2295
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   1695
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   9495
            _ExtentX        =   16748
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
         Begin VB.Label Label7 
            Caption         =   "Entries"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   2100
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "Total Scheduled"
            Height          =   255
            Left            =   3360
            TabIndex        =   24
            Top             =   2100
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Receipt Not Scheduled"
         Height          =   2055
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   9855
         Begin MSComctlLib.ListView ListView1 
            Height          =   1695
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   9495
            _ExtentX        =   16748
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
      Begin VB.Frame Frame4 
         Height          =   1455
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   9855
         Begin VB.TextBox txtAccountNo 
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
            Left            =   5160
            TabIndex        =   28
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox txtBankNo 
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
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPickerDateBanked 
            Height          =   375
            Left            =   9360
            TabIndex        =   22
            Top             =   360
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Format          =   57147393
            CurrentDate     =   38189
         End
         Begin VB.TextBox txtDateBanked 
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
            Left            =   7560
            TabIndex        =   20
            Top             =   360
            Width           =   1815
         End
         Begin VB.ComboBox cboTemplateCode 
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
            Left            =   4440
            TabIndex        =   18
            Top             =   840
            Width           =   5175
         End
         Begin VB.TextBox txtReference 
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
            Height          =   375
            Left            =   1200
            TabIndex        =   10
            Top             =   840
            Width           =   1935
         End
         Begin VB.TextBox txtDetails 
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
            Left            =   1800
            TabIndex        =   1
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Date Banked"
            Height          =   195
            Left            =   8040
            TabIndex        =   21
            Top             =   120
            Width           =   945
         End
         Begin VB.Label Label2 
            Caption         =   "Template Code"
            Height          =   255
            Left            =   3240
            TabIndex        =   19
            Top             =   900
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Account No"
            Height          =   195
            Left            =   5880
            TabIndex        =   15
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Reference No"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   900
            Width           =   1455
         End
         Begin VB.Label lbPolicyNo 
            AutoSize        =   -1  'True
            Caption         =   "Bank No"
            Height          =   195
            Left            =   480
            TabIndex        =   9
            Top             =   120
            Width           =   630
         End
         Begin VB.Label lbNames 
            Caption         =   "Account Details"
            Height          =   255
            Left            =   2640
            TabIndex        =   8
            Top             =   105
            Width           =   1215
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
         Height          =   855
         Left            =   120
         TabIndex        =   3
         Top             =   -120
         Width           =   9855
         Begin VB.CommandButton cmdPrintletter 
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
            Left            =   1920
            Picture         =   "frmALISMBankReceipts.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   615
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
            Height          =   495
            Left            =   2520
            Picture         =   "frmALISMBankReceipts.frx":0974
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   615
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
            Left            =   1320
            Picture         =   "frmALISMBankReceipts.frx":0EA6
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   615
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
            Left            =   720
            Picture         =   "frmALISMBankReceipts.frx":13D8
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   615
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
            Picture         =   "frmALISMBankReceipts.frx":190A
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "frmALISMBankReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCode As ADODB.Recordset, strcode, iAccountNo As String, bAddNew As Boolean
Public rsreceipt As New clsALISBankReceipt

Private Sub cboTemplateCode_GotFocus()
        strSQL = "Select * from ALISPLetterTemplate, ALISPLetterReceipient Where ALISPLetterReceipient.Banker = '1' and ALISPLetterTemplate.ReceipientCode = ALISPLetterReceipient.ReceipientCode "
        SelectTemplateGotFocus
End Sub

Private Sub cboTemplateCode_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub cboTemplateCode_LostFocus()
        strSQL = "Select * from ALISPLetterTemplate Where ALISPLetterTemplate.TemplateDescription = '" & Screen.ActiveForm.cboTemplateCode.Text & "';"
        selectTemplateLostFocus
End Sub

Private Sub cmdAddNew_Click()
        Set rsreceipt = New clsALISBankReceipt
        rsreceipt.addSCHEDULE
        Set rsreceipt = Nothing
              
End Sub

Private Sub cmdCancel_Click()
        Set rsreceipt = New clsALISBankReceipt
        rsreceipt.cancelSCHEDULE
        Set rsreceipt = Nothing

End Sub

Private Sub cmdPrintletter_Click()
        If frmALISMBankReceipt.txtReference.Text <= "" Then Exit Sub
        Load frmALISRBankReceipts
        frmALISRBankReceipts.Show 1, Me
End Sub

Private Sub cmdUpdate_Click()
        Set rsreceipt = New clsALISBankReceipt
        rsreceipt.updateSCHEDULE
        Set rsreceipt = Nothing
End Sub


Private Sub DTPickerDateBanked_Change()
    With Screen.ActiveForm
        .txtDateBanked.Text = .DTPickerDateBanked.Value
    End With
End Sub

Private Sub Form_Activate()
    enableSButtons
    disableALLRECORD
    
    Set rsreceipt = New clsALISBankReceipt
    rsreceipt.loadSCHEDULEDRECORD
    rsreceipt.GetBankedReceipts
    rsreceipt.getRECEIPT
    Set rsreceipt = Nothing
End Sub


