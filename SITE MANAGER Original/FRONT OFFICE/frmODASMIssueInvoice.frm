VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmODASMIssueInvoice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Invoice"
   ClientHeight    =   7980
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7875
   Icon            =   "frmODASMIssueInvoice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog HelpCommonDialog 
      Left            =   720
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.Frame Frame3 
         Height          =   1695
         Left            =   120
         TabIndex        =   15
         Top             =   5640
         Width           =   7335
         Begin VB.TextBox txtPriceExclusive 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   5160
            TabIndex        =   19
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtVATAmount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   5160
            TabIndex        =   18
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox txtPriceInclusive 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   5160
            TabIndex        =   17
            Top             =   1200
            Width           =   1935
         End
         Begin VB.TextBox txtRemark 
            BackColor       =   &H00FFFFC0&
            Height          =   1275
            Left            =   840
            MaxLength       =   120
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label9 
            Caption         =   "InVoice Amount"
            Height          =   255
            Left            =   3840
            TabIndex        =   23
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   "VAT Amount"
            Height          =   255
            Left            =   3840
            TabIndex        =   22
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label21 
            Caption         =   "Amount Incl"
            Height          =   255
            Left            =   3840
            TabIndex        =   21
            Top             =   1230
            Width           =   1215
         End
         Begin VB.Label Label30 
            Caption         =   "Remark"
            ForeColor       =   &H00004040&
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Invoice Sent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2775
         Left            =   120
         TabIndex        =   13
         Top             =   2880
         Width           =   7335
         Begin MSComctlLib.ListView ListView1 
            Height          =   2415
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4260
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
      Begin VB.Frame Frame2 
         Caption         =   "Client Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   7335
         Begin VB.TextBox txtCompanyName 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   1800
            TabIndex        =   11
            Top             =   240
            Width           =   3975
         End
         Begin VB.TextBox txtAccountNo 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   720
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtCurrentPeriod 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   5760
            TabIndex        =   9
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   " Name"
            ForeColor       =   &H00004040&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Invoice Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   7335
         Begin VB.TextBox txtInvoiceNo 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   1560
            TabIndex        =   4
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox txtInvoiceDate 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   4800
            TabIndex        =   3
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox txtInvoiceDescription 
            BackColor       =   &H00FFFFC0&
            Height          =   1035
            Left            =   1560
            MaxLength       =   120
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   840
            Width           =   5535
         End
         Begin VB.Label Label29 
            Caption         =   "InVoice No"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "Invoice Date"
            Height          =   255
            Left            =   3600
            TabIndex        =   6
            Top             =   390
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Description"
            ForeColor       =   &H00004040&
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   840
            Width           =   975
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClear 
         Caption         =   "Clear the &Screen"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnumm 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuShow 
      Caption         =   "&Show/View"
      Begin VB.Menu mnuRegisteredClients 
         Caption         =   "Registered Clients"
      End
      Begin VB.Menu mnuKHJGGFDHJ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowQuotations 
         Caption         =   "Show Quotations"
      End
      Begin VB.Menu mnuExtraInfo 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExtraInformation 
         Caption         =   "Extra Inform"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help System"
      Begin VB.Menu mnuHow 
         Caption         =   "How to use this System"
         Shortcut        =   ^{F1}
      End
   End
End
Attribute VB_Name = "frmODASMIssueInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsINVOICE As clsODASMAccounts


Private Sub cboVATRate_GotFocus()
        selectVATRATE_GotFocus
End Sub

Private Sub cboVATRate_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub cboVATRate_LostFocus()
        selectVATRate_LostFocus
        Set rsINVOICE = New clsODASMAccounts
        rsINVOICE.calculateVAT
        Set rsINVOICE = Nothing
End Sub

Private Sub Form_Activate()
        
        Set rsINVOICE = New clsODASMAccounts
        rsINVOICE.loadRECORD
        rsINVOICE.LoadDEFAULT
        rsINVOICE.loadINSTALLMENT
        rsINVOICE.calculateVAT
        rsINVOICE.updatePRICEExclusive
        rsINVOICE.updateVAT
        rsINVOICE.updatePRICEinclusive
        disableALLRECORD
        
        showBRIEFRECEIPTS
        showBRIEFINVOICESsenT
        showBRIEFINVOICESRECeived
        rsINVOICE.calculateTOTALRECEIPTS
        rsINVOICE.calculateInvoicesSend
        Set rsINVOICE = Nothing
        showBRIEFINACCOUNT
        showINVOICEitems
    Set rsINVOICE = Nothing
End Sub

Private Sub Form_Initialize()
        Set rsINVOICE = New clsODASMAccounts
End Sub


Private Sub Form_Unload(cancel As Integer)
        If NewRecord = True Then
            cancel = True
            MsgBox "Data entry in progress. Click Refresh to Cancel", vbCritical
        Else
            cancel = False
        End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        
        'If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then

        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView1.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                End If
            Next i
            

        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub Form_Load()
        OpenConnection
End Sub

Private Sub Form_Terminate()
       Set rsINVOICE = Nothing
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        
        'If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then

        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView2.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView2.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView2.ListItems(i).Checked = False
                End If
            Next i
            
            
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ListView3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView3.SortKey = ColumnHeader.Index - 1
    ListView3.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView3_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        
        'If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then

        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView3.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView3.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView3.ListItems(i).Checked = False
                End If
            Next i
            
            
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ListView4_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView4.SortKey = ColumnHeader.Index - 1
    ListView4.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView4_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        
        'If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then

        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView4.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView4.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView4.ListItems(i).Checked = False
                End If
            Next i
            
            
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ListView5_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView5.SortKey = ColumnHeader.Index - 1
    ListView5.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView5_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        
        'If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then

        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView5.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView5.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView5.ListItems(i).Checked = False
                End If
            Next i
            
            frmODASMAccounts.txtPaymentMethod.Text = Item.Text

        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuExtraInformation_Click()
        Load frmODASMInformation
        frmODASMInformation.Show 1, Me
End Sub


Private Sub mnuHow_Click()
            Me.HelpCommonDialog.DialogTitle = "Using the Main System"
            Me.HelpCommonDialog.HelpFile = App.HelpFile
            Me.HelpCommonDialog.HelpContext = 71
            Me.HelpCommonDialog.HelpCommand = cdlHelpContext
            Me.HelpCommonDialog.ShowHelp

End Sub
