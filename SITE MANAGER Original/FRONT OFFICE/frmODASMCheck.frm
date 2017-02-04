VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmODASMCheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process Cheque"
   ClientHeight    =   7890
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "frmODASMCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6975
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   11895
      Begin VB.Frame Frame5 
         Caption         =   "Company's Banks"
         Height          =   2535
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   5175
         Begin MSComctlLib.ListView ListView1 
            Height          =   2055
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   3625
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
      Begin VB.Frame Frame2 
         Caption         =   "Voucher Details"
         Height          =   2415
         Left            =   5400
         TabIndex        =   19
         Top             =   2640
         Width           =   6255
         Begin VB.TextBox txtCostCenter 
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
            Left            =   1200
            TabIndex        =   43
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txtCostCenterDescription 
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
            Left            =   2400
            TabIndex        =   42
            Top             =   1440
            Width           =   3495
         End
         Begin VB.TextBox txtRemark 
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1200
            ScrollBars      =   2  'Vertical
            TabIndex        =   38
            Top             =   1800
            Width           =   4695
         End
         Begin VB.TextBox txtChequeEntryNo 
            Alignment       =   1  'Right Justify
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
            Left            =   4440
            TabIndex        =   33
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtAmountDue 
            Alignment       =   1  'Right Justify
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
            Left            =   1200
            TabIndex        =   30
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox txtVoucherNo 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   1200
            TabIndex        =   23
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtDocumentNo 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   4440
            TabIndex        =   22
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtVoucherDate 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   1200
            TabIndex        =   21
            Top             =   690
            Width           =   1575
         End
         Begin VB.TextBox txtPaymentFlag 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   4440
            TabIndex        =   20
            Top             =   690
            Width           =   1455
         End
         Begin VB.Label Label13 
            Caption         =   "Cost Center"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   1455
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Remark"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   1875
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Amount"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1110
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "Item No"
            Height          =   255
            Left            =   3120
            TabIndex        =   34
            Top             =   1110
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Voucher No"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Document #"
            Height          =   255
            Left            =   3120
            TabIndex        =   26
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Voucher Date"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Payment Flag"
            Height          =   255
            Left            =   3120
            TabIndex        =   24
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2535
         Left            =   5400
         TabIndex        =   6
         Top             =   120
         Width           =   6255
         Begin VB.TextBox txtAccountNo 
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
            Left            =   4440
            TabIndex        =   40
            Top             =   1110
            Width           =   1455
         End
         Begin VB.TextBox txtChequeAmount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   1200
            TabIndex        =   35
            Top             =   2040
            Width           =   1575
         End
         Begin VB.TextBox txtBankName 
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
            Left            =   1920
            TabIndex        =   32
            Top             =   720
            Width           =   3975
         End
         Begin VB.TextBox txtBankNo 
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
            Left            =   1200
            TabIndex        =   31
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtPayeeDetails 
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
            Left            =   1200
            TabIndex        =   11
            Top             =   1605
            Width           =   4695
         End
         Begin VB.TextBox txtChequeDate 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   4440
            TabIndex        =   10
            Top             =   225
            Width           =   1215
         End
         Begin VB.TextBox txtChequeNo 
            BackColor       =   &H00FFC0C0&
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
            Left            =   1200
            TabIndex        =   9
            Top             =   225
            Width           =   1695
         End
         Begin VB.TextBox txtStatus 
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
            Left            =   1200
            TabIndex        =   8
            Top             =   1110
            Width           =   1575
         End
         Begin VB.TextBox txtNoOfEntries 
            Alignment       =   1  'Right Justify
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
            Left            =   4440
            TabIndex        =   7
            Top             =   2040
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker DTPickerChequeDate 
            Height          =   315
            Left            =   5640
            TabIndex        =   12
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            Format          =   57016321
            CurrentDate     =   37945
         End
         Begin VB.Label Label10 
            Caption         =   "Account No"
            Height          =   255
            Left            =   3240
            TabIndex        =   41
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Chk Amount"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Payee "
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1665
            Width           =   975
         End
         Begin VB.Label Label16 
            Caption         =   "Bank No"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   735
            Width           =   735
         End
         Begin VB.Label Label21 
            Caption         =   "Chk Date"
            Height          =   255
            Left            =   3240
            TabIndex        =   16
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   "Cheque No"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Status"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "Entries"
            Height          =   255
            Left            =   3240
            TabIndex        =   13
            Top             =   2100
            Width           =   735
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Similar Requisitions"
         Height          =   2415
         Left            =   120
         TabIndex        =   4
         Top             =   2640
         Width           =   5175
         Begin MSComctlLib.ListView ListView3 
            Height          =   2055
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   3625
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
      Begin VB.Frame Frame4 
         Caption         =   "Cheque Entries"
         Height          =   1815
         Left            =   120
         TabIndex        =   2
         Top             =   5040
         Width           =   11535
         Begin MSComctlLib.ListView ListView2 
            Height          =   1455
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   2566
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
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMCheck.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMCheck.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMCheck.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMCheck.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMCheck.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMCheck.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMCheck.frx":236E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   1164
      ButtonWidth     =   3069
      ButtonHeight    =   1005
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New &Record "
            Key             =   "N"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Edit/Change "
            Key             =   "E"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Search/Find "
            Key             =   "S"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Refresh "
            Key             =   "R"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print &Preview "
            Key             =   "P"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Help"
            Key             =   "F"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComDlg.CommonDialog HelpCommonDialog 
         Left            =   9360
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         HelpCommand     =   11
         HelpContext     =   72
         HelpFile        =   "REGHELP.HLP"
         PrinterDefault  =   0   'False
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
Attribute VB_Name = "frmODASMCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCHECK As clsALISCheque

Private Sub Form_Initialize()
        Set rsCHECK = New clsALISCheque
End Sub

Private Sub Form_Terminate()
        Set rsCHECK = Nothing
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
            
            frmODASMCheck.txtChequeEntryNo.Text = Item.Text

            rsCHECK.loadCHEQUE
            
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
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
            
            frmODASMCheck.txtBankNo.Text = Item.Text
            frmODASMCheck.txtBankName.Text = Item.SubItems(3)
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
            
            frmODASMCheck.txtVoucherNo.Text = Item.Text
            rsCHECK.loadVOUCHER
            rsCHECK.calculateTOTALPAID

        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
        
        With frmODASMCheck
        
        Select Case Button.Key
                Case "N"
                    Select Case Button.Caption
                    
                    Case "New &Record "
                            If editRECORD Then Exit Sub
                            NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4:
                            bmakePAYMENT = True
                            breversePAYMENT = False
                            rsCHECK.addRECORD
                            rsCHECK.enableRECORD
                    Case "&Save Record "
                    
                            bsaveRECORD = False
                            rsCHECK.UpdateALLRECORDS
                                    
                            If bsaveRECORD = True Then
                                .Toolbar1.Buttons(3).Caption = "&NEXT ITEM ": .Toolbar1.Buttons(3).Image = 2
                                .Toolbar1.Buttons(4).Caption = "FINISH"
                                disableALLRECORD
                            End If
                    
                    Case Else
                        Exit Sub
                    End Select
    
        Case "E"
            Select Case Button.Caption
                Case "&Edit/Change "
                
                Case "&NEXT ITEM "
                        
                        NewRecord = True: .Toolbar1.Buttons(2).Caption = "&Save Record ": .Toolbar1.Buttons(2).Image = 4
                        bmakePAYMENT = True
                        breversePAYMENT = False
                        rsCHECK.addRECORD
                        rsCHECK.enableRECORD
                        
                Case "&Save Record "

                        bsaveRECORD = False
                        rsCHECK.validateRECORD
                        
                        If bsaveRECORD = True Then
                                rsCHECK.UpdateALLRECORDS
                                If bsaveRECORD = False Then
                                          .Toolbar1.Buttons(2).Caption = "New &Record "
                                          .Toolbar1.Buttons(3).Caption = "&NEXT ITEM "
                                          .Toolbar1.Buttons(4).Caption = "FINISH"
                                          disableALLRECORD
                                End If
                        End If
                
                Case Else
            End Select
        
        Case "S"
            Select Case Button.Caption
                Case "FINISH"
                    .Toolbar1.Buttons(2).Caption = "New &Record "
                    .Toolbar1.Buttons(2).Image = 2
                    .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                    .Toolbar1.Buttons(3).Image = 5
                    NewRecord = False: editRECORD = False: clearALLRECORD
                End Select
        Case "R"
            If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                .Toolbar1.Buttons(2).Caption = "New &Record "
                .Toolbar1.Buttons(2).Image = 2
                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                .Toolbar1.Buttons(3).Image = 5
                NewRecord = False: editRECORD = False: clearALLRECORD
        Case "P"
                
        Case "F"
     
            Me.HelpCommonDialog.DialogTitle = "Using the Main System"
            Me.HelpCommonDialog.HelpFile = App.HelpFile
            Me.HelpCommonDialog.HelpContext = 71
            Me.HelpCommonDialog.HelpCommand = cdlHelpContext
            Me.HelpCommonDialog.ShowHelp

        Case Else
            Exit Sub
        End Select
        
        
End With
Exit Sub
err:
    ErrorMessage

End Sub


Private Sub cmdAddNew_Click()
        bmakePAYMENT = True
        breversePAYMENT = False
        rsCHECK.addRECORD
      
End Sub


Private Sub cmdCancel_Click()
'        rsCHECK.Cancelrecord
        bmakePAYMENT = False
        breversePAYMENT = False

End Sub





Private Sub Command1_Click()
        rsCHECK.contructDATA

End Sub

Private Sub DTPickerChequeDate_Change()
        rsCHECK.ChangeDATE
End Sub

Private Sub Form_Activate()
        
                disableALLRECORD

        If bApproveCheque = True Or bAuthorizeCheque = True Then
                'rsCHECK.loadRECORD
        Else
                bmakePAYMENT = True
                breversePAYMENT = False
                showCOYBankAccounts
                rsCHECK.loadVOUCHER
                rsCHECK.calculateTOTALPAID
                GetVouchers
                showALLCHKENTRIES
                loadCostCenter
        End If
        

End Sub

Private Sub Form_Load()
        OpenConnection
End Sub


Private Sub Form_Unload(cancel As Integer)
        bmakePAYMENT = False
        breversePAYMENT = False
End Sub


Private Sub txtAmountPaid_LostFocus()
        rsCHECK.checkSTATUS
End Sub

Private Sub txtChequeNo_LostFocus()
        If NewRecord = True Then
                rsCHECK.countENTRIES
        End If
End Sub
