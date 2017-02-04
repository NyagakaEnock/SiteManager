VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmODASMCheckIssuance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheque Issuance"
   ClientHeight    =   7650
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "frmODASMCheckIssuance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6855
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   11655
      Begin VB.Frame Frame5 
         Caption         =   "Related Cheques"
         Height          =   1575
         Left            =   120
         TabIndex        =   34
         Top             =   2160
         Width           =   6615
         Begin MSComctlLib.ListView ListView2 
            Height          =   1215
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   2143
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
         Caption         =   "Cheques Issued This Period"
         Height          =   3015
         Left            =   0
         TabIndex        =   32
         Top             =   3720
         Width           =   11535
         Begin MSComctlLib.ListView ListView1 
            Height          =   2655
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   4683
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
         Height          =   3615
         Left            =   6840
         TabIndex        =   15
         Top             =   120
         Width           =   4695
         Begin VB.TextBox txtVoucherNo 
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
            Left            =   1800
            TabIndex        =   37
            Top             =   651
            Width           =   2415
         End
         Begin VB.TextBox txtIssuedBy 
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
            Left            =   1800
            TabIndex        =   29
            Top             =   3120
            Width           =   2415
         End
         Begin VB.TextBox txtIssuanceNo 
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
            Left            =   1800
            TabIndex        =   27
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox txtComment 
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
            Left            =   1800
            TabIndex        =   25
            Top             =   2706
            Width           =   2415
         End
         Begin VB.TextBox txtDateCollected 
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
            Left            =   1800
            TabIndex        =   21
            Top             =   1062
            Width           =   2175
         End
         Begin VB.TextBox txtIdentityNo 
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
            Left            =   1800
            TabIndex        =   20
            Top             =   2295
            Width           =   2415
         End
         Begin VB.TextBox txtCollectedBy 
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
            Left            =   1800
            TabIndex        =   17
            Top             =   1473
            Width           =   2415
         End
         Begin VB.ComboBox cboIdType 
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
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   16
            Top             =   1884
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker DTPickerDateCollected 
            Height          =   315
            Left            =   3960
            TabIndex        =   22
            Top             =   1062
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55508993
            CurrentDate     =   37945
         End
         Begin VB.Label Label11 
            Caption         =   "Voucher No"
            Height          =   255
            Left            =   600
            TabIndex        =   38
            Top             =   681
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Issued By"
            Height          =   255
            Left            =   600
            TabIndex        =   30
            Top             =   3150
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Issuance No"
            Height          =   255
            Left            =   600
            TabIndex        =   28
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Remark"
            Height          =   255
            Left            =   600
            TabIndex        =   26
            Top             =   2736
            Width           =   1095
         End
         Begin VB.Label Label21 
            Caption         =   "Date Collected"
            Height          =   255
            Left            =   600
            TabIndex        =   24
            Top             =   1092
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Identity #"
            Height          =   255
            Left            =   600
            TabIndex        =   23
            Top             =   2325
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Identity Type"
            Height          =   255
            Left            =   600
            TabIndex        =   19
            Top             =   1914
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Collected by"
            Height          =   255
            Left            =   600
            TabIndex        =   18
            Top             =   1503
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2055
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   6615
         Begin VB.TextBox txtDocumentNo 
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
            Left            =   5640
            TabIndex        =   41
            Top             =   720
            Width           =   855
         End
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
            Left            =   5640
            TabIndex        =   39
            Top             =   240
            Width           =   855
         End
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
            Left            =   1080
            TabIndex        =   36
            Top             =   1200
            Width           =   855
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
            TabIndex        =   31
            Top             =   687
            Width           =   4575
         End
         Begin VB.TextBox txtChequeAmount 
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
            Left            =   4320
            TabIndex        =   8
            Top             =   1635
            Width           =   2175
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
            Left            =   1920
            TabIndex        =   7
            Top             =   1200
            Width           =   4575
         End
         Begin VB.TextBox txtChequeNo 
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
            Left            =   1080
            TabIndex        =   6
            Top             =   225
            Width           =   1815
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
            Left            =   1080
            TabIndex        =   5
            Top             =   1635
            Width           =   1815
         End
         Begin VB.TextBox txtCurrentPeriod 
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
            Left            =   3600
            TabIndex        =   4
            Top             =   225
            Width           =   975
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
            Left            =   1080
            TabIndex        =   3
            Top             =   687
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Cost Center"
            Height          =   255
            Left            =   4680
            TabIndex        =   40
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Cheque Amount"
            Height          =   255
            Left            =   3000
            TabIndex        =   14
            Top             =   1665
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Payee "
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1185
            Width           =   975
         End
         Begin VB.Label Label16 
            Caption         =   "Bank No"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label15 
            Caption         =   "Cheque No"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   278
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Status"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1665
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Period"
            Height          =   255
            Left            =   3000
            TabIndex        =   9
            Top             =   255
            Width           =   1095
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
            Picture         =   "frmODASMCheckIssuance.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMCheckIssuance.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMCheckIssuance.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMCheckIssuance.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMCheckIssuance.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMCheckIssuance.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMCheckIssuance.frx":236E
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
      Begin VB.TextBox txtInstallmentTotal 
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
         Left            =   10800
         TabIndex        =   42
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSComDlg.CommonDialog HelpCommonDialog 
         Left            =   9360
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         HelpCommand     =   11
         HelpContext     =   72
         HelpFile        =   "REGHELP.HLP"
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
Attribute VB_Name = "frmODASMCheckIssuance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCOLLECTED As clsALISCheque

Private Sub Form_Initialize()
        Set rsCOLLECTED = New clsALISCheque
End Sub

Private Sub Form_Terminate()
        Set rsCOLLECTED = Nothing
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
                            rsCOLLECTED.clearCHEQUE
                            rsCOLLECTED.enableISSUANCE
                    Case "&Save Record "
                    
                            bsaveRECORD = False
                            rsCOLLECTED.UpdateChecksIssued
                            If bsaveRECORD = True Then
                                bsaveRECORD = False
                                Button.Caption = "New &Record ": Button.Image = 2
                                .Toolbar1.Buttons(4).Caption = "FINISH": .Toolbar1.Buttons(3).Caption = "&NEXT ITEM ": .Toolbar1.Buttons(3).Image = 2
                                  disableALLRECORD
                            End If
                    
                    Exit Sub
                    End Select
    
        Case "E"
            Select Case Button.Caption
                Case "&Edit/Change "
                
                Case "&Save Record "

                        bsaveRECORD = False
                        rsCOLLECTED.validateRECORD
                        
                        If bsaveRECORD = True Then
                                rsCOLLECTED.UpdateALLRECORDS
                                If bsaveRECORD = False Then
                                          .Toolbar1.Buttons(2).Caption = "New &Record "
                                          .Toolbar1.Buttons(3).Caption = "&NEXT ITEM "
                                          .Toolbar1.Buttons(4).Caption = "FINISH"
                                          disableALLRECORD
                                End If
                        End If
                
                Case "&NEXT ITEM "
                            .Toolbar1.Buttons(2).Caption = "&Save Record ": .Toolbar1.Buttons(2).Image = 4
                            rsCOLLECTED.clearCHEQUE
                            rsCOLLECTED.enableISSUANCE
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
                bmakePAYMENT = False
                breversePAYMENT = False
                bissueCHECKS = False


        
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

Private Sub cboIdType_GotFocus()
    SelectIDTypeGotFocus
End Sub

Private Sub cboIdType_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboIdType_LostFocus()
    SelectIDTypeLostFocus
End Sub

Private Sub DTPickerChequeDate_Change()
        rsCOLLECTED.ChangeDATE
End Sub

Private Sub Form_Activate()
        bmakePAYMENT = False
        breversePAYMENT = False
        bissueCHECKS = True
        
        If bissueCHECKS = True Then
                rsCOLLECTED.loadAPPROVEDCHECKS
        End If
        
        Me.txtCurrentPeriod = CurrentPeriod
        GetChequesISSUEDTHISPERIOD
        GetChequesRELATED
        disableALLRECORD
End Sub

Private Sub Form_Unload(cancel As Integer)
        bmakePAYMENT = False
        breversePAYMENT = False
End Sub

Private Sub txtAmountPaid_LostFocus()
        rsCOLLECTED.checkSTATUS
End Sub

Private Sub saveINSTALLMENT()
On Error GoTo err

    With frmODASMVoucher
    
            Set rsSAVE = New Recordset
            strSQL = "SELECT * from ODASMInstallment where InvoiceNo = '" & .cboDocumentNo.Text & " ' and ContractNo = '" & .txtLPONo.Text & "'; "
            rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        
            If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
            
            rsSAVE!Status = "REQ-PREPARED"
            rsSAVE!StatusDate = Date
            
            If CDbl(.txtInvoiceBalance.Text) = 0 Then
                rsSAVE!PaymentFlag = "Y"
                rsSAVE!Requisitioned = "Y"
                rsSAVE!DateRequisitioned = Date
            Else:
                rsSAVE!PaymentFlag = "P"
                rsSAVE!Requisitioned = "P"
                rsSAVE!DateRequisitioned = Date
            End If
            
            rsSAVE!PaymentDue = CDbl(.txtInvoiceBalance)
            rsSAVE!Balance = CDbl(.txtInvoiceBalance)
            rsSAVE!vOUCHERnO = .txtVoucherNo
            rsSAVE!VoucherDate = CDate(.txtRequisitionDate.Text)
        
            rsSAVE.Update
            rsSAVE.Requery
    End With
Exit Sub
err:
    UpdateErrorMessage
End Sub
