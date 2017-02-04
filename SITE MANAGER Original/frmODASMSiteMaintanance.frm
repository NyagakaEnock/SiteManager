VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmODASMSiteMaintanance 
   Caption         =   "SITE MAINTANANCE"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   10755
   Icon            =   "frmODASMSiteMaintanance.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmODASMSiteMaintanance.frx":0442
   ScaleHeight     =   5040
   ScaleWidth      =   10755
   Begin VB.Frame Frame12 
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   10215
      Begin VB.Frame Frame2 
         Caption         =   "Maintanance DONE"
         Height          =   855
         Left            =   480
         TabIndex        =   20
         Top             =   1560
         Width           =   3495
         Begin VB.OptionButton optNo 
            Caption         =   "NO"
            Height          =   255
            Left            =   2520
            TabIndex        =   22
            Top             =   360
            Width           =   615
         End
         Begin VB.OptionButton optYES 
            Caption         =   "YES"
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.TextBox txtMaintananceNo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2160
         TabIndex        =   18
         Top             =   1200
         Width           =   1815
      End
      Begin VB.ComboBox cboStaff 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   5760
         TabIndex        =   17
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtLeaseDuration 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6480
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtCommencementDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtExpiryDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8280
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtSiteDetails 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   10
         Top             =   720
         Width           =   7935
      End
      Begin VB.TextBox txtSiteNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtDoneBy 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7200
         TabIndex        =   7
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtNarration 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   5760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   1680
         Width           =   4335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Maintanances"
         Height          =   1695
         Left            =   480
         TabIndex        =   2
         Top             =   2400
         Width           =   9615
         Begin MSComctlLib.ListView ListView1 
            Height          =   1335
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   2355
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
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
      Begin VB.Label Label1 
         Caption         =   "Maintanance No."
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   4080
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Expiry Date"
         Height          =   255
         Left            =   7320
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Site No"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Notes"
         Height          =   255
         Left            =   4080
         TabIndex        =   6
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Siet Details"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Maintanance Done By"
         Height          =   255
         Left            =   4080
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   1164
      ButtonWidth     =   3069
      ButtonHeight    =   1005
      TextAlignment   =   1
      ImageList       =   "ImageList1(1)"
      DisabledImageList=   "ImageList1(1)"
      HotImageList    =   "ImageList1(1)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New &Record "
            Key             =   "N"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Edit/Change "
            Key             =   "E"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Search/Find "
            Key             =   "S"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Refresh "
            Key             =   "R"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print &Preview "
            Key             =   "P"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Help"
            Key             =   "F"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Index           =   1
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMSiteMaintanance.frx":0784
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMSiteMaintanance.frx":0DFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMSiteMaintanance.frx":1340
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMSiteMaintanance.frx":1792
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMSiteMaintanance.frx":1AAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMSiteMaintanance.frx":2126
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMSiteMaintanance.frx":27A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMSiteMaintanance.frx":2BF2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComDlg.CommonDialog HelpCommonDialog 
         Left            =   10560
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         HelpCommand     =   11
         HelpContext     =   72
         HelpFile        =   "REGHELP.HLP"
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   0
      Left            =   0
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
            Picture         =   "frmODASMSiteMaintanance.frx":326C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMSiteMaintanance.frx":38E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMSiteMaintanance.frx":3D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMSiteMaintanance.frx":4052
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMSiteMaintanance.frx":46CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMSiteMaintanance.frx":4D46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMSiteMaintanance.frx":5198
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmODASMSiteMaintanance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTERMINATION As clsODASTerminationLandlord, MyCommonData As clsCommonData

Private Sub cboTerminationCode_Gotfocus()
        SelectTerminationReasonGotFocus
End Sub

Private Sub cboTerminationCode_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub cboTerminationCode_LostFocus()
        selectTerminationReasonLostFocus
End Sub

Private Sub Form_Activate()
        Set rsTERMINATION = New clsODASTerminationLandlord
        Set MyCommonData = New clsCommonData
        disableALLRECORD
        enableALLRECORD
        rsTERMINATION.LoadDefaultDetails
End Sub

Private Sub Form_Terminate()
        Set rsTERMINATION = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo err

        If NewRecord Then
            Cancel = True
            MsgBox "System cannot close this form when sending of Notice is in Process...", vbCritical + vbOKOnly, "Interruption of Job Brief creation process"
        Else
            Cancel = False
            Set rsTERMINATION = Nothing
        End If
Exit Sub
err:
ErrorMessage
End Sub

Private Sub optNo_Click()
With Me
    .txtNarration.Text = ""
    .cboStaff.Text = ""
    .txtDoneBy.Text = ""
    .Label6.Caption = "Reasons"
End With
End Sub

Private Sub optYES_Click()
With Me
    .txtNarration = ""
    .txtDoneBy.Text = ""
    .cboStaff.Text = ""
    .Label6.Caption = "Comments"
End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
    With Me
        Select Case Button.Key
            Case "N"
                Select Case Button.Caption
                    Case "New &Record "
                            .Frame2.Enabled = True
                            .txtNarration.Locked = False: .txtNarration.Enabled = True: .optNO.Enabled = True: .optYES.Enabled = True
                            .txtDoneBy.SetFocus: NewRecord = True
                            Button.Caption = "&Save Record": Button.Image = 5: .Toolbar1.Buttons(4).Caption = "Cancel": .Toolbar1.Buttons(4).Image = 2
                    Case "&Save Record"
                            rsTERMINATION.GoodData
                            If bSaveRECORD = True Then
                                    rsTERMINATION.saveMaintanance
                                    If bSaveRECORD = False Then
                                        disableALLRECORD
                                    End If
                            End If
                           NewRecord = False: Button.Caption = "New &Record ": Button.Image = 3: .Toolbar1.Buttons(4).Caption = "&Search/Find ": .Toolbar1.Buttons(4).Image = 7
                End Select
            Case "E"
                 Select Case Button.Caption
                    Case "Edit &Change "
                         If NewRecord Then Exit Sub
                    Case "Save &Change "
                End Select
            Case "S"
                Select Case Button.Caption
                    Case "&Search/Find "
                    Case "Cancel"
                End Select
            Case "R"
                    If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                    .Toolbar1.Buttons(2).Caption = "New &Record "
                    .Toolbar1.Buttons(2).Image = 2
                    .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                    .Toolbar1.Buttons(3).Image = 5
                    NewRecord = False: EditRecord = False: MyCommonData.ClearTheScreen

            Case "H"
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 32
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
        End Select
    End With
Exit Sub
err:
ErrorMessage
End Sub

