VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmODASMLeaseAgreement 
   Caption         =   "Terminate Contract "
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   10755
   Icon            =   "frmODASMLeaseAgreement.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmODASMLeaseAgreement.frx":0442
   ScaleHeight     =   6360
   ScaleWidth      =   10755
   Begin VB.Frame Frame12 
      Height          =   5655
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   10575
      Begin VB.TextBox txtLeaseDuration 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   5760
         TabIndex        =   38
         Top             =   240
         Width           =   255
      End
      Begin MSComCtl2.DTPicker DTPickerTerminationDate 
         Height          =   315
         Left            =   9840
         TabIndex        =   37
         Top             =   1680
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   22609921
         CurrentDate     =   38365
      End
      Begin MSComCtl2.DTPicker DTPickerNoticeDate 
         Height          =   315
         Left            =   5760
         TabIndex        =   36
         Top             =   1680
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   22609921
         CurrentDate     =   38365
      End
      Begin VB.TextBox txtCommencementDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   4440
         TabIndex        =   33
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtExpiryDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   7560
         TabIndex        =   32
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtPhysicalLocation 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   6000
         TabIndex        =   31
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox txtRentRecovered 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   8040
         TabIndex        =   30
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox txtRecoveryRatio 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   4560
         TabIndex        =   28
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox txtRentPaid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   1680
         TabIndex        =   25
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtTerminationDate 
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
         Height          =   315
         Left            =   7560
         TabIndex        =   23
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtLandLord 
         BackColor       =   &H00FFFFC0&
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
         Left            =   7560
         TabIndex        =   21
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtNoticeDate 
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
         Height          =   315
         Left            =   4560
         TabIndex        =   20
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtPlotNo 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1680
         TabIndex        =   17
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtTerminatedBy 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1680
         TabIndex        =   16
         Top             =   2160
         Width           =   1815
      End
      Begin VB.ComboBox cboTerminationCode 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   4560
         TabIndex        =   12
         Top             =   2160
         Width           =   5535
      End
      Begin VB.TextBox txtAgreementDate 
         BackColor       =   &H00FFFFC0&
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
         Left            =   4560
         TabIndex        =   1
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtSignedBy 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   1680
         TabIndex        =   0
         Top             =   1200
         Width           =   1815
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
         Left            =   1680
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   3120
         Width           =   8415
      End
      Begin VB.TextBox txtContractNo 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtPlotName 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1680
         TabIndex        =   5
         Top             =   720
         Width           =   4335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Terminated Contracts"
         Height          =   1695
         Left            =   240
         TabIndex        =   6
         Top             =   3840
         Width           =   9615
         Begin MSComctlLib.ListView ListView1 
            Height          =   1335
            Left            =   120
            TabIndex        =   7
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
      Begin VB.Label Label15 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   3600
         TabIndex        =   35
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Expiry Date"
         Height          =   255
         Left            =   6600
         TabIndex        =   34
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Recovery"
         Height          =   255
         Left            =   7200
         TabIndex        =   29
         Top             =   2670
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Recovery Ratio"
         Height          =   375
         Left            =   3600
         TabIndex        =   27
         Top             =   2610
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Rent Paid"
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   2670
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Termination Date"
         Height          =   375
         Left            =   6600
         TabIndex        =   24
         Top             =   1650
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Land Lord"
         Height          =   255
         Left            =   6600
         TabIndex        =   22
         Top             =   1230
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Notice Date"
         Height          =   255
         Left            =   3600
         TabIndex        =   19
         Top             =   1710
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Plot No"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   1710
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Notes"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Plot Name"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Terminated By"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   2190
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Date Signed"
         Height          =   255
         Left            =   3600
         TabIndex        =   11
         Top             =   1230
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Signed By"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   1230
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   " Reason"
         Height          =   255
         Left            =   3600
         TabIndex        =   9
         Top             =   2190
         Width           =   1455
      End
      Begin VB.Label lblRelationshipCode 
         Caption         =   "Contract No"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   270
         Width           =   1335
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   39
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
            Key             =   "H"
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
               Picture         =   "frmODASMLeaseAgreement.frx":0784
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMLeaseAgreement.frx":0DFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMLeaseAgreement.frx":1340
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMLeaseAgreement.frx":1792
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMLeaseAgreement.frx":1AAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMLeaseAgreement.frx":2126
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMLeaseAgreement.frx":27A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMLeaseAgreement.frx":2BF2
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
            Picture         =   "frmODASMLeaseAgreement.frx":326C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMLeaseAgreement.frx":38E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMLeaseAgreement.frx":3D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMLeaseAgreement.frx":4052
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMLeaseAgreement.frx":46CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMLeaseAgreement.frx":4D46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMLeaseAgreement.frx":5198
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmODASMLeaseAgreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTERMINATION As clsODASTerminationLandlord, MyCommonData As clsCommonData

Private Sub cboTerminationCode_Click()
    Me.txtRecoveryRatio.SetFocus
End Sub

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
        rsTERMINATION.loadRECORD
        rsTERMINATION.loadDEFAULTS
End Sub
Private Sub Form_Terminate()
        Set rsTERMINATION = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo err

        If NewRecord Then
            Cancel = True
            MsgBox "System cannot close this form when the generation of lease Agreement is in Process...", vbCritical + vbOKOnly, "Interruption of Job Brief creation process"
        Else
            Cancel = False
            Set rsTERMINATION = Nothing
        End If
Exit Sub
err:
ErrorMessage
        
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
    With Me
        Select Case Button.Key
            Case "N"
                Select Case Button.Caption
                    Case "New &Record "
                            rsTERMINATION.enableRECORD
                            NewRecord = True
                            Button.Caption = "&Save Record": Button.Image = 5: .Toolbar1.Buttons(4).Caption = "Cancel": .Toolbar1.Buttons(4).Image = 2
                    Case "&Save Record"
                            rsTERMINATION.ValidateRECORD
                            If bSaveRECORD = True Then
                                    rsTERMINATION.saveRecord
                                    rsTERMINATION.UpdateLease
                                    getCONTRACTSTerminated
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
                                If .txtContractNo.Text = Empty Then
                                MsgBox "There is NO Current Record to Edit. Please Search and Display a Record First...!", vbCritical + vbOKOnly, ""
                               .txtContractNo.SetFocus
                                Else
                                rsTERMINATION.enableRECORD
'                                disableButtons
                               .txtContractNo.Locked = True
                                Button.Caption = "Save &Changes ": Button.Image = 5
                                EditRecord = True
                            End If
                    Case "Save &Change "
                        rsTERMINATION.saveRecord
                        EditRecord = False: Button.Caption = "Edit &Change ": Button.Image = 6
                End Select
            Case "S"
                Select Case Button.Caption
                    Case "&Search/Find "
                    Case "Cancel"
                            clearALLRECORD
                            disableALLRECORD
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
                .HelpCommonDialog.HelpContext = 27
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
        End Select
    End With
Exit Sub
err:
ErrorMessage
End Sub
