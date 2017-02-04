VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmODASMPrepareNotice 
   Caption         =   "Prepare Notice"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   10785
   Icon            =   "frmODASMPrepareNotice.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmODASMPrepareNotice.frx":0442
   ScaleHeight     =   6465
   ScaleWidth      =   10785
   Begin VB.Frame Frame12 
      Height          =   5655
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   10455
      Begin VB.Frame Frame4 
         Caption         =   "Renewal Period"
         Height          =   1695
         Left            =   7200
         TabIndex        =   35
         Top             =   2040
         Width           =   3135
         Begin MSComCtl2.UpDown UpDownRenewalDuration 
            Height          =   255
            Left            =   600
            TabIndex        =   43
            Top             =   1080
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtStartDate 
            BackColor       =   &H00FFC0C0&
            Height          =   285
            Left            =   1200
            TabIndex        =   41
            Text            =   " "
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtYear 
            BackColor       =   &H00FFC0C0&
            Height          =   285
            Left            =   840
            TabIndex        =   39
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtRenewalEndDate 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            TabIndex        =   37
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txtMonth 
            BackColor       =   &H00FFC0C0&
            Height          =   285
            Left            =   120
            TabIndex        =   36
            Top             =   1080
            Width           =   495
         End
         Begin MSComCtl2.DTPicker DTPickerStartDate 
            Height          =   315
            Left            =   2760
            TabIndex        =   42
            Top             =   360
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            Format          =   76349441
            CurrentDate     =   38365
         End
         Begin VB.Label Label6 
            Caption         =   "Starting Date"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Months     Years              End Date"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   720
            Width           =   2895
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Renewable BillBoards"
         Height          =   1695
         Left            =   240
         TabIndex        =   32
         Top             =   2040
         Width           =   4095
         Begin VB.CheckBox chkAll 
            Caption         =   "Renew All?"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   1815
         End
         Begin MSComctlLib.ListView ListView3 
            Height          =   1095
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   1931
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Reasons for Notice"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   4440
         TabIndex        =   27
         Top             =   2040
         Width           =   2655
         Begin VB.OptionButton optTermination 
            Caption         =   "Termination"
            Height          =   255
            Left            =   1320
            TabIndex        =   30
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optRenewal 
            Caption         =   "Renewal"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   1095
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
            Height          =   555
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Label1 
            Caption         =   "Notes"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   31
            Top             =   720
            Width           =   615
         End
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
         Left            =   6360
         TabIndex        =   25
         Top             =   240
         Width           =   255
      End
      Begin MSComCtl2.DTPicker DTPickerNoticeDate 
         Height          =   315
         Left            =   10080
         TabIndex        =   24
         Top             =   1680
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   76349441
         CurrentDate     =   38365
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
         Left            =   4200
         TabIndex        =   21
         Top             =   240
         Width           =   2175
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
         Left            =   8040
         TabIndex        =   20
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtPhysicalLocation 
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
         Left            =   6600
         TabIndex        =   19
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox txtLandLord 
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
         Left            =   8040
         TabIndex        =   17
         Top             =   1200
         Width           =   2295
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
         Left            =   8040
         TabIndex        =   16
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtPlotNo 
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
         Left            =   1440
         TabIndex        =   13
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtTerminatedBy 
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
         Left            =   4320
         TabIndex        =   12
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtAgreementDate 
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
         Left            =   4320
         TabIndex        =   1
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtSignedBy 
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
         Left            =   1440
         TabIndex        =   0
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtContractNo 
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
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtPlotName 
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
         Left            =   1440
         TabIndex        =   4
         Top             =   720
         Width           =   5175
      End
      Begin VB.Frame Frame1 
         Caption         =   "Contracts Notices Prepared For"
         Height          =   1695
         Left            =   120
         TabIndex        =   5
         Top             =   3840
         Width           =   10215
         Begin MSComctlLib.ListView ListView1 
            Height          =   1335
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   9975
            _ExtentX        =   17595
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
         Left            =   3360
         TabIndex        =   23
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Expiry Date"
         Height          =   255
         Left            =   6840
         TabIndex        =   22
         Top             =   270
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Land Lord"
         Height          =   255
         Left            =   6840
         TabIndex        =   18
         Top             =   1230
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Notice Date"
         Height          =   255
         Left            =   6840
         TabIndex        =   15
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Plot No"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1710
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Plot Name"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Prepared By"
         Height          =   255
         Left            =   3360
         TabIndex        =   10
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Date Signed"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   1230
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Signed By"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1230
         Width           =   1335
      End
      Begin VB.Label lblRelationshipCode 
         Caption         =   "Contract No"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   270
         Width           =   1335
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1230
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   2170
      ButtonWidth     =   3307
      ButtonHeight    =   1005
      TextAlignment   =   1
      ImageList       =   "ImageList1(1)"
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
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Help System  "
            Key             =   "H"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Index           =   1
         Left            =   6720
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
               Picture         =   "frmODASMPrepareNotice.frx":0784
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMPrepareNotice.frx":0DFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMPrepareNotice.frx":1340
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMPrepareNotice.frx":1792
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMPrepareNotice.frx":1AAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMPrepareNotice.frx":2126
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMPrepareNotice.frx":27A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMPrepareNotice.frx":2BF2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComDlg.CommonDialog HelpCommonDialog 
         Left            =   10200
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         HelpCommand     =   11
         HelpContext     =   72
         HelpFile        =   "REGHELP.HLP"
      End
   End
End
Attribute VB_Name = "frmODASMPrepareNotice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsNOTICE As clsODASTerminationLandlord, MyCommonData As clsCommonData

Private Sub cboTerminationCode_Gotfocus()
        SelectTerminationReasonGotFocus
End Sub

Private Sub cboTerminationCode_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub cboTerminationCode_LostFocus()
        selectTerminationReasonLostFocus
End Sub

Private Sub chkAll_Click()
On Error GoTo err
With Me
    j = .ListView3.ListItems.Count
    If j = 0 Or .ListView3.View <> lvwReport Then
        .chkAll.Value = 0: Exit Sub
    Else
        If .chkAll.Value = 1 Then
            For i = 1 To j
                .ListView3.ListItems(i).Checked = True
                
            Next i
            k = .ListView3.ListItems.Count
        ElseIf .chkAll.Value = 0 Then
            For i = 1 To j
                .ListView3.ListItems(i).Checked = False
            Next i
            k = 0
        End If
    End If
End With
Exit Sub
err:
    ErrorMessage

End Sub

Private Sub DTPickerStartDate_CloseUp()
On Error GoTo err
    With Me
            .txtStartDate.Text = .DTPickerStartDate.Value
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub Form_Activate()
        Set rsNOTICE = New clsODASTerminationLandlord
        Set MyCommonData = New clsCommonData
        disableALLRECORD
        LeasedMasts
        rsNOTICE.loadRECORD
        rsNOTICE.loadDEFAULTS
End Sub

Private Sub Form_Terminate()
        Set rsNOTICE = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo err

        If NewRecord Then
            Cancel = True
            MsgBox "System cannot close this form when preparation of Notice is in Process...", vbCritical + vbOKOnly, "Interruption of Job Brief creation process"
        Else
            Cancel = False
            Set rsNOTICE = Nothing
        End If
Exit Sub
err:
ErrorMessage
End Sub

Private Sub optRenewal_Click()
On Error GoTo err
    With Me
         .txtStartDate.Enabled = True
         .DTPickerStartDate.Enabled = True
        .txtNarration.Text = Empty
        If DateDiff("M", Date, Format(.txtExpiryDate.Text, "MMMM dd,yyyy")) > 3 Then
        MsgBox "This Contract is not due for renewal", vbCritical + vbOKOnly, "Notice Preparation"
        .optTermination.SetFocus
        Else
        .txtNarration.Text = "Renewal of Contract"
        MsgBox "Go Ahead And Renew The Site,Adjust The Start Date First!!", vbCritical + vbOKOnly, "Notice Preparation"
        .txtStartDate.Text = Date
        End If
    End With
Exit Sub
err:
ErrorMessage
End Sub

Private Sub optTermination_Click()
With Me
    .txtStartDate.Enabled = False
    .DTPickerStartDate.Enabled = False
    .txtNarration.Text = Empty
    .txtNarration.Text = "Termination of Contract"
End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
    With Screen.ActiveForm
        Select Case Button.Key
            Case "N"
                Select Case Button.Caption
                    Case "New &Record "
                            rsNOTICE.enableRECORD
                            NewRecord = True
                            Button.Caption = "&Save Record": Button.Image = 5: .Toolbar1.Buttons(4).Caption = "Cancel": .Toolbar1.Buttons(4).Image = 2
                    Case "&Save Record"
                            rsNOTICE.ValidateRECORD
                            If bSaveRECORD = True Then
                                       k = 0: j = .ListView3.ListItems.Count
                                        For i = 1 To j
                                            If .ListView3.ListItems(i).Checked = True Then
                                            k = k + 1
                                            End If
                                        Next i
                                        
                                        If k = 0 Then
                                                MsgBox ("Please select one or more BillBoards to Lease!"), vbCritical + vbOKOnly, "Lease Preparation"
                                        
                                        Else

                                            rsNOTICE.SaveRecord
                                            updateLeasedPlotMasts
                                            getCONTRACTSNotices
                                            If bSaveRECORD = False Then
                                                disableALLRECORD
                                            End If
                            
                         NewRecord = False:  Button.Caption = "New &Record ": Button.Image = 3: .Toolbar1.Buttons(4).Caption = "&Search/Find ": .Toolbar1.Buttons(4).Image = 7
                    End If
                    End If
               End Select
            Case "E"
                Select Case Button.Caption
                    Case "&Edit/Change "
                         If NewRecord Then Exit Sub
                                If .txtContractNo.Text = Empty Then
                                MsgBox "There is NO Current Record to Edit. Please Search and Display a Record First...!", vbCritical + vbOKOnly, ""
                               .txtContractNo.SetFocus
                                Else
                                rsNOTICE.enableRECORD
                               .txtContractNo.Locked = True
                                Button.Caption = "Save &Changes ": Button.Image = 5
                                editRECORD = True
                            End If
                    Case "Save &Changes "
                        rsNOTICE.SaveRecord
                       editRECORD = False: Button.Caption = "Edit &Change ": Button.Image = 6
                End Select
            Case "S"
                Select Case Button.Caption
                Case "&Search/Find "
                    Set rsFindRecord = New ADODB.Recordset
                    INQUIRY = InputBox("Enter  the contract number to search and display...", "Search Values")
                    rsFindRecord.Open "SELECT * FROM ODASMLeaseAgreement WHERE ContractNo = '" & INQUIRY & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                    If rsFindRecord.EOF And rsFindRecord.BOF Then
                        MsgBox "System could not match the requested record. Either it is deleted or currently missing", vbInformation + vbOKOnly + vbDefaultButton1, "Missing Records"
                    Else
                        .txtNarration.Text = rsFindRecord!Narration & ""
                        .txtNoticeDate.Text = rsFindRecord!NoticeDate & ""
                        If SchedulingMain.txtTask.Text = "N6" Then
                            .txtRecoveryRatio.Text = rsFindRecord!RecoveryRatio & ""
                            .txtRentRecovered.Text = rsFindRecord!RentRecovered & ""
                            .txtTerminationDate.Text = rsFindRecord!TerminationDate & ""
                            .cboTerminationCode.Text = rsFindRecord!TerminationCode & ""
                            .txtSignedBy.Text = rsFindRecord!SignedBy
                        End If
                         bSaveRECORD = False
                    End If
                Case "Cancel"
                        clearALLRECORD
'                        enableButtons
                        disableALLRECORD
                End Select
            Case "R"
                    If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                    .Toolbar1.Buttons(2).Caption = "New &Record "
                    .Toolbar1.Buttons(2).Image = 2
                    .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                    .Toolbar1.Buttons(3).Image = 5
                    NewRecord = False: editRECORD = False: MyCommonData.ClearTheScreen

            Case "H"
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 23
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
        End Select
    End With
Exit Sub
err:
ErrorMessage
End Sub
Private Sub updateLeasedPlotMasts()

On Error GoTo err
    With frmODASMPrepareNotice
    Dim search As String
    
        j = .ListView3.ListItems.Count
        For i = 1 To j
            If .ListView3.ListItems(i).Checked = True Then
            search = .ListView3.ListItems(i).Text
            
            Set rsFindRecord = New ADODB.Recordset
            rsFindRecord.Open "SELECT * FROM ODASPPlotMast WHERE MastNo = '" & search & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
            If rsFindRecord.RecordCount = 0 Then Exit Sub
                rsFindRecord!Renewal = "Y"
           
            rsFindRecord.Update
            End If
        Next i
    End With
Exit Sub
err:
ErrorMessage
End Sub

Private Sub txtMonth_LostFocus()
On Error GoTo err
    With Me
        If .txtMonth.Text = Empty Then
            .txtMonth.SetFocus
            
            ElseIf .txtRenewalEndDate.Text <> "" Then
                .txtRenewalEndDate.Text = DateAdd("M", .txtMonth.Text, Format(.txtRenewalEndDate.Text, "MMMM dd,yyyy"))
            ElseIf .txtRenewalEndDate.Text = "" Then
            .txtRenewalEndDate.Text = DateAdd("M", .txtMonth.Text, Format(.txtExpiryDate.Text, "MMMM dd,yyyy"))
            .txtRenewalEndDate.Text = DateAdd("d", -1, Format(.txtRenewalEndDate.Text, "MMMM dd,yyyy"))
        End If
    End With
Exit Sub
err:
ErrorMessage

End Sub

Private Sub txtYear_LostFocus()
On Error GoTo err
    With Me
        If .txtYear.Text = "" Then Exit Sub
            If txtRenewalEndDate.Text = "" Then
                 .txtRenewalEndDate.Text = DateAdd("YYYY", .txtYear.Text, Format(.txtExpiryDate.Text, "MMMM dd,yyyy"))
                 .txtRenewalEndDate.Text = DateAdd("d", -1, Format(.txtRenewalEndDate.Text, "MMMM dd,yyyy"))
            Else
                 .txtRenewalEndDate.Text = DateAdd("YYYY", .txtYear.Text, Format(.txtRenewalEndDate.Text, "MMMM dd,yyyy"))
            End If
    End With
Exit Sub
err:
ErrorMessage

End Sub

Private Sub UpDown1_Change()

End Sub

Private Sub UpDownRenewalDuration_Change()
On Error GoTo err
        With Me
               If .optRenewal.Value = True Then
                .txtYear.Text = .UpDownRenewalDuration.Value
                .txtRenewalEndDate.Text = DateAdd("YYYY", .UpDownRenewalDuration.Value, Format(.txtStartDate.Text, "MMMM dd,yyyy"))
                Else:
                .txtYear.Text = .UpDownRenewalDuration.Value
                End If
        End With

Exit Sub

err:
    ErrorMessage
End Sub
