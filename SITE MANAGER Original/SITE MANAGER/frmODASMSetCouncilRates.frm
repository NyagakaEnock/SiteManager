VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmODASMCouncilRates 
   Caption         =   "COUNCIL RATES"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   1170
   ClientWidth     =   10665
   Icon            =   "frmODASMSetCouncilRates.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmODASMSetCouncilRates.frx":0442
   ScaleHeight     =   5985
   ScaleWidth      =   10665
   Begin VB.Frame Frame12 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   10455
      Begin VB.TextBox txtJobBriefNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   7920
         TabIndex        =   33
         Top             =   4800
         Width           =   2415
      End
      Begin VB.TextBox txtSiteNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3840
         TabIndex        =   32
         Top             =   4800
         Width           =   3015
      End
      Begin VB.TextBox txtPayMode 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1440
         TabIndex        =   30
         Top             =   4800
         Width           =   1335
      End
      Begin VB.TextBox txtMediaCode 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1440
         TabIndex        =   29
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox txtMediaSize 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   3840
         TabIndex        =   28
         Top             =   4320
         Width           =   3015
      End
      Begin VB.TextBox txtAmount 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   7560
         TabIndex        =   27
         Top             =   4320
         Width           =   2775
      End
      Begin VB.TextBox txtPeriod 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   6000
         TabIndex        =   19
         Top             =   3240
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   5640
         TabIndex        =   18
         Top             =   3240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Format          =   55640065
         CurrentDate     =   38400
      End
      Begin VB.TextBox txtTown 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Top             =   3240
         Width           =   2655
      End
      Begin VB.TextBox txtTownCode 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   3240
         Width           =   855
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
         Left            =   3840
         TabIndex        =   11
         Top             =   3240
         Width           =   1815
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
         TabIndex        =   10
         Top             =   3240
         Width           =   2775
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
         Left            =   3840
         TabIndex        =   9
         Top             =   3840
         Width           =   3015
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
         Left            =   240
         TabIndex        =   7
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox txtDispatchedBy 
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
         Left            =   7560
         TabIndex        =   0
         Top             =   3840
         Width           =   2775
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
         Left            =   1440
         TabIndex        =   2
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Frame Frame1 
         Caption         =   "LAND RATES"
         Height          =   2655
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   10215
         Begin MSComctlLib.ListView ListView1 
            Height          =   2295
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   4048
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
      Begin VB.Label Label13 
         Caption         =   "JobBrief No"
         Height          =   255
         Left            =   6960
         TabIndex        =   34
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Site No"
         Height          =   255
         Left            =   2880
         TabIndex        =   31
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Amount"
         Height          =   255
         Left            =   6960
         TabIndex        =   26
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Media Size"
         Height          =   255
         Left            =   2880
         TabIndex        =   25
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Payment Mode"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Media Code"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Duration in Months"
         Height          =   255
         Left            =   6000
         TabIndex        =   20
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Town"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Expiry Date"
         Height          =   255
         Left            =   7560
         TabIndex        =   12
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Plot No"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Plot Name"
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Set By"
         Height          =   255
         Left            =   6960
         TabIndex        =   5
         Top             =   3840
         Width           =   615
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   10665
      _ExtentX        =   18812
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
            Caption         =   "&Help System  "
            Key             =   "H"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   10080
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
               Picture         =   "frmODASMSetCouncilRates.frx":0784
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMSetCouncilRates.frx":0DFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMSetCouncilRates.frx":1340
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMSetCouncilRates.frx":1792
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMSetCouncilRates.frx":1AAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMSetCouncilRates.frx":2126
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMSetCouncilRates.frx":27A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmODASMSetCouncilRates.frx":2BF2
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
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   495
      Left            =   4680
      TabIndex        =   23
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   495
      Left            =   4680
      TabIndex        =   22
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "frmODASMCouncilRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSEND As clsODASCouncilRates, MyCommonData As clsCommonData

Private Sub cboPaymentMode_GotFocus()
    SelectModeGotFocus
End Sub

Private Sub cboPaymentMode_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboPaymentMode_LostFocus()
    selectModeLostFocus
End Sub

Private Sub DTPickerRateStartDate_Change()
    Set rsSEND = New clsODASCouncilRates
    rsSEND.calcDUEDATE
    Set rsSEND = Nothing
End Sub

Private Sub Form_Activate()
        Set MyCommonData = New clsCommonData
        Set rsSEND = New clsODASCouncilRates
        disableALLRECORD
        ShowSITESWITHRATES
        ListALLSITES
        ListALLCOUNCILACCOUNTS
        listALLJOBBRIEFITEMS
        frmODASMCouncilRates.Frame2.Enabled = False
End Sub


Private Sub Form_Load()
        OpenODBCConnection
End Sub

Private Sub Form_Terminate()
        Set rsSEND = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Set rsSEND = Nothing
End Sub



Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'''On Error GoTo err
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'''On Error GoTo err
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
            
            frmODASMCouncilRates.txtSiteNo.Text = Item.Text
            frmODASMCouncilRates.txtPlotName.Text = Item.SubItems(1)
            Set rsSEND = New clsODASCouncilRates
            rsSEND.loadRECORD
            listALLJOBBRIEFITEMS
            Set rsSEND = Nothing
            
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ListView4_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
''On Error GoTo err
    ListView4.SortKey = ColumnHeader.Index - 1
    ListView4.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView4_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'''On Error GoTo err
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
            
            frmODASMCouncilRates.txtJobBriefItemNo.Text = Item.Text
            Set rsSEND = New clsODASCouncilRates
            rsSEND.loadJOBBRIEFITEMS
            rsSEND.loadSTARTDATE
            rsSEND.loadCOUNCILRATES
            rsSEND.calcDUEDATE

            Set rsSEND = Nothing

        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ListView5_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'''On Error GoTo err
    ListView5.SortKey = ColumnHeader.Index - 1
    ListView5.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView5_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'''On Error GoTo err
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
            
            frmODASMCouncilRates.txtAccountNo.Text = Item.Text
            
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub mnuViewCouncilRates_Click()
        Load frmODASVCouncilRates
        frmODASVCouncilRates.Show 1, Me
End Sub

Private Sub mnuViewSchedule_Click()
        Load frmODASMRateSchedule
        frmODASMRateSchedule.Show 1, Me
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'''On Error GoTo err
    With Me
        Set rsSEND = New clsODASCouncilRates

        Select Case Button.Key
            Case "N"
                enableALLRECORD
                NewRecord = True
                frmODASMCouncilRates.Frame2.Enabled = True

                Select Case Button.Caption
                    Case "New &Record "
                            enableALLRECORD
                            Button.Caption = "&Save Record": Button.Image = 5: .Toolbar1.Buttons(4).Caption = "Cancel": .Toolbar1.Buttons(4).Image = 2
                    Case "&Save Record"
                            rsSEND.ValidateRECORD
                            If bSaveRECORD = True Then
                                    rsSEND.SaveRECORD
                                    rsSEND.updateSITE
                                    
                                    ShowSITESWITHRATES
                                    ListALLSITES

                                    disableALLRECORD
                                    Button.Caption = "New &Record ": Button.Image = 3: .Toolbar1.Buttons(4).Caption = "&Search/Find ": .Toolbar1.Buttons(4).Image = 7
                            End If
                End Select
            Case "E"
                 Select Case Button.Caption
                    Case "Edit &Change "
                         If NewRecord Then Exit Sub
                              
                                Button.Caption = "Save &Changes ": Button.Image = 5
                                EditRecord = True
'                            End If
                    Case "Save &Change "
                            'rsSEND.ValidateRECORD
                            If bSaveRECORD = True Then
                                    'rsSEND.SaveCouncilRates
                                    'rsSEND.SaveRECORD
                            Button.Caption = "Edit &Change ": Button.Image = 6
                            End If
                End Select
            Case "S"
                Select Case Button.Caption
                    Case "&Search/Find "
                    Case "Cancel"
                            clearALLRECORD
'                            enableButtons
                            disableALLRECORD
                End Select
            Case "R"
                    If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                    .Toolbar1.Buttons(2).Caption = "New &Record "
                    .Toolbar1.Buttons(2).Image = 3
                    .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                    .Toolbar1.Buttons(3).Image = 6
                    .Toolbar1.Buttons(4).Caption = "&Search/Find "
                    .Toolbar1.Buttons(4).Image = 7
                    NewRecord = False: EditRecord = False: MyCommonData.ClearTheScreen

            Case "H"
        End Select
                    
        Set rsSEND = Nothing

    End With
Exit Sub
err:
ErrorMessage
End Sub


