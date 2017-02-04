VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmODASMNoticeAcknoledgement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JOB RENEWALS"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   615
   ClientWidth     =   11040
   Icon            =   "frmODASMPrepareAcknoledgement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   11040
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10680
      Top             =   240
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
            Picture         =   "frmODASMPrepareAcknoledgement.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPrepareAcknoledgement.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPrepareAcknoledgement.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPrepareAcknoledgement.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPrepareAcknoledgement.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPrepareAcknoledgement.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPrepareAcknoledgement.frx":236E
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
      Width           =   11040
      _ExtentX        =   19473
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
            Caption         =   "&Help System  "
            Key             =   "H"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComDlg.CommonDialog HelpCommonDialog 
         Left            =   11040
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         HelpCommand     =   11
         HelpContext     =   72
         HelpFile        =   "REGHELP.HLP"
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   10935
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   4560
         TabIndex        =   22
         Top             =   2040
         Width           =   6255
         Begin VB.TextBox txtAcknoledgedBy 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   3960
            TabIndex        =   38
            Top             =   1320
            Width           =   2175
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   255
            Left            =   1800
            TabIndex        =   36
            Top             =   1320
            Width           =   255
            _ExtentX        =   423
            _ExtentY        =   450
            _Version        =   393216
            Max             =   1000
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtMonths 
            BackColor       =   &H00FFC0C0&
            Height          =   285
            Left            =   1200
            TabIndex        =   35
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox txtJobBriefItemNo 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   1200
            TabIndex        =   28
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtSiteDetails 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   2880
            TabIndex        =   27
            Top             =   600
            Width           =   3255
         End
         Begin VB.TextBox txtJobBriefNo 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   1200
            TabIndex        =   26
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtStartDate 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   1200
            TabIndex        =   25
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtEndDate 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   3960
            TabIndex        =   24
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox txtproduct 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   3720
            TabIndex        =   23
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label9 
            Caption         =   "End Date"
            Height          =   255
            Left            =   3000
            TabIndex        =   39
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "Received By"
            Height          =   255
            Left            =   2880
            TabIndex        =   37
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "JobBrief No"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "JobBrief Item"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Start Date"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Months"
            Height          =   255
            Left            =   2160
            TabIndex        =   31
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label lblproduct 
            Caption         =   "Product"
            Height          =   255
            Left            =   3000
            TabIndex        =   30
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "Renew By"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   1320
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3615
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   4335
         Begin VB.TextBox txtAccountName 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1800
            TabIndex        =   18
            Top             =   480
            Width           =   2415
         End
         Begin VB.ComboBox cboAccountType 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   240
            TabIndex        =   17
            Top             =   480
            Width           =   1575
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   2535
            Left            =   120
            TabIndex        =   21
            Top             =   960
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   4471
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Account Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1800
            TabIndex        =   20
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label Label10 
            Caption         =   "Account Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1935
         Left            =   4560
         TabIndex        =   3
         Top             =   120
         Width           =   6255
         Begin VB.TextBox txtEmail 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4320
            TabIndex        =   9
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtAddress2 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1440
            TabIndex        =   8
            Top             =   615
            Width           =   4695
         End
         Begin VB.TextBox txtAddress1 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1440
            TabIndex        =   7
            Top             =   1020
            Width           =   4695
         End
         Begin VB.TextBox txtFax 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4200
            TabIndex        =   6
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox txtCompanyName 
            BackColor       =   &H00FFFFC0&
            Height          =   330
            Left            =   1440
            TabIndex        =   5
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox txtContactPerson 
            BackColor       =   &H00FFFFC0&
            Height          =   330
            Left            =   1440
            TabIndex        =   4
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label Label12 
            Caption         =   "E- Mail:"
            Height          =   195
            Left            =   3600
            TabIndex        =   15
            Top             =   240
            Width           =   705
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Physical Address"
            Height          =   315
            Left            =   120
            TabIndex        =   14
            Top             =   1020
            Width           =   1425
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Postal Address"
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   615
            Width           =   1305
         End
         Begin VB.Label Label16 
            Caption         =   "Fax No."
            Height          =   315
            Left            =   3600
            TabIndex        =   12
            Top             =   1440
            Width           =   585
         End
         Begin VB.Label Label11 
            Caption         =   "Company Name"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   "Contact Person"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1440
            Width           =   1335
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1575
         Left            =   120
         TabIndex        =   2
         Top             =   3840
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   2778
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
         BackColor       =   -2147483643
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
      Begin VB.Menu mnuCurrent 
         Caption         =   "Current Settings"
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
Attribute VB_Name = "frmODASMNoticeAcknoledgement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ListCurrentSettings()
On Error GoTo err
With Me

.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Client", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Site Details", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Start Date", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Expiry Date", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "Renewal Period", .ListView1.Width / 5

.ListView1.View = lvwReport: .ListView1.Visible = True
Dim MyList As ListItem

If Not IsEmpty(.txtCompanyName.Text) Then
    Set MyList = .ListView1.ListItems.Add(, , CStr(Trim(.txtCompanyName.Text)))
End If
If Not IsEmpty(.txtSiteDetails.Text) Then
    MyList.SubItems(1) = CStr(Trim(.txtSiteDetails.Text))
End If
If Not IsEmpty(.txtStartDate.Text) Then
    MyList.SubItems(2) = CStr(Trim(.txtStartDate.Text))
End If
If Not IsEmpty(.txtEndDate.Text) Then
    MyList.SubItems(3) = CStr(Trim(.txtEndDate.Text))
End If
If Not IsEmpty(.txtMonths.Text) Then
    MyList.SubItems(4) = CStr(Trim(.txtMonths.Text)) & " MONTHS"
End If

End With
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub GetCurrentSettings()
On Error GoTo err
With Me

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Exporter Code", .ListView1.Width / 8 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Exporter Name", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 3.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Postal Address", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Town/City", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "Country", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Phone Number", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Fax/Telex", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Email", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "PIN No", .ListView1.Width / 8 ', lvwColumnCenter

.ListView1.View = lvwReport: .ListView1.Visible = True
End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cboAccountType_Click()
    Me.txtAccountName.SetFocus
End Sub

Private Sub cboAccountType_GotFocus()
With Me
        If .cboAccountType.ListCount <> 0 Then Exit Sub
        .cboAccountType.Clear
        AttachSQL = "SELECT (ODASPAccountType.AccountTypeDescription)as selectfield,ODASPAccountType.* FROM ODASPAccountType ;"
        AttachDropDowns
End With
End Sub

Private Sub cboAccountType_LostFocus()
On Error GoTo err
    With Me
                .ListView2.ListItems.Clear
                .ListView2.ColumnHeaders.Clear
                
                .ListView2.ColumnHeaders.Add , , "JBI No", .ListView2.Width / 4
                .ListView2.ColumnHeaders.Add , , "Media", .ListView2.Width / 4 ', lvwColumnCenter
                .ListView2.ColumnHeaders.Add , , "Town", .ListView2.Width / 4
                .ListView2.ColumnHeaders.Add , , "Comm. Date", .ListView2.Width / 4 ', lvwColumnCenter
                .ListView2.ColumnHeaders.Add , , "Expiry Date", .ListView2.Width / 4

                .ListView2.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Set rsLIST = New ADODB.Recordset
                    rsLIST.Open "SELECT AT.*,A.*,JB.*,(JBI.CommencementDate)as SDate,(JBI.ExpiryDate)as EndDate,JBI.* FROM ODASPAccountType AT,ODASPAccount A,ODASMJobBrief JB, ODASMJobBriefItems JBI WHERE AT.AccountTypeDescription = '" & .cboAccountType.Text & "' and AT.AccountType = A.AccountType and A.AccountNo = JB.AccountNo and JB.JobBriefNo = JBI.JobBriefNo AND JB.Closed = 'Y' and (JBI.NoticeReceived = 'N' or JBI.NoticeReceived is null);", cnCOMMON, adOpenKeyset, adLockOptimistic
                    If rsLIST.EOF And rsLIST.BOF Then Exit Sub
                    
                    .txtAccountName.Text = rsLIST!AccountTypeDescription
                    .cboAccountType.Text = rsLIST!AccountType
                Dim MyList As ListItem
                          
                While Not rsLIST.EOF
                
                If DateDiff("M", Date, rsLIST!EndDate) > 3 Then GoTo Continue
                    Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!JobBriefItemNo))
                        If Not IsNull(rsLIST!MediaCode) Then
                            MyList.SubItems(1) = CStr(rsLIST!MediaCode)
                        End If
                        If Not IsNull(rsLIST!Town) Then
                            MyList.SubItems(2) = CStr(rsLIST!Town)
                        End If
                        If Not IsNull(rsLIST!SDate) Then
                            MyList.SubItems(3) = Format(rsLIST!SDate, "dd/mm/yyyy")
                        End If
                        If Not IsNull(rsLIST!EndDate) Then
                            MyList.SubItems(4) = Format(rsLIST!EndDate, "dd/mm/yyyy")
                        End If
Continue:
                     rsLIST.MoveNext
                Wend
                Set MyList = Nothing

End With
Exit Sub
err:
ErrorMessage
End Sub

Private Sub cboCountry_Click()
If Not NewRecord And Not EditRecord Then Exit Sub
With Me
End With
End Sub

Private Sub cboCountry_gotfocus()
If Not NewRecord And Not EditRecord Then Exit Sub
With Screen.ActiveForm
    AttachSQL = "SELECT CountryName AS SelectField FROM ParamCountries ORDER BY CountryName;"
    .cboCountry.Clear
    MyCommonData.AttachDropDown
End With

End Sub

Private Sub Form_Activate()
With Me
    .Frame4.Enabled = False
    .Frame2.Enabled = False
    If .txtJobBriefItemNo.Text <> Empty Then
    loadBriefItemDETAILS
    End If
End With
End Sub

Private Sub Form_Initialize()
Set MyCommonData = New clsCommonData

End Sub

Private Sub Form_Resize()
On Error GoTo err
With Me
    .Frame3.Height = .Height - (8505 - 7215)
    .ListView1.Height = .Height - (8505 - 4335)
    .Frame3.Width = .Width - (12000 - 11895)
    .ListView1.Width = .Width - (12000 - 11655)
End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo err

        If NewRecord Then
            Cancel = True
            MsgBox "System cannot close this form when acknowledgement of notice is in Process...", vbCritical + vbOKOnly, "Interruption of Job Brief creation process"
        Else
            Cancel = False
        End If
Exit Sub
err:
ErrorMessage

End Sub

Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
With Me
    .txtAcknoledgedBy.Text = CurrentUserName
    .txtJobBriefItemNo.Text = Item
    loadBriefItemDETAILS
    .txtStartDate.Text = Item.SubItems(3)
    .txtEndDate.Text = Item.SubItems(4)
End With
End Sub

Private Sub mnuClear_Click()
    MyCommonData.ClearTextFields
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Function ValidRecord() As Boolean
With Me
    If .txtMonths.Text = Empty Then
        strMessage = "Required Renewal Period...!"
        .txtMonths.SetFocus
    If .txtJobBriefItemNo.Text = Empty Then
        strMessage = "Required Job Brief Item's Details...!"
        .txtJobBriefItemNo.SetFocus
    End If
    Else
        ValidRecord = True
    End If
    If Not ValidRecord Then
        MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
    End If
End With
End Function
Private Sub loadBriefItemDETAILS()
On Error GoTo err
    With Me
        Set rsFindRecord = New ADODB.Recordset
        rsFindRecord.Open "SELECT JB.*,(JBI.CommencementDate)as SDate,(JBI.ExpiryDate)as EDate,JBI.*,A.* FROM ODASMJobBrief JB, ODASMJobBriefItems JBI,ODASPAccount A WHERE JB.JobBriefNo = JBI.JobBriefNo and JBI.JobBriefItemNo = '" & .txtJobBriefItemNo.Text & "' and JB.AccountNo = A.AccountNo", cnCOMMON, adOpenKeyset, adLockOptimistic
        .txtJobBriefNo.Text = rsFindRecord!JobBriefNo
        .txtproduct.Text = rsFindRecord!ProductCode
        .txtCompanyName.Text = rsFindRecord!CompanyName
        .txtSiteDetails.Text = rsFindRecord!PhysicalLocation & " IN " & rsFindRecord!Town
        If Not IsNull(rsFindRecord!SDate) Then
        .txtStartDate.Text = rsFindRecord!SDate
        End If
        If Not IsNull(rsFindRecord!EDate) Then
        .txtEndDate.Text = rsFindRecord!EDate
        End If
        If Not IsNull(rsFindRecord!EmailAddress) Then
        .txtEmail.Text = rsFindRecord!EmailAddress
        End If
        If Not IsNull(rsFindRecord!PhysicalAddress) Then
        .txtAddress1.Text = rsFindRecord!PhysicalAddress
        End If
        If Not IsNull(rsFindRecord!PostalAddress) Then
        .txtAddress2.Text = rsFindRecord!PostalAddress
        End If
        If Not IsNull(rsFindRecord!FAxNo) Then
        .txtFax.Text = rsFindRecord!FAxNo
        End If
        If Not IsNull(rsFindRecord!ContactPerson) Then
        .txtContactPerson.Text = rsFindRecord!ContactPerson
        End If
    End With
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuHow_Click()
Call HelpUsingEnterprise

End Sub

Private Sub HelpUsingEnterprise()
On Error GoTo err
With Screen.ActiveForm
    
    .HelpCommonDialog.DialogTitle = "Using the Main System"
    .HelpCommonDialog.HelpFile = App.HelpFile
    .HelpCommonDialog.HelpContext = 5
    .HelpCommonDialog.HelpCommand = cdlHelpContext
    .HelpCommonDialog.ShowHelp
    
End With
Exit Sub
err:
    MsgBox err.Number & vbCrLf & vbCrLf & err.Description
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
With Me
Select Case Button.Key
Case "N"
    Select Case Button.Caption
    Case "New &Record "
        If EditRecord Then Exit Sub
        .Frame2.Enabled = True: .Frame4.Enabled = True
        MyCommonData.ClearTextFields: .cboAccountType.SetFocus
        NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4
        .cboAccountType.SetFocus
    Case "Re&new?"
        .Frame2.Enabled = True: txtMonths.Locked = True: .UpDown2.Enabled = True: .txtMonths.SetFocus: Button.Caption = "&Save Record ": Button.Image = 4
    Case "&Save Record "
        If ValidRecord Then
            Set rsNewRecord = New ADODB.Recordset
            rsNewRecord.Open "SELECT * FROM ODASMJobBriefItems WHERE JobBriefItemNo = '" & .txtJobBriefItemNo.Text & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
            If IsNull(rsNewRecord!NoticeReceived) Then
                rsNewRecord!NoticeReceived = "Y"
                rsNewRecord!NoticeApproved = "N"
                rsNewRecord!NoticeAuthorized = "N"
                rsNewRecord!NoticeReceivedBy = CurrentUserName
                rsNewRecord!NoticeReceivedDate = Format(Date, "MMMM dd,yyyy")
                rsNewRecord!Status = "NOTICE-RECEIVED"
                rsNewRecord!RenewalPeriod = .txtMonths.Text
            Else
                rsNewRecord!Status = "BRIEF-RENEWED"
                rsNewRecord!RenewalPeriod = .txtMonths.Text
                rsNewRecord!expirydate = DateAdd("M", .txtMonths.Text, rsNewRecord!expirydate)
            
            End If
                rsNewRecord.Update
                .txtEndDate.Text = rsNewRecord!expirydate
            Set rsNewRecord = Nothing
            NewRecord = False: Button.Caption = "New &Record ": Button.Image = 2
            Call ListCurrentSettings
        End If
        
        
    Case Else
        Exit Sub
    End Select
Case "E"
    Select Case Button.Caption
        Case "&Edit/Change "
            If NewRecord Then Exit Sub
            EditRecord = True
        Case "Save &Changes "
        If EditRecord Then
            EditRecord = False
        End If
    Case Else
        Exit Sub
    End Select
Case "S"
    Case "R"
            If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                .Toolbar1.Buttons(2).Caption = "New &Record "
                .Toolbar1.Buttons(2).Image = 2
                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                .Toolbar1.Buttons(3).Image = 5
                NewRecord = False: EditRecord = False: MyCommonData.ClearTheScreen
    Case "P"
    Case "H"
            .HelpCommonDialog.DialogTitle = "Using the Main System"
            .HelpCommonDialog.HelpFile = App.HelpFile
            .HelpCommonDialog.HelpContext = 38
            .HelpCommonDialog.HelpCommand = cdlHelpContext
            .HelpCommonDialog.ShowHelp
End Select
End With
Exit Sub
err:
    ErrorMessage

End Sub

Private Sub UpDown1_Change()
With Screen.ActiveForm
    .txtYears.Text = .UpDown1.Value
End With
End Sub

Private Sub UpDown2_Change()
With Me
    .txtMonths.Text = .UpDown2.Value
End With
End Sub
