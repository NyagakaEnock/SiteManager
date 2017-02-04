VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSiteRegistration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Site Registration"
   ClientHeight    =   6810
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11160
   Icon            =   "frmSiteRegistration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   120
      TabIndex        =   19
      Top             =   6000
      Width           =   10935
      Begin VB.TextBox txtNotes 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   6720
         MultiLine       =   -1  'True
         TabIndex        =   41
         Top             =   240
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker dtpContractEnd 
         Height          =   255
         Left            =   4440
         TabIndex        =   39
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   57933825
         CurrentDate     =   38282
      End
      Begin MSComCtl2.DTPicker dtpContractStart 
         Height          =   255
         Left            =   1440
         TabIndex        =   38
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         CalendarBackColor=   16777215
         Format          =   57933825
         CurrentDate     =   38282
      End
      Begin VB.Label Label17 
         Caption         =   "Notes"
         Height          =   375
         Left            =   6120
         TabIndex        =   42
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Contract Ends"
         Height          =   255
         Left            =   3120
         TabIndex        =   40
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Contract begins"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1695
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
            Picture         =   "frmSiteRegistration.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSiteRegistration.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSiteRegistration.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSiteRegistration.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSiteRegistration.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSiteRegistration.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSiteRegistration.frx":236E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
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
   Begin VB.Frame Frame3 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   11175
      Begin VB.ComboBox cboBBNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   7800
         TabIndex        =   43
         Top             =   480
         Width           =   3255
      End
      Begin VB.Frame Frame4 
         Height          =   1215
         Left            =   7680
         TabIndex        =   18
         Top             =   2160
         Width           =   3375
         Begin VB.ComboBox cboElecPeriod 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   2280
            TabIndex        =   37
            Top             =   720
            Width           =   990
         End
         Begin VB.TextBox txtElecInterval 
            BackColor       =   &H00FFC0C0&
            Height          =   285
            Left            =   840
            TabIndex        =   29
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtElecCharge 
            BackColor       =   &H00FFC0C0&
            Height          =   285
            Left            =   840
            TabIndex        =   28
            Top             =   240
            Width           =   2415
         End
         Begin VB.CheckBox chkIlluminate 
            Caption         =   "Iluminated"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "Interval"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label15 
            Caption         =   "Charge"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Rent rates"
         Height          =   1215
         Left            =   3960
         TabIndex        =   17
         Top             =   2160
         Width           =   3615
         Begin VB.ComboBox cboRentPeriod 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   2520
            TabIndex        =   36
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtRentInterval 
            BackColor       =   &H00FFC0C0&
            Height          =   285
            Left            =   960
            TabIndex        =   27
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtRentCharge 
            BackColor       =   &H00FFC0C0&
            Height          =   285
            Left            =   960
            TabIndex        =   26
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label14 
            Caption         =   "Interval"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label12 
            Caption         =   "Charge"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            Height          =   15
            Left            =   240
            TabIndex        =   30
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Council rates"
         Height          =   1215
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   3735
         Begin VB.ComboBox cboCouncilPeriod 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   2640
            TabIndex        =   35
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtCouncilInterval 
            BackColor       =   &H00FFC0C0&
            Height          =   285
            Left            =   960
            TabIndex        =   23
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtCouncilCharge 
            BackColor       =   &H00FFC0C0&
            Height          =   285
            Left            =   960
            TabIndex        =   22
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label10 
            Caption         =   "Interval"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Charge"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.ComboBox cboCountry 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   4680
         TabIndex        =   13
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H00FFC0C0&
         Height          =   360
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   4
         Top             =   480
         Width           =   5055
      End
      Begin VB.TextBox txtPhysicalAddress 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   3
         Top             =   1080
         Width           =   9615
      End
      Begin VB.ComboBox cboTown 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   2
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ComboBox cboLandLord 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8040
         TabIndex        =   1
         Top             =   1560
         Width           =   3015
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1815
         Left            =   120
         TabIndex        =   14
         Top             =   3480
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   3201
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
         BackColor       =   16777152
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
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   11040
         X2              =   240
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label1 
         Caption         =   "Size"
         Height          =   255
         Left            =   7800
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "SiteCode"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Landlord"
         Height          =   315
         Left            =   7080
         TabIndex        =   9
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label Label5 
         Caption         =   "Country"
         Height          =   315
         Left            =   3960
         TabIndex        =   8
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Site Name"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2520
         TabIndex        =   7
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Physical Address"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Town/City:"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   825
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
Attribute VB_Name = "frmSiteRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim billbno




Private Sub GetCurrentSettings()
'''On Error GoTo Err
With frmODASPSiteRegistration

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Brokers Code", .ListView1.Width / 8 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Brokers Name", .ListView1.Width / 8 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Postal Address", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Branch", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Town/City", .ListView1.Width / 10
.ListView1.ColumnHeaders.Add , , "Country", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Contact Person", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Contact Title", .ListView1.Width / 10
.ListView1.ColumnHeaders.Add , , "Phone Number", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Email", .ListView1.Width / 6

.ListView1.View = lvwReport: .ListView1.Visible = True
End With
Exit Sub
err:
    ErrorMessage
End Sub



Private Sub cboBBNo_GotFocus()
If Not NewRecord And Not EditRecord Then Exit Sub
With frmODASPSiteRegistration
        Set rsCOMBO = cnCOMMON.Execute("SELECT * FROM AdvertBBDetails;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    .cboBBNo.Clear
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!Length) And rsCOMBO!Length <> "" Then
        If Not IsNull(rsCOMBO!Width) And rsCOMBO!Width <> "" Then
        .cboBBNo.AddItem rsCOMBO!Length + " x " + rsCOMBO!Width
        billbno = rsCOMBO!BillBoardNo
        End If
        End If
    rsCOMBO.MoveNext
    Loop
'    .cboPaymentMethod.AddItem "<Add New>"
    
OUTS:
    Set rsCOMBO = Nothing
End With
Exit Sub
err:
    ErrorMessage

End Sub

Private Sub cboBBNo_LostFocus()
'''On Error GoTo Err
With frmODASPSiteRegistration
.cboBBNo.Text = billbno
End With
Exit Sub
err:
   ErrorMessage
End Sub

Private Sub cboCouncilPeriod_Click()
frmODASPSiteRegistration.ListView1.SetFocus
End Sub

Private Sub cboCouncilPeriod_GotFocus()
If Not NewRecord And Not EditRecord Then Exit Sub
With frmODASPSiteRegistration
    AttachSQL = "SELECT Intervalname AS SelectField FROM ParamIntervals ORDER BY Intervalname;"
    .cboCountry.Clear
    MyCommonData.AttachDropDown
End With

End Sub

Private Sub cboCountry_Click()
frmODASPSiteRegistration.ListView1.SetFocus
End Sub

Private Sub cboCountry_gotfocus()
If Not NewRecord And Not EditRecord Then Exit Sub
With frmODASPSiteRegistration
    AttachSQL = "SELECT Country AS SelectField FROM ParamCountries ORDER BY Country;"
    .cboCountry.Clear
    MyCommonData.AttachDropDown
End With

End Sub



Private Sub cboElecPeriod_Click()
frmODASPSiteRegistration.ListView1.SetFocus
End Sub

Private Sub cboElecPeriod_GotFocus()
If Not NewRecord And Not EditRecord Then Exit Sub
With frmODASPSiteRegistration
    AttachSQL = "SELECT Intervalname AS SelectField FROM ParamIntervals ORDER BY Intervalname;"
    .cboCountry.Clear
    MyCommonData.AttachDropDown
End With

End Sub

Private Sub cboLandLord_Click()
frmODASPSiteRegistration.ListView1.SetFocus
End Sub

Private Sub cboLandLord_GotFocus()
If Not NewRecord And Not EditRecord Then Exit Sub
With frmODASPSiteRegistration
    AttachSQL = "SELECT Surname AS SelectField FROM AdvertSiteLords ORDER BY Surname;"
    .cboLandLord.Clear
    MyCommonData.AttachDropDown
End With
End Sub

Private Sub cboLandLord_LostFocus()
With frmODASPSiteRegistration
Set rsFindRecord = New ADODB.Recordset
  rsFindRecord.Open "SELECT Landlordno FROM Advertsitelords WHERE Surname = '" & Trim(.cboLandLord.Text) & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
    If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub: Set rsFindRecord = Nothing
    
      .cboLandLord.Text = rsFindRecord!LandLordNo
  
    Set rsFindRecord = Nothing
End With
End Sub

Private Sub cboRentPeriod_Click()
frmODASPSiteRegistration.ListView1.SetFocus
End Sub

Private Sub cboRentPeriod_GotFocus()
If Not NewRecord And Not EditRecord Then Exit Sub
With frmODASPSiteRegistration
    AttachSQL = "SELECT Intervalname AS SelectField FROM ParamIntervals ORDER BY Intervalname;"
    .cboCountry.Clear
    MyCommonData.AttachDropDown
End With

End Sub

Private Sub cboTown_Click()
frmODASPSiteRegistration.ListView1.SetFocus
End Sub

Private Sub cboTown_gotfocus()
If Not NewRecord And Not EditRecord Then Exit Sub
With frmODASPSiteRegistration
    AttachSQL = "SELECT Town AS SelectField FROM ODASPTown ORDER BY Town;"
    .cboTown.Clear
    MyCommonData.AttachDropDown
End With

End Sub

Private Sub Combo3_Change()

End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub


Private Sub chkIlluminate_Click()
With frmODASPSiteRegistration
If Not NewRecord And Not EditRecord Then Exit Sub
 If .chkIlluminate.Value = 1 Then
   .Frame4.Enabled = False
   Else
   .Frame4.Enabled = True
 End If
End With
End Sub

Private Sub Form_Activate()
  Call GetCurrentSettings

End Sub

Private Sub Form_Initialize()
Set MyCommonData = New clsCommonData
'Set mycommondata New clsCommonData

End Sub

Private Sub Form_Load()
'Call OpenConnection
End Sub

Private Sub mnuClear_Click()
    MyCommonData.ClearTextFields
End Sub

Private Sub mnuClose_Click()
    Unload frmODASPSiteRegistration
End Sub

Private Sub mnuCurrent_Click()
    Call ShowCurrentSettings

End Sub
Private Sub ShowCurrentSettings()
'''On Error GoTo Err
With frmODASPSiteRegistration
Screen.MousePointer = vbHourglass

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Brokers Code", .ListView1.Width / 8 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Brokers Name", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Postal Address", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Branch", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Town/City", .ListView1.Width / 10
.ListView1.ColumnHeaders.Add , , "Country", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Contact Person", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Contact Title", .ListView1.Width / 3.5
.ListView1.ColumnHeaders.Add , , "Phone Number", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Email", .ListView1.Width / 5

.ListView1.View = lvwReport: .ListView1.Visible = True

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT ParamInsuranceBrokers.* FROM ParamInsuranceBrokers ORDER BY ParamInsuranceBrokers.BrokersCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem, NCount As Double

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Screen.MousePointer = vbDefault: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!BrokersCode))
    
    If Not IsNull(rsLIST!BrokersCode) Then
        MyList.SubItems(1) = CStr(rsLIST!BrokersName)
        MyList.SubItems(2) = CStr(rsLIST!PhysicalAddress)
        MyList.SubItems(3) = CStr(rsLIST!postaladdress)
        MyList.SubItems(4) = CStr(rsLIST!Branch)
        MyList.SubItems(5) = CStr(rsLIST!towncity)
        MyList.SubItems(6) = CStr(rsLIST!country)
        MyList.SubItems(7) = CStr(rsLIST!ContactPerson)
        MyList.SubItems(8) = CStr(rsLIST!ContactTitle)
        MyList.SubItems(9) = CStr(rsLIST!TelephoneNo)
        MyList.SubItems(10) = CStr(rsLIST!Email)

    End If

    rsLIST.MoveNext
    
Wend

    Set MyList = Nothing: Set rsLIST = Nothing: Screen.MousePointer = vbDefault
        
End With
Exit Sub
err:
If err.Number = 3265 Then
    Resume Next
Else
    Screen.MousePointer = vbDefault
    ErrorMessage
End If
End Sub
Private Function ValidRecord() As Boolean
With frmODASPSiteRegistration
    If .txtCode.Text = Empty Then
        strMessage = "Site code required...!"
        .txtCode.SetFocus
    ElseIf .txtName.Text = Empty Then
        strMessage = "Site name required...!"
        .txtName.SetFocus
    ElseIf .cboBBNo = Empty Then
        strMessage = "Billboard number required...!"
        .cboBBNo.SetFocus
    ElseIf .txtPhysicalAddress.Text = Empty Then
        strMessage = "Required Physical Address...!"
        .txtPhysicalAddress.SetFocus
'    ElseIf .txtAddress2.Text = Empty Then
'        strMessage = "Required Postal Address...!"
'        .txtAddress2.SetFocus
'    ElseIf .cboTown.Text = Empty Then
'        strMessage = "Required Town/City...!"
'        .cboTown.SetFocus
'    ElseIf .cboCountry.Text = Empty Then
'        strMessage = "Required Country...!"
'        .cboCountry.SetFocus
'    ElseIf .txtPerson.Text = Empty Then
'        strMessage = "Required contact Person...!"
'        .txtPerson.SetFocus
'    ElseIf .cboTitle.Text = Empty Then
'        strMessage = "Required Contact Title...!"
'        .cboTitle.SetFocus
'    ElseIf .txtPhone.Text = Empty Then
'        strMessage = "Required Telephone...!"
'        .txtPhone.SetFocus
'    ElseIf .txtEmail.Text = Empty Then
'        strMessage = "Required Email Address...!"
'        .txtEmail.SetFocus

    Else
        ValidRecord = True
    End If
    If Not ValidRecord Then
        MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
    End If
End With
End Function

Private Sub Text8_Change()

End Sub

Private Sub Text10_Change()

End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
''On Error GoTo Err
        
        With frmODASPSiteRegistration
        
        Select Case Button.Key
        Case "N"
            Select Case Button.Caption
            Case "New &Record "
                If EditRecord Then Exit Sub
                MyCommonData.ClearTextFields: .ListView1.ListItems.Clear: .ListView2.ListItems.Clear: .txtName.SetFocus
                NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4
                .txtName.SetFocus
                .txtQuotationNo.Text = AutoPurchaseOrderNo
                .dtpQuotationDate.Value = MyCurrentDate
                
            Case "&Save Record "
                If NewRecord Then
                        If ValidRecord Then
                                SaveRECORD
                            .Toolbar1.Buttons(3).Caption = "FINISH"
                        End If
                End If
                
             Case "NE&XT ITEM"
                  
            Case Else
                Exit Sub
            End Select
            
        Case "E"
            Select Case Button.Caption
            Case "&Edit/Change "
                If NewRecord Then Exit Sub
                        If .txtCode.Text = Empty Then
                            MsgBox "There is NO Current Record to Edit. Please Search and Display a Record First...!", vbCritical + vbOKOnly, ""
                           .txtCode.SetFocus
                        Else
                           .txtCode.Locked = True
                            Button.Caption = "Save &Changes ": Button.Image = 4
                            EditRecord = True
                        End If
            Case "Save &Changes "
                If EditRecord Then
                    If ValidRecord Then
                            SaveRECORD
                            Set rsEditRecord = Nothing
                            .txtCode.Locked = False: EditRecord = False: Button.Caption = "&Edit/Change ": Button.Image = 5
                    End If
                End If
             Case "FINISH"
                 If ValidMainRecord Then
                        .Toolbar1.Buttons(2).Caption = "New &Record "
                        .Toolbar1.Buttons(2).Image = 2
                        .Toolbar1.Buttons(3).Image = 5
                        .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                     
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
            Load frmRptAdvertPrintOut
            frmRptAdvertPrintOut.Show 1, frmODASPSiteRegistration
        Case "F"
             
             
        Case Else
            Exit Sub
        End Select
        End With
Exit Sub
err:
    ErrorMessage

End Sub


