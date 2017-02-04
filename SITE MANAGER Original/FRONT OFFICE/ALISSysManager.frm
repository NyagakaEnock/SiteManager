VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form ALISSysManager 
   Caption         =   "A.L.I.S. PLUS ~ [SETTINGS AND SYSTEM INFORMATION]"
   ClientHeight    =   7200
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10995
   Icon            =   "ALISSysManager.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "ALISSysManager.frx":0442
   ScaleHeight     =   7200
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   210
      Left            =   2160
      TabIndex        =   5
      Top             =   6960
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   370
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   6945
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   3704
            MinWidth        =   3704
            Picture         =   "ALISSysManager.frx":5B80
            Text            =   "System Manager"
            TextSave        =   "System Manager"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11360
            MinWidth        =   11360
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3598
            MinWidth        =   3598
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "01/06/2004"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ALISSysManager.frx":5FD2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6135
      Left            =   2880
      TabIndex        =   1
      Top             =   720
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   10821
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
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView TreeView2 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   10821
      _Version        =   393217
      Indentation     =   531
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   1058
      ButtonWidth     =   3334
      ButtonHeight    =   1005
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Company Details"
            Key             =   "COY"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Staff Information"
            Key             =   "EMP"
            ImageIndex      =   30
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Options/Settings"
            Key             =   "OPT"
            ImageIndex      =   28
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "User Guide/Help"
            Key             =   "HLP"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   8880
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   40
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":62EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":6606
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":6C80
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":72FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":7974
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":7FEE
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":8668
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":8CE2
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":935C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":99D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":A050
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":A6CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":AD44
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":B3BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":BA38
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":BD52
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":C06C
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":C4BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":C910
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":CD62
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":D1B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":D606
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":DA58
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":DEAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":E2FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":E74E
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":EBA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":EFF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":F444
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":F89E
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":FCF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":10142
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":10594
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":109E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":10E38
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":1128A
               Key             =   ""
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":116DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":11B2E
               Key             =   ""
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":11E48
               Key             =   ""
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ALISSysManager.frx":1259A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtTASK 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   375
         Left            =   10800
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuActionClearData 
         Caption         =   "&Clear Datasheet"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActionPrintSelect 
         Caption         =   "Print &Selected Records"
      End
      Begin VB.Menu mnui8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActionPrintAll 
         Caption         =   "Print &All Displayed"
      End
   End
   Begin VB.Menu mnuparameterreports 
      Caption         =   "&Parameter Reports"
      Begin VB.Menu mnagents 
         Caption         =   "&Agents"
         Begin VB.Menu mnuagents 
            Caption         =   "Agents Details"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuagentsbenefits 
            Caption         =   "Agents Benefits"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnu7 
         Caption         =   "-"
      End
      Begin VB.Menu mnupaymentreceipts 
         Caption         =   "&Payment Receipts"
         Begin VB.Menu mnuagentspay 
            Caption         =   "Agents Pay SetUp"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnureceipts 
            Caption         =   "Receipts SetUp"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuj9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclaims 
         Caption         =   "&Claims"
         Begin VB.Menu mnurequirements 
            Caption         =   "General Requirements"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuclaimsconfiguration 
            Caption         =   "Claims Configuration"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuclaimscauses 
            Caption         =   "Causes of Claims"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuReinstatementTypes 
            Caption         =   "Reinstatement Types"
         End
         Begin VB.Menu mnuReinstatementRates 
            Caption         =   "Reinstatement Rates"
         End
         Begin VB.Menu mnuPaidupSetup 
            Caption         =   "Paid up Setup"
         End
      End
      Begin VB.Menu mnu0o 
         Caption         =   "-"
      End
      Begin VB.Menu mnuloans 
         Caption         =   "&Loans"
         Begin VB.Menu mnuloansetup 
            Caption         =   "Loan SetUp/Type"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuloanapprovers 
            Caption         =   "Loan Approvers"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuloanoperationstype 
            Caption         =   "Loan Operations Type"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuPersonalAccident 
            Caption         =   "Personal Accident"
         End
      End
      Begin VB.Menu mnu0p 
         Caption         =   "-"
      End
      Begin VB.Menu mnuperiodssetup 
         Caption         =   "&Periods SetUp"
         Begin VB.Menu mnuperiods 
            Caption         =   "Periods SetUp"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnusurrender 
            Caption         =   "Surrender SetUp"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnu06 
         Caption         =   "-"
      End
      Begin VB.Menu mnnucompannyinfo 
         Caption         =   "&Company Information"
         Begin VB.Menu mnucompanydetails 
            Caption         =   "Company Details"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuratetable 
            Caption         =   "Rate Table"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnujointagetables 
            Caption         =   "Joint Age Tables"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnubabkssetup 
            Caption         =   "Banks SetUp"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuproductsetup 
            Caption         =   "Products' SetUp"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnurelationshipcodes 
            Caption         =   "Relationship Codes"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuempmaster 
            Caption         =   "Employees Master"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnubranch 
            Caption         =   "Branch"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnulastnumbers 
            Caption         =   "Last Numbers"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnud3 
         Caption         =   "-"
      End
      Begin VB.Menu mnumoreparameters 
         Caption         =   "&More Parameters"
         Begin VB.Menu mnucompanymaster 
            Caption         =   "Company Master"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuemployer 
            Caption         =   "Employer Information"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuaccperiods 
            Caption         =   "Accounts Periods"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnupayintervals 
            Caption         =   "Payment Intervals"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnucurrencies 
            Caption         =   "Currency"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnumoffice 
            Caption         =   "M Office"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnucompanybranches 
            Caption         =   "Company Branches"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnupaymethod 
            Caption         =   "Payment Method"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnucountries 
            Caption         =   "Countries"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnucities 
            Caption         =   "Cities"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuenquiry 
            Caption         =   "Enquiry"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnutaxes 
            Caption         =   "Taxes"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnufeesservices 
            Caption         =   "Fees Services"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnutittls 
            Caption         =   "Tittles"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnutown 
            Caption         =   "Town"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnucompanydepts 
            Caption         =   "Company Departments"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu mnuHelpSystem 
      Caption         =   "&Help System"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Help &Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "Search for &Help On..."
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnu4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "A&bout the System..."
      End
   End
End
Attribute VB_Name = "ALISSysManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SysManage As clsSysManager

Private Sub Form_Activate()
    SysManage.GetMainStructure
    Me.ProgressBar1.Visible = False
End Sub

Private Sub Form_Load()
    Call OpenConnection
    Set SysManage = New clsSysManager
    SysManage.NewTreeSetup
End Sub

Private Sub Form_Resize()
    SysManage.ResizeControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo err
If NewRecord Or EditRecord Then MsgBox "Data Entry or Edit in Progress! No Work was Done!", vbInformation + vbOKOnly, "Screen Unload": Cancel = 1: Exit Sub
    
    If MsgBox("Are You Sure You Want to Quit System Management Console", vbQuestion + vbYesNo + vbMsgBoxHelpButton, "Shut Down", "ALISHELP.HLP", 246) = vbNo Then Cancel = True: Exit Sub
    
    Call UpdateLogoutRecord
    
    End
    
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo err
    If Me.ListView1.ListItems.Count = 0 Then Exit Sub

    If Button = 2 Then
        PopupMenu mnuActions, , , , mnuActionClearData
    End If
    
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuaccperiods_Click()
On Error GoTo err
accperiods = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuActionClearData_Click()
On Error GoTo err
    With Me
    If .ListView1.ListItems.Count = 0 Then Exit Sub
        
        If MsgBox("Clear All the Currently Displayed Records?", vbQuestion + vbYesNo + vbDefaultButton2, "Clear Datasheet") = vbNo Then Exit Sub
        .ListView1.ListItems.Clear
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuagents_Click()
On Error GoTo err
agents = True
    Load frmParamReports
    frmParamReports.Show 1, ALISSysManager
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuagentsbenefits_Click()
On Error GoTo err
agentsbenefits = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuagentspay_Click()
On Error GoTo err
agentspay = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnubabkssetup_Click()
On Error GoTo err
bankssetup = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage

End Sub

Private Sub mnubranch_Click()
On Error GoTo err
  
    Load frmParamBranch
    frmParamBranch.Show 1, ALISSysManager
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnucities_Click()
On Error GoTo err
cities = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuclaimscauses_Click()
On Error GoTo err
claimcauses = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuclaimsconfiguration_Click()
On Error GoTo err
claimconfig = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnucompanybranches_Click()
On Error GoTo err
companybranch = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage

End Sub

Private Sub mnucompanydepts_Click()
On Error GoTo err
companydepts = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage

End Sub

Private Sub mnucompanydetails_Click()
On Error GoTo err
    companydetails = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage

End Sub

Private Sub mnucompanymaster_Click()
On Error GoTo err
companymaster = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage

End Sub

Private Sub mnucountries_Click()
On Error GoTo err
countries = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnucurrencies_Click()
On Error GoTo err
currencies = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuemployer_Click()
On Error GoTo err
employers = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage

End Sub

Private Sub mnuempmaster_Click()
On Error GoTo err
empmaster = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage

End Sub

Private Sub mnuenquiry_Click()
On Error GoTo err
enquiry = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage

End Sub

Private Sub mnufeesservices_Click()
On Error GoTo err
feesservices = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuFileClose_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    SysManage.HelpAbout
End Sub

Private Sub mnuHelpContents_Click()
    SysManage.HelpContents
End Sub

Private Sub mnuHelpIndex_Click()
    SysManage.HelpIndex
End Sub
'==

Private Sub mnujointagetables_Click()
On Error GoTo err
jointage = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage

End Sub

Private Sub mnulastnumbers_Click()
On Error GoTo err
lastnumbers = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage

End Sub

Private Sub mnuLeterCategory_Click()
On Error GoTo err
    Load frmALISPLetterCategory
    frmALISPLetterCategory.Show 1, Me
    Exit Sub
err:
    ErrorMessage

End Sub

Private Sub mnuLeterReceipient_Click()
On Error GoTo err
    Load frmALISPLetterReceipient
    frmALISPLetterReceipient.Show 1, Me
    Exit Sub
err:
    ErrorMessage

End Sub



Private Sub mnuLetterTemplate_Click()
    Load frmALISPLetterTemplate
    frmALISPLetterTemplate.Show 1, Me
    Exit Sub
End Sub

Private Sub mnuloanapprovers_Click()
On Error GoTo err
loanapprovers = True
    Load frmParamReports
    frmParamReports.Show 1, ALISSysManager
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuloanoperationstype_Click()
On Error GoTo err
loanoptype = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuloansetup_Click()
On Error GoTo err
loantype = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnumoffice_Click()
On Error GoTo err
moffice = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage

End Sub

Private Sub mnupayintervals_Click()
On Error GoTo err
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnupaymethod_Click()
On Error GoTo err
paymethods = True
    Load frmParamReports
    frmParamReports.Show 1, ALISSysManager
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuperiods_Click()
On Error GoTo err
periodssetup = True
    Load frmParamReports
    frmParamReports.Show 1, ALISSysManager
    Exit Sub
err:
    ErrorMessage

End Sub

Private Sub mnuPersonalAccident_Click()
On Error GoTo err
    Load frmALISPAccident
    frmALISPAccident.Show 1
    Exit Sub
err:
    ErrorMessage

End Sub

Private Sub mnuproductsetup_Click()
On Error GoTo err
  
    Load frmParamProduct
    frmParamProduct.Show 1, ALISSysManager
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuratetable_Click()
On Error GoTo err
ratetable = True
    Load frmParamReports
    frmParamReports.Show 1, ALISSysManager
    Exit Sub
err:
    ErrorMessage

End Sub

Private Sub mnureceipts_Click()
On Error GoTo err
receipts = True
    Load frmParamReports
    frmParamReports.Show 1, ALISSysManager
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnurelationshipcodes_Click()
On Error GoTo err
  
    Load frmParamRelationships
    frmParamRelationships.Show 1, ALISSysManager
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnurequirements_Click()
On Error GoTo err
claimsreq = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnusurrender_Click()
On Error GoTo err
surrendersetup = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage

End Sub

Private Sub mnutaxes_Click()
On Error GoTo err
taxes = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnutittls_Click()
On Error GoTo err
tittles = True
    Load frmParamReports
    frmParamReports.Show 1
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnutown_Click()

On Error GoTo err
    
    Load frmParamTown
    frmParamTown.Show 1

Exit Sub
err:
    ErrorMessage

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
With ALISSysManager
    Select Case Button.Key
    Case "COY"
        .txtTASK.Text = "C1"
        Load frmCompany
        frmCompany.Show 1, ALISSysManager
    Case "EMP"
        .txtTASK.Text = "C4"
        Load frmEmployeesPersonal
        frmEmployeesPersonal.Show 1, ALISSysManager
    Case "OPT"
        .txtTASK.Text = "C1"
        Load frmSettings
        frmSettings.Show 1, ALISSysManager
    Case "HLP"
        SysManage.HelpIndex
    Case Else
        Exit Sub
    End Select
End With
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub TreeView2_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo err
With ALISSysManager
    Select Case Node.Key
    Case "C1"
        .txtTASK.Text = "C1"
        Load frmALISPCoyDetails
        frmALISPCoyDetails.Show 1, ALISSysManager
    Case "C2"
        .txtTASK.Text = "C2"
        Load frmEmployeesPersonal
        frmEmployeesPersonal.Show 1, ALISSysManager
    Case "C3"
        .txtTASK.Text = "C3"
        Load frmALISPProduct
        frmALISPProduct.Show 1, ALISSysManager
    Case "C4"
        .txtTASK.Text = "C4"
        Load frmALISPClientSetup
        frmALISPClientSetup.Show 1, ALISSysManager
    Case "C5"
        .txtTASK.Text = "C5"
        Load frmALISPRateTable
        frmALISPRateTable.Show 1, ALISSysManager
    Case "C6"
        .txtTASK.Text = "C6"
        Load frmALISPJointAge
        frmALISPJointAge.Show 1, ALISSysManager
    Case "C7"
        .txtTASK.Text = "C7"
        Load frmALISPBanks
        frmALISPBanks.Show 1, ALISSysManager
    Case "C8"
        .txtTASK.Text = "C8"
        Load frmALISPRelationship
        frmALISPRelationship.Show 1, ALISSysManager
    Case "C9"
        .txtTASK.Text = "C9"
        Load frmALISPLastNumber
        frmALISPLastNumber.Show 1, ALISSysManager
    Case "C10"
        .txtTASK.Text = "C10"
        Load frmALISPInstallments
        frmALISPInstallments.Show 1, ALISSysManager
        
    Case "R1"
        .txtTASK.Text = "R1"
        Load frmALISPDefaults
        frmALISPDefaults.Show 1, ALISSysManager
    Case "R2"
        .txtTASK.Text = "R2"
        Load frmALISPSystemRates
        frmALISPSystemRates.Show 1, ALISSysManager
    Case "R3"
        .txtTASK.Text = "R3"
        Load frmALISPBenefit
        frmALISPBenefit.Show 1, ALISSysManager
    Case "R4"
        .txtTASK.Text = "R4"
        Load frmALISPPlanBenefit
        frmALISPPlanBenefit.Show 1, ALISSysManager
    Case "R5"
        .txtTASK.Text = "R5"
        Load frmALISPDiscount
        frmALISPDiscount.Show 1, ALISSysManager
    Case "R6"
        .txtTASK.Text = "R6"
        Load frmALISPDepartment
        frmALISPDepartment.Show 1, ALISSysManager
  
    Case "R3"
        .txtTASK.Text = "R3"
        Load frmALISPBenefit
        frmALISPBenefit.Show 1, ALISSysManager
    Case "R4"
        .txtTASK.Text = "R4"
        Load frmALISPPlanBenefit
        frmALISPPlanBenefit.Show 1, ALISSysManager
    Case "R5"
        .txtTASK.Text = "R5"
        Load frmALISPDiscount
        frmALISPDiscount.Show 1, ALISSysManager
    
    Case "M1"
        .txtTASK.Text = "M1"
        Load frmALISPUWDecision
        frmALISPUWDecision.Show 1, ALISSysManager
    Case "M2"
        .txtTASK.Text = "M2"
        Load frmALISPRatingType
        frmALISPRatingType.Show 1, ALISSysManager
    Case "M3"
        .txtTASK.Text = "M3"
        Load frmALISPMedical
        frmALISPMedical.Show 1, ALISSysManager
    
    Case "M4"
        .txtTASK.Text = "M4"
        Load frmALISPDoctor
        frmALISPDoctor.Show 1, ALISSysManager
    
    Case "T1"
        .txtTASK.Text = "T1"
        Load frmALISPReferenceDefaults
        frmALISPReferenceDefaults.Show 1, ALISSysManager
    Case "T2"
        .txtTASK.Text = "T2"
'        Load frmALISPRatingType
'        frmALISPRatingType.Show 1, ALISSysManager
    Case "T3"
        .txtTASK.Text = "T3"
        Load frmALISPDoctor
        frmALISPDoctor.Show 1, ALISSysManager
        
    Case "G1"
        .txtTASK.Text = "G1"
        Load frmALISPCreditor
        frmALISPCreditor.Show 1, ALISSysManager
    Case "G2"
        .txtTASK.Text = "G2"
        Load frmALISPAgentsPay
        frmALISPAgentsPay.Show 1, ALISSysManager
    Case "G3"
        .txtTASK.Text = "G3"
        Load frmALISPAgentBenefits
        frmALISPAgentBenefits.Show 1, ALISSysManager
        
    Case "Y1"
        .txtTASK.Text = "Y1"
        Load frmALISPAgentsPay
        frmALISPAgentsPay.Show 1, ALISSysManager
    Case "Y2"
        .txtTASK.Text = "Y2"
        Load frmALISPReceipt
        frmALISPReceipt.Show 1, ALISSysManager
        
    Case "Z1"
        .txtTASK.Text = "Z1"
        Load frmALISPPeriod
        frmALISPPeriod.Show 1, ALISSysManager
    Case "Z2"
        .txtTASK.Text = "Z2"
        Load frmALISPSurrender
        frmALISPSurrender.Show 1, ALISSysManager
        
    Case "S1"
        .txtTASK.Text = "S1"
        Load frmALISPClaim
        frmALISPClaim.Show 1, ALISSysManager
    Case "S2"
        .txtTASK.Text = "S2"
        Load frmALISPClaimCfg
        frmALISPClaimCfg.Show 1, ALISSysManager
    Case "S3"
        .txtTASK.Text = "S3"
        Load frmALISPClaimCauses
        frmALISPClaimCauses.Show 1, ALISSysManager
    Case "S4"
        .txtTASK.Text = "S4"
        Load frmALISPLoanType
        frmALISPLoanType.Show 1, ALISSysManager
    Case "S5"
        .txtTASK.Text = "S5"
        Load frmALISPLoanApprover
        frmALISPLoanApprover.Show 1, ALISSysManager
    Case "S6"
        .txtTASK.Text = "S6"
        Load frmALISPLoanOperationType
        frmALISPLoanOperationType.Show 1, ALISSysManager
    Case "S7"
        .txtTASK.Text = "S7"
        Load frmALISPAccident
        frmALISPAccident.Show 1, ALISSysManager
    Case "S8"
        .txtTASK.Text = "S8"
        Load frmALISPReinstatementType
        frmALISPReinstatementType.Show 1, ALISSysManager
    Case "S9"
        .txtTASK.Text = "S9"
        Load frmALISPReinstatement
        frmALISPReinstatement.Show 1, ALISSysManager
    Case "S10"
        .txtTASK.Text = "S10"
        Load frmALISPPaidup
        frmALISPPaidup.Show 1, ALISSysManager

    Case "D1"
        .txtTASK.Text = "D1"
        Load frmParamEmployers
        frmParamEmployers.Show 1, ALISSysManager
    Case "D2"
        .txtTASK.Text = "D2"
        Load frmParamCountries
        frmParamCountries.Show 1, ALISSysManager
    Case "D3"
        .txtTASK.Text = "D3"
        Load frmParamCities
        frmParamCities.Show 1, ALISSysManager
    Case "D4"
        .txtTASK.Text = "D4"
        Load frmParamCurrencies
        frmParamCurrencies.Show 1, ALISSysManager
    Case "D5"
        .txtTASK.Text = "D5"
        Load frmParamTitles
        frmParamTitles.Show 1, ALISSysManager
    Case "D6"
        .txtTASK.Text = "D6"
        Load frmParamPayMethods
        frmParamPayMethods.Show 1, ALISSysManager
    Case "D7"
        .txtTASK.Text = "D7"
        Load frmParamFeeServices
        frmParamFeeServices.Show 1, ALISSysManager
    Case "D8"
        .txtTASK.Text = "D8"
        Load frmParamTaxes
        frmParamTaxes.Show 1, ALISSysManager
    Case "D9"
        .txtTASK.Text = "D9"
        Load frmParamAccPeriods
        frmParamAccPeriods.Show 1, ALISSysManager
    Case "D10"
        .txtTASK.Text = "D10"
        Load frmParamPayIntervals
        frmParamPayIntervals.Show 1, ALISSysManager
    Case "D11"
        .txtTASK.Text = "D11"
       Load frmALISPOtherInquiry
       frmALISPOtherInquiry.Show 1, Me
    
    Case "P0"
        .txtTASK.Text = "P0"
        Load frmSettings
        frmSettings.Show 1, ALISSysManager
    Case "P1"
        .txtTASK.Text = "P1"
        Load frmOfficeSettings
        frmOfficeSettings.Show 1, ALISSysManager
    Case "P2"
        .txtTASK.Text = "P2"
        SysManage.FindDefaultCurrency
    Case "P3"
        .txtTASK.Text = "P3"
        SysManage.FindLocalCurrency
    Case "P4"
        .txtTASK.Text = "P4"
        SysManage.FindVATRate
    Case "P5"
        .txtTASK.Text = "P5"
        SysManage.FindCountryCode
    Case "P6"
        .txtTASK.Text = "P6"
        SysManage.FindAreaCode
    Case "P7"
        .txtTASK.Text = "P7"
        SysManage.FindDefaultPayMethod
    Case "P8"
        .txtTASK.Text = "P8"
        SysManage.FindMSOffice
    Case "V1"
        .txtTASK.Text = "V1"
        SysManage.FindEmployees
    Case "V2"
        .txtTASK.Text = "V2"
        SysManage.FindEmployers
    Case "V3"
        .txtTASK.Text = "V3"
        SysManage.FindCurrencies
    Case "V4"
        .txtTASK.Text = "V4"
        SysManage.FindTaxesRates
    Case "V5"
        .txtTASK.Text = "V5"
        SysManage.FindFeeServices
    Case "V6"
        .txtTASK.Text = "V6"
        SysManage.FindPaymentMethods
    Case "V7"
        .txtTASK.Text = "V7"
        SysManage.FindMyCompany
    Case "V8"
        .txtTASK.Text = "V8"
        SysManage.FindDepartments
    Case "V9"
        .txtTASK.Text = "V9"
        SysManage.FindCompanyBranch
    'Document Management
    Case "Doc1"
          .txtTASK.Text = "Doc1"
          Load frmALISPLetterCategory
          frmALISPLetterCategory.Show 1, Me
    Case "Doc2"
          .txtTASK.Text = "Doc2"
          Load frmALISPLetterReceipient
         frmALISPLetterReceipient.Show 1, Me
    Case "Doc3"
          .txtTASK.Text = "Doc3"
          Load frmALISPLetterTemplate
          frmALISPLetterTemplate.Show 1, Me
          
    Case Else
        Exit Sub
    End Select
    
End With
Exit Sub
err:
    ErrorMessage
End Sub
