VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form SysManMain 
   Caption         =   "SETTINGS AND SYSTEM INFORMATION"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   Icon            =   "SysManMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "SysManMain.frx":0442
   ScaleHeight     =   7200
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
            Object.Width           =   4304
            MinWidth        =   4304
            Picture         =   "SysManMain.frx":5B80
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7832
            MinWidth        =   7832
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   4304
            MinWidth        =   4304
            TextSave        =   "19/08/2003"
         EndProperty
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
            Picture         =   "SysManMain.frx":5FD2
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
      AllowReorder    =   -1  'True
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
               Picture         =   "SysManMain.frx":62EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":6606
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":6C80
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":72FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":7974
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":7FEE
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":8668
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":8CE2
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":935C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":99D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":A050
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":A6CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":AD44
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":B3BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":BA38
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":BD52
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":C06C
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":C4BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":C910
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":CD62
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":D1B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":D606
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":DA58
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":DEAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":E2FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":E74E
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":EBA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":EFF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":F444
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":F89E
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":FCF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":10142
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":10594
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":109E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":10E38
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":1128A
               Key             =   ""
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":116DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":11B2E
               Key             =   ""
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":11E48
               Key             =   ""
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SysManMain.frx":1259A
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
End
Attribute VB_Name = "SysManMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SysManage As clsSysManager

Private Sub Form_Activate()
    SysManage.GetMainStructure
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
If NewRecord Or EditRecord Then MsgBox "Data Entry or Edit in Progress! No Work was Done!", vbInformation + vbOKOnly, "Screen Unload": Cancel = 1: Exit Sub
    Call UpdateLogoutRecord
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
With SysManMain
    Select Case Button.Key
    Case "COY"
        .txtTASK.Text = "C1"
        Load frmCompany
        frmCompany.Show 1, SysManMain
    Case "EMP"
        .txtTASK.Text = "C4"
        Load frmEmployeesPersonal
        frmEmployeesPersonal.Show 1, SysManMain
    Case "OPT"
        .txtTASK.Text = "C1"
        Load frmSettings
        frmSettings.Show 1, SysManMain
    Case "HLP"
        SysManage.HelpIndex
    Case Else
        Exit Sub
    End Select
End With
    Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub TreeView2_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo Err
With SysManMain
    Select Case Node.Key
    Case "C1"
        .txtTASK.Text = "C1"
        Load frmCompany
        frmCompany.Show 1, SysManMain
    Case "C2"
        .txtTASK.Text = "C2"
        Load frmCompanyBranch
        frmCompanyBranch.Show 1, SysManMain
    Case "C3"
        .txtTASK.Text = "C3"
        Load frmCompanyDepartments
        frmCompanyDepartments.Show 1, SysManMain
    Case "C4"
        .txtTASK.Text = "C4"
        Load frmEmployeesPersonal
        frmEmployeesPersonal.Show 1, SysManMain
    Case "D1"
        .txtTASK.Text = "D1"
        Load frmParamEmployers
        frmParamEmployers.Show 1, SysManMain
    Case "D2"
        .txtTASK.Text = "D2"
        Load frmParamCountries
        frmParamCountries.Show 1, SysManMain
    Case "D3"
        .txtTASK.Text = "D3"
        Load frmParamCities
        frmParamCities.Show 1, SysManMain
    Case "D4"
        .txtTASK.Text = "D4"
        Load frmParamCurrencies
        frmParamCurrencies.Show 1, SysManMain
    Case "D5"
        .txtTASK.Text = "D5"
        Load frmParamTitles
        frmParamTitles.Show 1, SysManMain
    Case "D6"
        .txtTASK.Text = "D6"
        Load frmParamPayMethods
        frmParamPayMethods.Show 1, SysManMain
    Case "D7"
        .txtTASK.Text = "D7"
        Load frmParamFeeServices
        frmParamFeeServices.Show 1, SysManMain
    Case "D8"
        .txtTASK.Text = "D8"
        Load frmParamTaxes
        frmParamTaxes.Show 1, SysManMain
    Case "D9"
        .txtTASK.Text = "D9"
        Load frmParamAccPeriods
        frmParamAccPeriods.Show 1, SysManMain
    Case "D10"
        .txtTASK.Text = "D10"
        Load frmParamPayIntervals
        frmParamPayIntervals.Show 1, SysManMain
    Case "P0"
        .txtTASK.Text = "P0"
        Load frmSettings
        frmSettings.Show 1, SysManMain
    Case "P1"
        .txtTASK.Text = "P1"
        Load frmOfficeSettings
        frmOfficeSettings.Show 1, SysManMain
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
    Case Else
        Exit Sub
    End Select
    
End With
Exit Sub
Err:
    ErrorMessage
End Sub
