VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmPWMainMenu 
   BackColor       =   &H00FF0000&
   Caption         =   "PAYWELL PLUS PAYROLL SYSTEM"
   ClientHeight    =   5640
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8880
   Icon            =   "frmPWMainMenu.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "PAYWELL PLUS HELP"
      FontBold        =   -1  'True
      HelpCommand     =   11
      HelpFile        =   "C:\Picasso\DEVELOPMENT\Help\PAYWELL.HLP"
   End
   Begin MSComctlLib.StatusBar stbr 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5265
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Computer Name:"
            TextSave        =   "Computer Name:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Log on Date:"
            TextSave        =   "Log on Date:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "log on Time:"
            TextSave        =   "log on Time:"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Company Master"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "PAY Warning"
            ImageIndex      =   44
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Loan Processing"
            ImageIndex      =   36
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Employees Benefits"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Employees Deductions"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Close Period"
            ImageIndex      =   38
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Overtime Processing"
            ImageIndex      =   39
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Casual Processing"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Fringe Processing"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "frmcoinage"
            ImageIndex      =   40
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cash Payment"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Access the User Guide/Help Files"
            ImageIndex      =   18
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.TextBox txtCurrentUsers 
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   120
         Width           =   2295
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   45
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0368
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":03C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0424
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0482
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":04E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":053E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":059C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":05FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0658
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":06B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0714
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0772
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":07D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":082E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":088C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":08EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0948
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":09A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0A04
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0A62
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0AC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0B1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0B7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0C38
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0C96
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0D52
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0DB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0ECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0F28
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0F86
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":0FE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":1042
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":10A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":10FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":115C
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":11BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":1218
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":1276
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":12D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPWMainMenu.frx":1332
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileAdmin 
         Caption         =   "Administration"
         Begin VB.Menu mnusetcompany 
            Caption         =   "Set Default Company"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnusetcompanydash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileAdminUsers 
            Caption         =   "&Users Registration"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu10 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileAdmingroups 
            Caption         =   "User &Groups Edit"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu900 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileAdminSpecial 
            Caption         =   "User &Rights Setup"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu30 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileAdminManage 
            Caption         =   "User Management"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuMAN 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileAdminAccPeriod 
            Caption         =   "Payroll &Periods Setup"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu40 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileAdminActived 
            Caption         =   "Real-Time Users &Monitor"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu50 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileAdminBackup 
            Caption         =   "Backup / Restore Data"
            Checked         =   -1  'True
            Visible         =   0   'False
         End
         Begin VB.Menu mnudashbackup 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuUsersList 
            Caption         =   "Users List"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnu60 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePassword 
         Caption         =   "Change &Password "
         Shortcut        =   ^P
      End
      Begin VB.Menu mnu70 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileOffice 
         Caption         =   "M&S Office Settings"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu80 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close Active Screen"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu90 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCloseAll 
         Caption         =   "Close All &Screens"
      End
      Begin VB.Menu mnu100 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLogOut 
         Caption         =   "Log &Out"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnu110 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit System"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuSystemParameters 
      Caption         =   "&System Parameters"
      Begin VB.Menu mnuCountry 
         Caption         =   "Countries Master"
         Begin VB.Menu mnuCountriesMaster 
            Caption         =   "Countries Master"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashcountries 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCities 
            Caption         =   "Cities Master"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnudashcountry 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBankCodes 
         Caption         =   "Bank Master"
         Begin VB.Menu mnuBankCodesMaster 
            Caption         =   "Bank Master"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashbankcodes 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBranchMaster 
            Caption         =   "Branch Master"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnudash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompanyDetails 
         Caption         =   "Company Master"
         Begin VB.Menu mnuDetails 
            Caption         =   "Details Master"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashdetails 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDepartments 
            Caption         =   "Departments Master"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudash2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTitlesDesignations 
            Caption         =   "TitlesMaster"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashtitles 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDesignationsMatser 
            Caption         =   "Designations Master"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashgrades 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGradesMaster 
            Caption         =   "Grades Master"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashdesignations 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBranchesstations 
            Caption         =   "Branches Master"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashstations 
            Caption         =   "-"
         End
         Begin VB.Menu mnuStationsMaster 
            Caption         =   "Stations Master"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSepPayslipParameters 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPayslipParameters1 
            Caption         =   "Payslip Parameters"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashperiods 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCompanyPeriods 
            Caption         =   "Payroll Periods Master"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnudashovertimerates 
         Caption         =   "-"
      End
      Begin VB.Menu mnuotherparameter 
         Caption         =   "Tax Tables"
         Begin VB.Menu mnuPAYEtables 
            Caption         =   "PAYE Tables"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashnhif 
            Caption         =   "-"
         End
         Begin VB.Menu mnunhiftables 
            Caption         =   "NHIF tables"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashpayetables 
            Caption         =   "-"
         End
         Begin VB.Menu mnuNSSFtables 
            Caption         =   "NSSF tables"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashNSSF 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnuseparatorcar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMotorCars 
         Caption         =   "Motor Rates Setup"
         Begin VB.Menu mnuSaloonMaster 
            Caption         =   "Saloon Rates"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashsaloon 
            Caption         =   "-"
         End
         Begin VB.Menu mnupickuprates 
            Caption         =   "Pick-up Rates"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashPickups 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLandRoverRates 
            Caption         =   "Land-Rover Rates"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuMasterEmployee 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPayslipParameters 
         Caption         =   "Benefit and Deductions Master "
      End
      Begin VB.Menu mnudash7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmployeesMaster1 
         Caption         =   "Employees Master"
         Begin VB.Menu mnuEmployeesRecords 
            Caption         =   "Employees' Records Master"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashdependants 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDpendants 
            Caption         =   "Employee Dependants master"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashstatus 
            Caption         =   "-"
         End
         Begin VB.Menu mnuStatusDefaultsMaster 
            Caption         =   "Employees' Status Defauts Master"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashdefaults 
            Caption         =   "-"
         End
         Begin VB.Menu mnuStstusReasons 
            Caption         =   "Employees' Status Reasons Master"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashleft 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEmployeeStatusMaster 
            Caption         =   "Employees' Status Master"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuRecurrencies 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecurrentEarnings 
            Caption         =   "Employees' Benefits  Master"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashassigncars 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCarManagement 
            Caption         =   "Car Management Master"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashrecords 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecurrentdeductions 
            Caption         =   "Employees'  Deductions Master"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuotherparameters 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGovtInterest 
         Caption         =   "Government Rates Master"
         Begin VB.Menu mnuPension1 
            Caption         =   "Pension Master"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuPension 
            Caption         =   "-"
         End
         Begin VB.Menu mnuInterest 
            Caption         =   "Housing Rates  Master"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnudash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDenomination 
         Caption         =   "Denomination Setup"
      End
      Begin VB.Menu mnudashcoinagesetup 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPayrollcotrlos 
         Caption         =   "Payroll Control Master"
         Begin VB.Menu mnupayrollcotrol11 
            Caption         =   "Payroll Control1"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnupayrollcotrol1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPayrollControl2 
            Caption         =   "Payroll Control2"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuPaymentCycle 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPPaymentCycle 
            Caption         =   "Payment Cycle"
         End
         Begin VB.Menu mnuPaymentMethodSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPaymentMethod 
            Caption         =   "Payment Method"
         End
      End
      Begin VB.Menu mnudashpayslip 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTaxAuthoitygile 
         Caption         =   "Miscellaneous"
         Begin VB.Menu mnuLowInteresttables 
            Caption         =   "Low Interest tables"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashlow 
            Caption         =   "-"
         End
         Begin VB.Menu mnuKRA 
            Caption         =   "Authority File"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuNHIH 
            Caption         =   "-"
         End
         Begin VB.Menu mnuNHIFDefaults 
            Caption         =   "NHIFDefaults"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuDefaults 
            Caption         =   "-"
         End
         Begin VB.Menu mnuNssfDefaults 
            Caption         =   "NSSF Defaults"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu mnuProcessing 
      Caption         =   "&Payroll Processing"
      Begin VB.Menu mnuEmployeesdeductionsmaintain 
         Caption         =   "Maintain Employees Benefits"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnudashmaintaindeductions 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMaintainEmployeeDeductions 
         Caption         =   "Maintain Employee Deductions"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnubenefitsmainataindash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLowInterest 
         Caption         =   "Fringe Benefit Processing"
      End
      Begin VB.Menu mnudashovertime 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoanManagement 
         Caption         =   "Loan Management"
         Begin VB.Menu LoanIssuanceHeader 
            Caption         =   "-"
         End
         Begin VB.Menu mnuIssueLoan 
            Caption         =   "Loan Issuance"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuLoanIssuanceSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLoanRepayment 
            Caption         =   "Loan Repayment"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnudashsite 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPayrollProcessing 
         Caption         =   "Payroll Processing"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnudassh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCoinage 
         Caption         =   "Coinage Analysis"
      End
      Begin VB.Menu mnudashpaye 
         Caption         =   "-"
      End
      Begin VB.Menu mnumonthlyPaymentTransactions 
         Caption         =   "Monthly Payment Transactions"
         Begin VB.Menu mnuCashTransactions 
            Caption         =   "Cash Transactions"
            Checked         =   -1  'True
            Shortcut        =   {F9}
         End
         Begin VB.Menu mnudashcheckpayments 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCheckPayments 
            Caption         =   "Check Payments"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashCasualpaymentstransactions 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPAYE 
            Caption         =   "PAYE Transactions"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashpayepayments 
            Caption         =   "-"
         End
         Begin VB.Menu mnuNHIFTransactions 
            Caption         =   "NHIF Transactions"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashnhiftrasactions 
            Caption         =   "-"
         End
         Begin VB.Menu mnuNSSFTransactions 
            Caption         =   "NSSF Transactions"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuOtherPayments 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMOtherPayment 
            Caption         =   "Other Payments"
         End
      End
      Begin VB.Menu mnuYtddash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClosePeriod 
         Caption         =   "Close Payroll Period"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnudash15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuYearProcessing 
         Caption         =   "Year End Processing"
         Begin VB.Menu mnuCloseYear 
            Caption         =   "Close Year"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudashpandbtaxdeductioncard 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPb 
            Caption         =   "View Tax Deduction Card"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu mnuSystemUtilities 
      Caption         =   "System &Utilities"
      Begin VB.Menu mnuactivate 
         Caption         =   "Activate/Deactivate Benefits/Deductions"
      End
      Begin VB.Menu mnudashactivate 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalaryAdjustment 
         Caption         =   "Salary Adjustment"
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnusystemReportsSettings 
         Caption         =   "System Reports Settings"
      End
      Begin VB.Menu mnusystemReportsSettingsdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadReports 
         Caption         =   "System  Reports Viewer"
      End
   End
   Begin VB.Menu mnuNewports 
      Caption         =   "&New Reports"
      Begin VB.Menu mnu8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCoapmyReport 
         Caption         =   "Company Report"
      End
      Begin VB.Menu mnu7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCoyReportSum 
         Caption         =   "Company Report Summary"
      End
      Begin VB.Menu mnu6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNHIFReport 
         Caption         =   "NHIF Report"
      End
      Begin VB.Menu mnu5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNssfReport 
         Caption         =   "NSSF Report"
      End
      Begin VB.Menu mnu4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBenefitsReport 
         Caption         =   "Benefits Report"
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeductionsReport 
         Caption         =   "Deductions Report"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCoinageReport 
         Caption         =   "Coinage Report"
      End
      Begin VB.Menu mni13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCoinageSummary 
         Caption         =   "Coinage Summary"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPaymentRegister 
         Caption         =   "Payment Register"
      End
      Begin VB.Menu mnu12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMonthlyPayment 
         Caption         =   "Monthly Payments"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPayrollReport 
         Caption         =   "Payroll Report"
      End
      Begin VB.Menu mnu9 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuTileVertically 
         Caption         =   "Tile Vertically"
      End
      Begin VB.Menu dash42 
         Caption         =   "-"
      End
      Begin VB.Menu mnu 
         Caption         =   "Tile Horizontally"
      End
      Begin VB.Menu dash43 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "Cascade"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About Paywell Plus"
      End
      Begin VB.Menu mnuHelpPaywellPlus 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPaywellGuide 
         Caption         =   "Paywell Plus Guide"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmPWMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Activate()
        Call UpdateActivates
End Sub

Private Sub MDIForm_Load()
On Error GoTo err

Call OpenPayroll

Dim com As String, rstcom As ADODB.Recordset
        
        com = "select * from ParamCompanydetails;"
        
        Set rstcom = New Recordset
        rstcom.Open com, cnPAY, adOpenKeyset, adLockOptimistic

With rstcom
    If .EOF And .BOF Then GoTo loopcom
    
    If .RecordCount > 1 Then
            MsgBox "You have more than one Company in your database", vbCritical + vbOKOnly, "Company Master"
    End
    End If
End With

loopcom:

    With frmMain
        .WindowState = 2
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error GoTo err

    Dim msg As String
    msg = MsgBox("Are you sure you want to quit Paywell Plus", vbQuestion + vbYesNo, "Quit Paywell Plus")
    If msg = vbYes Then
    Call UpdateLogout
    
    Cancel = False
   
    End
    Else
    Cancel = True
    End If
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnoOvertime_Click()
On Error GoTo err
    frmOvertime.Show
    'frmTransDailyOvertime.Show
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnu_Click()
Me.Arrange (vbtilehorizontally)
End Sub

Private Sub mnuAccountingPeriodsSetup_Click()
On Error GoTo err
frmAccountingPeriods.Show
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuAbout_Click()
On Error GoTo err
    frmAbout.Show
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuAccountingTypes_Click()
On Error GoTo err
    frmperiod.Show
Exit Sub
err:
ErrorMessage

End Sub

Private Sub mnuactivate_Click()
On Error GoTo err
    frmActive.Show
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuBankCodesMaster_Click()
On Error GoTo err
    frmBankDetails.SSTab1.Tab = 0
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuBankRegister_Click()
On Error GoTo err
    Dim pay As String
    Dim rstpay As ADODB.Recordset

pay = "select TransPaye.*, ParamemployeesRecords.* from TransPaye,ParamemployeesRecords where TransPaye.employeepayno=ParamemployeesRecords.employeepayno;"
Set rstpay = New Recordset
rstpay.Open pay, cnPAY, adOpenKeyset, adLockOptimistic
Set rptDepartments.DataSource = rstpay
rptDepartments.Show
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuBankSummary_Click()
On Error GoTo err
Dim pay As String
Dim rstpay As ADODB.Recordset
pay = "select TransPaye.*, ParamemployeesRecords.* from TransPaye,ParamemployeesRecords where TransPaye.employeepayno=ParamemployeesRecords.employeepayno and ParamemployeesRecords.PayMethod='" & "B" & "';"
Set rstpay = New Recordset
rstpay.Open pay, cnPAY, adOpenKeyset, adLockOptimistic
Set rptBankPaymentsRegister.DataSource = rstpay
rptBankPaymentsRegister.Show
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuBenefits_Click()
'Dim emp As String
'Dim rstemp As ADODB.Recordset
'emp = "Select * from ParamEmployeesRecords;"
'Set rstemp = New Recordset
'rstemp.Open emp, cnPay, adOpenKeyset, adLockOptimistic
'
'Dim ben As String
'Dim rstben As ADODB.Recordset
'ben = "SELECT AssignBenefits.EmployeePayNo,AssignBenefits.ParameterCode,ParamPayslip1.Description,ParamPayslip1.ParameterCode AS Expr1,AssignBenefits.RecurrentAmount From AssignBenefits, ParamPayslip1 Where AssignBenefits.ParameterCode = ParamPayslip1.ParameterCode GROUP BY AssignBenefits.EmployeePayNo,AssignBenefits.ParameterCode,AssignBenefits.RecurrentAmount,ParamPayslip1.ParameterCode,ParamPayslip1.Description;"
'
'Set rstben = New Recordset
'rstben.Open ben, cnPay, adOpenKeyset, adLockOptimistic
'Set rptMiracle.DataSource = rstben
'rptMiracle.Show
On Error GoTo err
rptBenefits.Show
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuBenefitsReport_Click()
On Error GoTo err
        frmPWRCoyBenefits.Show
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuBranchesstations_Click()
On Error GoTo err
    frmBranchesAndStations.SSTab1.Tab = 0
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuBranchMaster_Click()
On Error GoTo err
frmBankDetails.SSTab1.Tab = 1
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuCarManagement_Click()
On Error GoTo err
    frmTransCar.Show
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuCascade_Click()
Me.Arrange (vbCascade)
End Sub

Private Sub mnuCashRegister_Click()
On Error GoTo err
Dim pay As String
Dim rstpay As ADODB.Recordset
pay = "select TransPaye.*, ParamemployeesRecords.* from TransPaye,ParamemployeesRecords where TransPaye.employeepayno=ParamemployeesRecords.employeepayno and ParamemployeesRecords.PayMethod='" & "C" & "';"
Set rstpay = New Recordset
rstpay.Open pay, cnPAY, adOpenKeyset, adLockOptimistic
Set rptCashPaymentsRegister.DataSource = rstpay
rptCashPaymentsRegister.Show
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuCashTransactions_Click()
On Error GoTo err
        frmCashPayment.Show
Exit Sub

err:
        ErrorMessage
End Sub

Private Sub mnuCasualProcessing_Click()
On Error GoTo err
            frmCasual.Show
Exit Sub

err:
        ErrorMessage
End Sub

Private Sub mnucasualtransactions_Click()
On Error GoTo err
        frmCasualPayment.Show
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuChabgePassword_Click()
On Error GoTo err
    Load frmChangePassword
    frmChangePassword.Show
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuCheckPayments_Click()
On Error GoTo err
    frmCheckPayment.Show
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuCities_Click()
On Error GoTo err
    frmCountriesCities.SSTab1.Tab = 1
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuClose_Click()

On Error Resume Next
Dim MyForm As Form
Set MyForm = ActiveForm
    Unload MyForm

End Sub

Private Sub mnuCloseAllScreens_Click()
On Error Resume Next
Dim MyForm As Form
Do Until ActiveForm Is Nothing
Set MyForm = ActiveForm
    Unload MyForm
Loop

End Sub

Private Sub mnuClosePeriod_Click()

  frmSecuClosingPeriod.Show
'frmClosePeriod.Show
End Sub

Private Sub mnuCloseYear_Click()
On Error GoTo err
frmSecuClosingYear.Show
'    frmCloseYear.Show
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuCoapmyReport_Click()
On Error GoTo err
        frmPWRCoyPayrollNew.Show
Exit Sub

err:
    ErrorMessage

End Sub

Private Sub mnuCoinage_Click()
On Error GoTo err
        frmCoinage.Show
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuCoinageAnalysis_Click()
'open the coin table and set the report to pick the data from it
On Error GoTo err
Dim coin As String
Dim rstcoin As ADODB.Recordset
coin = "select coin.*, TotalDenomination.*  from coin, TotalDenomination;"
Set rstcoin = New Recordset
rstcoin.Open coin, cnPAY, adOpenKeyset, adLockOptimistic
Dim toto As Currency
    With rstcoin
         Set rptCoinage.DataSource = rstcoin
    End With

    rptCoinage.Show

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuCompanyPayslip_Click()
On Error GoTo err
    DataEnvironment1.rscmdCompany_Grouping.Open
    rptCompany.Show
    DataEnvironment1.rscmdCompany_Grouping.Requery
    
    DataEnvironment1.rscmdCompany_Grouping.Close

Exit Sub

err:
    If err.Number = 3705 Then Resume Next
        ErrorMessage
End Sub

Private Sub mnuCoinageReport_Click()
On Error GoTo err
        frmPWRCoinage.Show
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuCoinageSummary_Click()
On Error GoTo err
    frmPWRCoyCoinageSum.Show
Exit Sub

err:
    ErrorMessage

End Sub

Private Sub mnuCompanyPeriods_Click()
On Error GoTo err
    frmAccountingPeriods.Show
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuDeductionPerEmployee_Click()
On Error GoTo err
        rptDed.Show
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnudeductions_Click()
On Error GoTo err
'open the paye tranaction table
    Dim pay As String
    Dim rstpay As ADODB.Recordset

    pay = "select TransPaye.*, ParamemployeesRecords.*,Transcasual.* from TransPaye,ParamemployeesRecords,Transcasual where TransPaye.employeepayno=ParamemployeesRecords.employeepayno and Transcasual.EmployeePayNo=Transcasualpayment.employeepayno;"
    
    Set rstpay = New Recordset
    rstpay.Open pay, cnPAY, adOpenKeyset, adLockOptimistic
    
    Set rptDeductions.DataSource = rstpay
        rptDeductions.Show
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuCountriesMaster_Click()
On Error GoTo err
    frmCountriesCities.SSTab1.Tab = 0
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuCoyReportSum_Click()
On Error GoTo err
        frmPWRCoyPayrollSum.Show
Exit Sub

err:
    ErrorMessage


End Sub

Private Sub mnuDeductionsReport_Click()
On Error GoTo err
        frmPWRCoyDeductions.Show
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuDenomination_Click()
On Error GoTo err
        frmPAYDCodes.Show
Exit Sub

err:
        ErrorMessage
End Sub

Private Sub mnuDepartments_Click()
        frmDepartments.Show
End Sub

Private Sub mnuDesignationsMatser_Click()
On Error GoTo err
        frmTitlesDesignations.SSTab1.Tab = 1
Exit Sub

err:
        ErrorMessage
End Sub

Private Sub mnuDetails_Click()
On Error GoTo err
    frmCompany.Show
Exit Sub

err:
    ErrorMessage
End Sub



Private Sub mnuEmployeesdeductionsmaintain_Click()
On Error GoTo err
        'frmMaintainEmployeesBenefitsDeductions.Show
        frmBenefits.Show
Exit Sub

err:
        ErrorMessage
End Sub

Private Sub mnuEmployeesNSSF_Click()
On Error GoTo err
    DataEnvironment1.rscmdNSSF_Grouping.Open

    rptEmployeeNSSF.Show

    DataEnvironment1.rscmdNSSF_Grouping.Requery
    DataEnvironment1.rscmdNSSF_Grouping.Close

Exit Sub

err:
    If err.Number = 3705 Then Resume Next
            ErrorMessage
End Sub

Private Sub mnuEmployeesRecords_Click()
        frmEmployees.Show
End Sub

Private Sub mnuEmployeeStatusMaster_Click()
On Error GoTo err
    
    frmStatus.SSTab1.Tab = 0

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuEmployerCertificate_Click()
On Error GoTo err
DataEnvironment1.rscmdPayslip_Grouping.Open
rptp10.Show
DataEnvironment1.rscmdPayslip_Grouping.Requery
DataEnvironment1.rscmdPayslip_Grouping.Close
Exit Sub
err:
If err.Number = 3705 Then Resume Next
ErrorMessage
End Sub

Private Sub mnuExit_Click()
On Error Resume Next
If MsgBox("Are you sure you want to shut down the system?", vbYesNo + vbQuestion + vbDefaultButton2, "Shut Down") = vbYes Then
   Call UpdateLogout
    End
Else
    Exit Sub
End If

End Sub

Private Sub mnuFileAdminAccPeriod_Click()
On Error GoTo err
    Load frmAccountingPeriods
    frmAccountingPeriods.Show
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuFileAdminActived_Click()
On Error GoTo err
    Dim rsMON As ADODB.Recordset
    Set rsMON = New ADODB.Recordset
    Dim DT As Date
    DT = Date
    
    rsMON.Open "SELECT SysUserLog.*,SysUserRegister.Username,SysUserRegister.Surname,SysUserRegister.Firstname FROM SysUserLog,SysUSerRegister WHERE SysUserRegister.UserName=SysUserLog.UserName AND SysUserLog.LoginDate LIKE '%" & Trim(DT) & "%' AND SysUserLog.LogoutDate IS NULL AND SysUserLog.LogoutTime IS NULL;", cnPAY, adOpenKeyset, adLockOptimistic
    
    Load frmSystemActiveUsers
    With frmSystemActiveUsers
        .Show
        If rsMON.EOF And rsMON.BOF Then Exit Sub
        Set .DataGrid1.DataSource = rsMON
    End With
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuFileAdmingroups_Click()
On Error GoTo err
    Load frmUsersParameters
    frmUsersParameters.Show
    frmUsersParameters.SSTab1.Tab = 1
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuFileAdminManage_Click()
On Error GoTo err
    Load frmUsersParameters
    frmUsersParameters.Show
    frmUsersParameters.SSTab1.Tab = 2
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuFileAdminSpecial_Click()
On Error GoTo err
    Load frmSpecialRights
    frmSpecialRights.Show
    frmSpecialRights.SSTab1.Tab = 1
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuFileAdminUsers_Click()
On Error GoTo err
    Load frmUsersParameters
    frmUsersParameters.Show
    frmUsersParameters.SSTab1.Tab = 0
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuFileClose_Click()
On Error Resume Next
Dim MyForm As Form
Set MyForm = ActiveForm
    Unload MyForm
End Sub

Private Sub mnuFileCloseAll_Click()
On Error Resume Next
Dim MyForm As Form
Do Until ActiveForm Is Nothing
Set MyForm = ActiveForm
    Unload MyForm
Loop

End Sub

Private Sub mnuFileExit_Click()
On Error Resume Next
If MsgBox("Are you sure you want to shut down the system?", vbYesNo + vbQuestion + vbDefaultButton2, "Shut Down") = vbYes Then
   Call UpdateLogout
    End
Else
    Exit Sub
End If
End Sub

Private Sub mnuFileLogOut_Click()
On Error Resume Next
    If MsgBox("Are you sure you want to end your session?", vbQuestion + vbYesNo, "Log Out") = vbYes Then
        Call UpdateLogout
        Call Unloadform
            Me.mnuFileAdmin.Enabled = True
            Me.mnuReports.Enabled = True
            Me.Hide
        Load frmLogin
        frmLogin.Show
    Else
        Exit Sub
    End If
End Sub

Private Sub mnuFilePassword_Click()
On Error GoTo err
    Load frmChangePassword
    frmChangePassword.Show
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuGradesMaster_Click()
On Error GoTo err
    frmTitlesDesignations.SSTab1.Tab = 2
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuHoliday_Click()
On Error GoTo err
frmOvertimeRates.SSTab1.Tab = 1
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuInterest_Click()
frmBenefitsInterests.SSTab1.Tab = 1
End Sub

Private Sub mnuIssueLoan_Click()
On Error GoTo err
        frmLoanManagement.Show
Exit Sub
err:
ErrorMessage

End Sub

Private Sub mnuKRA_Click()
On Error GoTo err
        frmFringe.SSTab1.Tab = 1
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuLascReport_Click()
On Error GoTo err
Dim pay As String
Dim rstpay As ADODB.Recordset
pay = "select TransPaye.*, ParamemployeesRecords.* from TransPaye,ParamemployeesRecords where TransPaye.employeepayno=ParamemployeesRecords.employeepayno and ParamemployeesRecords.PayMethod='" & "C" & "';"
Set rstpay = New Recordset
rstpay.Open pay, cnPAY, adOpenKeyset, adLockOptimistic
Set rptLASC.DataSource = rstpay
rptLASC.Show
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuLandRoverRates_Click()
On Error GoTo err
    frmCar.SSTab1.Tab = 2
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuLASCTables_Click()
On Error GoTo err
frmNHIFPAYELASCTables.SSTab1.Tab = 2
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuLoadReports_Click()
On Error GoTo err
    Dim rsREPORT As ADODB.Recordset, ReportPath As String
    Set rsREPORT = New ADODB.Recordset
    rsREPORT.Open "SELECT * FROM ParamReportPath;", cnPAY, adOpenKeyset, adLockOptimistic
    
    With rsREPORT
    If .RecordCount = 0 Then
        If MsgBox("No Settings for the Report Printing programs are available!" & vbCrLf & "Choose OK to Set it up!", vbOKCancel + vbInformation, "Report Settings") = vbCancel Then Exit Sub
        Load frmDefaultProgram
        frmDefaultProgram.Show
        frmDefaultProgram.drvPath.SetFocus
    Else
        ReportPath = !ReportPath
        Dim RetVal
        RetVal = Shell(ReportPath, vbNormalFocus)
    End If
    End With
    Exit Sub
err:
If err.Number = 94 Then
    If MsgBox("No Settings for the Report Printing programs are available!" & vbCrLf & "Choose OK to Set it up!", vbOKCancel + vbInformation, "Report Settings") = vbCancel Then Exit Sub
    Load frmDefaultProgram
    frmDefaultProgram.Show
    frmDefaultProgram.drvPath.SetFocus
Else
    ErrorMessage
End If

End Sub



Private Sub mnuMonthlyDeductions_Click()
        frmRecurrentDeductions.Show
End Sub

Private Sub mnuLoanPaymentHistory_Click()
On Error GoTo err
        rptLoanPaymentHistory.Show
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnulogout_Click()
On Error Resume Next
    If MsgBox("Are you sure you want to end your session?", vbQuestion + vbYesNo, "Log Out") = vbYes Then
        Call UpdateLogout
        Call Unloadform
            Me.mnuFile.Enabled = True
            Me.mnuReports.Enabled = True
            Me.Hide
        Load frmLogin
        frmLogin.Show
    Else
        Exit Sub
    End If
End Sub

Private Sub mnuLoanRepayment_Click()
On Error GoTo err
        frmLoanPayment.Show
Exit Sub

err:
    ErrorMessage

End Sub

Private Sub mnuLowInterest_Click()
On Error GoTo err
        frmLowInterest.Show
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuLowInteresttables_Click()
On Error GoTo err
        frmFringe.SSTab1.Tab = 0
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuMaintainEmployeeDeductions_Click()
On Error GoTo err
    'frmMaintainEmployeesDeductions.Show
    frmDeductions.Show
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuMinutesEquivalent_Click()
On Error GoTo err
    frmOvertimeRates.SSTab1.Tab = 2
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuMortgageManagement_Click()
frmMortgage.Show
End Sub

Private Sub mnuNetPayments_Click()
'open the employee table
On Error GoTo err
'open the paye tranaction table
Dim pay As String
Dim rstpay As ADODB.Recordset
pay = "select TransPaye.*, ParamemployeesRecords.* from TransPaye,ParamemployeesRecords where TransPaye.employeepayno=ParamemployeesRecords.employeepayno ;"
Set rstpay = New Recordset
rstpay.Open pay, cnPAY, adOpenKeyset, adLockOptimistic
Set rptNetPayments.DataSource = rstpay
rptNetPayments.Show
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuNhif_Click()
On Error GoTo err
Dim pay As String
Dim rstpay As ADODB.Recordset
pay = "select TransPaye.*, ParamemployeesRecords.* from TransPaye,ParamemployeesRecords where TransPaye.employeepayno=ParamemployeesRecords.employeepayno and ParamemployeesRecords.PayMethod='" & "C" & "';"
Set rstpay = New Recordset
rstpay.Open pay, cnPAY, adOpenKeyset, adLockOptimistic
Set rptnhif.DataSource = rstpay
rptnhif.Show
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuMonthlyPayment_Click()
On Error GoTo err
        frmPWRPayments.Show
Exit Sub

err:
        ErrorMessage
End Sub

Private Sub mnuMOtherPayment_Click()
On Error GoTo err
        frmPWPayments.Show
Exit Sub

err:
        ErrorMessage

End Sub

Private Sub mnuNHIFDefaults_Click()
On Error GoTo err
        frmPWPNhifDefaults.Show
Exit Sub

err:
        ErrorMessage

End Sub

Private Sub mnuNHIFReport_Click()
On Error GoTo err
        frmPWRCoyNhif.Show
Exit Sub

err:
        ErrorMessage
End Sub

Private Sub mnunhiftables_Click()
On Error GoTo err
        frmNHIFPAYELASCTables.SSTab1.Tab = 1
Exit Sub

err:
        ErrorMessage
End Sub

Private Sub mnuNHIFTransactions_Click()
On Error GoTo err
        frmNHIFTransactions.Show
Exit Sub

err:
        ErrorMessage
End Sub

Private Sub mnuNssf_Click()
On Error GoTo err
'Dim pay As String
'Dim rstpay As ADODB.Recordset
'pay = "select TransPaye.*, ParamemployeesRecords.*, ctrlpayrollcontrol1.* from TransPaye,ParamemployeesRecords,ctrlpayrollcontrol1 where TransPaye.employeepayno=ParamemployeesRecords.employeepayno and Transpaye.AccountPeriod=ctrlpayrollcontrol1.CurrentPeriod;"
'Set rstpay = New Recordset
'rstpay.Open pay, cnPay, adOpenKeyset, adLockOptimistic
'Set rptNSSF.DataSource = rstpay
'rptNSSF.Show

DataEnvironment1.rscmdNSSF_Grouping.Open
'DataEnvironment1.rscmdNSSF_Grouping.Fields("AccountPeriod")
rptNSSF.Show

DataEnvironment1.rscmdNSSF_Grouping.Requery
DataEnvironment1.rscmdNSSF_Grouping.Close

Exit Sub
err:
If err.Number = 3705 Then Resume Next
ErrorMessage
End Sub

Private Sub mnuNssfDefaults_Click()
On Error GoTo err
        frmPWPNssfDefaults.Show
Exit Sub
err:
ErrorMessage

End Sub

Private Sub mnuNssfReport_Click()
On Error GoTo err
        frmPWRCoyNssf.Show
Exit Sub

err:
    ErrorMessage


End Sub

Private Sub mnuNSSFtables_Click()
On Error GoTo err
frmNHIFPAYELASCTables.SSTab1.Tab = 3
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuNSSFTransactions_Click()
On Error GoTo err
        frmNSSFNHIFTransactions.SSTab1.Tab = 1
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuOutstandingLoans_Click()
On Error GoTo err

'Open the Loans file and the ParamEmployeesRecords and save the data
Dim EMP As String
Dim rstEMP As ADODB.Recordset
EMP = "select ParamEmployeesRecords.*, Transloan.* from ParamEmployeesRecords, Transloan where ParamEmployeesRecords.employeepayno=TransLoan.employeepayno"
Set rstEMP = New Recordset
rstEMP.Open EMP, cnPAY, adOpenKeyset, adLockOptimistic

Set rptOutstandingLoans.DataSource = rstEMP
rptOutstandingLoans.Show
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuOvertimeRatesP_Click()
On Error GoTo err
frmOvertimeRates.SSTab1.Tab = o
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuPayCertificate_Click()
On Error GoTo err

DataEnvironment1.rscmdPayslip_Grouping.Open
rptP10A.Show
DataEnvironment1.rscmdPayslip_Grouping.Requery
DataEnvironment1.rscmdPayslip_Grouping.Close

'Dim pay As String, rstpay As ADODB.Recordset
'pay = "select sum(GrossAmount) as total,ParamEmployeesRecords.*,ParamCompanyDetails.*, from Transpayecopy,ParamEmployeesRecords,ParamCompanyDetails where transpayecopy.employeepayno=ParamEmployeesRecords.Employeepayno;"
'Set rstpay = New Recordset
'rstpay.Open pay, cnPay, adOpenKeyset, adLockOptimistic
'Set rptp10asample.DataSource = rstpay
    
Exit Sub
err:
If err.Number = 3705 Then Resume Next
ErrorMessage
End Sub

Private Sub mnuPAYE_Click()
On Error GoTo err
        frmPAYE.Show
Exit Sub

err:
        ErrorMessage
End Sub

Private Sub mnuPAYECollections_Click()
On Error GoTo err
rptP12A.Show
Exit Sub
err:
If err.Number = 3705 Then Resume Next
ErrorMessage
End Sub

Private Sub mnuPAYEtables_Click()
On Error GoTo err
        frmNHIFPAYELASCTables.SSTab1.Tab = 0
Exit Sub

err:
        ErrorMessage
End Sub

Private Sub mnuPaymentMethod_Click()
On Error GoTo err:
    frmPaymentMethod.Show

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuPaymentRegister_Click()
On Error GoTo err
        frmPWRPaymentRegister.Show
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuPayrollControl2_Click()
On Error GoTo err
    frmPayrollControls.SSTab1.Tab = 1
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnupayrollcotrol11_Click()
        frmPayrollControls.SSTab1.Tab = 0
End Sub

Private Sub mnuPayrollDetai_Click()
On Error GoTo err
'open the paye tranaction table
Dim pay As String
Dim rstpay As ADODB.Recordset
pay = "select TransPaye.*, ParamemployeesRecords.* from TransPaye,ParamemployeesRecords where TransPaye.employeepayno=ParamemployeesRecords.employeepayno ;"
Set rstpay = New Recordset
rstpay.Open pay, cnPAY, adOpenKeyset, adLockOptimistic
Set rptPayrollListing.DataSource = rstpay
rptPayrollListing.Show
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuPayrollPeripdssetup_Click()
On Error GoTo err
    frmAccountingPeriods.Show
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuPayrollProcessing_Click()
On Error GoTo err
        frmWarning.Show
Exit Sub

err:
        ErrorMessage
End Sub

Private Sub mnuPaysliDepartmentwise_Click()
On Error GoTo err
        frmPayslipDept.Show
Exit Sub

err:
        ErrorMessage
End Sub

Private Sub mnupayslip_Click()
On Error GoTo err

DataEnvironment1.cnPayslip.Open

DataEnvironment1.rscmdTest_Grouping.Open

rptPAYECASH.Show

DataEnvironment1.rscmdTest_Grouping.Requery
DataEnvironment1.rscmdTest_Grouping.Close

Exit Sub
err:
If err.Number = 3705 Then Resume Next
ErrorMessage

End Sub

Private Sub mnuPayrollReport_Click()
On Error GoTo err
        frmPWRCoyPayslip.Show
Exit Sub

err:
        ErrorMessage
End Sub

Private Sub mnupayslipparameters_Click()
On Error GoTo err
        frmPayslipParameters.Show
Exit Sub

err:
        ErrorMessage
End Sub

Private Sub mnuPayslipScreens_Click()
On Error GoTo err
        frmpayslipscreen.Show
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuPayslipSummary_Click()
On Error GoTo err
frmLoadBranch.Show
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuPayslipSummaryReport_Click()
On Error GoTo err
'open the paye tranaction table
Dim pay As String
Dim rstpay As ADODB.Recordset
pay = "select TransPaye.*, ParamemployeesRecords.* from TransPaye,ParamemployeesRecords where TransPaye.employeepayno=ParamemployeesRecords.employeepayno ;"
Set rstpay = New Recordset
rstpay.Open pay, cnPAY, adOpenKeyset, adLockOptimistic
Set rptPayslipSummary.DataSource = rstpay
rptPayslipSummary.Show
Exit Sub
err:
ErrorMessage
End Sub



Private Sub mnuPayslipSummaryStationsWise_Click()
On Error GoTo err
frmLoadStation.Show
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuPayslipParameters1_Click()
On Error GoTo err
    frmPWPPayslipFormat.Show

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuPaywellGuide_Click()
On Error GoTo err
    Me.CommonDialog1.ShowHelp
Exit Sub

err:
        ErrorMessage
End Sub

Private Sub mnuPb_Click()
On Error GoTo err
frmTaxDeduction.Show
Exit Sub
err:
ErrorMessage
End Sub
Private Sub mnuPension1_Click()
On Error GoTo err
    frmBenefitsInterests.SSTab1.Tab = 0
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuPensionCertificate_Click()
On Error GoTo err
DataEnvironment1.rscmdPension_Grouping.Open

rptEmployeePension.Show

DataEnvironment1.rscmdPension_Grouping.Requery
DataEnvironment1.rscmdPension_Grouping.Close

Exit Sub
err:
If err.Number = 3705 Then Resume Next
ErrorMessage
End Sub

Private Sub mnuPensionReport_Click()
'On Error GoTo err
'Dim com As String, rstcom As ADODB.Recordset
'com = "select * from ctrlpayrollcontrol1;"
'Set rstcom = New Recordset
'rstcom.Open com, cnPay, adOpenKeyset, adLockOptimistic
'With rptPension
'    With .Sections(4)
'        .Controls("lblPeriod").Caption = rstcom!currentperiod
'   End With
'End With

Dim pay As String
Dim rstpay As ADODB.Recordset
pay = "select TransPaye.*, ParamemployeesRecords.*, ctrlpayrollcontrol1.* from TransPaye,ParamemployeesRecords,ctrlpayrollcontrol1 where TransPaye.employeepayno=ParamemployeesRecords.employeepayno and Transpaye.AccountPeriod=ctrlpayrollcontrol1.CurrentPeriod;"
Set rstpay = New Recordset
rstpay.Open pay, cnPAY, adOpenKeyset, adLockOptimistic
Set rptPension.DataSource = rstpay
rptPension.Show
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuPreviouspayslips_Click()
On Error GoTo err
frmPreviousPayslips.Show
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnupickuprates_Click()
On Error GoTo err
    frmCar.SSTab1.Tab = 1
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuRealTime_Click()
On Error GoTo err
    Dim rsMON As ADODB.Recordset
    Set rsMON = New ADODB.Recordset
    Dim DT As Date
    DT = Date
    
    rsMON.Open "SELECT SysUserLog.*,SysUserRegister.Username,SysUserRegister.Surname,SysUserRegister.Firstname FROM SysUserLog,SysUSerRegister WHERE SysUserRegister.UserName=SysUserLog.UserName AND SysUserLog.LoginDate LIKE '%" & Trim(DT) & "%' AND SysUserLog.LogoutDate IS NULL AND SysUserLog.LogoutTime IS NULL;", cnPAY, adOpenKeyset, adLockOptimistic
    
    Load frmSystemActiveUsers
    With frmSystemActiveUsers
        .Show
        If rsMON.EOF And rsMON.BOF Then Exit Sub
        Set .DataGrid1.DataSource = rsMON
    End With
    Exit Sub
err:
    ErrorMessage

End Sub

Private Sub mnuPPaymentCycle_Click()
On Error GoTo err:
    frmPaymentCycle.Show

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuRecurrentdeductions_Click()
On Error GoTo err
    frmRecurrentDeductions.Show
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuRecurrentEarnings_Click()
        frmAssignBenefits.Show
End Sub

Private Sub mnuSalaryAdjustment_Click()
On Error GoTo err
    frmSalaryadjustment.Show
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuSalaryHistory_Click()
DataEnvironment1.rscmdSalaryAdvance_Grouping.Open
rptSalaryHistory.Show
DataEnvironment1.rscmdSalaryAdvance_Grouping.Requery

DataEnvironment1.rscmdSalaryAdvance_Grouping.Close

Exit Sub
err:

If err.Number = 3705 Then Resume Next
    ErrorMessage
End Sub

Private Sub mnuSaloonMaster_Click()
On Error GoTo err
    frmCar.SSTab1.Tab = 0
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnusetcompany_Click()
On Error GoTo err
    frmSetCompany.Show
Exit Sub


err:
        ErrorMessage
End Sub




'Private Sub mnuSitemanagement_Click()
'On Error GoTo err
'    frmSite.Show
'Exit Sub
'err:
'errormessage
'End Sub

'Private Sub mnuStationsMaster_Click()
'On Error GoTo err
'        frmBranchesAndStations.SSTab1.Tab = 1
'Exit Sub
'err:
'errormessage
'End Sub

Private Sub mnuStatusDefaultsMaster_Click()
On Error GoTo err
frmStatus.SSTab1.Tab = 2
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuStstusReasons_Click()
On Error GoTo err
frmStatus.SSTab1.Tab = 1
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuTaxDeduction_Click()
On Error GoTo err
DataEnvironment1.rscmdPayslip_Grouping.Open
rptP9.Show
DataEnvironment1.rscmdPayslip_Grouping.Requery
DataEnvironment1.rscmdPayslip_Grouping.Close
Exit Sub
err:
If err.Number = 3705 Then Resume Next
ErrorMessage
End Sub

Private Sub mnusystemReportsSettings_Click()
On Error GoTo err
    frmDefaultProgram.Show
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuTileVertically_Click()
Me.Arrange (vbtilevertically)
End Sub

Private Sub mnuTitlesDesignations_Click()
On Error GoTo err
frmTitlesDesignations.SSTab1.Tab = 0
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuusergroup_Click()
On Error GoTo err
    frmUsersParameters.SSTab1.Tab = 1
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuUserManagement_Click()
On Error GoTo err
    frmUsersParameters.SSTab1.Tab = 2
    Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuUserRightsSetup_Click()
On Error GoTo err
    frmSpecialRights.SSTab1.Tab = 1
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuUsersList_Click()
On Error GoTo Myerr
Dim myusers As ADODB.Recordset, users As String
Set myusers = New Recordset
users = "SELECT *  FROM SysUserRegister;"
myusers.Open users, cnPAY, adOpenKeyset, adLockOptimistic
With rptusers
    Set .DataSource = myusers
    
End With
rptusers.Show
Exit Sub

Myerr:
    ErrorMessage
End Sub

Private Sub mnuUsersregister_Click()
On Error GoTo err
    fromuserparameters.SSTab1.Tab = 0
Exit Sub
err:
ErrorMessage
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Index
    Case 1
        Load frmCompany
        frmCompany.Show
    Case 2
        
        frmWarning.Show
    Case 3
        Exit Sub
    Case 4
       
        frmLoan.Show
    Case 5
        
        frmMaintainEmployeesBenefitsDeductions.Show
    Case 6
        Exit Sub
    Case 7
    
        frmMaintainEmployeesDeductions.Show
    Case 8
        frmClosePeriod.Show
    
    Case 9
        Exit Sub
    Case 10
        frmTransDailyOvertime.Show
    Case 11
        frmCasual.Show
    Case 12
        Exit Sub
    Case 13
        frmLowInterest.Show
    Case 14
        frmCoinage.Show
    Case 15
        Exit Sub
    Case 16
        frmCashPayment.Show
    Case 17
        CommonDialog1.ShowHelp
    Case Else
        Exit Sub
    End Select
    Exit Sub
err:
    ErrorMessage
End Sub
