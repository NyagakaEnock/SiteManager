VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SchedulingMain 
   AutoRedraw      =   -1  'True
   Caption         =   "SITE MONITORING CONSOLE"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11280
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SchedulingMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "SchedulingMain.frx":0442
   ScaleHeight     =   8160
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   1560
      Top             =   2760
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   7785
      Visible         =   0   'False
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2520
      Left            =   6120
      TabIndex        =   3
      Top             =   3960
      Visible         =   0   'False
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   4445
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   83558402
      CurrentDate     =   38401
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   7440
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7200
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   63
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":5B80
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":5FD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":616C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":821E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":176B9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":2117CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":212019
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":28A8DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":28AD2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":28B17F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":28B5D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":28BA23
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":28BE75
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":28C2CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":28C721
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":28CB73
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":28CFC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":28D417
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":28D869
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":28DCBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":28E10D
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":28E55F
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":28E879
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":28EFCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":28F645
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":2908F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":290C0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":291061
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":2914B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":291905
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":30E817
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":3BC3E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":3F197F
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":485D51
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":4861A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":486726
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":486B88
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":487820
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":48799B
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":487DED
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":48823F
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":48831A
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":48AEB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":48B511
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":48B963
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":48BDB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":48C207
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":48C521
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":48CB9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":48D215
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":48D88F
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":48DF09
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":48E583
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":48EBFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":48F277
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":48F8F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":48FF6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":4905E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":490C5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":4912D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":491953
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":491D80
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":49209A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7335
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   12938
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1080
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":492413
            Key             =   "C"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":492867
            Key             =   "O"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchedulingMain.frx":492CB9
            Key             =   "F"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   12938
      _Version        =   393217
      Indentation     =   531
      LineStyle       =   1
      Style           =   3
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImageList2"
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
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   7470
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   5470
            MinWidth        =   5470
            Picture         =   "SchedulingMain.frx":49310B
            Text            =   "OutDoor Systems"
            TextSave        =   "OutDoor Systems"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7057
            MinWidth        =   7057
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7057
            MinWidth        =   7057
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   7
            Alignment       =   2
            Enabled         =   0   'False
            Object.Width           =   7057
            MinWidth        =   7057
            TextSave        =   "KANA"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   1164
      ButtonWidth     =   4471
      ButtonHeight    =   1005
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Drugs' Dispenser"
            Key             =   "DD"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aquire SITE"
            Key             =   "AS"
            ImageIndex      =   42
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Voucher Preparation"
            Key             =   "VP"
            ImageIndex      =   44
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "MKN"
                  Text            =   "Make New Purchase Orders"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RDL"
                  Text            =   "Receive Shipments/Deliveries"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "PROPERTIES"
            Key             =   "PR"
            ImageIndex      =   35
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   9
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AB"
                  Text            =   "All Billboards"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "v"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BN"
                  Text            =   "Search Billboard by Number"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BT"
                  Text            =   "Search Billboard by Town"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BP"
                  Text            =   "Search Billboard by Plot name"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BL"
                  Text            =   "Search Billboard by Landlord name"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Payments Confirmation"
            Key             =   "PC"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "LEASE"
            Key             =   "L"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "NOTICE"
            Key             =   "N"
            ImageIndex      =   47
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "PRINT"
            Key             =   "PT"
            ImageIndex      =   56
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "C"
                  Text            =   "Contract"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CS"
                  Text            =   "Contract Schedule"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RS"
                  Text            =   "Rates Schedule"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ReS"
                  Text            =   "Rent Schedule"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Contract Termination"
            Key             =   "CT"
            ImageIndex      =   30
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "DR"
                  Text            =   "For Sites Expiring within a Date Range"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SD"
                  Text            =   "For Sites Expiring as at a Single Date"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.TextBox txtBB 
         Height          =   315
         Left            =   11280
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComDlg.CommonDialog HelpCommonDialog 
         Left            =   9240
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         HelpFile        =   "ODAS Help Facility.hlp"
      End
      Begin VB.TextBox txtJobBriefItemNo 
         Height          =   315
         Left            =   10920
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.TextBox txtSite 
         Height          =   315
         Left            =   10680
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtTask 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12840
         TabIndex        =   7
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileClearScreen 
         Caption         =   "Clear The &Screen"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuClearScreenSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpdateFlag 
         Caption         =   "Update Flag"
      End
      Begin VB.Menu mnujjhg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileEndSession 
         Caption         =   "&Log Out/End Session"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnu20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit / Quit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&System Settings"
      Begin VB.Menu mnuClientAgreementForm 
         Caption         =   "Client Contract Agreement Form"
      End
      Begin VB.Menu mnuKLMHJBGHBMJN 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdvertisingSites 
         Caption         =   "Sites Registration"
      End
      Begin VB.Menu mnuDFSD 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSiteLandlordsReg 
         Caption         =   "Landlords Registration"
      End
      Begin VB.Menu mnu78 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBillBoardDetails 
         Caption         =   "Bill Boards Date  &Schedule"
      End
      Begin VB.Menu mnullp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "A&bout the System..."
      End
   End
   Begin VB.Menu mnuActivities 
      Caption         =   "System &Activities"
      Begin VB.Menu mnusiteAcquisition 
         Caption         =   "Acquire Site"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuJKNHJG 
         Caption         =   "-"
      End
      Begin VB.Menu mnusiteproperties 
         Caption         =   "Site Properties"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuKLJGHVJBNHJG 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFreeSites 
         Caption         =   "Free Sites"
         Begin VB.Menu mnuBBSS 
            Caption         =   "BillBoards / StreetSigns"
         End
         Begin VB.Menu mnussss 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBBF 
            Caption         =   "Billboard sides / Faces"
         End
      End
      Begin VB.Menu mnuwew 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSitesReport 
         Caption         =   "Sites Report Based On Date"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuVoucherPreparation 
         Caption         =   "Voucher Preparation"
         Begin VB.Menu mnuPendingPayment 
            Caption         =   "Transactions Pending Payment"
         End
         Begin VB.Menu mnuVouchersPrepared 
            Caption         =   "Vouchers Prepared"
         End
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPaymentConfirmation 
         Caption         =   "Payment Confirmation"
         Begin VB.Menu mnuPaymentsPendingConfirmation 
            Caption         =   "Transactions Pending Confirmation"
         End
         Begin VB.Menu mnuPaymentsConfirmed 
            Caption         =   "Payments Confimed"
         End
      End
   End
   Begin VB.Menu mnuPayments 
      Caption         =   "&Leasing"
      Visible         =   0   'False
      Begin VB.Menu mnupreparelease 
         Caption         =   "Prepare lease"
      End
      Begin VB.Menu mnuJKHHJBKJKLJKLJKTYRDF 
         Caption         =   "-"
      End
      Begin VB.Menu mnusendnotice 
         Caption         =   "Send Notice"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports / Listings"
      Begin VB.Menu mnuAllFreeSites 
         Caption         =   "All Free Sites"
      End
      Begin VB.Menu mnulkiij 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNextRents 
         Caption         =   "Search Expiring Site"
         Begin VB.Menu mnuExpiringSitesAsAtASingleDate 
            Caption         =   "As at a Single Date"
         End
         Begin VB.Menu mnuExpiringSitesForADateRange 
            Caption         =   "Within a Date Range"
         End
         Begin VB.Menu MnuExpiredNotReneweds 
            Caption         =   "Expired Not Renewed"
         End
      End
      Begin VB.Menu mnuretretr 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSiteAlloacations 
         Caption         =   "Site Allocations"
      End
      Begin VB.Menu gdghfdhkjcf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLLs 
         Caption         =   "LandLords Listing"
      End
      Begin VB.Menu mnufdsfdsfds 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlots 
         Caption         =   "Site Details"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuwerewr 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAllSites 
         Caption         =   "ALL SITES (PLOTS)"
      End
      Begin VB.Menu gdgdvvc 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRoadReserve 
         Caption         =   "ALL SITES R.RESERVE"
      End
      Begin VB.Menu kkjuufhf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAllocations 
         Caption         =   "PLots  - Sites / Client Allocations"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnusdwqewq 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFreeBBs 
         Caption         =   "Querry Free BillBoards/ Faces"
      End
      Begin VB.Menu mnuqwewqe 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLandlordStatement 
         Caption         =   "Landlord Statement"
      End
      Begin VB.Menu mnuSepLansdlordStatement 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchToPay 
         Caption         =   "Outstanding Rent"
         Begin VB.Menu mnuOutstandingSingleDate 
            Caption         =   "As at a single Date"
         End
         Begin VB.Menu mnuOutstandingDateRange 
            Caption         =   "Within a Date Range"
         End
      End
      Begin VB.Menu mnuSepSitesPaid 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchSitesPaid 
         Caption         =   "Rent Paid"
      End
      Begin VB.Menu mnusads 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPaymentVouchers 
         Caption         =   "Payment Vouchers"
      End
      Begin VB.Menu mnuSep123 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSitesWithinCountyCouncils 
         Caption         =   "Sites Within Councils"
         Begin VB.Menu mnuAllSitesWithinCouncils 
            Caption         =   "All Sites"
         End
         Begin VB.Menu mnuFilterCouncil 
            Caption         =   "Filter Council"
         End
      End
      Begin VB.Menu mnuSepPaymentVouchers 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuSearchSiteByLRNo 
         Caption         =   "Search Site By LR No."
      End
      Begin VB.Menu mnuSepSearchSiteByLRNo 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchByLandLord 
         Caption         =   "Search By Landlord"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Hel&p"
      Begin VB.Menu mnuGene 
         Caption         =   "General Help Topics"
      End
      Begin VB.Menu mnuiiiii 
         Caption         =   "-"
      End
      Begin VB.Menu mnuspec 
         Caption         =   "Specific Help Guide"
      End
      Begin VB.Menu mnurtt 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIndex 
         Caption         =   "Help Index"
      End
      Begin VB.Menu mnujhjh 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "SchedulingMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyMainMenu As clsMainMenu
Dim rsOPERATION As clsODASOperation
Dim q, X
Private Sub Form_Activate()
On Error GoTo err
'If Not SystemActivated Then End
If bAlert = False Then
    MyLoginID = GetMyLoginID
    MyCurrentPeriod = TransactionPeriod
    MyCurrentDate = Format(Date, "MMMM dd,yyyy")
    'LoadReminders
    
End If
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub Form_Initialize()
On Error GoTo err
    Set MyMainMenu = New clsMainMenu
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub Form_Load()
On Error GoTo err

    Call OpenODBCConnection
    MyMainMenu.CreateMAINMENU
    MyMainMenu.GetMainStructure
    
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub Form_Resize()
On Error GoTo err
    With SchedulingMain
        .TreeView1.Height = .Height - (8880 - 7335)
        .ListView1.Height = .Height - (8880 - 7335)
        
        .ListView1.Width = .Width - (12000 - 8775)
    End With
Exit Sub
err:
If err.Number = 380 Then Resume Next
    ErrorMessage
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo err
    If MsgBox("Are you sure you want to quit this System!!", vbQuestion + vbYesNo + vbDefaultButton2 + vbMsgBoxHelpButton, "Quit System", "ALISELP.HLP", 246) = vbNo Then
        Cancel = 1: Exit Sub
    Else
        Cancel = 0
        Call UpdateLogoutRecord
        End
    End If
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
  Cancel = 1
End Sub

Private Sub ListView1_DblClick()
On Error GoTo err

    With SchedulingMain
            
            CurrentRecord = Trim(Me.ListView1.SelectedItem.Text)
            GlobalDepartmentCode = ""
            bQuotationApproval = False
            bQuotationAuthorization = False
            bSiteAPPROVAL = False
            bopenJOBBRIEF = False
            bcloseJobBrief = False
            bSiteAuthorization = False
            
            frmODASMOperation.txtApplicationNo.Text = CurrentRecord
            Set rsOPERATION = New clsODASOperation

            Select Case (.txtTask)
                Case "N2":
                        bsendnoticeAPPROVAL = True
                        rsOPERATION.checkAPPROVEDDISCHARGE
                        If bsendnoticeAPPROVAL = False Then Exit Sub
                Case "N3":
                        bsendnoticeAUTHORIZATION = True
                        rsOPERATION.checkAPPROVEDDISCHARGE
                        If bsendnoticeAUTHORIZATION = False Then Exit Sub
                Case "R2":
                        breceivenoticeAPPROVAL = True
                        rsOPERATION.checkAPPROVEDDISCHARGE
                        If breceivenoticeAPPROVAL = False Then Exit Sub
                Case "R3":
                        breceivenoticeAUTHORIZATION = True
                        rsOPERATION.checkAPPROVEDDISCHARGE
                        If breceivenoticeAUTHORIZATION = False Then Exit Sub
                Case "L1":
                        bopenJOBBRIEF = True
                        GlobalDepartmentCode = .ListView1.SelectedItem.SubItems(1)
                        rsOPERATION.checkAPPROVEDDISCHARGE
                        If bopenJOBBRIEF = False Then Exit Sub
                Case "C2":
                        bJobBriefApproval = True
                        rsOPERATION.checkAPPROVEDDISCHARGE
                        If bJobBriefApproval = False Then Exit Sub
                Case "C3":
                        bJobBriefAuthorization = True
                        rsOPERATION.checkAPPROVEDDISCHARGE
                        If bJobBriefAuthorization = False Then Exit Sub
                Case "K2"
                        frmODASRBBProperties.Show vbModal
                        Exit Sub
                Case "K5":
                        bSiteAPPROVAL = True
                        rsOPERATION.checkAPPROVEDDISCHARGE
                        If bSiteAPPROVAL = False Then Exit Sub
                Case "K51":
                        bSiteAPPROVAL = True
                        rsOPERATION.checkAPPROVEDDISCHARGE
                        If bSiteAPPROVAL = False Then Exit Sub
                Case "K52":
                        bSiteAuthorization = True
                        rsOPERATION.checkAPPROVEDDISCHARGE
                        If bSiteAuthorization = False Then Exit Sub
                Case "K6":
                        bSiteAuthorization = True
                        rsOPERATION.checkAPPROVEDDISCHARGE
                        If bSiteAuthorization = False Then Exit Sub
                Case "P2":
                        bQuotationApproval = True
                        rsOPERATION.checkAPPROVEDDISCHARGE
                        If bQuotationApproval = False Then Exit Sub
                Case "P3":
                        bQuotationAuthorization = True
                        rsOPERATION.checkAPPROVEDDISCHARGE
                        If bQuotationAuthorization = False Then Exit Sub
                Case "J2":
                        bPurchaseOrderAPPROVAL = True
                        rsOPERATION.checkAPPROVEDDISCHARGE
                        If bPurchaseOrderAPPROVAL = False Then Exit Sub
                Case "J3":
                        bPurchaseOrderAUTHORIZATION = True
                        rsOPERATION.checkAPPROVEDDISCHARGE
                        If bPurchaseOrderAUTHORIZATION = False Then Exit Sub
            End Select
            
            rsOPERATION.approveOPERATION
            Set rsOPERATION = Nothing

    End With

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub listView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        With SchedulingMain
                If Screen.ActiveForm.ListView1.ListItems.Count = 0 Or Screen.ActiveForm.ListView1.View <> lvwReport Then Exit Sub
                CurrentRecord = Empty
                Dim i, j, k
                j = Screen.ActiveForm.ListView1.ListItems.Count: k = 1
                
                If j = 0 Then Exit Sub
                
                For i = 1 To j
                    If Screen.ActiveForm.ListView1.ListItems(i).Text <> Item Then
                        Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                    End If
                Next i
                
                Select Case (.txtTask)
                Case "N1"
                        frmODASMPrepareNotice.txtContractNo.Text = Item.Text
                        frmODASMPrepareNotice.txtPlotNo.Text = Item.SubItems(1)
                        Load frmODASMPrepareNotice
                        frmODASMPrepareNotice.Show 1, Me
       
                Case "N4"
                        frmODASMSendNotice.txtContractNo.Text = Item.Text
                        frmODASMSendNotice.Toolbar1.Buttons(2).Caption = "Send Notice? ": .Toolbar1.Buttons(2).Image = 4
                        Load frmODASMSendNotice
                        frmODASMSendNotice.Show vbModal
                        CurrentRecord = Item.Text
                Case "N6"
                        frmODASMLeaseAgreement.txtContractNo.Text = Item.Text
                        Load frmODASMLeaseAgreement
                        frmODASMLeaseAgreement.Show 1, Me
                Case "N10"
                        frmODASMSendNotice.txtContractNo.Text = Item
                        Load frmODASRContractRenewalLetter
                        frmODASRContractRenewalLetter.Show vbModal
                Case "N12"
                        frmODASMAllocation.txtContractNo = Item.Text
                        frmODASMAllocation.txtPlotNo = Item.SubItems(1)
                        CurrentRecord = Item.Text
                        loadMoreDETAILS
                        bPlotRenewal = True
                        Load frmODASMAllocation
                        frmODASMAllocation.Show 1, Me
                Case "N13"
                        frmODASMLeaseRenewal.txtContractNo.Text = Item
                        Load frmODASRContract
                        frmODASRContract.Show vbModal
                Case "K91"
                        CurrentRecord = Item.Text
                        Load frmODASRContract1
                        frmODASRContract1.Show 1, Me
                Case "K92"
                        CurrentRecord = Item.Text
                        Load frmRptNACADAComp
                        frmRptNACADAComp.Show 1, Me
                Case "K93"
                        CurrentRecord = Item.Text
                        frmODASMContractEditing.txtContractNo.Text = Item
                        Load frmODASMContractEditing
                        frmODASMContractEditing.Show 1, Me
                 Case "K95"
                        CurrentRecord = Item.Text
                        Load frmRPTODASEditedContract
                        frmRPTODASEditedContract.Show 1, Me
                Case "K32"
                        CurrentRecord = Item.Text
                        Load frmODASRptSiteRegistration
                        frmODASRptSiteRegistration.Show 1, Me
                Case "K20"
                        frmODASMCouncilRates.txtTownCode.Text = Item
                        frmODASMCouncilRates.txtTown.Text = Item.SubItems(1)
                        Load frmODASMCouncilRates
                        frmODASMCouncilRates.Show 1, Me
                Case "K21"
                        CurrentRecord = Item
                        'INPQRY2 = InputBox("Please enter the year for which you want to view the installments", "Yearly Installments")
                        'If Len(INPQRY2) = 0 Or INPQRY2 = Empty Then
                           ' MsgBox "Either the system could not find matching records or the operation was cancelled...", vbCritical + vbOKOnly, "Missing Records"
                       ' Else
                            Load frmODASRRentPaymentInstallments
                            frmODASRRentPaymentInstallments.Show vbModal
                        'End If
                Case "K22"
                        frmODASMAllocation.txtContractNo.Text = Item
                        frmODASMAllocation.txtPlotNo.Text = Item.SubItems(1)
                        frmODASMAllocation.txtNames.Text = Item.SubItems(5)
                        Load frmODASMAllocation
                        frmODASMAllocation.Show vbModal
                 Case "X6"
                        frmODASPLoadBillBoardPhoto.txtTransactionNo.Text = Item
                        Load frmODASPLoadBillBoardPhoto
                        frmODASPLoadBillBoardPhoto.Show vbModal
                Case "K24"
                        CurrentRecord = Item
                        Load frmODASRagreementSchedule
                        frmODASRagreementSchedule.Show vbModal
                Case "K25"
                        CouncilForm = False
                        INPQRY2 = Trim(Item.SubItems(2))
                        CurrentRecord = Item
                        CurrentRecord1 = Item.SubItems(1)
                        Load frmODASRRatesSchedule
                        frmODASRRatesSchedule.Show vbModal
                        
                Case "K1": frmODASMSiteRegistration.txtTownCode = Item
                        Load frmODASMSiteRegistration
                        frmODASMSiteRegistration.Show 1, Me
                
                Case "K2": frmODASPAssignProperties.txtSiteNo = Item
                        frmODASPAssignProperties.txtMedia.Text = Item.SubItems(5)
                        Load frmODASPAssignProperties
                        frmODASPAssignProperties.Show 1, Me
                
                Case "K4": frmODASMAllocation.txtPlotNo = Item.Text
                        CurrentRecord = Item.Text
                        Load frmODASMAllocation
                        frmODASMAllocation.Show 1, Me
                Case "K7"
                        Load frmODASRAuthorizedLeases
                        frmODASRAuthorizedLeases.Show 1, Me
                Case "K150"
                        Load frmODASMFreeSite
                        frmODASMFreeSite.txtSiteNo.Text = Item.Text
                        frmODASMFreeSite.txtSiteDetails.Text = Item.SubItems(1)
                        frmODASMFreeSite.txtPlotNo.Text = Item.SubItems(2)
                        frmODASMFreeSite.txtPlotName.Text = Item.SubItems(3)
                        frmODASMFreeSite.Show vbModal
                Case "R11"
                        frmODASMNoticeAcknoledgement.txtJobBriefItemNo.Text = Item
                        frmODASMNoticeAcknoledgement.Toolbar1.Buttons(2).Caption = "Re&new?"
                        frmODASMNoticeAcknoledgement.txtMonths.Text = Item.SubItems(5)
                        frmODASMNoticeAcknoledgement.UpDown2.Max = Item.SubItems(5)
                        Load frmODASMNoticeAcknoledgement
                        frmODASMNoticeAcknoledgement.Show 1, Me
                Case "M8"
                        .txtJobBriefItemNo.Text = Item
                        .txtSite.Text = Item.SubItems(1)
                        .txtBB = Item.SubItems(6)
                        If .txtBB = "Y" Then
                            Load frmODASRMaitananceBSchedule
                            frmODASRMaitananceBSchedule.Show 1, Me
                        Else
                            Load frmODASRMaitananceSchedule
                            frmODASRMaitananceSchedule.Show 1, Me
                        End If
                Case "M6"
                        frmODASMSiteMaintanance.txtMaintananceNo.Text = Item
                        frmODASMSiteMaintanance.txtSiteNo.Text = Item.SubItems(1)
                        frmODASMSiteMaintanance.txtSiteDetails.Text = "Maintanance for " & Item.SubItems(3) & "IN " & Item.SubItems(6) & "For the SITE " & Item.SubItems(2)
                        Load frmODASMSiteMaintanance
                        frmODASMSiteMaintanance.Show 1, Me
                Case "Z2"
                        CurrentRecord = Item
                        Load frmODASMSiteSchedule
                        frmODASMSiteSchedule.Show vbModal
                Case "CR-1"
                         frmSitesReportGroupedByCouncils.strCouncilcode = Item.Text
                         Load frmSitesReportGroupedByCouncils
                         frmSitesReportGroupedByCouncils.Show 1, Me
                End Select
        End With
Exit Sub
err:
    ErrorMessage

End Sub
Private Sub ShowForm()
On Error GoTo err
    With Screen.ActiveForm.ListView1
            If Screen.ActiveForm.ListView1.ColumnHeaders(1).Text = "Quotation No" Then
                  If .Item.Checked = True Then
                   
                    If k = 0 Then k = 2
                    Dim DF: DF = CLng(k * 250)
                    
                    QuotationNumber = ListView1.SelectedItem
                   End If
                   
            ElseIf Screen.ActiveForm.ListView1.ColumnHeaders(1).Text = "Quotation Number" Then
                    
                    If k = 0 Then k = 2
                    Dim DS: DS = CLng(k * 250)
                    
            End If
                   
                    If k = 0 Then k = 2
                    Dim JB: JB = CLng(k * 250)
                    
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuDeleteSelected_Click()
End Sub
Private Sub mnuAllProductsSoldToday_Click()
On Error GoTo err
    With SchedulingMain
        Dim rsLIST As ADODB.Recordset
        Set rsLIST = New ADODB.Recordset
        Dim ThisDate As String
        ThisDate = Format(Date, "MMMM dd,yyyy")
    
            rsLIST.Open "SELECT * FROM PharmPointOfSale WHERE datecreated='" & Trim(ThisDate) & "' ORDER BY drugname;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
            If rsLIST.EOF And rsLIST.BOF Then
                MsgBox "No Sales Have Been Made Today!!", vbCritical + vbOKOnly, "Sales"
                Set rsFindRecord = Nothing: Exit Sub
            ElseIf rsLIST!DateCreated = 0 Then
                MsgBox "No Sales Have Been Made Today!!", vbCritical + vbOKOnly, "Sales"
                Set rsLIST = Nothing: Exit Sub
            Else
                Set rsLIST = Nothing
                
            End If
            
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuAdvertisingSites_Click()
Load frmODASMSiteRegistration
frmODASMSiteRegistration.Show vbModal

End Sub

Private Sub mnuAllFreeSites_Click()
Load frmRptODASAllFreeSites
frmRptODASAllFreeSites.Show vbModal
End Sub

Private Sub mnuAllocations_Click()
Load frmODASRPlotAllocations
frmODASRPlotAllocations.Show vbModal
End Sub

Private Sub mnuallsites_Click()
  Load frmRptODASAllSites
  frmRptODASAllSites.Show 1, Me
End Sub

Private Sub mnuAllSitesWithinCouncils_Click()
frmSitesReportGroupedByCouncils.strCouncilcode = ""
Load frmSitesReportGroupedByCouncils
frmSitesReportGroupedByCouncils.Show 1, Me
End Sub

Private Sub mnuBBF_Click()
    Me.txtTask = "Z2"
    bBillBoardFace = True
    CurrentRecord = InputBox("Please enter the billboard face number to search....", "Billboards search request")
    If Len(CurrentRecord) = 0 Then
        MsgBox ("Either the Operation was canceled or No billboard face Number was entered..."), vbCritical + vbOKOnly, "Missing Plot Number..."
    Else
        showALLSitesbyNumber
    End If
End Sub

Private Sub mnuBBSS_Click()
    Me.txtTask = "Z2"
    bBillBoard = True
    CurrentRecord = InputBox("Please enter the billboard number to search and assign properties...", "Billboards search request")
    If Len(CurrentRecord) = 0 Then
        MsgBox ("Either the Operation was canceled or No billboard Number was entered..."), vbCritical + vbOKOnly, "Missing Plot Number..."
    Else
        showALLSitesWithoutPropertiesbyNumber
    End If
End Sub

Private Sub mnuBillBoardDetails_Click()
   showALLBillBoardSchedule
End Sub

Private Sub mnuClientAgreementForm_Click()
    Load frmClientContractAgreement
    frmClientContractAgreement.Show 1, Screen.ActiveForm
End Sub

Private Sub mnuCreateJobCard_Click()
End Sub

Private Sub mnuCreateInvoice_Click()
On Error GoTo err
        Exit Sub
err:
   ErrorMessage
End Sub

Private Sub mnuCreateJobBrief_Click()
End Sub

Private Sub mnuClientRegistration_Click()
On Error GoTo err
    With SchedulingMain
    
    End With
    Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuExp_Click()
On Error GoTo err
    With Me
       Set rsFIND = New ADODB.Recordset
       rsFIND.Open "Select * From ODASPPlotMast;", cnCOMMON, adOpenKeyset, adLockOptimistic
       rsFIND.MoveFirst
       
       .ProgressBar1.Visible = True: .ProgressBar1.Min = 0: .ProgressBar1.Max = rsFIND.RecordCount
       Do While rsFIND.EOF <> True
            Set rsFindRecord = New ADODB.Recordset
            rsFindRecord.Open "Select * From ODASPPLotSite Where MastNo = '" & rsFIND!MastNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsFindRecord.RecordCount <> 0 Then
                 rsFindRecord.MoveFirst
                 
                 Do While rsFindRecord.EOF <> True
                     Set rsSAVE = New ADODB.Recordset
                     rsSAVE.Open "Select * From ODASMSiteSchedule Where SiteNo = '" & rsFindRecord!SiteNo & "' and ScheduleDate = '" & Format(rsFIND!expirydate, "MMMM dd,yyyy") & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                         If rsSAVE.RecordCount <> 0 Then
                             rsSAVE.Delete
                         End If
                     rsFindRecord.MoveNext
                 Loop
            End If
            rsFIND!expirydate = DateAdd("d", -1, rsFIND!expirydate)
            rsFIND.Update
            
        .ProgressBar1.Value = .ProgressBar1.Value + 1
        rsFIND.MoveNext
        Loop
        .ProgressBar1.Value = 0
        
            Set rsFindRecord1 = New ADODB.Recordset
                rsFindRecord1.Open "Select * From ODASMLeaseAgreement ;", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                rsFindRecord1.MoveFirst
                
                .ProgressBar1.Min = 0: .ProgressBar1.Max = rsFindRecord1.RecordCount
                Do While rsFindRecord1.EOF <> True
                    rsFindRecord1!expirydate = DateAdd("d", -1, rsFindRecord1!expirydate)
                    rsFindRecord1.Update
                .ProgressBar1.Value = .ProgressBar1.Value + 1
                rsFindRecord1.MoveNext
                Loop
                
                .ProgressBar1.Visible = False
    End With
Exit Sub
err:
ErrorMessage
End Sub

Private Sub MnuExpiredNotReneweds_Click()
    On Error GoTo err
    frmODASSitesToExpire.strReport = "ExpiredNotRenewed"
    Load frmODASSitesToExpire
    frmODASSitesToExpire.Show 1, Me
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuExpiringSitesAsAtASingleDate_Click()
On Error GoTo err
    frmODASSitesToExpire.strReport = "AsAtASingleDate"
    Load frmODASSitesToExpire
    frmODASSitesToExpire.Show 1, Me
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuExpiringSitesForADateRange_Click()
On Error GoTo err
    frmODASSitesToExpire.strReport = ""
    Load frmODASSitesToExpire
    frmODASSitesToExpire.Show 1, Me
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuFileClearScreen_Click()
On Error GoTo err
With Screen.ActiveForm
  .ListView1.ListItems.Clear
End With
Exit Sub
err:
   ErrorMessage
End Sub

Private Sub mnuFileEndSession_Click()
On Error GoTo err
If MsgBox("Are you sure you want to end your Current Session?", vbQuestion + vbYesNo, "End Session") = vbYes Then
    Call UpdateLogoutRecord
    SchedulingMain.Hide
        
    Load frmLogin
    frmLogin.Show 1
    
Else
    Exit Sub
End If
Exit Sub
err:
If err.Number = 400 Then Resume Next
    ErrorMessage
End Sub

Private Sub mnuFileExit_Click()
    Unload Screen.ActiveForm
End Sub

Private Sub mnuFileUtilBCP_Click()
On Error GoTo err

    AppPath = Trim(App.Path & "\DBUtility.exe")
    
    RetVal = Shell(AppPath, 1)
    
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuFilterCouncil_Click()
showALLCOUNCILS
Me.txtTask = "CR-1"
End Sub

Private Sub mnuFreeBBs_Click()
    CurrentRecord = InputBox("Please enter the town name in which to search to search free sites...", "Free sites search request")
    If Len(CurrentRecord) = 0 Then
        MsgBox ("Either the Operation was canceled or No Plot Number was entered..."), vbCritical + vbOKOnly, "Missing Plot Number..."
    Else
        showALLAvailableFaces
    End If
End Sub

Private Sub mnuFreeSoites_Click()

End Sub

Private Sub mnuGene_Click()
    With Me
    .HelpCommonDialog.DialogTitle = "Using the Main System"
    .HelpCommonDialog.HelpFile = App.HelpFile
    .HelpCommonDialog.HelpContext = 6
    .HelpCommonDialog.HelpCommand = cdlHelpContext
    .HelpCommonDialog.ShowHelp
    End With
End Sub

Private Sub mnuHelpAbout_Click()
    MyMainMenu.HelpAbout
End Sub

Private Sub mnuHelpContents_Click()
    MyMainMenu.HelpContents
End Sub

Private Sub mnuHelpIndex_Click()
    MyMainMenu.HelpIndex
End Sub

Private Sub mnuProjectCosting_Click()

End Sub

Private Sub mnuQuotationWriting_Click()
End Sub

Private Sub mnuLandlordStatement_Click()
    'strSEARCHSQL = "SELECT DISTINCT ME.EMployeeNo as [Employee No], ME.Surname + ' ' + ME.OtherNames as [NAME] FROM MEmployee ME INNER JOIN MTransaction MT ON ME.EMployeeNo=MT.EMployeeNo WHERE MT.CurrentPeriod>='" & strStartPeriod & "' AND MT.CurrentPeriod<='" & strEndPeriod & "'"
    strSEARCHSQL = "SELECT AccountNo as [LL No.],CompanyName as [Landlord] FROM ODASPAccount WHERE AccountType LIKE 'LLORD'"
    frmSearchRecord.load_form strSEARCHSQL, "SELECT LANDLORD"
    frmSearchRecord.Show vbModal, Me
    If CurrentRecord = "" Then
                    Exit Sub
    End If
    frmRStatement.strAccountNo = CurrentRecord
    Load frmRStatement
    frmRStatement.Show vbModal
End Sub

Private Sub mnuLLs_Click()
    Load frmODASRLandLords
    frmODASRLandLords.Show vbModal
End Sub

Private Sub mnuOutstandingDateRange_Click()
frmODASSearchSitesNotPaid.strReport = ""
Load frmODASSearchSitesNotPaid
frmODASSearchSitesNotPaid.Show vbModal
End Sub

Private Sub mnuOutstandingSingleDate_Click()
frmODASSearchSitesNotPaid.strReport = "PendingPaymentAsAtASingleDate"
Load frmODASSearchSitesNotPaid
frmODASSearchSitesNotPaid.Show vbModal
End Sub

Private Sub mnuPaymentsConfirmed_Click()
frmODASSearchSitesNotPaid.strReport = "PaymentsConfirmed"
Load frmODASSearchSitesNotPaid
frmODASSearchSitesNotPaid.Show vbModal
End Sub

Private Sub mnuPaymentsPendingConfirmation_Click()
frmODASSearchSitesNotPaid.strReport = "PendingConfirmation"
Load frmODASSearchSitesNotPaid
frmODASSearchSitesNotPaid.Show vbModal
End Sub

Private Sub mnuPaymentVouchers_Click()
Load frmUVouchersPrepared
frmUVouchersPrepared.Show 1, Me
End Sub

Private Sub mnuPendingPayment_Click()
frmODASSearchSitesNotPaid.strReport = "PendingPayment"
Load frmODASSearchSitesNotPaid
frmODASSearchSitesNotPaid.Show vbModal
End Sub

Private Sub mnuPlots_Click()
On Error GoTo err
    frmODASMSiteRegistration.txtPlotNo.Text = InputBox("Please enter the Plot Number (correctly)to search and display report", "Plot Details search Engine")
    If Len(frmODASMSiteRegistration.txtPlotNo.Text) = 0 Then
        MsgBox ("Either the Operation was canceled or No Plot Number was entered..."), vbCritical + vbOKOnly, "Missing Plot Number..."
    Else
        Load frmODASRPlotSites
        frmODASRPlotSites.Show vbModal
    End If
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnupreparelease_Click()
    setALLAcquiredSites
End Sub

Private Sub mnuReminders_Click()
                  AppPath = Trim(App.Path & "\messages.EXE")
                    RetVal = Shell(AppPath, 1)
End Sub

Private Sub mnuRentPaymentS_Click()
Load frmODASRRentPayments
frmODASRRentPayments.Show vbModal
End Sub

Private Sub mnuRoadReserve_Click()
       Load frmRptODASAllRoadSites
       frmRptODASAllRoadSites.Show 1, Me
End Sub

Private Sub mnuSearchByLandLord_Click()
On Error GoTo errMSG
CurrentRecord = InputBox("Enter LR No.. (Add % before or after the search text to locate other  matching values )")
If CurrentRecord = "" Then Exit Sub
strSQL = "SELECT     P.PlotNo, P.PlotName, P.LRNo, P.PhysicalLocation, P.TownCode, P.AnnualRent, P.CommencementDate, P.ExpiryDate, ODASPAccount.CompanyName FROM ODASPPlot AS P INNER JOIN ODASPAccount ON P.AccountNo = ODASPAccount.AccountNo WHERE CompanyName LIKE '" & CurrentRecord & "'"
FillList strSQL, Me.ListView1
Exit Sub
errMSG:
        ErrorMessage
End Sub

Private Sub mnuSearchSiteByLRNo_Click()
On Error GoTo errMSG
CurrentRecord = InputBox("Enter LR No.. (Add % before or after the search text to locate other  matching values )")
If CurrentRecord = "" Then Exit Sub
strSQL = "SELECT     P.PlotNo, P.PlotName, P.LRNo, P.PhysicalLocation, P.TownCode, P.AnnualRent, P.CommencementDate, P.ExpiryDate, ODASPAccount.CompanyName FROM ODASPPlot AS P INNER JOIN ODASPAccount ON P.AccountNo = ODASPAccount.AccountNo WHERE LRNo LIKE '" & CurrentRecord & "'"
FillList strSQL, Me.ListView1
Exit Sub
errMSG:
        ErrorMessage
End Sub

Private Sub mnuSearchSitesPaid_Click()
frmODASSearchSitesNotPaid.strReport = "PaymentsConfirmed"
Load frmODASSearchSitesNotPaid
frmODASSearchSitesNotPaid.Show vbModal
End Sub

Private Sub mnusendnotice_Click()
    getNOTICESAUTHORIZED
End Sub

Private Sub mnuSiteAlloacations_Click()
  Load frmODASRPlotAllocations
  frmODASRPlotAllocations.Show 1, Me
End Sub

Private Sub mnuSiteLandlordsReg_Click()
    Me.txtTask.Text = "K3"
    getLANDLORDTYPE
End Sub

Private Sub mnusiteproperties_Click()
    showALLSitesWithoutProperties
End Sub

Private Sub mnuSitesExpiring_Click()
  Load frmODASSitesToExpire
  frmODASSitesToExpire.Show vbModal
End Sub

Private Sub mnuSitesReport_Click()
Load frmUSelectDateRange
frmUSelectDateRange.Report = "SitesBasedOnDate"
frmUSelectDateRange.Show 1, Me
End Sub

Private Sub mnuspec_Click()
With Me
            If .txtTask.Text = "N2" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 14
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "X" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 24
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N3" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 13
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N1" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 23
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K5" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 14
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K6" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 13
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K20" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 12
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K25" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 11
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K21" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 10
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K9" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 9
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K24" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 8
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "R2" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 14
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "R3" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 13
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N4" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 26
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N5" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 25
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N12" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 29
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N13" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 30
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "S" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 24
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "G" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 40
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N7" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 24
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N8" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 24
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N9" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 24
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N10" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 24
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N11" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 24
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "R1" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 38
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "R11" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 36
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K1" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 17
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K3" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 18
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K4" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 15
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K22" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 16
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "Z2" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 39
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "M8" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 41
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            Else
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 6
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            End If
        End With
End Sub

Private Sub mnuUpdate_Click()
On Error GoTo err
    With Me
        Set rsFIND = New ADODB.Recordset
            rsFIND.Open "Select * From ODASPPlotMast order by ContractNo;", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                rsFIND.MoveFirst
                .ProgressBar1.Visible = True: .ProgressBar1.Min = 0: .ProgressBar1.Max = rsFIND.RecordCount
                Do While rsFIND.EOF <> True
                
                Set rsFindRecord = New ADODB.Recordset
                    rsFindRecord.Open "Select min(ODASMInstallment.PaymentDueDate)as DueDate,ODASMInstallment.ContractNo,ODASMInstallment.PaymentFlag, ODASMInstallment.PaymentDue From ODASMInstallment WHERE ODASMInstallment.PaymentFlag ='N' and ODASMInstallment.ContractNo = '" & rsFIND!ContractNo & "' Group BY ContractNo, PaymentDueDate, PaymentDue,PaymentFlag;", cnCOMMON, adOpenKeyset, adLockOptimistic
                        
                    If rsFindRecord.RecordCount <> 0 Then
                        rsFIND!RentDueDate = rsFindRecord!DueDate
                        rsFIND!RentDue = rsFindRecord!PaymentDue
                        rsFIND.Update
                    End If
                .ProgressBar1.Value = .ProgressBar1.Value + 1
                rsFIND.MoveNext
               Loop
            .ProgressBar1.Value = 0: .ProgressBar1.Visible = False
     End With
Exit Sub
err:
ErrorMessage
End Sub

Private Sub mnuUpdateFlag_Click()
      Load frmUpdatePaymentFlag
      frmUpdatePaymentFlag.Show 1, Me
End Sub

Private Sub mnuVouchersPrepared_Click()
Load frmODASSearchSitesNotPaid
frmODASSearchSitesNotPaid.strReport = "VouchersPrepared"
frmODASSearchSitesNotPaid.Show vbModal
End Sub

Private Sub Timer1_Timer()
With Me

    q = Timer1.Interval
    
    If Timer1.Interval = 60000 Then
        X = X + q
        If X >= (q * 5) Then
            X = 0
            Call Form_Load
        End If
    End If

End With
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
With SchedulingMain
    Select Case Button.Key
        Case "AS"
            showALLTOWNS
            .txtTask.Text = "K1"
        Case "LS"
        
        Case "L"
            setALLAcquiredSites
            .txtTask.Text = "K4"
        Case "N"
            getALLContracts
            .txtTask.Text = "N1"
        Case "PR"
            .txtTask.Text = "K2"
            showALLSitesWithoutProperties
        Case "PT"
            MsgBox "Hallo, what would you want to print? Click on the arrow pointing down, choose the approprate category and enter a search statement. Click OK or press enter once through. Thank you... ", vbInformation + vbOKOnly, "Instant Help"
        Case "VP"
            frmODASMVoucher.cboPaymentCode.Text = "RENT"
            Load frmODASMVoucher
            frmODASMVoucher.Show 1, Me
        Case "PC"
            frmODASMPaymentConfirmation.cboPaymentCode.Text = "RENT"
            Load frmODASMPaymentConfirmation
            frmODASMPaymentConfirmation.Show 1, Me
        Case "CT"
            Load frmODASMContractTermination
            frmODASMContractTermination.Show 1, Me
        Case "HP"
        
            If .txtTask.Text = "N2" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 14
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "X" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 24
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N3" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 13
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N1" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 23
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K5" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 14
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K6" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 13
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K20" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 12
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K25" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 11
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K21" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 10
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K9" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 9
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K24" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 8
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "R2" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 14
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "R3" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 13
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N4" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 26
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N5" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 25
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N12" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 29
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N13" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 30
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "S" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 24
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "G" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 40
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N7" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 24
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N8" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 24
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N9" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 24
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N10" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 24
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "N11" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 24
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "R1" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 38
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "R11" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 36
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K1" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 17
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K3" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 18
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K4" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 15
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "K22" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 16
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "Z2" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 39
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            ElseIf .txtTask.Text = "M8" Then
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 41
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            Else
                .HelpCommonDialog.DialogTitle = "Using the Main System"
                .HelpCommonDialog.HelpFile = App.HelpFile
                .HelpCommonDialog.HelpContext = 6
                .HelpCommonDialog.HelpCommand = cdlHelpContext
                .HelpCommonDialog.ShowHelp
            End If
        Case Else
            Exit Sub
    End Select
End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo err
Select Case ButtonMenu.Key
    Case "AB"
        Me.txtTask.Text = "K2"
        showALLSitesWithoutProperties
    Case "BN"
        Me.txtTask.Text = "K2"
        CurrentRecord = InputBox("Please enter the billboard number to search and assign properties...", "Billboards search request")
        If Len(CurrentRecord) = 0 Then
            MsgBox ("Either the Operation was canceled or No billboard Number was entered..."), vbCritical + vbOKOnly, "Missing Plot Number..."
        Else
            showALLSitesWithoutPropertiesbyNumber
        End If

    Case "BT"
        Me.txtTask.Text = "K2"
        CurrentRecord = InputBox("Please enter the town name in which to search billboard...", "Free sites search request")
        If Len(CurrentRecord) = 0 Then
            MsgBox ("Either the Operation was canceled or no town was entered..."), vbCritical + vbOKOnly, "Missing Plot Number..."
        Else
            showALLSitesWithoutPropertiesbyTown
        End If
    Case "BP"
        Me.txtTask.Text = "K2"
        CurrentRecord = InputBox("Please enter the plot name on which to search billboard...", "Free sites search request")
        If Len(CurrentRecord) = 0 Then
            MsgBox ("Either the Operation was canceled or no town was entered..."), vbCritical + vbOKOnly, "Missing Plot Number..."
        Else
            showALLSitesWithoutPropertiesbyPlotName
        End If
    Case "BL"
        Me.txtTask.Text = "K2"
        CurrentRecord = InputBox("Please enter the landlord's name of the plot to search...", "Free sites search request")
        If Len(CurrentRecord) = 0 Then
            MsgBox ("Either the Operation was canceled or no name was entered..."), vbCritical + vbOKOnly, "Missing Plot Number..."
        Else
            showALLSitesWithoutPropertiesbyLandlordName
        End If
    Case "C"
        Me.txtTask.Text = "K9"
        CurrentRecord = InputBox("Please enter the last digits of the contract number to search and print...", "Free sites search request")
        If Len(CurrentRecord) = 0 Then
            MsgBox ("Either the Operation was canceled or no number was entered..."), vbCritical + vbOKOnly, "Missing Plot Number..."
        Else
            getOneContract
        End If
    Case "CS"
        Me.txtTask.Text = "K24"
        CurrentRecord = InputBox("Please enter the last digits of the contract number to search and print schedule...", "Free sites search request")
        If Len(CurrentRecord) = 0 Then
            MsgBox ("Either the Operation was canceled or no number was entered..."), vbCritical + vbOKOnly, "Missing Plot Number..."
        Else
            getOneCURRENTLEASE
        End If
    Case "ReS"
        Me.txtTask.Text = "K21"
        'CurrentRecord = InputBox("Please enter the last digits of the contract number to search and print...", "Free sites search request")
        'If Len(CurrentRecord) = 0 Then
            'MsgBox ("Either the Operation was canceled or no number was entered..."), vbCritical + vbOKOnly, "Missing Plot Number..."
        'Else
         '   getOneContract
        'End If
    Case "RS"
        Me.txtTask.Text = "K25"
        CurrentRecord = InputBox("Please enter the Billboard number to search and print rates schedule...", "Free sites search request")
        If Len(CurrentRecord) = 0 Then
            MsgBox ("Either the Operation was canceled or no number was entered..."), vbCritical + vbOKOnly, "Missing Plot Number..."
        Else
            getOneALLsitesRatesPrepared
        End If
    Case "AB"
    Case "BN"
    Case "SD"
        'showALLPlotsToExpire
        Me.txtTask.Text = "LR"
        Load frmUSelectDateRange
        frmUSelectDateRange.Report = "ExpiringSitesAsAtASingleDate"
        frmUSelectDateRange.DTPickerStartDate.Enabled = False
        frmUSelectDateRange.Show 1, Me
      
    Case "DR"
         'showALLPlotsToExpire
          Me.txtTask.Text = "LR"
        Load frmUSelectDateRange
        frmUSelectDateRange.Report = "ExpiringSitesWithinADateRange"
        frmUSelectDateRange.DTPickerStartDate.Visible = True
        frmUSelectDateRange.Show 1, Me
        
       
    
    Case Else
        Exit Sub
    End Select
Exit Sub
err:
    ErrorMessage
End Sub


Private Sub showALLPlotsToExpire(Optional strReport As String = "")
                If strReport = "" Then
                        strSQL = "SELECT *  FROM ODASPPlot, ODASPAccount WHERE ODASPPLot.ExpiryDate >= '" & Format(frmODASSitesToExpire.txtStartDate.Text, "yyyy/mm/dd") & "' and ODASPPLot.ExpiryDate <= '" & Format(frmODASSitesToExpire.txtLastDate.Text, "yyyy/mm/dd") & "' and ODASPPlot.AccountNo = ODASPAccount.AccountNo ;"
                Else
                        strSQL = "SELECT *  FROM ODASPPlot, ODASPAccount WHERE  ODASPPLot.ExpiryDate <= '" & Format(frmODASSitesToExpire.txtLastDate.Text, "yyyy/mm/dd") & "' and ODASPPlot.AccountNo = ODASPAccount.AccountNo ;"
                End If

End Sub
Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
    SchedulingMain.ListView1.ListItems.Clear
    Node.Image = "C"
End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
    SchedulingMain.ListView1.ListItems.Clear
    Node.Image = "O"
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo err
    With SchedulingMain
            .txtTask.Text = Node.Key
            Select Case Node.Key
            
            Case "X"
                .txtTask.Text = "X"
            Case "S8"
                    .txtTask.Text = "S8"
                    showALLDEPARTMENTS
            Case "S13"
                    .txtTask.Text = "S13"
                    showALLTerminationReasons
            Case "S4"
                    .txtTask.Text = "S4"
                    showALLDEPARTMENTS
            Case "K1"
                    .txtTask.Text = "K1"
                    showALLTOWNS
            Case "K2"
                    .txtTask.Text = "K2"
                    showALLSitesWithoutProperties
            Case "K31"
                    .txtTask.Text = "K31"
                    Load frmODASPLandLord
                    frmODASPLandLord.Show 1, Me
            Case "K4"
                    .txtTask.Text = "K4"
                    setALLAcquiredSites
            Case "K41"
                   .txtTask.Text = "K41"
                   'AllowProcess
                   'If bAllowProcess = True Then
                    CurrentRecord = InputBox("Please Enter The Contract Number For Editing", "Contract Rent Inrement Editing")
                         If Len(CurrentRecord) = 0 Then
                               MsgBox ("Either the Operation was canceled or no name was entered..."), vbCritical + vbOKOnly, "Missing Invoice Number..."
                         Else
                               getContractToEdit
                         End If
                   
                   'End If
            Case "K5"
                    .txtTask.Text = "K5"
                    getALLAllocatedSites
            Case "K51"
                    .txtTask.Text = "K51"
                    getALLAllocatedSites
            Case "K52"
                    .txtTask.Text = "K52"
                    getALLApprovedSites
            Case "K6"
                    .txtTask.Text = "K6"
                    getALLApprovedSites
            Case "K13"
                    .txtTask.Text = "K13"
                    getLEASESDUEToEXPIRE
            Case "K14"
                    .txtTask.Text = "K14"
                    Load frmODASSearchSitesTerminated
                    frmODASSearchSitesTerminated.Show vbModal
                    SchedulingMain.ListView1.ListItems.Clear
             Case "K141"
                    .txtTask.Text = "K141"
                    SchedulingMain.ListView1.ListItems.Clear
                    Load frmODASSearchSitesTerminated
                    frmODASSearchSitesTerminated.Show vbModal
                    
            Case "K15"
                    .txtTask.Text = "K15"
                    getPLOTRentRateDues
            Case "K16"
                    .txtTask.Text = "K16"
                    Load frmODASSitesOnRoadReserve
                    frmODASSitesOnRoadReserve.Show vbModal
                    SchedulingMain.ListView1.ListItems.Clear
            Case "K150"
                    .txtTask.Text = "K150"
                    ListALLSitesToFree
            Case "K32"
                    .txtTask.Text = "K32"
                    SchedulingMain.ListView1.ListItems.Clear
                    Load frmODASSitesProperties
                    frmODASSitesProperties.Show vbModal
            Case "K34"
                    .txtTask.Text = "K34"
                    showALLSitesWithoutProperties
                    
            Case "K17"
                    .txtTask.Text = "K17"
                    showALLFreeSites
            Case "K18"
                    .txtTask.Text = "K18"
                    Load frmODASSearchSitesNotPaid
                    frmODASSearchSitesNotPaid.Show vbModal
                    SchedulingMain.ListView1.ListItems.Clear
             Case "K19"
                    .txtTask.Text = "K19"
                    Load frmODASSearchPaidSites
                    frmODASSearchPaidSites.Show vbModal
                    SchedulingMain.ListView1.ListItems.Clear
            Case "K11"
                    .txtTask.Text = "K11"
                    getSITESRateNotPaid
            Case "K12"
                    .txtTask.Text = "K12"
                    getSITESRateRentPaid
            Case "K20"
                    .txtTask.Text = "K20"
                     showALLCOUNCILS
            Case "K21"
                    .txtTask.Text = "K21"
                    getALLContracts
            Case "K22"
                    .txtTask.Text = "K22"
                    getCURRENTLEASES
            Case "K24"
                    .txtTask.Text = "K24"
                    getCURRENTLEASES
            Case "K25"
                    .txtTask.Text = "K25"
                    getALLsitesRatesPrepared
            Case "K23"
                    .txtTask.Text = "K23"
                    Load frmODASSearchSiteNewSites
                    frmODASSearchSiteNewSites.Show vbModal
                    SchedulingMain.ListView1.ListItems.Clear
            Case "K32"
                    .txtTask.Text = "K32"
                    showALLSitesWithoutProperties
             Case "K33"
                    .txtTask.Text = "K33"
                    Load frmODASSearchGeneral
                    frmODASSearchGeneral.Show vbModal
                    SchedulingMain.ListView1.ListItems.Clear
            Case "K8"
                    .txtTask.Text = "K8"
                    Load frmODASSitesToExpire
                    frmODASSitesToExpire.Show vbModal
                    SchedulingMain.ListView1.ListItems.Clear
            Case "K91"
                    .txtTask.Text = "K91"
                    getALLContracts
            Case "K92"
                    .txtTask.Text = "K92"
                    getALLNACADAContracts
            Case "K93"
                    .txtTask.Text = "K93"
                    getALLNACADAContracts
            Case "K94"
                    .txtTask.Text = "K94"
                    Load frmODASPClause
                    frmODASPClause.Show vbModal
                    SchedulingMain.ListView1.ListItems.Clear
            Case "K95"
                    .txtTask.Text = "K95"
                    getALLNACADAContracts
            Case "K32"
                    .txtTask.Text = "K32"
                    getALLContracts
            Case "K10"
                    .txtTask.Text = "K10"
                    getPLOTRentNotPaid
            Case "X5"
                    .txtTask.Text = "X5"
                    getALLSitesExpired
            
            Case "S2"
                    .txtTask.Text = "S2"
                    getALLOPERATIONS
            Case "N"
                    .txtTask.Text = ""
            Case "N1"
                    .txtTask.Text = "N1"
                    getALLContracts
            Case "N2"
                    .txtTask.Text = "N2"
                    NoticesPrepared
            Case "N3"
                    .txtTask.Text = "N3"
                    getNOTICESAPPROVED
            Case "N4"
                    .txtTask.Text = "N4"
                    getNOTICESAUTHORIZED
            Case "N5"
                    .txtTask.Text = "N5"
                    getNOTICESDISPATCHED
            Case "N6"
                    .txtTask.Text = "N6"
                    getCONTRACTSToTerminate
            Case "N7"
                    .txtTask.Text = "N7"
                    getNoticesPrepared
            Case "N8"
                    .txtTask.Text = "N8"
                    getNOTICESAPPROVED
            Case "N9"
                    .txtTask.Text = "N9"
                    getALLNoticesAuthorized
            Case "N10"
                    .txtTask.Text = "N10"
                    getAllNoticesSent
            Case "N11"
                    .txtTask.Text = "N11"
                    getNoticesReceived
            Case "N12"
                    .txtTask.Text = "N12"
                    getCONTRACTSToRenew
                    'setALLAcquiredSitesForRenewal
            Case "N13"
                     .txtTask.Text = "N13"
                    getCONTRACTSRenewed
            Case "R"
                    .txtTask.Text = ""
            Case "R1"
                    .txtTask.Text = "R1"
                    Load frmODASMNoticeAcknoledgement
                    frmODASMNoticeAcknoledgement.Show 1, Me
            Case "R2"
                    .txtTask.Text = "R2"
                    showNoticesReceived
            Case "R3"
                    .txtTask.Text = "R3"
                    showRenewalNoticesApproved
            Case "R11"
                    .txtTask.Text = "R11"
                    showNoticesAuthorized
            Case "P1"
                    .txtTask.Text = "P1"
                    showALLTOWNS
            Case "M1"
                    .txtTask.Text = "M1"
                    showMaintenancePROPERTIES
            Case "M2"
                    .txtTask.Text = "M2"
                    ShowAllWorksDueForMaintenance
            Case "M8"
                    .txtTask.Text = "M8"
                    showAllJobsCompleted
            Case "M3"
                    .txtTask.Text = "M3"
                    ShowAllWorksDueForMaintenanceONEMonth
            Case "M5"
                    .txtTask.Text = "M5"
                    ShowAllWorksDueForMaintenanceSpcificPeriod
            Case "M6"
                    .txtTask.Text = "M6"
                    ShowAllWorksDueForMaintenanceSpecificDate
            Case "G"
                .txtTask.Text = "G"
            Case "G1"
                 getALLContracts
                 .txtTask.Text = "K9"
            Case "G2"
                 Call ShowAllContractsExpiryOnSpecifiedDate
            Case "G3"
                 Call ShowAllCouncilFeeDueOnSpecificDate
            Case "G4"
                 Call ShowAllRentFeeDueOnSpecificDate
            Case "G5"
                 Call ShowAllValidEmptyBillBoards
            Case "G7"
                Call showALLLandlords
            Case "G9"
                .txtTask.Text = "G9"
                showALLSitesAllocated
'                 Call ShowAllSitesNotValid
            Case "G10"
               Call ShowSitesExpiryOnSpecificDate
            Case "G11"
                .txtTask.Text = "G11"
                showALLSitesReserved
            Case "G12"
                .txtTask.Text = "G12"
                showALLSitesToFree
            Case "G13"
                .txtTask.Text = "G13"
                showALLSitesUnAllocated
            Case "X2"
                .txtTask.Text = "X2"
                AllSitesOnRoadReserve
            Case "X3"
                .txtTask.Text = "X3"
                AllNonEagleStructures
            Case "X4"
                .txtTask.Text = "X4"
                AllPlotRents
            Case "X5"
                .txtTask.Text = "X5"
                RateSchedules
            Case "Z1"
                .txtTask.Text = "Z1"
                AllSiteSchedule
            Case "X6"
                .txtTask.Text = "X6"
                getALLApprovedMasts
            Case "Z2"
                .txtTask.Text = "Z2"
                bBillBoard = False: bBillBoardFace = False
                getALLsites
            Case "K101"
                        frmODASMVoucher.cboPaymentCode.Text = "RENT"
                        Load frmODASMVoucher
                        frmODASMVoucher.Show 1, Me
            Case "K102"
                        frmODASMPaymentConfirmation.cboPaymentCode.Text = "RENT"
                        Load frmODASMPaymentConfirmation
                        frmODASMPaymentConfirmation.Show 1, Me
            
            Case Else
                Exit Sub
            End Select
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ShowAllUnApprovedJobBriefs()
On Error GoTo err
    With Screen.ActiveForm
        .ListView1.ListItems.Clear
        .ListView1.ColumnHeaders.Clear
        
        .ListView1.ColumnHeaders.Add , , "Job Brief No", .ListView1.Width / 6.5
        .ListView1.ColumnHeaders.Add , , "Expected DOC", .ListView1.Width / 6.5
        .ListView1.ColumnHeaders.Add , , "Quotation No", .ListView1.Width / 6.5
        .ListView1.ColumnHeaders.Add , , "Brief Date", .ListView1.Width / 6.5
        .ListView1.ColumnHeaders.Add , , "Deadline Date", .ListView1.Width / 6.5
        .ListView1.ColumnHeaders.Add , , "ClientName", .ListView1.Width / 4.5
        .ListView1.ColumnHeaders.Add , , "Contact Person", .ListView1.Width / 4.5
        .ListView1.ColumnHeaders.Add , , "TelePhone", .ListView1.Width / 6.5
        .ListView1.ColumnHeaders.Add , , "Mobile No", .ListView1.Width / 6.5
        .ListView1.ColumnHeaders.Add , , "L.P.O No", .ListView1.Width / 6.5

        .ListView1.View = lvwReport
        
        Dim rsLIST As ADODB.Recordset
        Set rsLIST = New ADODB.Recordset
        
        rsLIST.Open "SELECT * FROM ODASMJobBrief, AdvertClients WHERE AdvertClients.CustomerID = ODASMJobBrief.CustomerNumber and (ODASMJobBrief.Approved IS NULL or ODASMJobBrief.Approved = 'N') ORDER BY ODASMJobBrief.JobBriefNo;", cnCOMMON, adOpenKeyset, adLockOptimistic
        
        Dim MyList As ListItem
        
        If rsLIST.EOF And rsLIST.BOF Then
            .ListView1.View = lvwList
            Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
            Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
        End If
        
        While Not rsLIST.EOF
        
        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!JobBriefNo))

        If Not IsNull(rsLIST!expectedDOC) Then
            MyList.SubItems(1) = CStr(rsLIST!expectedDOC)
        End If
         
        If Not IsNull(rsLIST!QuotationNumber) Then
            MyList.SubItems(2) = CStr(rsLIST!QuotationNumber)
        End If
        
        If Not IsNull(rsLIST!JobBriefDate) Then
            MyList.SubItems(3) = CStr(rsLIST!JobBriefDate)
        End If
        
        If Not IsNull(rsLIST!deadlineDate) Then
            MyList.SubItems(4) = CStr(rsLIST!deadlineDate)
        End If
        
        If Not IsNull(rsLIST!CompanyName) Then
            MyList.SubItems(5) = CStr(rsLIST!CompanyName)
        End If
        
        If Not IsNull(rsLIST!Contactname) Then
            MyList.SubItems(6) = CStr(rsLIST!Contactname)
        End If
        
        If Not IsNull(rsLIST!phone) Then
        MyList.SubItems(7) = CStr(rsLIST!phone)
        End If
        
        If Not IsNull(rsLIST!Mobilephone) Then
            MyList.SubItems(8) = CStr(rsLIST!Mobilephone)
        End If
        
        If Not IsNull(rsLIST!LPONo) Then
            MyList.SubItems(9) = CStr(rsLIST!LPONo)
        End If
    
    rsLIST.MoveNext
    
Wend

Set MyList = Nothing: Set rsLIST = Nothing
End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Public Sub ShowSitesExpiryOnSpecificDate()
On Error GoTo err
Dim Expiry As Variant
With SchedulingMain
Expiry = Format(Date, "MMMM dd,yyyy")
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Site No ", .ListView1.Width / 6.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Site Name", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Site Physical Address", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "City", .ListView1.Width / 6.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "LandLord Name", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "LandLord Physical Address", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract Start Date", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Contract Finish Date", .ListView1.Width / 5.5 ', lvwColumnCenter

INPQRY = InputBox("Please Enter the Date on Which to show Payments Due..." & vbCrLf & vbCrLf & "Format: - dd/MM/yyyy", "Enter Date", Date)

If Len(INPQRY) = 0 Then
    MsgBox "No Values Entered or the Operation has been Cancelled!!", vbCritical + vbOKOnly, "No Values"
    Exit Sub
End If

Dim q As Date, ThisDate As String

q = CDate(INPQRY): ThisDate = Format(q, "MMMM dd,yyyy")

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset


'rsLIST.Open "SELECT * FROM AdvertSites A,AdvertSiteLords B WHERE A.LandLordNo = B.LandLordNo AND A.ContractFinish = '" & ThisDate & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem
Exit Sub
If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))

    If Not IsNull(rsLIST!SiteName) Then
        MyList.SubItems(1) = CStr(rsLIST!SiteName)
    End If
    
    If Not IsNull(rsLIST!SitePhysicalAddress) Then
        MyList.SubItems(2) = CStr(rsLIST!SitePhysicalAddress)
    End If
        
    If Not IsNull(rsLIST!city) Then
        MyList.SubItems(3) = CStr(rsLIST!city)
    End If
    
    If Not IsNull(rsLIST!Surname) Then
        MyList.SubItems(4) = CStr(rsLIST!Surname)
    End If
    
    If Not IsNull(rsLIST!PhysicalAddress) Then
        MyList.SubItems(4) = CStr(rsLIST!PhysicalAddress)
    End If
        
    If Not IsNull(rsLIST!ContractStart) Then
        MyList.SubItems(5) = CStr(rsLIST!ContractStart)
    End If

    If Not IsNull(rsLIST!ContractFinish) Then
        MyList.SubItems(6) = CStr(rsLIST!ContractFinish)
    End If
    rsLIST.MoveNext
    
Wend

NewRecord = False
Set MyList = Nothing: Set rsLIST = Nothing
End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Public Sub ShowAllSitesNotValid()
On Error GoTo err
Dim Expiry As Variant
With SchedulingMain
Expiry = Format(Date, "MMMM dd,yyyy")
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Site No ", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Site Name", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Site Physical Address", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "City", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "LandLord Name", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "LandLord Physical Address", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract Start Date", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Contract Finish Date", .ListView1.Width / 5.5 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ODASPPlot A,ODASPAccount B WHERE A.AccountNo = B.AccountNo AND  A.Status IS NULL ;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))

    If Not IsNull(rsLIST!SiteName) Then
        MyList.SubItems(1) = CStr(rsLIST!SiteName)
    End If
    
    If Not IsNull(rsLIST!SitePhysicalAddress) Then
        MyList.SubItems(2) = CStr(rsLIST!SitePhysicalAddress)
    End If
        
    If Not IsNull(rsLIST!city) Then
        MyList.SubItems(3) = CStr(rsLIST!city)
    End If
    
    If Not IsNull(rsLIST!Surname) Then
        MyList.SubItems(4) = CStr(rsLIST!Surname)
    End If
    
    If Not IsNull(rsLIST!PhysicalAddress) Then
        MyList.SubItems(4) = CStr(rsLIST!PhysicalAddress)
    End If
        
    If Not IsNull(rsLIST!ContractStart) Then
        MyList.SubItems(5) = CStr(rsLIST!ContractStart)
    End If

    If Not IsNull(rsLIST!ContractFinish) Then
        MyList.SubItems(6) = CStr(rsLIST!ContractFinish)
    End If
   
    rsLIST.MoveNext
Wend
NewRecord = False

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Public Sub ShowAllcustomers()
On Error GoTo err
With SchedulingMain
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Customer Id", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Name", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contact Person", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contact Title", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Address", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "City", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Country", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Phone", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Email", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Fax", .ListView1.Width / 5.5 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertClients;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!customerid))

    If Not IsNull(rsLIST!CompanyName) Then
        MyList.SubItems(1) = CStr(rsLIST!CompanyName)
    End If
    
    If Not IsNull(rsLIST!Contactname) Then
        MyList.SubItems(2) = CStr(rsLIST!Contactname)
    End If
        
    If Not IsNull(rsLIST!ContactTitle) Then
        MyList.SubItems(3) = CStr(rsLIST!ContactTitle)
    End If
    
    If Not IsNull(rsLIST!Address) Then
        MyList.SubItems(4) = CStr(rsLIST!Address)
    End If
    
    If Not IsNull(rsLIST!city) Then
        MyList.SubItems(5) = CStr(rsLIST!city)
    End If
        
    If Not IsNull(rsLIST!Country) Then
        MyList.SubItems(6) = CStr(rsLIST!Country)
    End If

    If Not IsNull(rsLIST!PhysicalAddress) Then
        MyList.SubItems(7) = CStr(rsLIST!PhysicalAddress)
    End If
    
    If Not IsNull(rsLIST!phone) Then
        MyList.SubItems(8) = CStr(rsLIST!phone)
    End If
    
    If Not IsNull(rsLIST!Email) Then
        MyList.SubItems(9) = CStr(rsLIST!Email)
    End If
    
    If Not IsNull(rsLIST!Fax) Then
        MyList.SubItems(10) = CStr(rsLIST!Fax)
    End If
   
    rsLIST.MoveNext
    
Wend

NewRecord = False

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Public Sub showALLLandlords()
On Error GoTo err
With SchedulingMain
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "LandLord No", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Name", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 3 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "City", .ListView1.Width / 8 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Postal Address", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Telephone", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Mobile No", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "E-Mail", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Fax", .ListView1.Width / 5.5 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ODASPAccount WHERE AccountType = 'LLORD';", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!AccountNo))

    If Not IsNull(rsLIST!CompanyName) Then
        MyList.SubItems(1) = CStr(rsLIST!CompanyName)
    End If
    
    If Not IsNull(rsLIST!PhysicalAddress) Then
        MyList.SubItems(2) = CStr(rsLIST!PhysicalAddress)
    End If
        
    If Not IsNull(rsLIST!Towncity) Then
        MyList.SubItems(3) = CStr(rsLIST!Towncity)
    End If
    
    If Not IsNull(rsLIST!PostalAddress) Then
        MyList.SubItems(4) = CStr(rsLIST!PostalAddress)
    End If
    
    If Not IsNull(rsLIST!TelephoneNo) Then
        MyList.SubItems(5) = CStr(rsLIST!TelephoneNo)
    End If
        
    If Not IsNull(rsLIST!MobileNo) Then
        MyList.SubItems(6) = CStr(rsLIST!MobileNo)
    End If

    If Not IsNull(rsLIST!EmailAddress) Then
        MyList.SubItems(7) = CStr(rsLIST!EmailAddress)
    End If
    
    If Not IsNull(rsLIST!FAxNo) Then
        MyList.SubItems(8) = CStr(rsLIST!FAxNo)
    End If
rsLIST.MoveNext
    
Wend

NewRecord = False

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Public Sub ShowAllContractsNotPinned()
On Error GoTo err
With SchedulingMain

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Serial No", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Contract Start Date", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "Contract End Date", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Code", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Name", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Advert Cost", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Client Name", .ListView1.Width / 4

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertContractRequisition A,AdvertcontractRequisitionData B WHERE A.PurchaseOrderNo = B.PurchaseOrderNo AND  B.PaidStatus='" & "Y" & "' AND B.AllocationStatus IS NULL ORDER BY PurchaseOrderNo;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: .txtTotal.Text = 0: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SerialNo))
    
    If Not IsNull(rsLIST!ContractStartDate) Then
        MyList.SubItems(1) = CStr(rsLIST!ContractStartDate)
    End If
    
    If Not IsNull(rsLIST!ContractEndDate) Then
        MyList.SubItems(2) = CStr(rsLIST!ContractEndDate)
    End If
    
    If Not IsNull(rsLIST!ContractNo) Then
        MyList.SubItems(3) = CStr(rsLIST!ContractNo)
    End If
    
    If Not IsNull(rsLIST!Advcode) Then
        MyList.SubItems(4) = CStr(rsLIST!Advcode)
    End If
    
    If Not IsNull(rsLIST!AdvName) Then
        MyList.SubItems(4) = CStr(rsLIST!AdvName)
    End If
    
    If Not IsNull(rsLIST!AdvCost) Then
        MyList.SubItems(5) = FormatNumber(rsLIST!AdvCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
   If Not IsNull(rsLIST!ClientName) Then
        MyList.SubItems(6) = CStr(rsLIST!ClientName)
    End If
    
    
    rsLIST.MoveNext
    
Wend

'.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
'.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
.txtTotal.Text = .ListView1.ListItems.Count & "Contracts"

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Public Sub ShowAllValidEmptyBillBoards()
On Error GoTo err
Dim Expiry As Variant
With SchedulingMain
Expiry = Format(Date, "MMMM dd,yyyy")
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "BillBord No ", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "BillBord Details", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Location", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Faces Free", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Land Lord", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Start Date", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Expiry Date", .ListView1.Width / 6.5 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ODASPPlot P,ODASPTown T, ODASPAccount AC WHERE P.BillBoard = 'Y' and T.TownCode = P.TownCode and P.AccountNo = AC.AccountNo ;", cnCOMMON, adOpenKeyset, adLockOptimistic
If rsLIST.RecordCount = 0 Then Exit Sub

    rsLIST.MoveFirst
        Do While Not rsLIST.EOF
            Set rsFindRecord = New ADODB.Recordset
            rsFindRecord.Open "SELECT * FROM ODASPPlotSite PS, ODASPPlotMast PM WHERE PS.PlotNo = '" & rsLIST!PlotNo & "' and PS.Status = 'SITE-AVAILABLE' and PS.MastNo = PM.MastNo;", cnCOMMON, adOpenKeyset, adLockOptimistic
            If rsFindRecord.RecordCount = 0 Then GoTo SearchRecords
            
            Dim MyList As ListItem
            Set MyList = .ListView1.ListItems.Add(, , CStr(rsFindRecord!MastNo))
            
                If Not IsNull(rsFindRecord!MastDetails) Then
                    MyList.SubItems(1) = CStr(rsFindRecord!MastDetails)
                End If
                
                If Not IsNull(rsLIST!PhysicalLocation) And Not IsNull(rsLIST!Town) Then
                    MyList.SubItems(2) = CStr(rsLIST!PhysicalLocation) & " IN" & CStr(rsLIST!Town)
                End If
                    
                MyList.SubItems(3) = rsFindRecord.RecordCount
                
                If Not IsNull(rsLIST!CompanyName) Then
                    MyList.SubItems(4) = CStr(rsLIST!CompanyName)
                End If
                
                If Not IsNull(rsLIST!CommencementDate) Then
                    MyList.SubItems(5) = Format(rsLIST!CommencementDate, "dd/mm/yyyy")
                End If
            
                If Not IsNull(rsLIST!expirydate) Then
                    MyList.SubItems(6) = Format(rsLIST!expirydate, "dd/mm/yyyy")
                End If
                  
SearchRecords:
                rsLIST.MoveNext
                
            Loop

NewRecord = False

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Public Sub ShowAllRentFeeDueOnSpecificDate()
On Error GoTo err
With SchedulingMain
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Plot No", .ListView1.Width / 7
.ListView1.ColumnHeaders.Add , , "Plot Name", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Masts", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Sites", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Annual Rent", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Rent Paid", .ListView1.Width / 7
.ListView1.ColumnHeaders.Add , , "Balance", .ListView1.Width / 7
.ListView1.ColumnHeaders.Add , , "Rent Due on: ", .ListView1.Width / 6

.ListView1.View = lvwReport

INPQRY = InputBox("Please Enter the Date on Which to show Payments Due..." & vbCrLf & vbCrLf & "Format: - dd/MM/yyyy", "Enter Date", Date)

If Len(INPQRY) = 0 Then
    MsgBox "No Values Entered or the Operation has been Cancelled!!", vbCritical + vbOKOnly, "No Values"
    Exit Sub
End If

Dim q As Date, ThisDate As String, Balance As Currency

q = CDate(INPQRY): ThisDate = Format(q, "MMMM dd,yyyy")

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ODASPPlot WHERE RentDueDate ='" & Trim(ThisDate) & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: .txtTotal.Text = 0: Exit Sub
'End If
Else

While Not rsLIST.EOF
If CCur(rsLIST!AnnualRent) > CCur(rsLIST!RentPaid) Or CCur(rsLIST!RentPaid) = 0 Then

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PlotNo))
    
    If Not IsNull(rsLIST!PlotName) Then
        MyList.SubItems(1) = CStr(rsLIST!PlotName)
    End If
    
    If Not IsNull(rsLIST!NoOfMasts) Then
        MyList.SubItems(2) = CStr(rsLIST!NoOfMasts)
    End If
    
    If Not IsNull(rsLIST!NoofSites) Then
        MyList.SubItems(3) = CStr(rsLIST!NoofSites)
    End If
    
    If Not IsNull(rsLIST!AnnualRent) Then
        MyList.SubItems(4) = CCur(rsLIST!AnnualRent)
    End If
    
    If Not IsNull(rsLIST!RentPaid) Then
        MyList.SubItems(5) = CCur(rsLIST!RentPaid)
    End If
    Balance = CCur(rsLIST!AnnualRent) - CCur(rsLIST!RentPaid)
    MyList.SubItems(6) = Balance
    
    If Not IsNull(rsLIST!RentDueDate) Then
        MyList.SubItems(7) = CStr(rsLIST!RentDueDate)
    End If
      
    
    rsLIST.MoveNext
  End If
Wend

'.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
'.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
.txtTotal.Text = .ListView1.ListItems.Count & "Payments"

Set MyList = Nothing: Set rsLIST = Nothing
End If
End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub ShowAllCouncilFeeDueOnSpecificDate()
On Error GoTo err
With SchedulingMain


.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Plot No", .ListView1.Width / 7
.ListView1.ColumnHeaders.Add , , "Plot Name", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Masts", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Sites", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Annual Rate", .ListView1.Width / 7 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Rate Paid", .ListView1.Width / 7
.ListView1.ColumnHeaders.Add , , "Balance", .ListView1.Width / 7
.ListView1.ColumnHeaders.Add , , "Rate Due on: ", .ListView1.Width / 6


.ListView1.View = lvwReport

INPQRY = InputBox("Please Enter the Date on Which to show Payments Due..." & vbCrLf & vbCrLf & "Format: - dd/MM/yyyy", "Enter Date", Date)

If Len(INPQRY) = 0 Then
    MsgBox "No Values Entered or the Operation has been Cancelled!!", vbCritical + vbOKOnly, "No Values"
    Exit Sub
End If

Dim q As Date, ThisDate As String, Balance As Currency

q = CDate(INPQRY): ThisDate = Format(q, "MMMM dd,yyyy")

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ODASPPlot WHERE RateDueDate ='" & Trim(ThisDate) & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: .txtTotal.Text = 0: Exit Sub
'End If
Else

While Not rsLIST.EOF
If CCur(rsLIST!AnnualRate) > CCur(rsLIST!RatePaid) Or CCur(rsLIST!RatePaid) = 0 Then

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PlotNo))
    
    If Not IsNull(rsLIST!PlotName) Then
        MyList.SubItems(1) = CStr(rsLIST!PlotName)
    End If
    
    If Not IsNull(rsLIST!NoOfMasts) Then
        MyList.SubItems(2) = CStr(rsLIST!NoOfMasts)
    End If
    
    If Not IsNull(rsLIST!NoofSites) Then
        MyList.SubItems(3) = CStr(rsLIST!NoofSites)
    End If
    
    If Not IsNull(rsLIST!AnnualRate) Then
        MyList.SubItems(4) = CCur(rsLIST!AnnualRate)
    End If
    
    If Not IsNull(rsLIST!RatePaid) Then
        MyList.SubItems(5) = CCur(rsLIST!RatePaid)
    End If
    Balance = CCur(rsLIST!AnnualRate) - CCur(rsLIST!RatePaid)
    MyList.SubItems(6) = Balance
    
    If Not IsNull(rsLIST!RateDueDate) Then
        MyList.SubItems(7) = CStr(rsLIST!RateDueDate)
    End If
      
    
    rsLIST.MoveNext
  End If
Wend

'.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
'.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
.txtTotal.Text = .ListView1.ListItems.Count & "Payments"

Set MyList = Nothing: Set rsLIST = Nothing
End If
End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Public Sub ShowAllContractsExpiryOnSpecifiedDate()
On Error GoTo err
With SchedulingMain

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear
.ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Start Date", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "End Date", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "LandLord", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Plot No", .ListView1.Width / 5.5

.ListView1.View = lvwReport

INPQRY = InputBox("Please Enter the Date on Which to show contract Expiry..." & vbCrLf & vbCrLf & "Format: - dd/MM/yyyy", "Enter Date", Date)

If Len(INPQRY) = 0 Then
    MsgBox "No Values Entered or the Operation has been Cancelled!!", vbCritical + vbOKOnly, "No Values"
    Exit Sub
End If

Dim q As Date, ThisDate As String

q = CDate(INPQRY): ThisDate = Format(q, "MMMM dd,yyyy")

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ODASMLeaseAgreement,ODASPPlot  WHERE (ODASMLeaseAgreement.Terminated ='N' or ODASMLeaseAgreement.Terminated is null) AND ODASPPlot.ExpiryDate ='" & Trim(ThisDate) & "' AND ODASPPlot.PlotNo = ODASMLeaseAgreement.PlotNo;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: .txtTotal.Text = 0: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!ContractNo))
    
    If Not IsNull(rsLIST!CommencementDate) Then
        MyList.SubItems(1) = CStr(rsLIST!CommencementDate)
    End If
    
    If Not IsNull(rsLIST!expirydate) Then
        MyList.SubItems(2) = CStr(rsLIST!expirydate)
    End If
    
    If Not IsNull(rsLIST!AccountNo) Then
        MyList.SubItems(3) = CStr(rsLIST!AccountNo)
    End If
    
    If Not IsNull(rsLIST!PlotNo) Then
        MyList.SubItems(4) = CStr(rsLIST!PlotNo)
    End If
    rsLIST.MoveNext
    
Wend

.txtTotal.Text = .ListView1.ListItems.Count & " Contracts"

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Public Sub ShowAllValidContracts()
On Error GoTo err
Dim Expiry As Variant
With SchedulingMain
Expiry = Format(Date, "MMMM dd,yyyy")
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Serial No ", .ListView1.Width / 6.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract Start Date", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract Expiry Date", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 6.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Bill Board No", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Site Name", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "City", .ListView1.Width / 5 '
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Code", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Name", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Client Name", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM ContractSiteAllocationData WHERE ExpDate > '" & Expiry & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SerialNo))

    If Not IsNull(rsLIST!StartDate) Then
        MyList.SubItems(1) = Format(rsLIST!StartDate, "dd/mm/yyyy")
    End If
    
    If Not IsNull(rsLIST!EXPDate) Then
        MyList.SubItems(2) = Format(rsLIST!EXPDate, "dd/mm/yyyy")
    End If
        
    If Not IsNull(rsLIST!SiteNumber) Then
        MyList.SubItems(3) = CStr(rsLIST!SiteNumber)
    End If
    
    If Not IsNull(rsLIST!ContractNo) Then
        MyList.SubItems(4) = CStr(rsLIST!ContractNo)
    End If
    
    If Not IsNull(rsLIST!BillBoardNo) Then
        MyList.SubItems(5) = CStr(rsLIST!BillBoardNo)
    End If
        
    If Not IsNull(rsLIST!SiteName) Then
        MyList.SubItems(6) = CStr(rsLIST!SiteName)
    End If

    If Not IsNull(rsLIST!city) Then
        MyList.SubItems(7) = CStr(rsLIST!city)
    End If
    
    If Not IsNull(rsLIST!SitePhysicalAddress) Then
        MyList.SubItems(8) = CStr(rsLIST!SitePhysicalAddress)
    End If
    
    If Not IsNull(rsLIST!Advcode) Then
        MyList.SubItems(9) = CStr(rsLIST!Advcode)
    End If
    
    If Not IsNull(rsLIST!AdvName) Then
        MyList.SubItems(10) = CStr(rsLIST!AdvName)
    End If
    
    If Not IsNull(rsLIST!ClientName) Then
        MyList.SubItems(11) = CStr(rsLIST!ClientName)
    End If
    rsLIST.MoveNext
    
Wend

NewRecord = False

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
err:
If err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Private Sub LoadReminders()
On Error GoTo err
OpenODBCConnection
With Me
    Set rsFindRecord = New ADODB.Recordset
    rsFindRecord.Open "Select * From ODASMCouncilRateDue Where DueDate< '" & Format(Date, "MMMM dd,YYYY") & "' and (Paid = 'N' or Paid is null)", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub
        If MsgBox(rsFindRecord.RecordCount & " Site(s) have Rates due dates expired yet haven't been paid... Do you want to see the sites?", vbYesNo, "Council rate payments remainder") = vbYes Then
            getALLExpiredSites
            bAlert = True
        End If
End With
Exit Sub
err:
ErrorMessage

End Sub
Public Sub getALLExpiredSites()
On Error GoTo err
    
        With SchedulingMain
        
                .ListView1.ListItems.Clear
                .ListView1.ColumnHeaders.Clear
                
                .ListView1.ColumnHeaders.Add , , "Council", .ListView1.Width / 4.5 ', lvwColumnCenter
                .ListView1.ColumnHeaders.Add , , "Due Date", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Amount Payable", .ListView1.Width / 6
                .ListView1.ColumnHeaders.Add , , "Site Details", .ListView1.Width / 2

                .ListView1.View = lvwReport
                
                Dim rsLIST As ADODB.Recordset
                Dim MyList As ListItem
                
                Set rsFindRecord = New ADODB.Recordset
                rsFindRecord.Open "Select * From ODASMCouncilRateDue Where DueDate< '" & Format(Date, "MMMM dd,YYYY") & "' and (Paid = 'N' or Paid is null)", cnCOMMON, adOpenKeyset, adLockOptimistic
                
                rsFindRecord.MoveFirst
                While Not rsFindRecord.EOF
                    
                    If rsFindRecord!Face = "Y" Then
                
                        Set rsLIST = New ADODB.Recordset
                        strSQL = "SELECT *  FROM ODASPPlot P,ODASPCouncil C, ODASPPlotSite PS where PS.SiteNo = '" & rsFindRecord!SiteNo & "' and PS.PlotNo = P.PLotNo and P.CouncilCode = C.CouncilCode ;"
                        rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                            
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!Council))
                        
                            If Not IsNull(rsFindRecord!DueDate) Then
                                MyList.SubItems(1) = CStr(rsFindRecord!DueDate)
                            End If
                            
                            If Not IsNull(rsFindRecord!AmountDue) Then
                                MyList.SubItems(2) = CStr(rsFindRecord!AmountDue)
                            End If
                            
                            If Not IsNull(rsLIST!CommencementDate) Then
                                MyList.SubItems(3) = CStr(rsLIST!SiteDetails)
                            End If
                    
                    ElseIf rsFindRecord!BillBoard = "Y" Then
                               
                        Set rsLIST = New ADODB.Recordset
                        strSQL = "SELECT *  FROM ODASPPlot P,ODASPCouncil C, ODASPPlotMast PS where PS.MastNo = '" & rsFindRecord!SiteNo & "' and PS.PlotNo = P.PLotNo and P.CouncilCode = C.CouncilCode ;"
                        rsLIST.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                        
                        Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!Council))
                            
                            If Not IsNull(rsFindRecord!DueDate) Then
                                MyList.SubItems(1) = CStr(rsFindRecord!DueDate)
                            End If
                            
                            If Not IsNull(rsFindRecord!AmountDue) Then
                                MyList.SubItems(2) = CStr(rsFindRecord!AmountDue)
                            End If
                            
                            If Not IsNull(rsLIST!MastDetails) Then
                                MyList.SubItems(3) = CStr(rsLIST!MastDetails)
                            End If
                        End If
                    rsFindRecord.MoveNext
                    Wend
                Set MyList = Nothing: Set rsFindRecord = Nothing
        End With

Exit Sub

err:
        If err.Number = 3265 Then Resume Next
         ErrorMessage
End Sub

Public Sub UpdateContract()
On Error GoTo err
On Error Resume Next
With Me
        Set rsSAVE = New ADODB.Recordset
        strSQL = "SELECT * FROM ODASPPlot WHERE PlotNo = '" & CurrentRecord & "';"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic

        With rsSAVE
                    
                    With Me
                             rsSAVE!paid = ""
                                                                                       
                           
                            rsSAVE.Update
                            rsSAVE.Requery
                        
                    End With

        End With
End With
Exit Sub

err:
    ErrorMessage
End Sub
Public Sub AllowProcess()
On Error GoTo err
       
              CurrentUserName = CurrentUserName
              strSQL = "select * from UserMaster Where UserMaster.Username='" & CurrentUserName & "';"
                Set rsFindRecord = New ADODB.Recordset
                rsFindRecord.Open strSQL, cnSECURE, adOpenKeyset, adLockOptimistic
                    CurrentPic = rsFindRecord!StaffId
                strSQL = "select * from ODASPApprovers Where ODASPApprovers.StaffId='" & CurrentPic & "' and (OperationType='15' or OperationType='7' or OperationType='8');"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                
                If rsSAVE.EOF Or rsSAVE.BOF Then
                MsgBox "Rights Denied. You Do NOT have rights to Perform This Task. Consult your system administrator for assistance...!", vbCritical + vbOKOnly, "Rights Denied...!"
                bAllowProcess = False
                End If
                bAllowProcess = True
                Set rsSAVE = Nothing
                strSQL = Empty
        
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub UpdateTargetDate()

  Set rsFindRecord = New ADODB.Recordset
   rsFindRecord.Open "Select * From ODASPPlotMast", cnCOMMON, adOpenKeyset, adLockOptimistic
    If rsFindRecord.EOF And rsFindRecord.BOF Then
    End If
      Dim StartDate As Date
      Dim EndDate As Date
      Dim Targetdate As Date
      Dim Difference As String
      
     Do Until rsFindRecord.EOF
      StartDate = rsFindRecord!CommencementDate
      EndDate = rsFindRecord!expirydate
      Difference = DateDiff("M", Date, EndDate)
      Targetdate = DateAdd("M", -6, Format(EndDate, "MMMM dd,yyyy"))
      
     
     
        rsFindRecord!Targetdate = Targetdate
        
        rsFindRecord.Update
        
       rsFindRecord.MoveNext
    Loop
      
    Set rsFindRecord = Nothing
    


End Sub
