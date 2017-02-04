VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmReceivedQuotations 
   Caption         =   "IMPORT DECLARATION FORM AMMENDMENT"
   ClientHeight    =   7845
   ClientLeft      =   75
   ClientTop       =   630
   ClientWidth     =   11880
   Icon            =   "frmReceivedQuotations.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   11880
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10680
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
            Picture         =   "frmReceivedQuotations.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceivedQuotations.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceivedQuotations.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceivedQuotations.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceivedQuotations.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceivedQuotations.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceivedQuotations.frx":236E
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
      Width           =   11880
      _ExtentX        =   20955
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
            Object.ToolTipText     =   "Create and Save New Records"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Edit/Change "
            Key             =   "E"
            Object.ToolTipText     =   "Edit/Change and Update Existing Records"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Search/Find "
            Key             =   "S"
            Object.ToolTipText     =   "Search and Display Existing Records"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Refresh "
            Key             =   "R"
            Object.ToolTipText     =   "Refresh the Screen"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print &Preview "
            Key             =   "P"
            Object.ToolTipText     =   "Print Preview Form/Report"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Help System  "
            Key             =   "H"
            Object.ToolTipText     =   "Get Help Online"
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
   Begin VB.Frame Frame1 
      Height          =   7215
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   11895
      Begin VB.Frame Frame3 
         Caption         =   "General/Main Ammendment Details"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   5160
         TabIndex        =   25
         Top             =   240
         Width           =   6615
         Begin VB.TextBox txtAmmendNo 
            BackColor       =   &H00FFFFC0&
            Height          =   345
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtIndentNo 
            BackColor       =   &H00FFFFC0&
            Height          =   345
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtIDFNumber 
            BackColor       =   &H00FFFFC0&
            Height          =   345
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtIDFDate 
            BackColor       =   &H00FFFFC0&
            Height          =   345
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   600
            Width           =   1695
         End
         Begin VB.ComboBox cboCommodityDesc 
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1080
            TabIndex        =   28
            Top             =   1170
            Width           =   5415
         End
         Begin VB.ComboBox cboAmmendType 
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1080
            TabIndex        =   27
            Top             =   1620
            Width           =   2055
         End
         Begin VB.ComboBox cboAmmendPurp 
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3840
            TabIndex        =   26
            Top             =   1620
            Width           =   2655
         End
         Begin VB.Label Label23 
            Caption         =   "Ammend. No"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   5280
            TabIndex        =   50
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Indent No"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Comm. Desc"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1170
            Width           =   975
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   6480
            Y1              =   1050
            Y2              =   1050
         End
         Begin VB.Label Label5 
            Caption         =   "IDF Number"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1680
            TabIndex        =   35
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "IDF Issue Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   3480
            TabIndex        =   34
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Nature/Type"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1620
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Purpose"
            Height          =   255
            Left            =   3195
            TabIndex        =   32
            Top             =   1620
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Product Quantities/Sizes Ammended"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   5160
         TabIndex        =   2
         Top             =   2520
         Width           =   6615
         Begin VB.TextBox txtNetWeight 
            BackColor       =   &H00FFFFC0&
            Height          =   345
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   3512
            Width           =   2175
         End
         Begin VB.TextBox txtFOBValue 
            BackColor       =   &H00FFFFC0&
            Height          =   345
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   3512
            Width           =   2055
         End
         Begin VB.TextBox txtExchRate 
            BackColor       =   &H00FFFFC0&
            Height          =   345
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   3066
            Width           =   2175
         End
         Begin VB.TextBox txtCurrency 
            BackColor       =   &H00FFFFC0&
            Height          =   345
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   3066
            Width           =   2055
         End
         Begin VB.TextBox txtProductCode 
            BackColor       =   &H00FFFFC0&
            Height          =   345
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtProductName 
            BackColor       =   &H00FFFFC0&
            Height          =   345
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   480
            Width           =   4575
         End
         Begin VB.TextBox txtOrigDesc 
            BackColor       =   &H00FFFFC0&
            Height          =   585
            Left            =   1320
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   1488
            Width           =   5175
         End
         Begin VB.TextBox txtNewPackages 
            BackColor       =   &H00FFC0C0&
            Height          =   345
            Left            =   1320
            TabIndex        =   10
            Top             =   2174
            Width           =   2295
         End
         Begin VB.TextBox txtPackageType 
            BackColor       =   &H00FFFFC0&
            Height          =   345
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   2174
            Width           =   1815
         End
         Begin VB.TextBox txtUnitsPerPack 
            BackColor       =   &H00FFFFC0&
            Height          =   345
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   2620
            Width           =   975
         End
         Begin VB.TextBox txtPackageSize 
            BackColor       =   &H00FFFFC0&
            Height          =   345
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   2620
            Width           =   1335
         End
         Begin VB.TextBox txtQttyUnits 
            BackColor       =   &H00FFFFC0&
            Height          =   345
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   2620
            Width           =   975
         End
         Begin VB.TextBox txtNewDesc 
            BackColor       =   &H00FFFFC0&
            Height          =   465
            Left            =   1320
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   3960
            Width           =   5175
         End
         Begin VB.TextBox txtCommodityCode 
            BackColor       =   &H00FFFFC0&
            Height          =   345
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   1042
            Width           =   2295
         End
         Begin VB.TextBox txtSITC 
            BackColor       =   &H00FFFFC0&
            Height          =   345
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1042
            Width           =   2295
         End
         Begin VB.Label Label22 
            Caption         =   "Net Weight"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   48
            Top             =   3512
            Width           =   735
         End
         Begin VB.Label Label21 
            Caption         =   "FOB Value"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   3512
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   "Exch. Rate"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   44
            Top             =   3066
            Width           =   735
         End
         Begin VB.Label Label19 
            Caption         =   "Currency"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   3066
            Width           =   1215
         End
         Begin VB.Label Label10 
            Caption         =   "Product Code"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label11 
            Caption         =   "Product/Brand Name"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1920
            TabIndex        =   23
            Top             =   240
            Width           =   2775
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   6480
            Y1              =   926
            Y2              =   926
         End
         Begin VB.Label Label14 
            Caption         =   "Origin. Description"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1488
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   "Total Packages"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   2174
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "Package Type"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   20
            Top             =   2174
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Units per Package"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   2620
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Package Size"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   18
            Top             =   2620
            Width           =   975
         End
         Begin VB.Label Label17 
            Caption         =   "Qtty Units"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4800
            TabIndex        =   17
            Top             =   2620
            Width           =   735
         End
         Begin VB.Label Label18 
            Caption         =   "New Description"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   3960
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "Commodity Code"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1042
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "SITC"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   14
            Top             =   1042
            Width           =   375
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2055
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   3625
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
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4455
         Left            =   120
         TabIndex        =   39
         Top             =   2640
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   7858
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
         NumItems        =   0
      End
      Begin VB.Label Label2 
         Caption         =   "Particulars of Commodities Under Selected Indent"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   2400
         Width           =   4335
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileClear 
         Caption         =   "Clear The &Screen"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnujjks 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Checked         =   -1  'True
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuShow 
      Caption         =   "Show/&View"
      Begin VB.Menu mnuShowNewIndents 
         Caption         =   "New Indents &With IDF Numbers"
         Checked         =   -1  'True
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnudfdgs 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowAllIndents 
         Caption         =   "All Indents With &Ammendments"
         Shortcut        =   ^{F8}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help System"
      Begin VB.Menu mnuHelpUsing 
         Caption         =   "How to &Use this Screen"
         Checked         =   -1  'True
         Shortcut        =   ^{F1}
      End
   End
End
Attribute VB_Name = "frmReceivedQuotations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboAmmendPurp_Click()
If Not NewRecord And Not EditRecord Then Exit Sub
    Me.ListView2.SetFocus
End Sub

Private Sub cboAmmendPurp_GotFocus()
If Not NewRecord And Not EditRecord Then Exit Sub
With Me
    If .cboAmmendPurp.ListCount <> 0 Then Exit Sub
    AttachSQL = "SELECT PurposeDescription AS SelectField FROM ParamIDFAmmendPurp ORDER BY PurposeDescription;"
    If .cboAmmendPurp.Text = Empty Then
        .cboAmmendPurp.Clear
    End If
    MyCommonData.AttachDropDown
End With
End Sub

Private Sub cboAmmendPurp_LostFocus()
On Error GoTo Err
If Not NewRecord Or EditRecord Then Exit Sub
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT ParamIDFAmmendPurp.* FROM ParamIDFAmmendPurp WHERE ParamIDFAmmendPurp.PurposeDescription='" & Trim(.cboAmmendPurp.Text) & "';")
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .cboAmmendPurp.Text = Trim(rsFindRecord!AmmendPurpose & "")
        .ListView2.SetFocus
    End If
    Set rsFindRecord = Nothing
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub cboAmmendType_Click()
If Not NewRecord And Not EditRecord Then Exit Sub
    Me.cboAmmendPurp.SetFocus
End Sub

Private Sub cboAmmendType_GotFocus()
If Not NewRecord And Not EditRecord Then Exit Sub
With Me
    If .cboAmmendType.ListCount <> 0 Then Exit Sub
    AttachSQL = "SELECT AmmendName AS SelectField FROM ParamIDFAmmendType ORDER BY AmmendName;"
    If .cboAmmendType.Text = Empty Then
        .cboAmmendType.Clear
    End If
    MyCommonData.AttachDropDown
End With
End Sub

Private Sub cboAmmendType_LostFocus()
On Error GoTo Err
If Not NewRecord Or EditRecord Then Exit Sub
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT ParamIDFAmmendType.* FROM ParamIDFAmmendType WHERE ParamIDFAmmendType.AmmendName='" & Trim(.cboAmmendType.Text) & "';")
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .cboAmmendType.Text = Trim(rsFindRecord!AmmendType & "")
        .cboAmmendPurp.SetFocus
    End If
    Set rsFindRecord = Nothing
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub cboCommodityDesc_Click()
If Not NewRecord And Not EditRecord Then Exit Sub
    With Me
        .cboAmmendType.SetFocus
    End With
End Sub

Private Sub cboCommodityDesc_GotFocus()
If Not NewRecord And Not EditRecord Then Exit Sub
With Me
    If .cboCommodityDesc.ListCount <> 0 Then Exit Sub
    AttachSQL = "SELECT GoodsDescription AS SelectField FROM ParamInsuranceDesc ORDER BY CodeNumber;"
    If .cboCommodityDesc.Text = Empty Then
        .cboCommodityDesc.Clear
    End If
    MyCommonData.AttachDropDown
End With
End Sub

Private Sub cboCommodityDesc_LostFocus()
On Error GoTo Err
If Not NewRecord And Not EditRecord Then Exit Sub
With Me
If .cboCommodityDesc.Text = Empty Then Exit Sub
    If .cboAmmendType.Text = Empty Then
        .cboAmmendType.SetFocus
    Else
        If .cboAmmendPurp.Text = Empty Then
            .cboAmmendPurp.SetFocus
        Else
            If .txtProductCode.Text = Empty Then
                .ListView2.SetFocus
            Else
                .txtNewPackages.SetFocus
            End If
        End If
    End If
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub Form_Activate()
With Me
    If .ListView1.ListItems.Count = 0 Or .ListView1.View <> lvwReport Then
        Call GetAllNewIndents: Call GetCommodityUnderIndent
    End If
End With
End Sub

Private Sub ClearMainData1()
With Me
'    .txtIndentNo.Text = Empty
'    .txtIDFNumber.Text = Empty
'    .txtIDFDate.Text = Empty
'    .cboAmmendPurp.Clear
'    .cboAmmendType.Clear
'    .cboCommodityDesc.Clear
    .txtCommodityCode.Text = Empty
    .txtNewDesc.Text = Empty
    .txtNewPackages.Text = Empty
    .txtOrigDesc.Text = Empty
    .txtPackageSize.Text = Empty
    .txtPackageType.Text = Empty
    .txtProductCode.Text = Empty
    .txtProductName.Text = Empty
    .txtQttyUnits.Text = Empty
    .txtSITC.Text = Empty
    .txtUnitsPerPack.Text = Empty
End With
End Sub

Private Sub ClearMainData()
With Me
    .txtIndentNo.Text = Empty
    .txtIDFNumber.Text = Empty
    .txtIDFDate.Text = Empty
    .cboAmmendPurp.Clear
    .cboAmmendType.Clear
    .cboCommodityDesc.Clear
    .txtCommodityCode.Text = Empty
    .txtNewDesc.Text = Empty
    .txtNewPackages.Text = Empty
    .txtOrigDesc.Text = Empty
    .txtPackageSize.Text = Empty
    .txtPackageType.Text = Empty
    .txtProductCode.Text = Empty
    .txtProductName.Text = Empty
    .txtQttyUnits.Text = Empty
    .txtSITC.Text = Empty
    .txtUnitsPerPack.Text = Empty
End With
End Sub

Private Sub Form_Resize()
On Error GoTo Err
With Me
    .Frame1.Height = .Height - (8505 - 7215)
    .Frame4.Height = .Height - (8505 - 4575)
    .ListView2.Height = .Height - (8505 - 4455)
    .txtNewDesc.Height = .Height - (8505 - 465)
    .Frame1.Width = .Width - (12000 - 11895)
    .ListView1.Width = .Width - (12000 - 4935)
    .ListView2.Width = .Width - (12000 - 4935)
    .Frame3.Left = .ListView1.Left + .ListView1.Width + 100
    .Frame4.Left = .Frame3.Left
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo Err
With Me
j = .ListView1.ListItems.Count
If j = 0 Or .ListView1.View <> lvwReport Then Item.Checked = False: Exit Sub
If Not NewRecord And Not EditRecord Then Item.Checked = False: Exit Sub
    If Item.Checked = True Then
        Call ClearMainData
        For i = 1 To j
        If .ListView1.ListItems(i).Text <> Item Then
            .ListView1.ListItems(i).Checked = False
        End If
        Next i
        .txtIndentNo.Text = Item
        .txtIDFNumber.Text = Item.SubItems(1)
        .txtIDFDate.Text = Item.SubItems(2)
        .txtAmmendNo.Text = GetNextAmmendmentNo
        .cboCommodityDesc.SetFocus
        Call ShowCommodityUnderIndent
    ElseIf Item.Checked = False Then
        Call ClearMainData
    End If
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub GetCommodityUnderIndent()
On Error GoTo Err
With Me

.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Product Code", .ListView2.Width / 3  ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Product/Brand Name", .ListView2.Width / 0.9
.ListView2.ColumnHeaders.Add , , "Full Description", .ListView2.Width / 0.9
.ListView2.ColumnHeaders.Add , , "Total Packages", .ListView2.Width / 3#  ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Package Type", .ListView2.Width / 3.5  ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Units/Pack", .ListView2.Width / 3.5  ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Unit Size", .ListView2.Width / 3.5  ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Qtty Units", .ListView2.Width / 3.5  ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "S.I.T.C.", .ListView2.Width / 3.5  ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Commodity Code", .ListView2.Width / 3
.ListView2.ColumnHeaders.Add , , "Country of Origin", .ListView2.Width / 3
.ListView2.ColumnHeaders.Add , , "Currency", .ListView2.Width / 3.5  ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Exch. Rate", .ListView2.Width / 3  ', lvwColumnCenter

.ListView2.View = lvwReport: .ListView2.Visible = True
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub ShowCommodityUnderIndent()
On Error GoTo Err
With Me
Screen.MousePointer = vbHourglass

.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Product Code", .ListView2.Width / 3  ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Product/Brand Name", .ListView2.Width / 0.9
.ListView2.ColumnHeaders.Add , , "Full Description", .ListView2.Width / 0.9
.ListView2.ColumnHeaders.Add , , "Total Packages", .ListView2.Width / 3#  ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Package Type", .ListView2.Width / 3.5  ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Units/Pack", .ListView2.Width / 3.5  ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Unit Size", .ListView2.Width / 3.5  ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Qtty Units", .ListView2.Width / 3.5  ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "S.I.T.C.", .ListView2.Width / 3.5  ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Commodity Code", .ListView2.Width / 3
.ListView2.ColumnHeaders.Add , , "Country of Origin", .ListView2.Width / 3
.ListView2.ColumnHeaders.Add , , "Currency", .ListView2.Width / 3.5  ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Exch. Rate", .ListView2.Width / 3  ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "FOB Value", .ListView2.Width / 3.5  ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Net Weight", .ListView2.Width / 3  ', lvwColumnCenter

.ListView2.View = lvwReport: .ListView2.Visible = True

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT ParamCommodityBrands.*,IDFCommoditiesInfo.* FROM IDFCommoditiesInfo,ParamCommodityBrands WHERE ParamCommodityBrands.ProductCode=IDFCommoditiesInfo.ProductCode AND IDFCommoditiesInfo.IndentNo='" & Trim(.txtIndentNo.Text) & "' AND IDFCommoditiesInfo.Ammended IS NULL ORDER BY IDFCommoditiesInfo.ProductCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem, NCount As Double

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Screen.MousePointer = vbDefault: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!productcode))
    If Not IsNull(rsLIST!ProductName) Then
        MyList.SubItems(1) = CStr(rsLIST!ProductName)
    End If
    If Not IsNull(rsLIST!fulldescription) Then
        MyList.SubItems(2) = CStr(rsLIST!fulldescription)
    End If
    If Not IsNull(rsLIST!TotalPackages) Then
        MyList.SubItems(3) = FormatNumber(rsLIST!TotalPackages)
    End If
    If Not IsNull(rsLIST!PackageType) Then
        MyList.SubItems(4) = CStr(rsLIST!PackageType)
    End If
    If Not IsNull(rsLIST!unitsperpack) Then
        MyList.SubItems(5) = CStr(rsLIST!unitsperpack)
    End If
    If Not IsNull(rsLIST!packagesize) Then
        MyList.SubItems(6) = CStr(rsLIST!packagesize)
    End If
    If Not IsNull(rsLIST!QttyUnits) Then
        MyList.SubItems(7) = CStr(rsLIST!QttyUnits)
    End If
    If Not IsNull(rsLIST!sitc) Then
        MyList.SubItems(8) = CStr(rsLIST!sitc)
    End If
    If Not IsNull(rsLIST!CommodityCode) Then
        MyList.SubItems(9) = CStr(rsLIST!CommodityCode)
    End If
    If Not IsNull(rsLIST!Country) Then
        MyList.SubItems(10) = CStr(rsLIST!Country)
    End If
    If Not IsNull(rsLIST!Currency) Then
        MyList.SubItems(11) = CStr(rsLIST!Currency)
    End If
    If Not IsNull(rsLIST!exchrate) Then
        MyList.SubItems(12) = FormatNumber(rsLIST!exchrate, 5, vbUseDefault, vbUseDefault, vbTrue)
    End If
    If Not IsNull(rsLIST!fobvalue) Then
        MyList.SubItems(13) = FormatNumber(rsLIST!fobvalue, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    If Not IsNull(rsLIST!netweight) Then
        MyList.SubItems(14) = FormatNumber(rsLIST!netweight, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    rsLIST.MoveNext
    
Wend

Set MyList = Nothing: Set rsLIST = Nothing: Screen.MousePointer = vbDefault
End With
Exit Sub
Err:
If Err.Number = 3265 Then
    Resume Next
Else
    Screen.MousePointer = vbDefault
    ErrorMessage
End If
End Sub

Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo Err
With Me
j = .ListView2.ListItems.Count
If j = 0 Or .ListView2.View <> lvwReport Then Item.Checked = False: Exit Sub
If Not NewRecord And Not EditRecord Then Item.Checked = False: Exit Sub
If .cboAmmendType.Text = Empty Or .cboAmmendPurp.Text = Empty Then
    Item.Checked = False: MsgBox "Required Ammendment Nature and Purpose...!", vbCritical + vbOKOnly, "Data Validation"
Else
    If Item.Checked = True Then
        Call ClearMainData1
        For i = 1 To j
        If .ListView2.ListItems(i).Text <> Item Then
            .ListView2.ListItems(i).Checked = False
        End If
        Next i
        .txtProductCode.Text = Item
        .txtProductName.Text = Item.SubItems(1)
        .txtOrigDesc.Text = Item.SubItems(2)
        .txtNewPackages.Text = Item.SubItems(3)
        .txtPackageType.Text = Item.SubItems(4)
        .txtUnitsPerPack.Text = Item.SubItems(5)
        .txtPackageSize.Text = Item.SubItems(6)
        .txtQttyUnits.Text = Item.SubItems(7)
        .txtSITC.Text = Item.SubItems(8)
        .txtCommodityCode = Item.SubItems(9)
        .txtCurrency = Item.SubItems(11)
        .txtExchRate = Item.SubItems(12)
        .txtFOBValue = Item.SubItems(13)
        .txtNetWeight = Item.SubItems(14)
        .txtNewPackages.SetFocus
        If .cboAmmendType.Text = "QTS" Or .cboAmmendType.Text = "SZE" Then
            Load frmIDFAmmendSizes
            With frmIDFAmmendSizes
                .txtCommodityCode.Text = Me.txtCommodityCode.Text
                .txtProductCode.Text = Me.txtProductCode.Text
                .txtSITC.Text = Me.txtSITC.Text
                .Show 1, Me
            End With
        End If
    ElseIf Item.Checked = False Then
        Call ClearMainData1
    End If
End If
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub mnuFileClear_Click()
    MyCommonData.ClearTheScreen
End Sub

Private Sub mnuFileClose_Click()
    Unload Me
End Sub

Private Sub mnuHelpUsing_Click()
    Call mnuHelpUsing_Click
End Sub

Private Sub mnuShowNewIndents_Click()
With Me
    Call ShowAllNewIndents
End With
End Sub

Private Sub GetAllNewIndents()
On Error GoTo Err
With Me

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Indent No", .ListView1.Width / 3.5
.ListView1.ColumnHeaders.Add , , "I.D.F. Number", .ListView1.Width / 3  ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "I.D.F. Date", .ListView1.Width / 3  ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Prof. Invoice No", .ListView1.Width / 2.5  ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Prof. Invoice Date", .ListView1.Width / 2.5 ', lvwColumnCenter

.ListView1.View = lvwReport: .ListView1.Visible = True
End With
Exit Sub
Err:
    ErrorMessage
End Sub
Private Sub ShowAllNewIndents()
On Error GoTo Err
With Me
Screen.MousePointer = vbHourglass

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Indent No", .ListView1.Width / 3.5
.ListView1.ColumnHeaders.Add , , "I.D.F. Number", .ListView1.Width / 3  ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "I.D.F. Date", .ListView1.Width / 3  ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Prof. Invoice No", .ListView1.Width / 2.5  ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Prof. Invoice Date", .ListView1.Width / 2.5 ', lvwColumnCenter

.ListView1.View = lvwReport: .ListView1.Visible = True

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT IDFMainData.*,ProformaInvoiceRequest.ProfInvoiceNo,ProformaInvoiceRequest.ProfInvoiceDate FROM IDFMainData,ProformaInvoiceRequest WHERE ProformaInvoiceRequest.IndentNo=IDFMainData.IndentNo AND IDFMainData.Received IS NOT NULL AND IDFMainData.Ammended IS NULL ORDER BY IDFMainData.IndentNo;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem, NCount As Double

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Screen.MousePointer = vbDefault: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!indentno))
    If Not IsNull(rsLIST!idfnumber) Then
        MyList.SubItems(1) = CStr(rsLIST!idfnumber)
    End If
    If Not IsNull(rsLIST!idfdate) Then
        MyList.SubItems(2) = CStr(rsLIST!idfdate)
    End If
    If Not IsNull(rsLIST!ProfInvoiceNo) Then
        MyList.SubItems(3) = CStr(rsLIST!ProfInvoiceNo)
    End If
    If Not IsNull(rsLIST!ProfInvoiceDate) Then
        MyList.SubItems(4) = CStr(rsLIST!ProfInvoiceDate)
    End If
    rsLIST.MoveNext
    
Wend

Set MyList = Nothing: Set rsLIST = Nothing: Screen.MousePointer = vbDefault
        
End With
Exit Sub
Err:
If Err.Number = 3265 Then
    Resume Next
Else
    Screen.MousePointer = vbDefault
    ErrorMessage
End If
End Sub

Private Function ValidRecord() As Boolean
With Me
    For Each i In Me
    If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
        If i.Text = Empty Then
            MsgBox "One or More Required Fields Empty. Please Enter the Required Data...!", vbCritical + vbOKOnly, "Data Validation"
            i.SetFocus: ValidRecord = False: Exit Function
        End If
    End If
    Next i
    ValidRecord = True
End With
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err
With Me
Dim SQL1, SQL2
Select Case Button.Key
Case "N"
    Select Case Button.Caption
    Case "New &Record "
        If EditRecord Then Exit Sub
        Call ClearMainData1: .txtCommodityCode.SetFocus
        NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4
        If .ListView1.ListItems.Count = 0 Or .ListView1.View <> lvwReport Then
            Call ShowAllNewIndents
        End If
        If .txtIndentNo.Text = Empty Then
            .ListView1.SetFocus
        Else
            If .cboCommodityDesc.Text = Empty Then
                .cboCommodityDesc.SetFocus
            Else
                If .cboAmmendType.Text = Empty Then
                    .cboAmmendType.SetFocus
                Else
                    If .cboAmmendPurp.Text = Empty Then
                        .cboAmmendPurp.SetFocus
                    Else
                        If .txtProductCode.Text = Empty Then
                            .ListView2.SetFocus
                        Else
                            .txtNewPackages.SetFocus
                        End If
                    End If
                End If
            End If
        End If
    Case "&Save Record "
        If NewRecord Then
        If ValidRecord Then
            .txtAmmendNo.Text = GetNextAmmendmentNo
            'update idfcommodities information with ammended/new data
            j = .ListView2.ListItems.Count
            For i = 1 To j
            If .ListView2.ListItems(i).Checked = True Then
                EditSQL = "UPDATE IDFCommoditiesInfo SET Ammended='" & 1 & "',Currency='" & Trim(.txtCurrency.Text) & "',TotalPackages=" & CDbl(.txtNewPackages.Text) & ",ExchRate=" & CDbl(.txtExchRate.Text) & ",FOBValue=" & CDbl(.txtFOBValue.Text) & ",NetWeight=" & CDbl(.txtNetWeight.Text) & ",PackageType='" & Trim(.txtPackageType.Text) & "',UnitsPerPack=" & CDbl(.txtUnitsPerPack.Text) & ",PackageSize=" & CDbl(.txtPackageSize.Text) & ",QttyUnits='" & Trim(.txtQttyUnits.Text) & "',FullDescription='" & Trim(.txtNewDesc.Text) & "' WHERE IndentNo='" & Trim(.txtIndentNo.Text) & "' AND ProductCode='" & Trim(.ListView2.ListItems(i).Text) & "';"
                Set rsEditRecord = New ADODB.Recordset
                rsEditRecord.Open EditSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                Set rsEditRecord = Nothing
            End If
            Next i
            
            'copy original record into ammendment data file
            For i = 1 To j
            If .ListView2.ListItems(i).Checked = True Then
                SQL1 = "INSERT INTO IDFAmmendData(IndentNo,AmmendNo,ProductCode,PackageType,TotalPackages,UnitsPerPack,PackageSize,QttyUnits,FullDescription,Currency,ExchRate,FOBValue,NetWeight,CreatedBy,DateCreated,AccPeriod)"
                SQL2 = "VALUES('" & Trim(.txtIndentNo.Text) & "'," & CLng(.txtAmmendNo.Text) & ",'" & Trim(.ListView2.ListItems(i).Text) & "','" & Trim(.ListView2.ListItems(i).SubItems(4)) & "'," & CDbl(.ListView2.ListItems(i).SubItems(3)) & "," & CDbl(.ListView2.ListItems(i).SubItems(5)) & "," & CDbl(.ListView2.ListItems(i).SubItems(6)) & ",'" & Trim(.ListView2.ListItems(i).SubItems(7)) & "','" & Trim(.ListView2.ListItems(i).SubItems(2)) & "','" & Trim(.ListView2.ListItems(i).SubItems(11)) & "'," & CDbl(.ListView2.ListItems(i).SubItems(12)) & "," & CDbl(.ListView2.ListItems(i).SubItems(13)) & "," & CDbl(.ListView2.ListItems(i).SubItems(14)) & ",'" & Trim(CurrentUserName) & "','" & Trim(MyCurrentDate) & "','" & Trim(MyCurrentPeriod) & "');"
                NewSQL = SQL1 + SQL2
                Set rsNewRecord = New ADODB.Recordset
                rsNewRecord.Open NewSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                Set rsNewRecord = Nothing
            End If
            Next i
            
            'save main record if not yet saved
            If Not MainDataExists Then
                SQL1 = "INSERT INTO IDFAmmendMain(IndentNo,AmmendNo,AmmendType,DateAmmended,CommodityDesc,AmmendPurpose,StaffIDNo,CreatedBy,DateCreated,AccPeriod)"
                SQL2 = "VALUES('" & Trim(.txtIndentNo.Text) & "'," & CLng(.txtAmmendNo.Text) & ",'" & Trim(.cboAmmendType.Text) & "','" & Trim(MyCurrentDate) & "','" & Trim(.cboCommodityDesc.Text) & "','" & Trim(.cboAmmendPurp.Text) & "','" & Trim(GetCurrentStaffID) & "','" & Trim(CurrentUserName) & "','" & Trim(MyCurrentDate) & "','" & Trim(MyCurrentPeriod) & "');"
                NewSQL = SQL1 + SQL2
                Set rsNewRecord = New ADODB.Recordset
                rsNewRecord.Open NewSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                Set rsNewRecord = Nothing
            End If
            
            Call Remove2SelectedItems
            If .ListView2.ListItems.Count = 0 Then
                Call RemoveSelectedItems
                Set rsLineUpdate = New ADODB.Recordset
                rsLineUpdate.Open "UPDATE IDFMainData SET Ammended='" & 1 & "' WHERE IndentNo='" & Trim(.txtIndentNo.Text) & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
                Set rsLineUpdate = Nothing
            End If
            
            NewRecord = False: Button.Caption = "New &Record ": Button.Image = 2
            
        End If
        End If
    Case Else
        Exit Sub
    End Select
Case "E"
    Select Case Button.Caption
    Case "&Edit/Change "
    If NewRecord Then Exit Sub
        If .txtIndentNo.Text = Empty Then
            MsgBox "There is NO Current Record to Edit. Please Search and Display a Record First...!", vbCritical + vbOKOnly, ""
            .txtIndentNo.SetFocus
        Else
            .txtIndentNo.Locked = True
            Button.Caption = "Save &Changes ": Button.Image = 4
            EditRecord = True
        End If
    Case "Save &Changes "
        If EditRecord Then
        If ValidRecord Then
'                EditSQL = "UPDATE IDFAmmendMain SET IndentDate='" & Format(.txtIndentDate.Text, "MMMM dd,yyyy") & "',DateApplied='" & Format(.txtDateApplied.Text, "MMMM dd,yyyy") & "',ImporterCode='" & Trim(.cboImporterCode.Text) & "',ExporterCode='" & Trim(.cboExporterCode.Text) & "',PSIConfirm='" & Trim(.chkConfirmation.Value) & "',PSIAgency='" & Trim(.cboPSIAgency.Text) & "',InterventionCode='" & Trim(.txtInterventionCode.Text) & "',Prepaidamount=" & CDbl(txtPrePaidAmount.Text) & ",gokfee=" & CDbl(.txtGOKProcessingFee.Text) & ",receiptno='" & Trim(.txtReceiptNumber.Text) & "',SerialNo='" & Trim(.txtSerialNumber.Text) & "',ccrfno='" & Trim(.txtCCRFNumber.Text) & "',ChequeNo='" & Trim(.txtChecqueNumber.Text) & "',PriorApproval='" & Trim(.chkApproval.Value) & "' WHERE IndentNo='" & Trim(.txtIndentNo.Text) & "';"
            Set rsEditRecord = New ADODB.Recordset
            rsEditRecord.Open EditSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsEditRecord = Nothing
            .txtIndentNo.Locked = False: EditRecord = False: Button.Caption = "&Edit/Change ": Button.Image = 5
        End If
        End If
    Case Else
        Exit Sub
    End Select
Case "S"
    If NewRecord Or EditRecord Then Exit Sub
        INPQRY = InputBox("Please Enter the Indent Number for the Record to Search and Display...!!!", "Search for Record...")
        If Len(INPQRY) = 0 Then
            MsgBox "Required Search Parameter Missing or the Operation Was Cancelled...! No Work was Done!!!", vbCritical + vbOKOnly, "Missing Parameter"
            Exit Sub
        Else
            Set rsFindRecord = cnCOMMON.Execute("SELECT IDFAmmendMain.* FROM IDFAmmendMain WHERE IDFAmmendMain.IndentNo='" & Trim(INPQRY) & "';")
            If rsFindRecord.EOF And rsFindRecord.BOF Then
                MsgBox "Requested Record Missing or Has Been Deleted. Check your Entries to Ensure they are Accurately Spelt...!", vbOKOnly + vbExclamation, "Record NOT Found...!"
                Set rsFindRecord = Nothing: Exit Sub
            Else
                .txtIndentNo.Text = Trim(rsFindRecord!indentno & "")
            End If
            Set rsFindRecord = Nothing
        End If
Case "R"
    If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
        .Toolbar1.Buttons(2).Caption = "New &Record "
        .Toolbar1.Buttons(2).Image = 2
        .Toolbar1.Buttons(3).Caption = "&Edit/Change "
        .Toolbar1.Buttons(3).Image = 5: NewRecord = False: EditRecord = False: MyCommonData.ClearTheScreen
Case "P"
    If .txtIndentNo.Text = Empty Then
        MsgBox "Required Indent Number. Please Search and Display and Indent Before Printing the Import Declaration Form...!", vbCritical + vbOKOnly, "Data Validation"
        .txtIndentNo.SetFocus
    Else
        Load frmRptIDFAmmendment
        frmRptIDFAmmendment.Show 1, Me
    End If
Case "H"
    Call mnuHelpUsing_Click
Case Else
    Exit Sub
End Select
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Function MainDataExists() As Boolean
On Error GoTo Err
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT COUNT(IDFAmmendMain.IndentNo) AS TOTAL FROM IDFAmmendMain WHERE IndentNo='" & Trim(.txtIndentNo.Text) & "';")
    If Not NewRecord And Not EditRecord Then
        MainDataExists = False
    ElseIf IsNull(rsFindRecord!Total) = True Then
        MainDataExists = False
    ElseIf rsFindRecord!Total = 0 Then
        MainDataExists = False
    ElseIf rsFindRecord!Total >= 1 Then
        MainDataExists = True
    End If
    Set rsFindRecord = Nothing
End With
Exit Function
Err:
    ErrorMessage
End Function

Private Sub txtNewPackages_Change()
If Not NewRecord And Not EditRecord Then Exit Sub
With Me
    If .txtNewPackages.Text = Empty Then
        .txtNewDesc.Text = Empty
    Else
        .txtNewDesc.Text = FormatNumber(.txtNewPackages.Text, 0, vbUseDefault, vbUseDefault, vbTrue) & " " & Trim(.txtPackageType.Text) & " OF " & Trim(.txtProductName.Text) & " " & Trim(.txtUnitsPerPack.Text) & "/" & Trim(.txtPackageSize.Text) & Trim(.txtQttyUnits.Text)
    End If
End With
End Sub

Private Sub txtNewPackages_GotFocus()
If Not NewRecord And Not EditRecord Then Exit Sub
With Me
    If .txtNewPackages.Text = Empty Then
        .txtNewDesc.Text = Empty
    Else
        .txtNewDesc.Text = FormatNumber(.txtNewPackages.Text, 0, vbUseDefault, vbUseDefault, vbTrue) & " " & Trim(.txtPackageType.Text) & " OF " & Trim(.txtProductName.Text) & " " & Trim(.txtUnitsPerPack.Text) & "/" & Trim(.txtPackageSize.Text) & Trim(.txtQttyUnits.Text)
    End If
End With
End Sub

Private Sub txtNewPackages_KeyPress(KeyAscii As Integer)
If Not NewRecord And Not EditRecord Then KeyAscii = 0
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyDecpt Or KeyAscii = vbKeyBack Then
        Exit Sub
    Else
        MsgBox "Enter numeric values only", vbOKOnly + vbCritical, "Data Validation"
        KeyAscii = 0
        Beep
    End If
End Sub
