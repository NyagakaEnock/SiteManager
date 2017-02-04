VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRsContractRequisition 
   Caption         =   "Resource Scheduling-CONTRACT REQUISITION "
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10365
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRsContractRequisition.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   Picture         =   "frmRsContractRequisition.frx":0442
   ScaleHeight     =   7815
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   585
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   41
      Text            =   "***  CONTRACT REQUISITION  ***"
      Top             =   0
      Width           =   11895
   End
   Begin VB.CheckBox chkOrderNo 
      Height          =   255
      Left            =   0
      TabIndex        =   35
      Top             =   600
      Value           =   1  'Checked
      Width           =   205
   End
   Begin VB.TextBox txtOrderDescription 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   4680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   840
      Width           =   7215
   End
   Begin VB.TextBox txtOrderNO 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   15
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtOrderDate 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   6240
      TabIndex        =   8
      Top             =   1560
      Width           =   5535
      Begin VB.ComboBox cboAdvCode 
         BackColor       =   &H00FFC0C0&
         Height          =   345
         Left            =   1320
         TabIndex        =   52
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   3960
         TabIndex        =   50
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtDuration 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   330
         Left            =   1320
         TabIndex        =   48
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox txtWidth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   3960
         TabIndex        =   46
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txtType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   330
         Left            =   3240
         TabIndex        =   42
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txtLenght 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   1320
         TabIndex        =   40
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox txtDataSource 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   4440
         Width           =   4095
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H80000000&
         Caption         =   "&REFRESH"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5160
         Width           =   1695
      End
      Begin VB.CommandButton cmdFinish 
         BackColor       =   &H80000000&
         Caption         =   "< FINIS&H >"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5160
         Width           =   1815
      End
      Begin VB.CommandButton cmdSAVE 
         BackColor       =   &H80000000&
         Caption         =   "ADD &NEW"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5160
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dtpPromised 
         Height          =   375
         Left            =   5175
         TabIndex        =   34
         Top             =   1440
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Format          =   64552961
         CurrentDate     =   37965
      End
      Begin MSComCtl2.DTPicker dtpRequired 
         Height          =   375
         Left            =   2775
         TabIndex        =   33
         Top             =   1440
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Format          =   64552961
         CurrentDate     =   37965
      End
      Begin VB.TextBox txtTotalCost 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   4080
         Width           =   4095
      End
      Begin VB.TextBox txtAdvCost 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3720
         Width           =   4095
      End
      Begin VB.ComboBox cboAdvName 
         BackColor       =   &H00FFC0C0&
         Height          =   345
         Left            =   1320
         TabIndex        =   11
         Top             =   2040
         Width           =   4095
      End
      Begin VB.TextBox txtDatePromised 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtDateRequired 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtPhysicalAddress 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox txtContactPerson 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   2295
      End
      Begin VB.ComboBox cboClientId 
         BackColor       =   &H00FFC0C0&
         Height          =   345
         Left            =   1440
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cboClientName 
         BackColor       =   &H00FFC0C0&
         Height          =   345
         Left            =   1440
         TabIndex        =   5
         Top             =   120
         Width           =   3975
      End
      Begin VB.Label Label15 
         Caption         =   "Total"
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Days"
         Height          =   375
         Left            =   3240
         TabIndex        =   49
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "Duration"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Width"
         Height          =   255
         Left            =   3360
         TabIndex        =   45
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   44
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "Type"
         Height          =   375
         Left            =   2760
         TabIndex        =   43
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Length"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Adv.  Code"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000002&
         BorderWidth     =   2
         X1              =   120
         X2              =   5400
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Label Label14 
         Caption         =   "Adv.  Cost"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Adv. Name"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000002&
         BorderWidth     =   2
         X1              =   120
         X2              =   5400
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label10 
         Caption         =   "End Date"
         Height          =   255
         Left            =   3120
         TabIndex        =   30
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Contr. Start Date"
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Physical Addr."
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1066
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Client ID."
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Client Name"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3600
      Left            =   0
      TabIndex        =   3
      Top             =   3840
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6350
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
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
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   7440
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   4657
            MinWidth        =   4657
            Picture         =   "frmRsContractRequisition.frx":5B80
            Text            =   "CONTRACTS REQ."
            TextSave        =   "CONTRACTS REQ."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   13829
            MinWidth        =   13829
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2893
            MinWidth        =   2893
            TextSave        =   "25/10/2004"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   635
      ButtonWidth     =   1138
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2040
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3598
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
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
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComCtl2.DTPicker dtpOrderDate 
      Height          =   375
      Left            =   4335
      TabIndex        =   26
      Top             =   840
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Format          =   64552961
      CurrentDate     =   37965
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   0
      X2              =   11880
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Re 
      BackStyle       =   0  'Transparent
      Caption         =   "Contract  Number"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Contract Description"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4680
      TabIndex        =   28
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Order Date"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2400
      TabIndex        =   27
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contract Particulars"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6240
      TabIndex        =   21
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "List of Current Registered Clients"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   1320
      Width           =   5295
   End
   Begin VB.Label lblInventory 
      BackStyle       =   0  'Transparent
      Caption         =   "List of Billboard Items"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   3600
      Width           =   5295
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileClearScreen 
         Caption         =   "Clear &Screen"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnujjjj 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuShow 
      Caption         =   "&Show/View"
      Begin VB.Menu mnuShowCurrentPurchaseOrders 
         Caption         =   "Show Pending Contract Requisitions"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu mnuDFd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowBillBoardItems 
         Caption         =   "Bill Board Items"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnujjnh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowRegisteredClients 
         Caption         =   "Current List of &Registered Clients"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Print Preview"
      Begin VB.Menu mnuPrintOrderForm 
         Caption         =   "Contract Requisiton Form"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help System"
      Begin VB.Menu mnuHelpUsing 
         Caption         =   "Using the &Point-Of-Sale"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmRsContractRequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyContracts As clsRsContractRequisition, OrderTotalSum, Ordate, PromDate, ReqDate

Private Sub cboAdvCode_Click()
Me.ListView2.SetFocus
End Sub

Private Sub cboAdvCode_GotFocus()
If Not NewRecord And Not EditRecord Then Exit Sub
 'On Error GoTo Err
With frmRsContractRequisition

If .cboAdvCode.ListCount <> 0 Then Exit Sub
    
    Set rsCOMBO = cnCOMMON.Execute("SELECT BillBoardNo FROM AdvertBBDetails WHERE BillBoardNo IS NOT NULL AND CostPriceStatus = '" & "Y" & "' AND Discontinued = '" & "N" & "' ORDER BY Name;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!BillBoardNo) And rsCOMBO!BillBoardNo <> "" Then
            .cboAdvCode.AddItem rsCOMBO!BillBoardNo
        End If
    rsCOMBO.MoveNext
    Loop
    
OUTS:
    Set rsCOMBO = Nothing
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub cboAdvCode_LostFocus()
'On Error GoTo Err
With frmRsContractRequisition

    If .cboAdvCode.Text = Empty Then Exit Sub
    
    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM AdvertBBDetails  WHERE BillBoardNo='" & Trim(.cboAdvCode.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .cboAdvName.Text = rsFindRecord!Name & ""
        .cboAdvCode.Text = rsFindRecord!BillBoardNo & ""
        .txtType.Text = rsFindRecord!TypeName & ""
        .txtLenght.Text = rsFindRecord!Length & ""
        .txtWidth.Text = rsFindRecord!Width & ""
        .txtDuration.Text = rsFindRecord!DurationName & ""
        .txtDays.Text = rsFindRecord!NoOfDays & ""
        .txtAdvCost.Text = GetBillBoardCost
        .txtTotalCost.Text = .txtAdvCost.Text
        
        
    End If
    
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub cboAdvName_Click()
Me.ListView2.SetFocus
End Sub

Private Sub cboAdvName_GotFocus()
If Not NewRecord And Not EditRecord Then Exit Sub
 'On Error GoTo Err
With frmRsContractRequisition

If .cboAdvName.ListCount <> 0 Then Exit Sub
    
    Set rsCOMBO = cnCOMMON.Execute("SELECT Name FROM AdvertBBDetails WHERE BillBoardNo IS NOT NULL AND CostPriceStatus = '" & "Y" & "' AND Discontinued = '" & "N" & "' ORDER BY Name;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!Name) And rsCOMBO!Name <> "" Then
            .cboAdvName.AddItem rsCOMBO!Name
        End If
    rsCOMBO.MoveNext
    Loop
    
OUTS:
    Set rsCOMBO = Nothing
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub cboAdvName_LostFocus()
'On Error GoTo Err
With frmRsContractRequisition

    If .cboAdvName.Text = Empty Then Exit Sub
    
    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM AdvertBBDetails  WHERE Name='" & Trim(.cboAdvName.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .cboAdvName.Text = rsFindRecord!Name & ""
        .cboAdvCode.Text = rsFindRecord!BillBoardNo & ""
        .txtType.Text = rsFindRecord!TypeName & ""
        .txtLenght.Text = rsFindRecord!Length & ""
        .txtWidth.Text = rsFindRecord!Width & ""
        .txtDuration.Text = rsFindRecord!DurationName & ""
        .txtDays.Text = rsFindRecord!NoOfDays & ""
        .txtAdvCost.Text = GetBillBoardCost
        .txtTotalCost.Text = .txtAdvCost.Text
        
        
    End If
    
End With
Exit Sub
Err:
    ErrorMessage
End Sub
Private Function GetBillBoardCost() As Variant
'On Error GoTo Err
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT BBCharges FROM AdvertPricing  WHERE BBNo = '" & Trim(.cboAdvCode.Text) & "' AND CostPriceStatus = '" & "Y" & "' ;")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetBillBoardCost = 0: Set rsFindRecord = Nothing
    ElseIf IsNull(rsFindRecord!BBCharges) = True Or rsFindRecord!BBCharges = "" Then
        GetBillBoardCost = 0: Set rsFindRecord = Nothing
    Else
        GetBillBoardCost = CDbl(rsFindRecord!BBCharges)
    End If
        
    Set rsFindRecord = Nothing
Exit Function
Err:
    ErrorMessage
    End With
End Function
Private Sub cboClientId_Click()
Me.ListView1.SetFocus
End Sub

Private Sub cboClientId_GotFocus()

If Not NewRecord And Not EditRecord Then Exit Sub
    'On Error GoTo Err
With frmRsContractRequisition

If .cboClientId.ListCount <> 0 Then Exit Sub
    
    Set rsCOMBO = cnCOMMON.Execute("SELECT CustomerId FROM AdvertClients WHERE CustomerId IS NOT NULL ORDER BY CustomerId;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!CustomerId) And rsCOMBO!CustomerId <> "" Then
            .cboClientId.AddItem rsCOMBO!CustomerId
        End If
    rsCOMBO.MoveNext
    Loop
    
OUTS:
    Set rsCOMBO = Nothing
End With
Exit Sub
Err:
    ErrorMessage
End Sub


Private Sub cboClientId_LostFocus()
'On Error GoTo Err
With frmRsContractRequisition

    If .cboClientId.Text = Empty Then Exit Sub
    
    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM AdvertClients WHERE CustomerId='" & Trim(.cboClientId.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .cboClientName.Text = rsFindRecord!CompanyName & ""
        .cboClientId.Text = rsFindRecord!CustomerId & ""
        .txtContactPerson.Text = rsFindRecord!ContactName & ""
        .txtPhysicalAddress.Text = rsFindRecord!PhysicalAddress & ""
        
        .txtDateRequired.SetFocus
    End If
    
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub cboClientName_Click()
Me.ListView1.SetFocus
End Sub

Private Sub cboClientName_GotFocus()
If Not NewRecord And Not EditRecord Then Exit Sub
    'On Error GoTo Err
With frmRsContractRequisition

If .cboClientName.ListCount <> 0 Then Exit Sub
    
    Set rsCOMBO = cnCOMMON.Execute("SELECT CompanyName FROM AdvertClients WHERE CompanyName IS NOT NULL ORDER BY CompanyName;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!CompanyName) And rsCOMBO!CompanyName <> "" Then
            .cboClientName.AddItem rsCOMBO!CompanyName
        End If
    rsCOMBO.MoveNext
    Loop
    
OUTS:
    Set rsCOMBO = Nothing
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub cboClientName_LostFocus()
'On Error GoTo Err
With frmRsContractRequisition

    If .cboClientName.Text = Empty Then Exit Sub
    
    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM AdvertClients WHERE CompanyName='" & Trim(.cboClientName.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .cboClientName.Text = rsFindRecord!CompanyName & ""
        .cboClientId.Text = rsFindRecord!CustomerId & ""
        .txtContactPerson.Text = rsFindRecord!ContactName & ""
        .txtPhysicalAddress.Text = rsFindRecord!PhysicalAddress & ""
        
        .txtDateRequired.SetFocus
    End If
    
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub cmdFinish_Click()

'On Error GoTo Err
With frmRsContractRequisition
    
    Ordate = Format(.txtOrderDate.Text, "MMMM dd,yyyy")
    ReqDate = Format(.txtDateRequired.Text, "MMMM dd,yyyy")
    PromDate = Format(.txtDatePromised.Text, "MMMM dd,yyyy")
    If ValidMainPurchase Then
    Set rsNewRecord = New ADODB.Recordset
    
    MySQL = "INSERT INTO AdvertContractRequisition (ClientName,PurchaseOrderNo,ClientCode,OrderDescription,ContactPerson,PhysicalAddress,StartDate,EndDate,Createdby,datecreated,accperiod) VALUES('" & Trim(.cboClientName.Text) & "','" & Trim(.txtOrderNO) & "','" & Trim(.cboClientId.Text) & "','" & Trim(.txtOrderDescription) & "','" & Trim(.txtContactPerson.Text) & "','" & Trim(.txtPhysicalAddress.Text) & "','" & Trim(ReqDate) & "','" & Trim(PromDate) & "','" & CurrentUserName & "','" & MyCurrentDate & "','" & MyCurrentPeriod & "')"
    
    rsNewRecord.Open MySQL, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    Set rsNewRecord = Nothing
    End If
    If MsgBox("Current Contract Data Entry Successfully Completed!! Do you want to Print the Contract Form?", vbYesNo + vbQuestion + vbDefaultButton1, "Print Contract Form?") = vbYes Then
        
        If .txtDataSource.Text = "CONTRACT RENEWAL" Then
            Call ShowRenewalContracts
        ElseIf .txtDataSource.Text = "NEW CONTRACT" Then
            Call mnuShowBillBoardItems_Click
        End If
        
        Load frmRPTContractRequisitionForm
        frmRPTContractRequisitionForm.Show 1, Me
        
    Else
    
        If .txtDataSource.Text = "CONTRACT RENEWAL" Then
            Call ShowRenewalContracts
        ElseIf .txtDataSource.Text = "NEW CONTRACT" Then
            Call mnuShowBillBoardItems_Click
        End If
        
       Call GetOrderTotalSum
       OrderTotalSum = GetOrderTotalSum
       .txtTotalCost.Text = OrderTotalSum
       
       Set rsEditRecord = New ADODB.Recordset
       rsEditRecord.Open "UPDATE AdvertContractRequisition SET TotalCost = " & OrderTotalSum & " WHERE PurchaseOrderNo = '" & .txtOrderNO.Text & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
       Set rsEditRecord = Nothing
       
       

    End If
    .cmdSAVE.Caption = "ADD &NEW"
    .cmdFinish.Enabled = False
    NewRecord = False
    
End With

Exit Sub
Err:
    UpdateErrorMessage
End Sub
Private Function GetOrderTotalSum() As Currency
'On Error GoTo Err
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT Sum(AdvCost) As Total FROM AdvertContractRequisitionData WHERE PurchaseOrderno = '" & Trim(.txtOrderNO.Text) & "'")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetOrderTotalSum = Empty: Exit Function
    ElseIf IsNull(rsFindRecord!Total) = True Or rsFindRecord!Total = "" Then
        GetOrderTotalSum = Empty: Exit Function
    Else
        GetOrderTotalSum = FormatNumber(rsFindRecord!Total, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
          
    Set rsFindRecord = Nothing
    
End With
Exit Function
Err:
    ErrorMessage
End Function


Private Sub cmdRefresh_Click()
'On Error GoTo Err
With Me

    If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Screen Refresher") = vbCancel Then Exit Sub
        NewRecord = False: EditRecord = False
      Call ClearTextFields
      Call ClearDataSheets
        
        .cmdSAVE.Enabled = True
        .cmdSAVE.Caption = "ADD &NEW"
        .cmdFinish.Enabled = False
        .cmdRefresh.Enabled = True
        
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub cmdSAVE_Click()
'On Error GoTo Err
With Me
'If EditRecord Then Exit Sub
If CurrentUserName = "administrator" Then
    MsgBox "SORRY!! You cannot Add Purchase Order Data when Logged-On as System Administrator! Please Log Out and Use a Registered Staff Member's Profile!!!", vbCritical + vbOKOnly, "Wrong Profile"
    Exit Sub
Else
    Select Case .cmdSAVE.Caption
    Case "ADD &NEW"
        
        .txtDataSource.Text = "CONTRACT RENEWAL"
        NewRecord = True
        
        If .txtDataSource.Text = "CONTRACT RENEWAL" Or .txtDataSource.Text = Empty Then
            Call ShowRenewalContracts: Call ShowClients
        ElseIf .txtDataSource.Text = "NEW CONTRACT" Then
            Call mnuShowBillBoardItems_Click: Call ShowClients
        End If
        
        MyContracts.AddNewRecord
        Call GetStaffIdNo
        
    Case "&SAVE RECORD"
    
        If NewRecord Then
            If ValidRecord Then
            
                MyContracts.SavePurchaseData
                Call RemoveCurrentListItem
                
            End If
        End If
        
    Case "&NEXT ITEM"
          
           NewRecord = True
           If EditRecord Then Exit Sub
           Call ClearForNewBBItem
           .cmdSAVE.Caption = "&SAVE RECORD"
           
        
    Case Else
        Exit Sub
    End Select
End If
End With
Exit Sub
Err:
    ErrorMessage
End Sub
Public Sub ShowRenewalContracts()
'On Error GoTo Err
Dim Today As Variant
Today = Format(Date, "MMMM dd,yy")
With frmRsContractRequisition
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Advert Code ", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Name", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Type", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Lenght", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Width", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Duration", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Days", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Client Code", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Client Name", .ListView1.Width / 4


.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertContractRequisition A,AdvertContractRequisitionData B WHERE A.PurchaseOrderNo = B.PurchaseOrderNo AND A.RenewalApprovalStatus = '" & "Y" & "' AND B.RenewalApprovalStatus = '" & "Y" & "' ORDER BY B.PurchaseOrderNo ;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!AdvCode))

    If Not IsNull(rsLIST!AdvName) Then
        MyList.SubItems(1) = CStr(rsLIST!AdvName)
    End If
    
    If Not IsNull(rsLIST!AdvType) Then
        MyList.SubItems(2) = Trim(CStr(rsLIST!AdvType))
        End If
        
    If Not IsNull(rsLIST!AdvLength) Then
        MyList.SubItems(3) = Trim(CStr(rsLIST!AdvLength))
        End If
    
    If Not IsNull(rsLIST!AdvWidth) Then
        MyList.SubItems(4) = CStr(rsLIST!AdvWidth)
    End If
    
    If Not IsNull(rsLIST!Duration) Then
        MyList.SubItems(5) = CStr(rsLIST!Duration)
    End If
    
    If Not IsNull(rsLIST!Days) Then
        MyList.SubItems(6) = Trim(CStr(rsLIST!Days))
    End If
    
    If Not IsNull(rsLIST!ClientCode) Then
        MyList.SubItems(7) = Trim(CStr(rsLIST!ClientCode))
    End If
    
    If Not IsNull(rsLIST!ClientName) Then
        MyList.SubItems(8) = Trim(CStr(rsLIST!ClientName))
    End If
    
    rsLIST.MoveNext
    
Wend

.txtDataSource.Text = "CONTRACT RENEWAL"
Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub


Private Function ValidMainPurchase() As Boolean
'On Error GoTo Err
With Me
    If .txtOrderNO.Text = Empty Then
        strMessage = "Required Purchase Order Number!"
        .txtOrderNO.SetFocus
    ElseIf .txtOrderDate.Text = Empty Then
        strMessage = "Required Purchase Order Date!!"
        .txtOrderDate.SetFocus
    ElseIf .cboClientId.Text = Empty Then
        strMessage = "Required Client ID!!"
        .cboClientId.SetFocus
    ElseIf .cboClientName.Text = Empty Then
        strMessage = "Required Client Name!!"
        .cboClientName.SetFocus
    ElseIf .cboAdvName.Text = Empty Then
        strMessage = "Required Advertisement Name!!"
        .cboAdvName.SetFocus
    ElseIf .txtDatePromised.Text = Empty Then
        strMessage = "Required Promised Date!!"
        .txtDatePromised.SetFocus
    ElseIf .txtDateRequired.Text = Empty Then
        strMessage = "Please enter the date the order is required!!"
        .txtDateRequired.SetFocus
      
    ElseIf .txtAdvCost.Text = Empty Then
        strMessage = "Required Advert Cost!!"
        .txtAdvCost.SetFocus
    Else
        ValidMainPurchase = True
    End If
    If Not ValidMainPurchase Then
        MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
    End If
End With
Exit Function
Err:
    ErrorMessage
End Function
Private Sub ClearTextFields()
For Each i In Screen.ActiveForm
    If TypeOf i Is TextBox And i.Name <> "txtTitle" Then
        i.Text = Empty
    End If
    If TypeOf i Is ComboBox Then
        i.Clear
    End If
    If TypeOf i Is Image Then
        i.Picture = LoadPicture("")
    End If
Next i
End Sub

Private Sub ClearDataSheets()
For Each i In Screen.ActiveForm
    If TypeOf i Is ListView Then
        i.ListItems.Clear
    End If
Next i
End Sub
Private Function ValidRecord() As Boolean
With Me
    For Each i In Me
    If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
        If i.Text = Empty And i.Name <> "txtTotalCost" And i.Name <> "txtOrderDescription" Then
            MsgBox "All the Fields are Required. Please Enter the Missing Data...!", vbCritical + vbOKOnly, "Data Validation"
            i.SetFocus: ValidRecord = False: Exit Function
        End If
    End If
    Next i
    ValidRecord = True
End With
End Function
Private Function RecordExists() As Boolean
'On Error GoTo Err
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT COUNT(SerialNO) AS TCount FROM ContractRequisitionData WHERE AdvCode='" & Trim(.cboAdvCode.Text) & "' AND PurchaseOrderNo='" & Trim(.txtOrderNO.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        RecordExists = False: Set rsFindRecord = Nothing
    ElseIf rsFindRecord!tcount = 0 Then
        RecordExists = False: Set rsFindRecord = Nothing
    ElseIf rsFindRecord!tcount >= 1 Then
        RecordExists = True: Set rsFindRecord = Nothing
    Else
        RecordExists = False
    End If
    
End With
Exit Function
Err:
    ErrorMessage
End Function

Private Sub dtpOrderDate_CloseUp()
If Not NewRecord And Not EditRecord Then Exit Sub
    If Me.dtpOrderDate.Value > Date Then
        MsgBox "Wrong Date! Order Date Cannot be Later than the Date Today!!", vbCritical + vbOKOnly, "Invalid Date"
        Me.txtOrderDate.Text = Empty: Me.txtOrderDate.SetFocus
    Else
        Me.txtOrderDate.Text = Me.dtpOrderDate.Value
        Me.ListView1.SetFocus
    End If
End Sub

Private Sub dtpPromised_CloseUp()
If Not NewRecord And Not EditRecord Then Exit Sub
    If Me.dtpPromised.Value < Date Then
        MsgBox "Wrong Date! Date Promised Cannot be Earlier than the Date Today!!", vbCritical + vbOKOnly, "Invalid Date"
        Me.txtDatePromised.Text = Empty: Me.txtDatePromised.SetFocus
    Else
        Me.txtDatePromised.Text = Me.dtpPromised.Value
        Me.cboAdvName.SetFocus
    End If
End Sub

Private Sub dtpRequired_CloseUp()
If Not NewRecord And Not EditRecord Then Exit Sub
    If Me.dtpRequired.Value < Date Then
        MsgBox "Wrong Date! Date Required Cannot be Earlier than the Date Today!!", vbCritical + vbOKOnly, "Invalid Date"
        Me.txtDateRequired.Text = Empty: Me.txtDateRequired.SetFocus
    Else
        Me.txtDateRequired.Text = Me.dtpRequired.Value
        Me.txtDatePromised.SetFocus
    End If
End Sub

Private Sub Form_Activate()
If Me.ListView2.ListItems.Count = 0 Then
    MyContracts.GetPurchaseStructure
End If
If Me.ListView1.ListItems.Count = 0 Then
    MyContracts.GetSupplyStructure
End If
End Sub

Private Sub Form_Initialize()
    Set MyContracts = New clsRsContractRequisition
End Sub

Private Sub Form_Load()
    Me.dtpOrderDate.Value = Date
    Me.dtpPromised.Value = Date
    Me.dtpRequired.Value = Date
End Sub

Private Sub Form_Resize()
With Me
    .ListView1.Width = .Width - (12000 - 6135)
    .ListView2.Width = .Width - (12000 - 6135)
    .ListView1.Height = .Height - (8505 - 3615)
    .Frame1.Height = .Height - (8505 - 5655)
    .Frame1.Left = .ListView1.Width + 100
    .txtTitle.Width = .Width - (12000 - 11895)
    .Label3.Left = .Frame1.Left
End With
End Sub

Private Sub Form_Terminate()
    Set MyContracts = Nothing
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo Err
If Not NewRecord And Not EditRecord Then Item.Checked = False: Exit Sub
If Me.ListView1.ListItems.Count = 0 Or Me.ListView1.View <> lvwReport Then Exit Sub
    
    Dim i, j, k
    j = Me.ListView1.ListItems.Count
    
    If j = 0 Then Exit Sub
    
    For i = 1 To j
        If Me.ListView1.ListItems(i).Checked = True Then
            If Me.ListView1.ListItems(i).Text <> Item Then
                Me.ListView1.ListItems(i).Checked = False
            End If
        End If
    Next i
    
    If Item.Checked = True Then
    
        Me.cboAdvCode.Text = Item
        Call cboAdvCode_LostFocus
        
    ElseIf Item.Checked = False Then
       Call ClearForNewBBItem
      
    

        
    End If
    
Exit Sub
Err:
    ErrorMessage
End Sub
Private Sub ClearForNewBBItem()
On Error GoTo Err
With Me
.cboAdvCode.Text = ""
.cboAdvName.Text = ""
.txtType.Text = ""
.txtLenght.Text = ""
.txtWidth.Text = ""
.txtDuration.Text = ""
.txtDays.Text = ""
.txtAdvCost.Text = ""
.txtTotalCost.Text = ""
End With
Exit Sub
Err:
ErrorMessage
End Sub

Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo Err
If Not NewRecord And Not EditRecord Then Item.Checked = False: Exit Sub
If Me.ListView2.ListItems.Count = 0 Or Me.ListView2.View <> lvwReport Then Exit Sub
    
    Dim i, j, k
    j = Me.ListView2.ListItems.Count
    
    If j = 0 Then Exit Sub
    
    For i = 1 To j
        If Me.ListView2.ListItems(i).Text <> Item Then
            Me.ListView2.ListItems(i).Checked = False
        End If
    Next i
    
    If Item.Checked = True Then
        Me.cboClientId.Text = Item
        Call cboClientId_LostFocus
    ElseIf Item.Checked = False Then
        Me.cboClientId.Text = ""
        Me.cboClientName.Text = ""
        Me.txtContactPerson.Text = ""
        Me.txtPhysicalAddress.Text = ""
        Me.txtDateRequired.Text = ""
        Me.txtDatePromised.Text = ""
    End If
    
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub mnuFileClearScreen_Click()
    Call ClearTextFields
    Call ClearDataSheets
End Sub

Private Sub mnuFileClose_Click()
    Unload Me
End Sub


Private Sub mnuPrintCoveringLetter_Click()

End Sub

Private Sub mnuFileLogOut_Click()

End Sub

Private Sub mnuPrintOrderForm_Click()
'On Error GoTo Err
With Me
    INPQRY = InputBox("Please Enter Contract No!!", "Contract Requisition Form")
    
    If Len(INPQRY) = 0 Then
        MsgBox "No Values Entered or Operation Was Cancelled! No Work Will Be Done!!"
        Exit Sub
    Else
        
    
        ThisProduct = (INPQRY)
    
        Set rsFindRecord = cnCOMMON.Execute("SELECT *  FROM AdvertContractRequisition A,AdvertContractRequisitionData B WHERE A.Purchaseorderno='" & ThisProduct & "';")
        
        If rsFindRecord.EOF And rsFindRecord.BOF Then
            MsgBox "The Contract Number Does Not Exist Or Entry is Not Correct!!", vbCritical + vbOKOnly, "Invalid Product Name"
            Set rsFindRecord = Nothing: Exit Sub
       ElseIf Not ValidOrderNumber Then
            MsgBox "The Contract Number: " & " [" & ThisProduct & "] " & "does not exist or has been deleted! No work was done!!", vbCritical + vbOKOnly, "Invalid Contract Number"
             Set rsFindRecord = Nothing: Exit Sub
       Else
               
            SelectedProduct = (INPQRY)
            
            
            Load frmRPTContractRequisitionForm
            frmRPTContractRequisitionForm.Show 1, Me
        End If
    End If

End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Function ValidOrderNumber() As Boolean
'On Error GoTo Err
    Set rsFindRecord = cnCOMMON.Execute("SELECT COUNT(PurchaseOrderNo) AS TOrders FROM AdvertContractRequisition WHERE PurchaseOrderNO='" & ThisProduct & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        ValidOrderNumber = False
    ElseIf rsFindRecord!torders = 0 Then
        ValidOrderNumber = False
    ElseIf rsFindRecord!torders >= 1 Then
        ValidOrderNumber = True
    Else
        ValidOrderNumber = False
    End If
    
    Set rsFindRecord = Nothing
    
Exit Function
Err:
    ErrorMessage
End Function

Private Sub mnuShowCurrentPurchaseOrders_Click()

'On Error GoTo Err
With frmRsContractRequisition
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear
'If Not NewRecord And Not EditRecord Then: Exit Sub
    
    
.ListView1.ColumnHeaders.Add , , "Advert Code", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Serial No", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Name", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Type", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Length", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Width", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Cost", .ListView1.Width / 6 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Duration", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Start Date", .ListView1.Width / 5.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "End Date", .ListView1.Width / 5.5 ', lvwColumnCenter
ListView1.ColumnHeaders.Add , , "Paid", .ListView1.Width / 5.5 ', lvwColumnCenter


Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertContractRequisition A, AdvertContractRequisitionData B WHERE A.PurchaseOrderNo=B.PurchaseOrderNo AND B.PaidStatus IS  NULL ORDER BY B.SerialNO;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!AdvCode))
    
    If Not IsNull(rsLIST!SerialNo) Then
        MyList.SubItems(1) = CStr(rsLIST!SerialNo)
    End If
    
    
    If Not IsNull(rsLIST!PurchaseOrderNo) Then
        MyList.SubItems(2) = CStr(rsLIST!PurchaseOrderNo)
    End If
    
    If Not IsNull(rsLIST!AdvName) Then
        MyList.SubItems(3) = CStr(rsLIST!AdvName)
    End If
    
    If Not IsNull(rsLIST!AdvType) Then
        MyList.SubItems(4) = CStr(rsLIST!AdvType)
    End If
    
    If Not IsNull(rsLIST!AdvLength) Then
        MyList.SubItems(5) = CStr(rsLIST!AdvLength)
    End If
    
    If Not IsNull(rsLIST!AdvWidth) Then
        MyList.SubItems(6) = CStr(rsLIST!AdvWidth)
    End If
    
    
    If Not IsNull(rsLIST!AdvCost) Then
        MyList.SubItems(7) = FormatNumber(rsLIST!AdvCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!Duration) Then
        MyList.SubItems(8) = CStr(rsLIST!Duration)
    End If
    
    If Not IsNull(rsLIST!ContractStartDate) Then
        MyList.SubItems(9) = CStr(rsLIST!ContractStartDate)
    End If
    
    If Not IsNull(rsLIST!ContractEndDate) Then
        MyList.SubItems(10) = CStr(rsLIST!ContractEndDate)
    End If
          
    If IsNull(rsLIST!PaidStatus) Then
        MyList.SubItems(11) = CStr("NO")
    ElseIf Not IsNull(rsLIST!PaidStatus) Then
        If rsLIST!PaidStatus = "Y" Then
            MyList.SubItems(11) = CStr("YES")
        Else
            MyList.SubItems(11) = CStr("NO")
        End If
    End If
    
    rsLIST.MoveNext
    
Wend

.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
.ListView1.ColumnHeaders(6).Alignment = lvwColumnRight
.ListView1.View = lvwReport
If .ListView1.ListItems.Count = 0 Or .ListView1.View <> lvwReport Then Exit Sub

    Dim i, j, k
    j = .ListView1.ListItems.Count

    .ListView1.SelectedItem.Checked = False

    If .ListView1.ListItems.Count = 0 Or .ListView1.View <> lvwReport Then Exit Sub

    
Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub


Private Sub mnuShowBillBoardItems_Click()

'On Error GoTo Err
With frmRsContractRequisition
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Advert Code ", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Name", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Advert Type", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Length", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Width", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Weight", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Duration", .ListView1.Width / 5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Days", .ListView1.Width / 5 ', lvwColumnCenter

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertBBDetails  WHERE costpricestatus='" & "Y" & "'    AND BillBoardNo IS NOT NULL AND Discontinued='" & "N" & "' ORDER BY BillBoardNo ;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!BillBoardNo))

    If Not IsNull(rsLIST!Name) Then
        MyList.SubItems(1) = CStr(rsLIST!Name)
    End If
    
    If Not IsNull(rsLIST!TypeName) Then
        MyList.SubItems(2) = Trim(CStr(rsLIST!TypeName))
    End If
    
    If Not IsNull(rsLIST!Length) Then
        MyList.SubItems(3) = Trim(CStr(rsLIST!Length))
    End If
    
    If Not IsNull(rsLIST!Width) Then
        MyList.SubItems(4) = Trim(CStr(rsLIST!Width))
    End If
    
    If Not IsNull(rsLIST!Weight) Then
        MyList.SubItems(5) = Trim(CStr(rsLIST!Weight))
    End If
    
    If Not IsNull(rsLIST!DurationName) Then
        MyList.SubItems(6) = Trim(CStr(rsLIST!DurationName))
    End If
    
    If Not IsNull(rsLIST!NoOfDays) Then
        MyList.SubItems(7) = Trim(CStr(rsLIST!NoOfDays))
    End If
    
    
            
    rsLIST.MoveNext
    
Wend

.txtDataSource.Text = "NEW CONTRACT"
Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub mnuShowRegisteredClients_Click()

'On Error GoTo Err
With frmRsContractRequisition
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Client Code", .ListView2.Width / 4.5
.ListView2.ColumnHeaders.Add , , "Client Name", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contact Person", .ListView2.Width / 4.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Physical Address", .ListView2.Width / 2.5 ', lvwColumnCenter

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertClients WHERE Customerid IS NOT NULL ORDER BY CompanyName;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!CustomerId))

    If Not IsNull(rsLIST!CompanyName) Then
        MyList.SubItems(1) = CStr(rsLIST!CompanyName)
    End If
    
    If Not IsNull(rsLIST!ContactName) Then
        MyList.SubItems(2) = CStr(rsLIST!ContactName)
    End If
    
    If Not IsNull(rsLIST!PhysicalAddress) Then
        MyList.SubItems(3) = Trim(CStr(rsLIST!PhysicalAddress))
    End If
    
    
    rsLIST.MoveNext
    
Wend

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage

End Sub
Private Sub ShowClients()
'On Error GoTo Err
With frmRsContractRequisition
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Client Code", .ListView2.Width / 4.5
.ListView2.ColumnHeaders.Add , , "Client Name", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contact Person", .ListView2.Width / 4.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Physical Address", .ListView2.Width / 2.5 ', lvwColumnCenter

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertClients WHERE Customerid IS NOT NULL ORDER BY CompanyName;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!CustomerId))

    If Not IsNull(rsLIST!CompanyName) Then
        MyList.SubItems(1) = CStr(rsLIST!CompanyName)
    End If
    
    If Not IsNull(rsLIST!ContactName) Then
        MyList.SubItems(2) = CStr(rsLIST!ContactName)
    End If
    
    If Not IsNull(rsLIST!PhysicalAddress) Then
        MyList.SubItems(3) = Trim(CStr(rsLIST!PhysicalAddress))
    End If
    
    
    rsLIST.MoveNext
    
Wend

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:

If Err.Number = 3265 Then Resume Next

End Sub
  
Private Sub Text5_GotFocus()
    Me.txtOrderNO.SetFocus
End Sub

Private Sub txtOrderDescription_LostFocus()
    Me.txtOrderDescription.Text = StrConv(Me.txtOrderDescription.Text, vbProperCase)
End Sub

Private Sub txtOrderNO_GotFocus()

If Me.txtOrderNO.Text <> Empty Or Not NewRecord Then Exit Sub
    If Me.chkOrderNo.Value = 1 Then
        Me.txtOrderNO.Text = MyContracts.AutoPurchaseOrderNo
        Me.txtOrderDescription.SetFocus
    Else
        Me.txtOrderNO.SetFocus
    End If

End Sub


Private Sub txtOrderNO_KeyPress(KeyAscii As Integer)
'On Error GoTo Err
With Me
If NewRecord Then
    If .chkOrderNo.Value = 1 Then
        KeyAscii = 0
    Else
        Exit Sub
    End If
Else
    If KeyAscii = vbKeyReturn Then
        MyContracts.FindOrderDetails
    Else
        Exit Sub
    End If
End If
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub txtOrderNO_LostFocus()
    If NewRecord Then
        Me.txtOrderDate.Text = Date
        Me.txtOrderDescription.SetFocus
    End If
End Sub








Private Sub GetStaffIdNo()
'On Error GoTo Err
Dim StaffIdNo As Variant
With frmRsContractRequisition

    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM Paramempmaster A,Adminuserregister B WHERE A.Staffidno=B.Staffidno AND B.UserName='" & CurrentUserName & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        
        StaffIdNo = rsFindRecord!StaffIdNo & ""
             
        
    End If
    
End With
Exit Sub
Err:
    ErrorMessage
End Sub



Public Sub RemoveCurrentListItem()
'On Error GoTo Err
With Me
Dim i, j, k
   j = .ListView1.ListItems.Count: i = 1
     If j = 0 Then Exit Sub
     
     For i = 1 To j
      If .ListView1.ListItems(i).Checked = True Then
         .ListView1.ListItems.Remove (i): Exit Sub
      End If
    Next i
End With
Exit Sub
Err:
   ErrorMessage
End Sub
