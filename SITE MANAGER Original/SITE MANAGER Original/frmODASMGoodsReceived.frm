VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmODASMReceiveGoods 
   Caption         =   "Receive GOODS"
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmODASMGoodsReceived.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   Picture         =   "frmODASMGoodsReceived.frx":0442
   ScaleHeight     =   7815
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboOrderNO 
      BackColor       =   &H00FFFFC0&
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
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox txtOrderDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Height          =   390
      Left            =   4560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   840
      Width           =   7215
   End
   Begin VB.TextBox txtOrderDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Height          =   390
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   6960
      TabIndex        =   26
      Top             =   1695
      Width           =   4815
      Begin VB.TextBox txtNewCost 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   3600
         TabIndex        =   49
         Top             =   3480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtDosesOrdered 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2362
         Width           =   2415
      End
      Begin VB.TextBox txtDosesReceived 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   17
         Top             =   3400
         Width           =   2415
      End
      Begin VB.TextBox txtShippDate 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   1335
         TabIndex        =   9
         Top             =   1030
         Width           =   1335
      End
      Begin VB.TextBox txtDateReceived 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   3495
         TabIndex        =   10
         Top             =   1030
         Width           =   1215
      End
      Begin VB.TextBox txtDeliveryNote 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtInvoiceNumber 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   575
         Width           =   1575
      End
      Begin VB.TextBox txtDelDate 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtInvoiceDate 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   8
         Top             =   575
         Width           =   1215
      End
      Begin VB.TextBox txtSerialNO 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   345
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtDrugName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   345
         Left            =   2040
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox txtDrugCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   345
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1800
         Width           =   1215
      End
      Begin VB.ComboBox cboUnitsReceived 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   345
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   3876
         Width           =   1095
      End
      Begin VB.TextBox txtQuantityReceived 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         TabIndex        =   18
         Top             =   3876
         Width           =   2415
      End
      Begin VB.CommandButton cmdChange 
         BackColor       =   &H80000000&
         Caption         =   "&CHANGE"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   5040
         Width           =   2175
      End
      Begin VB.CommandButton cmdSAVE 
         BackColor       =   &H80000000&
         Caption         =   "&NEW"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   5040
         Width           =   2295
      End
      Begin VB.TextBox txtQuantityOrdered 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2838
         Width           =   2415
      End
      Begin VB.ComboBox cboUnitsOrdered 
         BackColor       =   &H00FFFFC0&
         Height          =   345
         Left            =   3600
         TabIndex        =   16
         Top             =   2838
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpDelDate 
         Height          =   375
         Left            =   3240
         TabIndex        =   36
         Top             =   120
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19791873
         CurrentDate     =   37965
      End
      Begin MSComCtl2.DTPicker dtpInvoiceDate 
         Height          =   375
         Left            =   3240
         TabIndex        =   37
         Top             =   575
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19791873
         CurrentDate     =   37965
      End
      Begin MSComCtl2.DTPicker dtpDateReceived 
         Height          =   375
         Left            =   3240
         TabIndex        =   42
         Top             =   1030
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19791873
         CurrentDate     =   37965
      End
      Begin MSComCtl2.DTPicker dtpShippDate 
         Height          =   375
         Left            =   1080
         TabIndex        =   43
         Top             =   1030
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19791873
         CurrentDate     =   37965
      End
      Begin VB.Label Label21 
         Caption         =   "S.N"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   48
         Top             =   1560
         Width           =   615
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000002&
         X1              =   120
         X2              =   4680
         Y1              =   4322
         Y2              =   4322
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000002&
         X1              =   120
         X2              =   4680
         Y1              =   3284
         Y2              =   3284
      End
      Begin VB.Label Label20 
         Caption         =   "Prod Cost."
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   2355
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Prod Cost."
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   3405
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Ship. Date"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1030
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Rec."
         Height          =   255
         Left            =   2760
         TabIndex        =   44
         Top             =   1030
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "Del. Note"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Invoice No"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   575
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "Date"
         Height          =   255
         Left            =   2760
         TabIndex        =   39
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   "Date"
         Height          =   255
         Left            =   2760
         TabIndex        =   38
         Top             =   575
         Width           =   375
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000002&
         X1              =   120
         X2              =   4680
         Y1              =   1485
         Y2              =   1485
      End
      Begin VB.Label Label13 
         Caption         =   "Prod Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   35
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Prod Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Qty Rec."
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   3870
         Width           =   975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000002&
         X1              =   120
         X2              =   4680
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Label Label12 
         Caption         =   "Qty Ord."
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   2838
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000002&
         X1              =   120
         X2              =   4680
         Y1              =   2246
         Y2              =   2246
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1695
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "Check an Order Record to Display the Details"
      Top             =   1695
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   2990
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
      BackColor       =   16777215
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
      TabIndex        =   23
      Top             =   7440
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   4304
            MinWidth        =   4304
            Picture         =   "frmODASMGoodsReceived.frx":5B80
            Text            =   "Purchase Orders"
            TextSave        =   "Purchase Orders"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   13829
            MinWidth        =   13829
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3246
            MinWidth        =   3246
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
      Height          =   570
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
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
         TabIndex        =   24
         Text            =   "*** RECEIVE L.P. ORDERS ***"
         Top             =   0
         Width           =   11895
      End
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   3735
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Select a Record Here by Checking or Double-Clicking on it!"
      Top             =   3960
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6588
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
      BackColor       =   16777215
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Items/Commodities/Drugs Under Current Order"
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
      TabIndex        =   32
      Top             =   3360
      Width           =   5295
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   0
      X2              =   11880
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order No"
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
      TabIndex        =   31
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Order Description"
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
      Left            =   4560
      TabIndex        =   29
      Top             =   600
      Width           =   1935
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
      Left            =   2520
      TabIndex        =   28
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Order Shipment/Delivery Particulars"
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
      Left            =   6960
      TabIndex        =   27
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "List of Current/New Purchase Order(s)"
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
      TabIndex        =   25
      Top             =   1440
      Width           =   5295
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileClearScreen 
         Caption         =   "Clear &Screen"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuShow 
      Caption         =   "Show/&View"
      Begin VB.Menu mnuShowAllNew 
         Caption         =   "All New Purchase &Orders"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnujn 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowExpectedToday 
         Caption         =   "Orders &Expected/Due Today"
      End
      Begin VB.Menu mnullk 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowOverdue 
         Caption         =   "&Overdue/Delayed Orders"
      End
      Begin VB.Menu mnujjjhb 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowRefresh 
         Caption         =   "&Refresh the Screen"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "Print &Preview"
      Begin VB.Menu mnuPrintListReceived 
         Caption         =   "List of Order Items &Received"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnukkk 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintNotReceived 
         Caption         =   "Order Items &Not Delivered/Received"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help System"
      Begin VB.Menu mnuHelpUsing 
         Caption         =   "Using the &Shipment Receiving Assistant"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmODASMReceiveGoods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyReceiver As clsODASReceiveGoods

Private Sub cboOrderNO_Click()
If Not NewRecord Then Exit Sub
    MyReceiver.ClearTextFields
    Me.txtDeliveryNote.SetFocus
End Sub

Private Sub cboOrderNO_GotFocus()
If Not NewRecord Then Exit Sub
    MyReceiver.AttachNewPurchaseOrders
End Sub

Private Sub cboOrderNO_LostFocus()
If Not NewRecord Then Exit Sub
    MyReceiver.FindDetailsByPurchaseOrderNo
End Sub

Private Sub cmdSAVE_Click()
'On Error GoTo Err
With Me
If EditRecord Then Exit Sub
If CurrentUserName = "administrator" Then
    MsgBox "SORRY!! You cannot Receive Orders when Logged on as System Administrator! Please Log Out and Use a Registered Staff Member's Profile!!!", vbCritical + vbOKOnly, "Wrong Profile"
    Exit Sub
Else
    Select Case .cmdSAVE.Caption
    Case "&NEW"
        NewRecord = True
        MyReceiver.ShowNewPurchaseOrders: MyReceiver.ShowCurrentPendingOrdersData
        MyReceiver.AddNewRecord
    Case "&SAVE RECORD"
        If NewRecord Then
            If ValidReceived Then
                MyReceiver.SaveReceivedData
            End If
        End If
    Case Else
        Exit Sub
    End Select
End If
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Function ValidReceived() As Boolean
'On Error GoTo Err
With Me
    If .cboOrderNO.Text = Empty Then
        strMessage = "Required Purchase Order Number!"
        .cboOrderNO.SetFocus
    ElseIf .txtDrugCode.Text = Empty Then
        strMessage = "Required Drug/Item Code!!"
        .txtDrugCode.SetFocus
    ElseIf .txtDrugName.Text = Empty Then
        strMessage = "Required Name of Drugs!!"
        .txtDrugName.SetFocus
    ElseIf .txtSerialNO.Text = Empty Then
        strMessage = "Required Serial Number!!"
        .txtSerialNO.SetFocus
    ElseIf .txtDeliveryNote.Text = Empty Then
        strMessage = "Required Delivery Note Number!!"
        .txtDeliveryNote.SetFocus
    ElseIf .txtDelDate.Text = Empty Then
        strMessage = "Required Date of Delivery!!"
        .txtDelDate.SetFocus
    ElseIf .txtShippDate.Text = Empty Then
        strMessage = "Required shipping Date!!"
        .txtShippDate.SetFocus
    ElseIf .txtDateReceived.Text = Empty Then
        strMessage = "Required Date Received!!"
        .txtDateReceived.SetFocus
    ElseIf .txtQuantityOrdered.Text = Empty Then
        strMessage = "Required Quantity Ordered!!"
        .txtQuantityOrdered.SetFocus
    ElseIf .txtQuantityReceived.Text = Empty Then
        strMessage = "Required Quantity Received!!"
        .txtQuantityReceived.SetFocus
'    ElseIf .txtManufactureDate.Text = Empty Then
'        StrMessage = "Required Date of Manufacture!!"
'        .txtManufactureDate.SetFocus
'    ElseIf .txtExpiryDate.Text = Empty Then
'        StrMessage = "Required Expiry Date!!"
'        .txtExpiryDate.SetFocus
    ElseIf .cboUnitsOrdered.Text = Empty Then
        strMessage = "Required Units of Quantity!!"
        .cboUnitsOrdered.SetFocus
    ElseIf .cboUnitsReceived.Text = Empty Then
        strMessage = "Required Units of Quantity!!"
        .cboUnitsReceived.SetFocus
    Else
        ValidReceived = True
    End If
    If Not ValidReceived Then
        MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
    End If
End With
Exit Function
Err:
    ErrorMessage
End Function

Private Sub dtpDateReceived_CloseUp()
'On Error GoTo Err
If Not NewRecord And Not EditRecord Then Exit Sub
With Me
    If .dtpDateReceived.Value > Date Then
        MsgBox "Invalid Date!! The Date Received Cannot be in the Future!!!", vbCritical + vbOKOnly, "Wrong Date"
        .txtDateReceived.Text = Empty: .txtDateReceived.SetFocus
    Else
        .txtDateReceived.Text = .dtpDateReceived.Value
'        .txtManufactureDate.SetFocus
    End If
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub dtpDelDate_CloseUp()
'On Error GoTo Err
If Not NewRecord And Not EditRecord Then Exit Sub
With Me
    If .dtpDelDate.Value > Date Then
        MsgBox "Invalid Date!! Delivery Note Date Cannot be in the Future!!!", vbCritical + vbOKOnly, "Wrong Date"
        .txtDelDate.Text = Empty: .txtDelDate.SetFocus
    Else
        .txtDelDate.Text = .dtpDelDate.Value
        .txtInvoiceNumber.SetFocus
    End If
End With
Exit Sub
Err:
    ErrorMessage
End Sub


Private Sub dtpInvoiceDate_CloseUp()
'On Error GoTo Err
If Not NewRecord And Not EditRecord Then Exit Sub
With Me
    If .dtpInvoiceDate.Value > Date Then
        MsgBox "Invalid Date!! The Invoice Date Cannot be in the Future!!!", vbCritical + vbOKOnly, "Wrong Date"
        .txtInvoiceDate.Text = Empty: .txtInvoiceDate.SetFocus
    Else
        .txtInvoiceDate.Text = .dtpInvoiceDate.Value
        .txtShippDate.SetFocus
    End If
End With
Exit Sub
Err:
    ErrorMessage
End Sub



Private Sub dtpShippDate_CloseUp()
'On Error GoTo Err
If Not NewRecord And Not EditRecord Then Exit Sub
With Me
    If .dtpShippDate.Value > Date Then
        MsgBox "Invalid Date!! Shipping Date Cannot be in the Future!!!", vbCritical + vbOKOnly, "Wrong Date"
        .txtShippDate.Text = Empty: .txtShippDate.SetFocus
    Else
        .txtShippDate.Text = .dtpShippDate.Value
        .txtDateReceived.SetFocus
    End If
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub Form_Activate()
If Me.ListView1.ListItems.Count = 0 Then
    MyReceiver.GetPurchaseStructure
End If
If Me.ListView2.ListItems.Count = 0 Then
    MyReceiver.GetRecordsStructure
End If
End Sub

Private Sub Form_Initialize()
    Set MyReceiver = New clsODASReceiveGoods
End Sub

Private Sub Form_Load()
With Me
    .dtpDateReceived.Value = Date
'    .dtpExpiryDate.Value = Date
'    .dtpManufactureDate.Value = Date
    .dtpShippDate.Value = Date
    .dtpDelDate.Value = Date
    .dtpInvoiceDate.Value = Date
End With
End Sub

Private Sub Form_Resize()
With Me
    .ListView1.Width = .Width - (12000 - 6855)
    .ListView2.Width = .Width - (12000 - 6855)
    .ListView2.Height = .Height - (8505 - 3735)
    .Text5.Width = .Width - (12000 - 11895)
    .Frame1.Height = .Height - (8505 - 5655)
    .Frame1.Left = .ListView1.Width + 100
    .txtOrderDescription.Width = .Width - (12000 - 7215)
    .Label3.Left = .Frame1.Left
End With
End Sub

Private Sub Form_Terminate()
    Set MyReceiver = Nothing
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo Err

If Me.ListView1.ListItems.Count = 0 Or Me.ListView1.View <> lvwReport Then Item.Checked = False: Exit Sub
    
    Dim i, j, k
    j = Me.ListView1.ListItems.Count
    
    If j = 0 Then Exit Sub
    
    For i = 1 To j
        If Me.ListView1.ListItems(i).Text <> Item Then
            Me.ListView1.ListItems(i).Checked = False
        End If
    Next i
    
    If Item.Checked = True Then
        CurrentOrder = Item
        Me.cboOrderNO.Text = Item
        Me.txtOrderDate.Text = Item.SubItems(1)
        Me.txtOrderDescription.Text = MyReceiver.GetMyOrderDescription
        MyReceiver.ShowItemsInCurrentOrder
        Me.ListView2.SetFocus
    ElseIf Item.Checked = False Then
        Me.cboOrderNO.Clear
        Me.txtOrderDate.Text = Empty
        Me.txtOrderDescription.Text = Empty
        Me.ListView2.ListItems.Clear
    End If
    
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
        Me.txtSerialNO.Text = Item
            
        If Me.txtDeliveryNote.Text = Empty Or Me.txtInvoiceNumber.Text = Empty Then
            
            Set rsFindRecord = New ADODB.Recordset
            rsFindRecord.Open "SELECT * FROM PharmPurchaseORdersData WHERE PurchaseOrderNo='" & Trim(Me.cboOrderNO.Text) & "' AND ReceivedStatus IS NOT NULL ORDER BY SerialNo;", cnCOMMON, adOpenKeyset, adLockOptimistic
            
            If rsFindRecord.EOF And rsFindRecord.BOF Then
                Set rsFindRecord = Nothing: Me.txtDeliveryNote.SetFocus
            Else
                rsFindRecord.MoveLast
                
                Me.txtDeliveryNote.Text = rsFindRecord!delnoteno & ""
                Me.txtDelDate.Text = rsFindRecord!DelDate & ""
                Me.txtInvoiceDate.Text = rsFindRecord!invoicedate & ""
                Me.txtInvoiceNumber.Text = rsFindRecord!invoiceno & ""
                Me.txtDateReceived.Text = rsFindRecord!datereceived & ""
                Me.txtShippDate.Text = rsFindRecord!shippingdate & ""
                
                Me.txtQuantityReceived.SetFocus
                
            End If
            
            
        End If
        
        MyReceiver.ClearForNewRecord
        
        Me.txtSerialNO.Text = Item
        Me.txtDrugCode.Text = Item.SubItems(1)
        Me.txtDrugName.Text = Item.SubItems(2)
        Me.txtQuantityOrdered.Text = Item.SubItems(3)
'        Me.txtQuantityReceived.Text = Item.SubItems(3)
        Me.cboUnitsOrdered.Text = Item.SubItems(4)
        Me.cboUnitsReceived.Text = Item.SubItems(4)
        Me.txtDosesOrdered.Text = Item.SubItems(6)
'        Me.txtDosesReceived.Text = Item.SubItems(5)
        Me.txtNewCost.Text = Item.SubItems(5)
'        Me.txtDosesReceived.SetFocus
        Me.txtDosesReceived.Locked = False
        
    ElseIf Item.Checked = False Then
    
        MyReceiver.ClearForNewRecord
        
    End If
    
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub mnuShowAllNew_Click()
    MyReceiver.ShowNewPurchaseOrders
    Me.ListView1.SetFocus
End Sub

Private Sub mnuShowExpectedToday_Click()
    MyReceiver.ShowPurchaseOrdersDueToday
    Me.ListView1.SetFocus
End Sub

Private Sub mnuShowOverdue_Click()
    MyReceiver.ShowPurchaseOrdersOverDue
    Me.ListView1.SetFocus
End Sub

Private Sub mnuShowRefresh_Click()
'On Error GoTo Err
With Me
If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Screen Refresher") = vbCancel Then Exit Sub
    NewRecord = False: EditRecord = False
    MyReceiver.ClearTheScreen
    
    .cmdSAVE.Enabled = True
    .cmdSAVE.Caption = "&NEW"
    .cmdChange.Enabled = True
    .cmdChange.Caption = "&CHANGE"
    
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text5_GotFocus()
    Me.cboOrderNO.SetFocus
End Sub

Private Sub txtDosesReceived_GotFocus()
 'On Error GoTo Err
If Not NewRecord And Not EditRecord Then Exit Sub
With Me
    .txtDosesReceived.SelStart = 0
    .txtDosesReceived.SelLength = Len(.txtDosesReceived.Text)
    Me.txtDosesReceived.Text = CCur(Me.txtNewCost.Text * Me.txtQuantityReceived.Text)
        
End With
Exit Sub
Err:
If Err.Number = 13 Then
MsgBox "Please Enter The Quantity Received First ", vbInformation, "P.O Receive"
Else
ErrorMessage
End If
End Sub

Private Sub txtDosesReceived_LostFocus()
If Not NewRecord And Not EditRecord Then Exit Sub
If Me.txtDosesReceived.Text = Empty Then Exit Sub
'    Call ComputeTotalQuantity
'    Me.txtManufactureDate.SetFocus
End Sub

Private Sub txtExpiryDate_Change()

End Sub

Private Sub txtManufactureDate_Change()

End Sub

