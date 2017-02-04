VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmODASMPurchaseOrders 
   Caption         =   "General Inventory- PURCHASE ORDERS SYSTEM"
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
   Icon            =   "frmODASMPurchaseOrders.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   Picture         =   "frmODASMPurchaseOrders.frx":0442
   ScaleHeight     =   7815
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkOrderNo 
      Height          =   255
      Left            =   0
      TabIndex        =   37
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
      Height          =   5655
      Left            =   6240
      TabIndex        =   6
      Top             =   1680
      Width           =   5535
      Begin VB.ComboBox cboStaffIdNo 
         BackColor       =   &H00FFC0C0&
         Height          =   345
         Left            =   1320
         TabIndex        =   48
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtProductDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   2160
         Width           =   4095
      End
      Begin VB.TextBox cboShippingMethod 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   330
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox cboSupplierCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   330
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   120
         Width           =   3975
      End
      Begin VB.TextBox txtTotalDoses 
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   4080
         TabIndex        =   43
         Top             =   2688
         Width           =   1095
      End
      Begin VB.TextBox txtDataSource 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   5280
         Width           =   495
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   375
         Left            =   5160
         Max             =   1
         Min             =   500
         TabIndex        =   41
         Top             =   2688
         Value           =   1
         Width           =   255
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
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5040
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
         Left            =   1815
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5040
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
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4320
         Width           =   3615
      End
      Begin MSComCtl2.DTPicker dtpPromised 
         Height          =   315
         Left            =   5175
         TabIndex        =   36
         Top             =   1560
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57671681
         CurrentDate     =   37965
      End
      Begin MSComCtl2.DTPicker dtpRequired 
         Height          =   315
         Left            =   2775
         TabIndex        =   35
         Top             =   1569
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57671681
         CurrentDate     =   37965
      End
      Begin VB.TextBox txtTotalCost 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3694
         Width           =   1335
      End
      Begin VB.TextBox txtDosageCost 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         TabIndex        =   10
         Top             =   3694
         Width           =   1455
      End
      Begin VB.ComboBox cboUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   345
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3191
         Width           =   1335
      End
      Begin VB.TextBox txtProductCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2688
         Width           =   1455
      End
      Begin VB.TextBox txtDatePromised 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1569
         Width           =   1215
      End
      Begin VB.TextBox txtDateRequired 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtFreightCharge 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1066
         Width           =   2535
      End
      Begin VB.TextBox txtShippingMethod 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   330
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Staff Id No"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   3285
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   3120
         TabIndex        =   40
         Top             =   2748
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Product Code"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2748
         Width           =   1335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000002&
         BorderWidth     =   2
         X1              =   120
         X2              =   5400
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label15 
         Caption         =   "Total"
         Height          =   255
         Left            =   3120
         TabIndex        =   34
         Top             =   3750
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "Product Cost"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   3754
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Units "
         Height          =   255
         Left            =   3120
         TabIndex        =   32
         Top             =   3236
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Product Ordered"
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   2140
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000002&
         BorderWidth     =   2
         X1              =   120
         X2              =   5400
         Y1              =   2072
         Y2              =   2072
      End
      Begin VB.Label Label10 
         Caption         =   "Promised"
         Height          =   255
         Left            =   3120
         TabIndex        =   30
         Top             =   1629
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Date Required"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1629
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Freight Charge"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1126
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Shipp. Method"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Supplier Code"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   165
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3615
      Left            =   0
      TabIndex        =   3
      Top             =   3735
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6376
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
      TabIndex        =   17
      Top             =   7440
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   4657
            MinWidth        =   4657
            Picture         =   "frmODASMPurchaseOrders.frx":5B80
            Text            =   "PURCHASE ORDERS"
            TextSave        =   "PURCHASE ORDERS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   13829
            MinWidth        =   13829
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2893
            MinWidth        =   2893
            TextSave        =   "15/11/2004"
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
      TabIndex        =   16
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
         TabIndex        =   18
         Text            =   "***  PURCHASE ORDER SYSTEM ***"
         Top             =   0
         Width           =   11895
      End
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1815
      Left            =   0
      TabIndex        =   4
      Top             =   1695
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3201
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
      Format          =   57671681
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
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "L.P.O  No"
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
      TabIndex        =   38
      Top             =   600
      Width           =   1695
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
      Left            =   4680
      TabIndex        =   28
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
      Left            =   2280
      TabIndex        =   27
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order Particulars"
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
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "List of Current / Accredited Suppliers"
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
      Top             =   1440
      Width           =   5295
   End
   Begin VB.Label lblInventory 
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory / Purchase Order Records"
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
      Top             =   3480
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
      Begin VB.Menu mnuFileLogOut 
         Caption         =   "&Log Out/End Session"
         Shortcut        =   ^L
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
      Caption         =   "&Show/View"
      Begin VB.Menu mnuShowCurrentPurchaseOrders 
         Caption         =   "Show Current Purchase Order(s)"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu mnuDFd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowPurchaseRecords 
         Caption         =   "Re-&Order Products/Items"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuoook 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowNewInventory 
         Caption         =   "New Inventory Products/Items"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnujjnh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowSuppliers 
         Caption         =   "Current List of &Suppliers"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnujjhn 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Print Preview"
      Begin VB.Menu mnuPrintOrderForm 
         Caption         =   "Purchase &Order Form"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnujjbn 
         Caption         =   "-"
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
Attribute VB_Name = "frmODASMPurchaseOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyPurchase As clsODASLPORequisition

Private Sub txtProductDescription_Click()
If Not NewRecord And Not EditRecord Then Exit Sub
    Me.txtTotalDoses.SetFocus
End Sub

Private Sub txtProductDescription_GotFocus()
If Not NewRecord And Not EditRecord Then Exit Sub
    MyPurchase.AttachDrugsList
End Sub

Private Sub txtProductDescription_LostFocus()
'If Not NewRecord And Not EditRecord Then Exit Sub
'    MyPurchase.GetDrugCode
End Sub

Private Sub cboDeliveryMethod_Click()
If Not NewRecord And Not EditRecord Then Exit Sub
    Me.txtDateRequired.SetFocus
End Sub

Private Sub cboDeliveryMethod_GotFocus()
If Not NewRecord And Not EditRecord Then Exit Sub
    MyPurchase.AttachDeliveryMethods
End Sub

Private Sub cboStaffIdNo_Click()
With frmODASMPurchaseOrders
.cmdRefresh.SetFocus
End With
End Sub

Private Sub cboStaffIdNo_GotFocus()
MyPurchase.AttachPharmacyStaff
End Sub

Private Sub cboStaffIdNo_LostFocus()
MyPurchase.GetStaffID
End Sub

Private Sub cboSupplierCode_Click()
If Not NewRecord And Not EditRecord Then Exit Sub
    Me.cboDeliveryMethod.SetFocus
End Sub

Private Sub cboSupplierCode_GotFocus()
If Not NewRecord And Not EditRecord Then Exit Sub
    MyPurchase.AttachSuppliers
End Sub

Private Sub cboSupplierCode_LostFocus()
If Not NewRecord And Not EditRecord Then Exit Sub
    MyPurchase.GetSupplierCode
End Sub

Private Sub cboUnits_Click()
'If Not NewRecord And Not EditRecord Then Exit Sub
'    Me.txtTotalCost.SetFocus
Me.cboUnits.Locked = False
End Sub

Private Sub cboUnits_GotFocus()
If Not NewRecord And Not EditRecord Then Exit Sub
    MyPurchase.AttachQuantityUnits
End Sub

Private Sub cmdChange_Click()

End Sub

Private Sub cmdFinish_Click()
On Error GoTo err
    If Not NewRecord Then
        If ValidMainPurchase Then
            MyPurchase.SaveMainPurchase
        End If
    End If
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo err
With Me

    If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Screen Refresher") = vbCancel Then Exit Sub
        NewRecord = False: EditRecord = False
        MyPurchase.ClearTheScreen
        
        .cmdSAVE.Enabled = True
        .cmdSAVE.Caption = "ADD &NEW"
        .cmdFinish.Enabled = False
        .cmdRefresh.Enabled = True
        
End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdSAVE_Click()

On Error GoTo err
With Me
If EditRecord Then Exit Sub
If CurrentUserName = "administrator" Then
    MsgBox "SORRY!! You cannot Add Purchase Order Data when Logged-On as System Administrator! Please Log Out and Use a Registered Staff Member's Profile!!!", vbCritical + vbOKOnly, "Wrong Profile"
    Exit Sub
Else
    Select Case .cmdSAVE.Caption
    Case "ADD &NEW"
    
        NewRecord = True
        MyPurchase.AddNewRecord
        
    Case "&SAVE RECORD"
    
        If NewRecord Then
            If ValidPurchase Then
            
                MyPurchase.SavePurchaseData
                
            End If
        End If
        
    Case Else
        Exit Sub
    End Select
End If
End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Function ValidMainPurchase() As Boolean
On Error GoTo err
With Me
    If .txtOrderNO.Text = Empty Then
        strMessage = "Required Purchase Order Number!"
        .txtOrderNO.SetFocus
    ElseIf .txtOrderDate.Text = Empty Then
        strMessage = "Required Purchase Order Date!!"
        .txtOrderDate.SetFocus
    ElseIf .cboSupplierCode.Text = Empty Then
        strMessage = "Required Supplier Code!!"
        .cboSupplierCode.SetFocus
    ElseIf .cboDeliveryMethod.Text = Empty Then
        strMessage = "Required Shipping Method!!"
        .cboDeliveryMethod.SetFocus
    ElseIf .txtProductDescription.Text = Empty Then
        strMessage = "Required Name of Drug Ordered!!"
        .txtProductDescription.SetFocus
    ElseIf .cboUnits.Text = Empty Then
        strMessage = "Required Units of Quantity!!"
        .cboUnits.SetFocus
    ElseIf .txtDatePromised.Text = Empty Then
        strMessage = "Required Promised Date!!"
        .txtDatePromised.SetFocus
    ElseIf .txtDateRequired.Text = Empty Then
        strMessage = "Please enter the date the order is required!!"
        .txtDateRequired.SetFocus
    ElseIf .txtProductCode.Text = Empty Then
        strMessage = "Required Drug Code!!"
        .txtProductCode.SetFocus
    ElseIf .txtFreightCharge.Text = Empty Then
        strMessage = "Required Freight Charges!!"
        .txtFreightCharge.SetFocus
'    ElseIf .txtQuantity.Text = Empty Then
'        StrMessage = "Required Quantity Ordered!!"
'        .txtQuantity.SetFocus
    ElseIf .txtDeliveryMethod.Text = Empty Then
        strMessage = "Required Shipping Method!!"
        .txtDeliveryMethod.SetFocus
    ElseIf .txtTotalCost.Text = Empty Then
        strMessage = "Required Total Cost!!"
        .txtTotalCost.SetFocus
    ElseIf .txtDosageCost.Text = Empty Then
        strMessage = "Required Unit Cost!!"
        .txtDosageCost.SetFocus
    Else
        ValidMainPurchase = True
    End If
    If Not ValidMainPurchase Then
        MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
    End If
End With
Exit Function
err:
    ErrorMessage
End Function

Private Function ValidPurchase() As Boolean
On Error GoTo err
With Me
    If .txtOrderNO.Text = Empty Then
        strMessage = "Required Purchase Order Number!"
        .txtOrderNO.SetFocus
    ElseIf .txtOrderDate.Text = Empty Then
        strMessage = "Required Purchase Order Date!!"
        .txtOrderDate.SetFocus
    ElseIf .cboSupplierCode.Text = Empty Then
        strMessage = "Required Supplier Code!!"
        .cboSupplierCode.SetFocus
    
    ElseIf .cboStaffIdNo.Text = Empty Then
        strMessage = "Required Staff Id No!!"
    ElseIf .cboDeliveryMethod.Text = Empty Then
        strMessage = "Required Shipping Method!!"
        .cboDeliveryMethod.SetFocus
    ElseIf .txtProductDescription.Text = Empty Then
        strMessage = "Required Name of Drug Ordered!!"
        .txtProductDescription.SetFocus
    ElseIf .cboUnits.Text = Empty Then
        strMessage = "Required Units of Quantity!!"
        .cboUnits.SetFocus
    ElseIf .txtDatePromised.Text = Empty Then
        strMessage = "Required Promised Date!!"
        .txtDatePromised.SetFocus
    ElseIf .txtDateRequired.Text = Empty Then
        strMessage = "Please enter the date the order is required!!"
        .txtDateRequired.SetFocus
    ElseIf .txtProductCode.Text = Empty Then
        strMessage = "Required Drug Code!!"
        .txtProductCode.SetFocus
    ElseIf .txtFreightCharge.Text = Empty Then
        strMessage = "Required Freight Charges!!"
        .txtFreightCharge.SetFocus
'    ElseIf .txtQuantity.Text = Empty Then
'        StrMessage = "Required Quantity Ordered!!"
'        .txtQuantity.SetFocus
    ElseIf .txtDeliveryMethod.Text = Empty Then
        strMessage = "Required Shipping Method!!"
        .txtDeliveryMethod.SetFocus
    ElseIf .txtTotalCost.Text = Empty Then
        strMessage = "Required Total Cost!!"
        .txtTotalCost.SetFocus
    ElseIf .txtDosageCost.Text = Empty Then
        strMessage = "Required Unit Cost!!"
        .txtDosageCost.SetFocus
    ElseIf RecordExists Then
        strMessage = "The Product Code: " & " [" & .txtProductCode.Text & "] already exists in the current order!!"
        MyPurchase.ClearForNewDrug: .txtProductDescription.SetFocus
    Else
        ValidPurchase = True
    End If
    If Not ValidPurchase Then
        MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
    End If
End With
Exit Function
err:
    ErrorMessage
End Function

Private Function RecordExists() As Boolean
On Error GoTo err
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT COUNT(SerialNO) AS TCount FROM PurchaseOrder WHERE DrugCode='" & Trim(.txtProductCode.Text) & "' AND PurchaseOrderNo='" & Trim(.txtOrderNO.Text) & "';")
    
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
err:
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
        Me.txtProductDescription.SetFocus
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
On Error GoTo err
        
'        If Me.ListView2.ListItems.Count = 0 Then
'            MyPurchase.GetPurchaseStructure
'        End If
'
'        If Me.ListView1.ListItems.Count = 0 Then
'            MyPurchase.GetSupplyStructure
'        End If
        ShowALLACTIVESuppliers
        LoadDEFAULTS
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub Form_Initialize()
    Set MyPurchase = New clsODASLPORequisition
End Sub

Private Sub LoadDEFAULTS()
On Error GoTo err
    With frmODASMPurchaseOrders
            .dtpOrderDate.Value = Date
            .dtpPromised.Value = Date
            .dtpRequired.Value = Date
            .txtDateRequired.Text = Date
            .txtDatePromised.Text = Date
            .txtOrderDate.Text = Date
    End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub Form_Resize()
With Me
    .ListView1.Width = .Width - (12000 - 6135)
    .ListView2.Width = .Width - (12000 - 6135)
    .ListView1.Height = .Height - (8505 - 3615)
    .Frame1.Height = .Height - (8505 - 5655)
    .Frame1.Left = .ListView1.Width + 100
    .Text5.Width = .Width - (12000 - 11895)
    .Label3.Left = .Frame1.Left
End With
End Sub

Private Sub Form_Terminate()
    Set MyPurchase = Nothing
End Sub

Private Sub ListView1_DblClick()
On Error GoTo err
        With Me
        If Not NewRecord And Not EditRecord Then: Exit Sub
            
            If .ListView1.ListItems.Count = 0 Or .ListView1.View <> lvwReport Then Exit Sub
            
            Dim i, j, k
            j = .ListView1.ListItems.Count
            
            .ListView1.SelectedItem.Checked = True
            
            .txtProductCode.Text = .ListView1.SelectedItem.Text
            .txtProductDescription.Text = .ListView1.SelectedItem.SubItems(1)
        '    .txtDosageType.Text = .ListView1.SelectedItem.SubItems(3)
            
            MyPurchase.GetDrugDataByCode
            
            For i = 1 To j
                If .ListView1.ListItems(i).Text <> Trim(.txtProductCode.Text) Then
                    .ListView1.ListItems(i).Checked = False
                End If
            Next i
            
        '    .txtQuantity.SetFocus
            
        End With
Exit Sub
err:
    ErrorMessage
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
    
        MyPurchase.ClearForNewDrug
        Me.txtProductCode.Text = Item
        Me.txtProductDescription.Text = Item.SubItems(1)
        Me.cboUnits.Text = Item.SubItems(2)
        Me.txtDosageCost.Text = Item.SubItems(2)
        MyPurchase.GetDrugDataByCode
        Me.txtTotalDoses.SetFocus
        
    ElseIf Item.Checked = False Then
    
        MyPurchase.ClearForNewDrug
        
    End If
    
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
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
        Me.cboSupplierCode.Text = Item
        MyPurchase.GetSupplierDataByCode
    ElseIf Item.Checked = False Then
        MyPurchase.ClearSupplierData
    End If
    
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub mnuFileClearScreen_Click()
    MyPurchase.ClearTheScreen
End Sub

Private Sub mnuFileClose_Click()
    Unload Me
End Sub


Private Sub mnuPrintCoveringLetter_Click()

End Sub

Private Sub mnuPrintOrderForm_Click()
On Error GoTo err
With Me
If .txtOrderNO.Text = Empty Then
    MsgBox "Required Purchase Order Number to Print the Report!!", vbCritical + vbOKOnly, "Missing Order No"
    .txtOrderNO.SetFocus
ElseIf Not ValidOrderNumber Then
    MsgBox "The Purchase Order Number: " & " [" & .txtOrderNO.Text & "] " & "does not exist or has been deleted! No work was done!!", vbCritical + vbOKOnly, "Invalid Order Number"
    .txtOrderNO.SetFocus: .txtOrderNO.SelStart = 0: .txtOrderNO.SelLength = Len(.txtOrderNO.Text)
Else
    Load frmRPTPurchaseOrderForm
    frmRPTPurchaseOrderForm.Show 1, Me
End If
End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Function ValidOrderNumber() As Boolean
On Error GoTo err
    Set rsFindRecord = cnCOMMON.Execute("SELECT COUNT(PurchaseOrderNo) AS TOrders FROM Orders WHERE PurchaseOrderNO='" & Trim(Me.txtOrderNO.Text) & "';")
    
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
err:
    ErrorMessage
End Function

Private Sub mnuShowCurrentPurchaseOrders_Click()
MyPurchase.ShowCurrentOrderRecordsFromMenu
End Sub

Private Sub mnuShowNewInventory_Click()
'If Not NewRecord Then Exit Sub
    MyPurchase.ShowNewProducts
End Sub

Private Sub mnuShowPurchaseRecords_Click()

If NewRecord Or EditRecord Then Exit Sub
    MyPurchase.ShowReOrderItems
End Sub

Private Sub mnuShowSuppliers_Click()
If NewRecord Or EditRecord Then Exit Sub
    MyPurchase.ShowPossibleSuppliers
End Sub

'Private Sub mnuViewRefresh_Click()
'On Error GoTo Err
'If NewRecord Or EditRecord Then Exit Sub
'With Me
'    If .ListView1.ListItems.Count = 0 Or .ListView1.View <> lvwReport Then Exit Sub
'        MyPurchase.ShowCurrentOrderRecordsFromMenu
'    If .ListView2.ListItems.Count = 0 Or .ListView2.View <> lvwReport Then Exit Sub
'        MyPurchase.ShowPossibleSuppliers
'End With
'Exit Sub
'Err:
'    ErrorMessage
'End Sub

Private Sub mnuViewShowAll_Click()
If NewRecord Or EditRecord Then Exit Sub
    MyPurchase.ShowPurchaseRecords
    MyPurchase.ShowPossibleSuppliers
End Sub

Private Sub Text5_GotFocus()
    Me.txtOrderNO.SetFocus
End Sub

Private Sub txtOrderDescription_LostFocus()
    Me.txtOrderDescription.Text = StrConv(Me.txtOrderDescription.Text, vbProperCase)
End Sub

Private Sub txtOrderNO1_GotFocus()
Me.txtOrderNO.Locked = False
End Sub

Private Sub txtOrderNO_GotFocus()
If Me.txtOrderNO.Text <> Empty Or Not NewRecord Then Exit Sub
    If Me.chkOrderNo.Value = 1 Then
        Me.txtOrderNO.Text = MyPurchase.AutoPurchaseOrderNo
        Me.txtOrderDescription.SetFocus
    Else
        Me.txtOrderNO.SetFocus
    End If
End Sub

Private Sub txtOrderNO_KeyPress(KeyAscii As Integer)
On Error GoTo err
With Me
If NewRecord Then
    If .chkOrderNo.Value = 1 Then
        KeyAscii = 0
    Else
        Exit Sub
    End If
Else
    If KeyAscii = vbKeyReturn Then
        MyPurchase.FindOrderDetails
    Else
        Exit Sub
    End If
End If
End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub txtOrderNO_LostFocus()
    If NewRecord Then
        Me.txtOrderDate.Text = Date
        Me.txtOrderDescription.SetFocus
    End If
End Sub

Private Sub ComputeTotalCost()
On Error GoTo err
If Not NewRecord And Not EditRecord Then Exit Sub
If Me.txtTotalDoses.Text = Empty Then Exit Sub

    Dim TDoses As Double, DCost As Double, TCost As Double
    
    TDoses = CDbl(Me.txtTotalDoses.Text): DCost = CDbl(Me.txtDosageCost.Text)
    
    TCost = TDoses * DCost
    
    Me.txtTotalCost.Text = FormatNumber(TCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    
    Me.cmdSAVE.SetFocus
    
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ComputeTotalQuantity()
'On Error GoTo Err
With frmODASMPurchaseOrders
    
    Set rsFindRecord = cnCOMMON.Execute("SELECT Quantity,QuantityUnits FROM ProductQuantitySetup WHERE DrugCode='" & Trim(.txtProductCode.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    ElseIf IsNull(rsFindRecord!quantity) = True Or rsFindRecord!quantity = "" Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
    
        Dim d, qd, q
        d = CDbl(.txtTotalDoses.Text)
        qd = CDbl(rsFindRecord!quantity)
        q = CDbl(d * qd)
        
'        .txtQuantity.Text = q
        .cboUnits.Text = rsFindRecord!QuantityUnits & ""
        
    End If
    
    Set rsFindRecord = Nothing
    
End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub txtStaffIdNo_Change()

End Sub

Private Sub txtTotalDoses_LostFocus()
If Not NewRecord And Not EditRecord Then Exit Sub
With Me
If .txtTotalDoses.Text = Empty Then Exit Sub
'    Call ComputeTotalQuantity
    Call ComputeTotalCost
End With
End Sub

Private Sub VScroll1_Change()
If Not NewRecord And Not EditRecord Then Exit Sub
With Me
    .txtTotalDoses.Text = .VScroll1.Value
    Call ComputeTotalQuantity
    Call ComputeTotalCost
End With
End Sub
