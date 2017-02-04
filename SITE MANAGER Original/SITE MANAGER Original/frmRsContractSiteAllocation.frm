VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmRsContractSiteAllocation 
   Caption         =   "Resource Scheduling - CONTRACTS SITE ALLOCATION"
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRsContractSiteAllocation.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   Picture         =   "frmRsContractSiteAllocation.frx":0442
   ScaleHeight     =   7815
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView ListView3 
      Height          =   1935
      Left            =   0
      TabIndex        =   41
      Top             =   5640
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3413
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
      NumItems        =   0
   End
   Begin VB.TextBox txtTitle 
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
      TabIndex        =   38
      Text            =   "***  SITE ALLOCATION ***"
      Top             =   0
      Width           =   11895
   End
   Begin VB.ComboBox cboOrderNO 
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
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtOrderDescription 
      Appearance      =   0  'Flat
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
      Height          =   390
      Left            =   4560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   960
      Width           =   7215
   End
   Begin VB.TextBox txtOrderDate 
      Appearance      =   0  'Flat
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
      Height          =   390
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   6960
      TabIndex        =   20
      Top             =   1680
      Width           =   4815
      Begin VB.TextBox txtClientPhysicalAd 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   2280
         TabIndex        =   55
         Top             =   4320
         Width           =   2295
      End
      Begin VB.TextBox txtContractExpiryDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   3240
         TabIndex        =   54
         Top             =   4800
         Width           =   1335
      End
      Begin VB.TextBox txtContractStartDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   330
         Left            =   1080
         TabIndex        =   52
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox txtDays 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   3240
         TabIndex        =   50
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox txtDuration 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   1080
         TabIndex        =   49
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox txtAdvWidth 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   330
         Left            =   3240
         TabIndex        =   44
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtAdvName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   1080
         TabIndex        =   43
         Top             =   2400
         Width           =   3495
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H8000000A&
         Caption         =   "&PRINT"
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
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   5280
         Width           =   1455
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H8000000A&
         Caption         =   "&REFRESH"
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
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   5280
         Width           =   1935
      End
      Begin VB.TextBox txtAdvCost 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtClientName 
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
         TabIndex        =   14
         Top             =   3840
         Width           =   3495
      End
      Begin VB.TextBox txtPhysicalAddress 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   1095
         TabIndex        =   9
         Top             =   1030
         Width           =   3615
      End
      Begin VB.TextBox txtSiteNo 
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
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtSiteName 
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
         TabIndex        =   7
         Top             =   575
         Width           =   1575
      End
      Begin VB.TextBox txtBBNo 
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
         Left            =   3480
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtCity 
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
         Left            =   3480
         TabIndex        =   8
         Top             =   575
         Width           =   1215
      End
      Begin VB.TextBox txtSerialNO 
         Alignment       =   2  'Center
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtAdvCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   345
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtClientCode 
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
         TabIndex        =   15
         Top             =   4320
         Width           =   1095
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
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5280
         Width           =   1455
      End
      Begin VB.TextBox txtAdvLength 
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
         TabIndex        =   13
         Top             =   2838
         Width           =   1215
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF8080&
         X1              =   120
         X2              =   4680
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label Label25 
         Caption         =   "ExpDate"
         Height          =   255
         Left            =   2520
         TabIndex        =   53
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label Label24 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   4800
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "Days"
         Height          =   255
         Left            =   2640
         TabIndex        =   48
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label22 
         Caption         =   "Duration"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Width"
         Height          =   255
         Left            =   2640
         TabIndex        =   46
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   45
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "S.N"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "Advertisement Cost."
         Height          =   255
         Left            =   2640
         TabIndex        =   36
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label19 
         Caption         =   "Client Name"
         Height          =   255
         Left            =   0
         TabIndex        =   35
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Physical Ad."
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1035
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Site No."
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Site Name"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   575
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "BB No."
         Height          =   255
         Left            =   2760
         TabIndex        =   31
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "City"
         Height          =   255
         Left            =   2760
         TabIndex        =   30
         Top             =   570
         Width           =   615
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000002&
         X1              =   120
         X2              =   4680
         Y1              =   1485
         Y2              =   1485
      End
      Begin VB.Label Label13 
         Caption         =   "Adv Name"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Adv Code"
         Height          =   255
         Left            =   1200
         TabIndex        =   28
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "ClientCode"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Length"
         Height          =   255
         Left            =   120
         TabIndex        =   24
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
      Top             =   1680
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
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   7560
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   4304
            MinWidth        =   4304
            Picture         =   "frmRsContractSiteAllocation.frx":5B80
            Text            =   "Site Allocation"
            TextSave        =   "Site Allocation"
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
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   635
      ButtonWidth     =   1138
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1815
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Select a Record Here by Checking or Double-Clicking on it!"
      Top             =   3600
      Width           =   6855
      _ExtentX        =   12091
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
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "List of available sites"
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
      TabIndex        =   42
      Top             =   5400
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Adverts under selected contract"
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
      TabIndex        =   26
      Top             =   3360
      Width           =   5295
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   0
      X2              =   11880
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Contract No"
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
      Top             =   720
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
      TabIndex        =   23
      Top             =   720
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
      TabIndex        =   22
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Particulars"
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
      TabIndex        =   21
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "List of current contracts"
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
         Caption         =   "All New Contracts"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnujn 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowExpectedToday 
         Caption         =   "Contracts Due Today"
      End
      Begin VB.Menu mnullk 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowOverdue 
         Caption         =   "&Overdue Contracts"
      End
      Begin VB.Menu mnuMGHJGGFDF 
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
         Caption         =   "List of Contract Allocated Adverts"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnukkk 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintNotReceived 
         Caption         =   "List of Contract Unallocated Adverts"
         Shortcut        =   ^Z
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
Attribute VB_Name = "frmRsContractSiteAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyReceiver As clsRsContractSiteAllocation, CurrentBBNo

Private Sub cboOrderNO_Click()
'If Not NewRecord Then Exit Sub
    MyReceiver.ClearTextFields
    
End Sub

Private Sub cboOrderNO_GotFocus()
'If Not NewRecord Then Exit Sub
    MyReceiver.AttachNewPurchaseOrders
End Sub

Private Sub cboOrderNO_LostFocus()
'If Not NewRecord Then Exit Sub
    MyReceiver.FindDetailsByPurchaseOrderNo
End Sub


Private Sub cmdRefresh_Click()
Call mnuShowRefresh_Click
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
            If ValidRecord Then
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
Private Function ValidRecord() As Boolean
With Me
    For Each i In Me
    If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
        If i.Text = Empty And i.Name <> "txtOrderDescription" Then
            MsgBox "All the Fields are Required. Please Enter the Missing Data...!", vbCritical + vbOKOnly, "Data Validation"
            i.SetFocus: ValidRecord = False: Exit Function
        End If
    End If
    Next i
    ValidRecord = True
End With
End Function

Private Sub Form_Activate()
If Me.ListView1.ListItems.Count = 0 Then
    MyReceiver.GetPurchaseStructure
End If
If Me.ListView2.ListItems.Count = 0 Then
    MyReceiver.GetRecordsStructure
End If

If Me.ListView3.ListItems.Count = 0 Then
   Call GetSiteStructure
 End If
End Sub

Private Sub Form_Initialize()
    Set MyReceiver = New clsRsContractSiteAllocation
End Sub
Private Sub GetSiteStructure()
With Me
'On Error GoTo Err
.ListView3.ListItems.Clear
.ListView3.ColumnHeaders.Clear

.ListView3.ColumnHeaders.Add , , "Site No", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Site Name ", .ListView3.Width / 4.5 ', lvwColumnCenter
.ListView3.ColumnHeaders.Add , , "Billboard Number", .ListView3.Width / 6 ', lvwColumnCenter
.ListView3.ColumnHeaders.Add , , "City", .ListView3.Width / 5 ', lvwColumnCenter
.ListView3.ColumnHeaders.Add , , "Physical Address", .ListView3.Width / 1 ', lvwColumnCenter
.ListView3.ColumnHeaders.Add , , "Allocated", .ListView3.Width / 1 ', lvwColumnCenter
.ListView3.ColumnHeaders.Add , , "Valid", .ListView3.Width / 1 ', lvwColumnCenter

.ListView3.View = lvwReport
End With
End Sub


Private Sub Form_Resize()
With Me
'    .ListView1.Width = .Width - (12000 - 6855)
'    .ListView2.Width = .Width - (12000 - 6855)
'    .ListView2.Height = .Height - (8505 - 3735)
'    .Text5.Width = .Width - (12000 - 11895)
'    .Frame1.Height = .Height - (8505 - 5655)
'    .Frame1.Left = .ListView1.Width + 100
'    .txtOrderDescription.Width = .Width - (12000 - 7215)
'    .Label3.Left = .Frame1.Left
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
        Me.txtContractStartDate.Text = Item.SubItems(2)
        Me.txtContractExpiryDate.Text = Item.SubItems(3)
        Me.txtClientCode.Text = Item.SubItems(4)
        Me.txtClientName.Text = Item.SubItems(5)
        Me.txtClientPhysicalAd.Text = Item.SubItems(6)
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
        
        MyReceiver.ClearTextFields
        
        Me.txtSerialNo.Text = Item
        Me.txtAdvCode.Text = Item.SubItems(1)
        Me.txtAdvName.Text = Item.SubItems(2)
        Me.txtAdvLength.Text = Item.SubItems(4)
        Me.txtAdvWidth.Text = Item.SubItems(5)
        Me.txtDuration.Text = Item.SubItems(6)
        Me.txtDays.Text = Item.SubItems(7)
        Me.txtAdvCost.Text = Item.SubItems(8)
        CurrentBBNo = Item.SubItems(1)
        Call GetAvailableSites

        
        
    ElseIf Item.Checked = False Then
    
        MyReceiver.ClearTextFields
        
    End If
    
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub ListView3_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo Err
 If Not NewRecord Then Exit Sub
 Me.ListView2.ListItems.Clear
If Me.ListView3.ListItems.Count = 0 Or Me.ListView3.View <> lvwReport Then Item.Checked = False: Exit Sub
    
    Dim i, j, k
    j = Me.ListView3.ListItems.Count
    
    If j = 0 Then Exit Sub
    
    For i = 1 To j
        If Me.ListView3.ListItems(i).Text <> Item Then
            Me.ListView3.ListItems(i).Checked = False
        End If
    Next i
    
    If Item.Checked = True Then
        CurrentOrder = Item
        Me.txtSiteNo.Text = Item
        Me.txtSiteName.Text = Item.SubItems(1)
        Me.txtBBNo.Text = Item.SubItems(2)
        Me.txtCity.Text = Item.SubItems(3)
        Me.txtPhysicalAddress.Text = Item.SubItems(4)
        Me.txtOrderDescription.Text = MyReceiver.GetMyOrderDescription
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



Private Sub mnuFileClearScreen_Click()
Call mnuShowRefresh_Click
End Sub

Private Sub mnuFileClose_Click()
Unload Me
End Sub

Private Sub mnuPrintListReceived_Click()
'On Error GoTo Err
With Me
    INPQRY = InputBox("Please Enter Contract No!!", "Allocated Adverts")
    
    If Len(INPQRY) = 0 Then
        MsgBox "No Values Entered or Operation Was Cancelled! No Work Will Be Done!!"
        Exit Sub
    Else
        
        Dim ThisProduct As String
        ThisProduct = (INPQRY)
    
        Set rsFindRecord = cnCOMMON.Execute("SELECT *  FROM ContractSiteAllocationData WHERE ContractNo='" & ThisProduct & "';")
        
        If rsFindRecord.EOF And rsFindRecord.BOF Then
            MsgBox "The Contract Number Does Not Exist Or Entry is Not Correct!!", vbCritical + vbOKOnly, "Invalid Product Name"
            Set rsFindRecord = Nothing: Exit Sub
        ElseIf rsFindRecord!DateCreated = 0 Then
            MsgBox "The Contract Number Does Not Exist Or Entry is Not Correct!!", vbCritical + vbOKOnly, "Invalid Date"
            Set rsFindRecord = Nothing: Exit Sub
        Else
            Set rsFindRecord = Nothing
            
            SelectedProduct = (INPQRY)
            
            
            Load frmRPTContractAllocatedAdverts
            frmRPTContractAllocatedAdverts.Show 1, Me
        End If
    End If
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub mnuPrintNotReceived_Click()
'On Error GoTo Err
With Me
    INPQRY = InputBox("Please Enter Contract No!!", "Not Allocated Adverts")
    
    If Len(INPQRY) = 0 Then
        MsgBox "No Values Entered or Operation Was Cancelled! No Work Will Be Done!!"
        Exit Sub
    Else
        
        Dim ThisProduct As String
        ThisProduct = (INPQRY)
        
    
        Set rsFindRecord = cnCOMMON.Execute("SELECT *  FROM AdvertContractRequisition A,AdvertContractRequisitionData B WHERE A.PurchaseOrderNo = B.PurchaseOrderNo AND B.PurchaseOrderNo='" & ThisProduct & "'AND B.AllocationStatus IS NULL;")
        
        If rsFindRecord.EOF And rsFindRecord.BOF Then
            MsgBox "The Contract Number Does Not Exist Or Entry is Not Correct!!", vbCritical + vbOKOnly, "Invalid Product Name"
            Set rsFindRecord = Nothing: Exit Sub
        ElseIf rsFindRecord!DateCreated = 0 Then
            MsgBox "The Contract Number Does Not Exist Or Entry is Not Correct!!", vbCritical + vbOKOnly, "Invalid Date"
            Set rsFindRecord = Nothing: Exit Sub
        Else
            Set rsFindRecord = Nothing
            
            SelectedProduct = (INPQRY)
            
            
            Load frmRPTContractUnAllocatedAdverts
            frmRPTContractUnAllocatedAdverts.Show 1, Me
        End If
    End If
End With
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
    MyReceiver.ClearTextFields
    .ListView1.ListItems.Clear
    .ListView2.ListItems.Clear
    .ListView3.ListItems.Clear
    .cmdSAVE.Enabled = True
    .cmdSAVE.Caption = "&NEW"
'    .cmdCHANGE.Enabled = True
'    .cmdCHANGE.Caption = "&CHANGE"
    
    
End With
Exit Sub
Err:
    ErrorMessage
End Sub


Private Sub Text5_GotFocus()
    Me.cboOrderNO.SetFocus
End Sub


Public Sub GetAvailableSites()
'On Error GoTo Err
Dim Today As Variant, SelectedBillBoardNo As String
Today = Format(Date, "MMMM dd,yy")
With frmRsContractSiteAllocation
.ListView3.ListItems.Clear
.ListView3.ColumnHeaders.Clear

.ListView3.ColumnHeaders.Add , , "Site No", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Site Name ", .ListView3.Width / 4 ', lvwColumnCenter
.ListView3.ColumnHeaders.Add , , "Billboard Number", .ListView3.Width / 6 ', lvwColumnCenter
.ListView3.ColumnHeaders.Add , , "City", .ListView3.Width / 5 ', lvwColumnCenter
.ListView3.ColumnHeaders.Add , , "Physical Address", .ListView3.Width / 4 ', lvwColumnCenter
.ListView3.ColumnHeaders.Add , , "Allocated", .ListView3.Width / 6.5 ', lvwColumnCenter
.ListView3.ColumnHeaders.Add , , "Valid", .ListView3.Width / 6.5 ', lvwColumnCenter

.ListView3.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertSites WHERE BBNo='" & Trim(.txtAdvCode.Text) & "' AND AllocationStatus  IS NULL  AND ValidStatus = '" & "Y" & "' AND Approvedstatus = '" & "Y" & "'AND Discontinued = '" & "N" & "' ORDER BY BBNO;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView3.View = lvwList
    Set MyList = .ListView3.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing
    
    Load frmRPTSiteStatistics
    frmRPTSiteStatistics.Show 1, Me
    Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!SiteNo))

    If Not IsNull(rsLIST!SiteName) Then
        MyList.SubItems(1) = CStr(rsLIST!SiteName)
    End If
    
    If Not IsNull(rsLIST!BBNo) Then
        MyList.SubItems(2) = CStr(rsLIST!BBNo)
    End If
    
    If Not IsNull(rsLIST!city) Then
        MyList.SubItems(3) = CStr(rsLIST!city)
    End If
    
    If Not IsNull(rsLIST!sitephysicalAddress) Then
        MyList.SubItems(4) = CStr(rsLIST!sitephysicalAddress)
    End If
    
               
    If IsNull(rsLIST!AllocationStatus) Then
        MyList.SubItems(5) = CStr("NO")
    ElseIf Not IsNull(rsLIST!AllocationStatus) Then
        If rsLIST!AllocationStatus = "Y" Then
            MyList.SubItems(5) = CStr("YES")
        Else
            MyList.SubItems(5) = CStr("NO")
        End If
    End If
    
    
    If IsNull(rsLIST!Validstatus) Then
        MyList.SubItems(6) = CStr("NO")
    ElseIf Not IsNull(rsLIST!Validstatus) Then
        If rsLIST!Validstatus = "Y" Then
            MyList.SubItems(6) = CStr("YES")
        Else
            MyList.SubItems(6) = CStr("NO")
        End If
    End If
    
    
    rsLIST.MoveNext
    
Wend

'.ListView3.ColumnHeaders(6).Alignment = lvwColumnRight
'.ListView3.ColumnHeaders(7).Alignment = lvwColumnRight

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

