VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReceiveCustomerPayments 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resource Scheduling -RECEIVE CUSTOMER'S PAYMENTS"
   ClientHeight    =   8625
   ClientLeft      =   -270
   ClientTop       =   2355
   ClientWidth     =   11910
   Icon            =   "frmReceiveCustomerPayments.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboReceivedFrom 
      Height          =   315
      Left            =   1680
      TabIndex        =   44
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txtContractNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8040
      TabIndex        =   43
      Top             =   3480
      Width           =   3495
   End
   Begin VB.TextBox txtTaxAmount 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5760
      TabIndex        =   41
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtTotalAmount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8040
      TabIndex        =   38
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CheckBox chkAllowDiscount 
      BackColor       =   &H80000009&
      Caption         =   "Allow Discount"
      Height          =   255
      Left            =   7800
      TabIndex        =   37
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtDiscountValue 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      TabIndex        =   36
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtDiscountedAmnt 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1680
      TabIndex        =   34
      Top             =   2280
      Width           =   2775
   End
   Begin VB.ComboBox cboDiscountAllowed 
      Height          =   315
      Left            =   6480
      TabIndex        =   32
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox cboTax 
      Height          =   315
      Left            =   9720
      TabIndex        =   29
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox txtSerialNo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10200
      TabIndex        =   28
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtBalanceDue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      TabIndex        =   26
      Top             =   1800
      Width           =   2175
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   56
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":0984
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":0EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":1408
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":194A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":1E8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":2026
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":21C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":235A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":24F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":268E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":2EC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":36F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":388C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":3A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":3D58
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":458A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":4724
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":48BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":4A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":4E0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":509C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":5196
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":5290
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":538A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":5484
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":59C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":5AD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":5BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":5CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":5E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":5F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":6032
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":6144
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":6256
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":6368
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":647A
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":658C
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":669E
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":67B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":68C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":69D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":6AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":6BF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":6D0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":6E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":6F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":7040
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":7152
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":7264
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":7376
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":7488
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":759A
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":76AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":77BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReceiveCustomerPayments.frx":78D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      Caption         =   "List Of Advertisements"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   6240
      TabIndex        =   20
      Top             =   4080
      Width           =   5535
      Begin MSComctlLib.ListView ListView2 
         Height          =   4095
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   7223
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   "List of Contracts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   19
      Top             =   4080
      Width           =   6015
      Begin VB.TextBox txtTotalItems 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9720
         TabIndex        =   27
         Top             =   1560
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4095
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   7223
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
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1680
      TabIndex        =   18
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   2760
      Width           =   2655
   End
   Begin VB.ComboBox cboPaymentMethod 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5640
      TabIndex        =   16
      Top             =   1320
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   255
      Left            =   11280
      TabIndex        =   13
      Top             =   2280
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Format          =   64421889
      CurrentDate     =   38210
   End
   Begin VB.TextBox txtExpiryDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9720
      TabIndex        =   12
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtChequeNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7920
      TabIndex        =   11
      Top             =   2280
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   9360
      TabIndex        =   8
      Top             =   1560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Format          =   64421889
      CurrentDate     =   38210
   End
   Begin VB.TextBox txtInvoiceDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7920
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtInvoiceNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9720
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   714
      ButtonWidth     =   2090
      ButtonHeight    =   556
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&NEW"
            Key             =   "N"
            ImageIndex      =   26
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&FINISH"
            Key             =   "F"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&EDIT"
            Key             =   "E"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&REFRESH"
            Key             =   "R"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&PRINT"
            Key             =   "P"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&SEARCH"
            Key             =   "S"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&VIEW"
            ImageIndex      =   36
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "CONTRACT NUMBER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   8040
      TabIndex        =   42
      Top             =   3240
      Width           =   3495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   40
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "TOTAL AMOUNT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   8040
      TabIndex        =   39
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Line Line5 
      X1              =   5760
      X2              =   7800
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Discount Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4440
      TabIndex        =   35
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Line Line4 
      X1              =   1680
      X2              =   4320
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Discounted Amnt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   33
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Discount Allowed (Days)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   31
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "TAX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9720
      TabIndex        =   30
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Line Line3 
      X1              =   5760
      X2              =   7800
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line2 
      X1              =   1680
      X2              =   4200
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      X1              =   1680
      X2              =   4320
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Balance Due"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4440
      TabIndex        =   25
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Payment Meth."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   24
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Received From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Memo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "EXPIRY DATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9720
      TabIndex        =   10
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "CHECK #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7920
      TabIndex        =   9
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "PAYMENT #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9720
      TabIndex        =   6
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7920
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Customer Receipts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   11535
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000008&
      Height          =   3375
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   11535
   End
   Begin VB.Menu mnuEditOptions 
      Caption         =   "EditOptions"
      Begin VB.Menu mnuEditMain 
         Caption         =   "EditMain"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
      End
   End
End
Attribute VB_Name = "frmReceiveCustomerPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim MyCommonData As clsCommonData, EditOption



Private Sub cboDepositTo_Change()

End Sub

Private Sub cboDiscountAllowed_Change()
On Error Resume Next
With Me
If .txtAmount.Text = "" Then Exit Sub
.txtDiscountValue.Text = FormatNumber(CDbl(.cboDiscountAllowed.Text) / 100 * (.txtAmount.Text), 2, vbUseDefault, vbUseDefault, vbTrue)
.txtDiscountedAmnt.Text = FormatNumber(CDbl(.txtAmount.Text) - (.txtDiscountValue.Text), 2, vbUseDefault, vbUseDefault, vbTrue)
End With
End Sub

Private Sub cboDiscountAllowed_Click()
Me.txtInvoiceDate.SetFocus
End Sub

Private Sub cboReceivedFrom_Click()
Me.cboPaymentMethod.SetFocus
End Sub

Private Sub cboTax_Click()
Me.ListView2.SetFocus
End Sub

Private Sub cboTax_GotFocus()
'On Error GoTo Err
'If Not NewRecord Or Not EditRecord Then Exit Sub
With Me
      If .txtAmount.Text = "" Then Exit Sub
        Set rsCOMBO = cnCOMMON.Execute("SELECT Description FROM ParamTaxes;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    .cboTax.Clear
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!Description) And rsCOMBO!Description <> "" Then
        .cboTax.AddItem rsCOMBO!Description
        End If
        
    rsCOMBO.MoveNext
    Loop
'    .cboPaymentMethod.AddItem "<Add New>"
    
OUTS:
    Set rsCOMBO = Nothing
End With
Exit Sub
Err:
    ErrorMessage

End Sub

Private Sub cboTax_LostFocus()
'On Error Resume Next
With Me

    Set rsFindRecord = cnCOMMON.Execute("SELECT *  FROM ParamTaxes  WHERE Description ='" & Trim(.cboTax.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        
        .cboTax.Text = rsFindRecord!TaxRate & ""
        .txtTaxAmount.Text = FormatNumber(CDbl(.txtTotalAmount.Text) * CDbl(.cboTax.Text) / 100, 2, vbUseDefault, vbUseDefault, vbTrue)
        .txtAmount.Text = FormatNumber(CDbl(.txtTotalAmount.Text) + CDbl(.txtTotalAmount.Text) * CDbl(.cboTax.Text) / 100, 2, vbUseDefault, vbUseDefault, vbTrue)
        
    End If
        
End With
Exit Sub

End Sub

Private Sub chkAllowDiscount_Click()
With Me
If .chkAllowDiscount.Value = 0 Then
   .cboDiscountAllowed.Enabled = False
ElseIf .chkAllowDiscount.Value = 1 Then
    .cboDiscountAllowed.Enabled = True
    End If
    End With
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
        Me.txtContractNo.Text = Item
        Me.cboReceivedFrom.Text = Item.SubItems(5)
        Me.txtAmount.Text = Item.SubItems(6)
        Me.txtTotalAmount = Item.SubItems(6)
        Me.txtExpiryDate = Item.SubItems(3)
        
        MyCommonData.ShowItemsInCurrentOrder
        Me.ListView2.SetFocus
   
    ElseIf Item.Checked = False Then
        Me.cboReceivedFrom.Clear
        Me.txtAmount.Text = Empty
        Me.ListView2.ListItems.Clear
    End If
    
Exit Sub
Err:
    ErrorMessage

End Sub

Private Sub mnuEditMain_Click()
    EditOption = 1
End Sub
Private Sub cboReceivedFrom_GotFocus()

'On Error GoTo Err
If Not NewRecord Then Exit Sub
With frmReceiveCustomerPayments
        Set rsCOMBO = cnCOMMON.Execute("SELECT CompanyName FROM AdvertClients A,AdvertcontractRequisition B WHERE A.CustomerId = B.ClientCode AND B.PaidStatus IS NULL ORDER BY A.CompanyName;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    .cboReceivedFrom.Clear
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!CompanyName) And rsCOMBO!CompanyName <> "" Then
        .cboReceivedFrom.AddItem rsCOMBO!CompanyName
        End If
        
    rsCOMBO.MoveNext
    Loop
'    .cboPaymentMethod.AddItem "<Add New>"
    
OUTS:
    Set rsCOMBO = Nothing
End With
Exit Sub
Err:
    ErrorMessage

End Sub


Private Sub cboReceivedFrom_LostFocus()
Dim ContractNo, ClientCode, Today As Variant
'On Error GoTo Err
With Me

    Set rsFindRecord = cnCOMMON.Execute("SELECT *  FROM AdvertContractRequisition WHERE ClientName ='" & Trim(.cboReceivedFrom.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        
'         .cboReceivedFrom.Text = rsFindRecord!PurchseOrderNo & ""
         .txtInvoiceDate = MyCurrentDate
'         .txtDueDate = rsFindRecord!StartDate & ""
'         .txtBalanceDue = rsFindRecord!Totalbalance & ""
    End If
    Screen.MousePointer = vbHourglass

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 5.5  ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Order Date", .ListView1.Width / 5.5  ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Start Date", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "End Date", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Client Code", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Client Name", .ListView1.Width / 4.5
.ListView1.ColumnHeaders.Add , , "Amount Due", .ListView1.Width / 5.5

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertContractRequisition WHERE ClientName = '" & Trim(.cboReceivedFrom.Text) & "' AND PaidStatus IS NULL AND ApprovedStatus = '" & "Y" & "' AND EndDate > '" & MyCurrentDate & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem, NCount As Double

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Screen.MousePointer = vbDefault: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!PurchaseOrderNo))
    
    If Not IsNull(rsLIST!DateCreated) Then
        MyList.SubItems(1) = CStr(rsLIST!DateCreated)
    End If
    If Not IsNull(rsLIST!StartDate) Then
        MyList.SubItems(2) = CStr(rsLIST!StartDate)
    End If
    
    If Not IsNull(rsLIST!EndDate) Then
        MyList.SubItems(3) = CStr(rsLIST!EndDate)
    End If
    
    If Not IsNull(rsLIST!ClientCode) Then
        MyList.SubItems(4) = CStr(rsLIST!ClientCode)
    End If
    
    If Not IsNull(rsLIST!ClientName) Then
        MyList.SubItems(5) = CStr(rsLIST!ClientName)
    End If
    If Not IsNull(rsLIST!TotalCost) Then
        MyList.SubItems(6) = FormatNumber(rsLIST!TotalCost, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    If Not IsNull(rsLIST!TotalBalance) Then
        MyList.SubItems(7) = FormatNumber(rsLIST!TotalBalance, 2, vbUseDefault, vbUseDefault, vbTrue)
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

Private Sub cboPaymentMethod_Click()
Me.txtAmount.SetFocus
End Sub



Private Sub Form_Load()
On Error Resume Next
Set MyCommonData = New clsCommonData
Call GetItemListStructure
Call GetCurrentInvoiceItems
Me.mnuEditOptions.Visible = False
Me.ListView1.Visible = True
Me.ListView2.Visible = True
Me.Frame1.Visible = True
Me.Frame2.Visible = True
Call chkAllowDiscount_Click

End Sub

Private Sub Form_Terminate()
'    Set MyCommonData = Nothing
    End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo Err
    If Save = True Or Edit = True Then
        MsgBox "Please there is Work going on, Refresh to continue", vbOKCancel + vbCritical
        Cancel = 1
    Else
        Found = False

    End If
Exit Sub

Err:
ErrorMessage
End Sub

Private Sub List1_Click()

End Sub

Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)

'On Error GoTo Err
With Me
j = .ListView2.ListItems.Count
If j = 0 Or .ListView2.View <> lvwReport Then Item.Checked = False: Exit Sub
'If Not NewRecord And Not EditRecord Then Item.Checked = False: Exit Sub
    If Item.Checked = True Then

        For i = 1 To j
        If .ListView2.ListItems(i).Text <> Item Then
            .ListView2.ListItems(i).Checked = False
        End If
        Next i
        .txtSerialNo.Text = Item
        .txtAmount.Text = Item.SubItems(3)
        .txtTotalAmount.Text = Item.SubItems(3)
        .txtDescription.Text = Item.SubItems(2)
        .txtContractNo.Text = Item.SubItems(8)
        Me.txtExpiryDate = Item.SubItems(5)
        .txtInvoiceDate.Text = Date
     
    ElseIf Item.Checked = False Then
'        Call ClearData
    End If
End With
Exit Sub
Err:
    ErrorMessage

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo Err
Dim SQL1, SQL2
Dim Taxed, Discounted, HasBalance As Variant
With Me
Select Case Button.Key
Case "N"
    Select Case Button.Caption
    Case "&NEW"
        If EditRecord Then Exit Sub
        NewRecord = True
        MyCommonData.ClearTextFields: .txtInvoiceNo = GetNextInvoiceNo
        Button.Caption = "&SAVE": Button.Image = 3
        .cboReceivedFrom.SetFocus
        
    Case "&SAVE"
         If .cboTax.Text = Empty Then
         Taxed = "N"
         .txtAmount.Text = 0
         Else
         Taxed = "Y"
         End If
         
         If .cboDiscountAllowed.Text = Empty Then
         Discounted = "N"
         .txtDiscountedAmnt.Text = 0
         .txtDiscountValue.Text = 0
         .cboDiscountAllowed.Text = 0
         Else
         Discounted = "Y"
         End If
         
         If .txtBalanceDue.Text = 0 Then
           HasBalance = "N"
           Else
           HasBalance = "Y"
        End If
         
         
        If NewRecord Then
        If ValidRecord Then
         'save payment
            
            SQL1 = "INSERT INTO AdvertCustomerPayments(Taxed,Discounted,PaymentNo,ContractNo,ClientName,AmountPaid,BalanceDue,DiscountedAmount,DiscountValue,Memo,PaymentMethod,TaxAmount,ChequeNo,TaxPercentage,DiscountPercentage,InitialAmount,CreatedBy,DateCreated,AccPeriod)"
            SQL2 = "VALUES('" & Taxed & "','" & Discounted & "','" & Trim(.txtInvoiceNo.Text) & "','" & Trim(.txtContractNo.Text) & "','" & Trim(.cboReceivedFrom.Text) & "'," & CCur(.txtAmount.Text) & "," & CCur(.txtBalanceDue.Text) & "," & CCur(.txtDiscountedAmnt.Text) & "," & CCur(.txtDiscountValue.Text) & ",'" & Trim(.txtDescription.Text) & "','" & Trim(.cboPaymentMethod.Text) & "'," & CCur(.txtTaxAmount.Text) & ",'" & Trim(.txtChequeNo.Text) & "','" & Trim(.cboTax.Text) & "','" & Trim(.cboDiscountAllowed.Text) & "'," & CCur(.txtTotalAmount.Text) & ",'" & CurrentUserName & "','" & MyCurrentDate & "','" & MyCurrentPeriod & "');"
            
            NewSQL = SQL1 + SQL2
            Set rsNewRecord = New ADODB.Recordset
            rsNewRecord.Open NewSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsNewRecord = Nothing
            
            Set rsEditRecord = New ADODB.Recordset
            rsEditRecord.Open "UPDATE AdvertContractRequisition SET BalanceStatus = '" & HasBalance & "',  PaidStatus = '" & "Y" & "' WHERE PurchaseOrderNo = '" & Trim(.txtContractNo.Text) & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsEditRecord = Nothing
            
            Set rsEditRecord = New ADODB.Recordset
            rsEditRecord.Open "UPDATE AdvertContractRequisitionData SET BalanceStatus = '" & HasBalance & "',PaidStatus = '" & "Y" & "' ,Balance = " & CCur(.txtBalanceDue.Text) & " WHERE PurchaseOrderNo = '" & Trim(.txtContractNo.Text) & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsEditRecord = Nothing
            
            Button.Caption = "&NEXT ITEM": Button.Image = 25: Call RemoveCurrentListItem
           End If
             End If
           
            
      Case "&NEXT ITEM"
            NewRecord = True: NextItem = True
            If EditRecord Then Exit Sub
            Call ClearForNextItem
            Button.Caption = "&SAVE": Button.Image = 3
                                    
         Case Else
        Exit Sub
    End Select
Case "E"
'   Select Case Button.Caption
'    Case "&EDIT"
'    If NewRecord Then Exit Sub
'        If .txtInvoiceNo.Text = Empty Then
'            MsgBox "There is NO Current Record to Edit. Please Search and Display a Record First...!", vbCritical + vbOKOnly, ""
'            .txtInvoiceNo.SetFocus
'        Else
'            PopupMenu mnuEditOptions, , , , mnuEditMain
'            If EditOption = Empty Then
'                Exit Sub
'            Else
'                .txtInvoiceNo.Locked = True
'                Button.Caption = "SAVE": Button.Image = 3
'                EditRecord = True
'            End If
'        End If
'    Case "SAVE"
'        If EditRecord Then
'        If ValidRecord Then
'            Select Case EditOption
'            Case 1
'                EditSQL = "UPDATE AccountsReceivableMainData SET JobDescription='" & Trim(.txtJobDescription.Text) & "'  WHERE InvoiceNo='" & Trim(.txtInvoiceNo.Text) & "';"
'                Set rsEditRecord = New ADODB.Recordset
'                rsEditRecord.Open EditSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
'                Set rsEditRecord = Nothing
'                .txtInvoiceNo.Locked = False: EditRecord = False: Button.Caption = "&EDIT": Button.Image = 5
'            Case 2
'                j = .ListView2.ListItems.Count
'                If j = 0 Or .ListView2.View <> lvwReport Then Exit Sub
'                For i = 1 To j
'
'                  NewSQL = "INSERT INTO AccountsReceivableMainData(DueDate,InvoiceNo,CustomerNo,CustomerName,JobDescription,StartDate,FinishDate,TransActionDate,CreatedBy,DateCreated,AccPeriod)VALUES('" & Trim(DueDate) & "','" & Trim(.txtInvoiceNo.Text) & "','" & Trim(.txtCustomerNo.Text) & "','" & Trim(.cboCustomerName.Text) & "','" & StrConv(.txtJobDescription.Text, vbProperCase) & "','" & Trim(StartDate) & "','" & Trim(FinishDate) & "','" & Trim(TransActionDate) & "','" & Trim(CurrentUserName) & "','" & Trim(MyCurrentDate) & "','" & Trim(MyCurrentPeriod) & "');"
'                  Set rsNewRecord = New ADODB.Recordset
'                  rsNewRecord.Open NewSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
'                  Set rsNewRecord = Nothing
'
'
''                    Set rsLineUpdate = New ADODB.Recordset
''                    rsLineUpdate.Open "UPDATE AccountReceivableRegister SET  WHERE IndentNo='" & Trim(.txtIndentNo.Text) & "' AND ProductCode='" & Trim(.ListView2.ListItems(i).Text) & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
''                    Set rsLineUpdate = Nothing
''
'                    .ListView2.ListItems.Clear
'                    If .ListView2.ListItems.Count = 0 Or .ListView2.View <> lvwReport Then
'                        Call RemoveCurrentListItem
'                    End If
'
'                    .txtInvoiceNo.Locked = False: EditRecord = False: Button.Caption = "&EDIT": Button.Image = 5
'                Next i
'            Case Else
'                Exit Sub
'            End Select
'        End If
'        End If
'    Case Else
'        Exit Sub
'    End Select
Case "S"
'    If NewRecord Or EditRecord Then Exit Sub
'    INPQRY = InputBox("Please Enter the Payment Number to Search and Display Record...!!!", "Enter Invoice Number...")
'    If Len(INPQRY) = 0 Then
'        MsgBox "Required Search Parameter Missing or the Operation Was Cancelled...! No Work was Done!!!", vbCritical + vbOKOnly, "Missing Parameter"
'        Exit Sub
'    Else
'        Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM AdvertCustomerPaymentsMain WHERE PaymentNo='" & Trim(INPQRY) & "';")
'        If rsFindRecord.EOF And rsFindRecord.BOF Then
'            MsgBox "Requested Record Missing or Has Been Deleted. Check your Entries to Ensure they are Accurately Spelt...!", vbOKOnly + vbExclamation, "Record NOT Found...!"
'            Set rsFindRecord = Nothing: Exit Sub
'        Else
'            .txtInvoiceNo.Text = Trim(rsFindRecord!PaymentNo & "")
'            .txtContractNo.Text = Trim(rsFindRecord!ContractNo & "")
'            .cboReceivedFrom.Text = Trim(rsFindRecord!ClientName & "")
'
'
'            Call ShowCurrentInvoiceItems
'            Call ShowItemListData
'        End If
'        Set rsFindRecord = Nothing
'    End If
Case "R"
    If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh ") = vbCancel Then Exit Sub
        .Toolbar1.Buttons(1).Caption = "&NEW"
        .Toolbar1.Buttons(1).Image = 26
        .Toolbar1.Buttons(2).Caption = "&FINISH"
        .Toolbar1.Buttons(2).Image = 22
        .Toolbar1.Buttons(3).Image = 1
        .Toolbar1.Buttons(3).Caption = "&EDIT"
        NewRecord = False: EditRecord = False: MyCommonData.ClearTheScreen
Case "P"
'      Load frmRPTExpenseAccounts
'      frmRPTExpenseAccounts.Show vbModal, Me

Case "H"

Case "F"
       'save the main data
            
            NewSQL = "INSERT INTO AdvertCustomerPaymentsMain(PaymentNo,ContractNo,Client,CreatedBy,DateCreated,AccPeriod)VALUES('" & Trim(.txtInvoiceNo.Text) & "','" & Trim(.txtContractNo.Text) & "','" & Trim(.cboReceivedFrom.Text) & "','" & Trim(CurrentUserName) & "','" & Trim(MyCurrentDate) & "','" & Trim(MyCurrentPeriod) & "');"
            Set rsNewRecord = New ADODB.Recordset
            rsNewRecord.Open NewSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsNewRecord = Nothing
            
                        
             NewRecord = False
             .Toolbar1.Buttons(1).Caption = "&NEW"
             .Toolbar1.Buttons(1).Image = 26
             .ListView2.ListItems.Clear
            


Case Else

    Exit Sub
    

End Select
End With


Exit Sub
Err:
   If Err.Number = -2147217900 Then
   MsgBox "The Payment I.D You Entered Already Exists, Please Enter A Different Payment I.D ", vbCritical, "Data Validation"""
   Screen.MousePointer = vbDefault
   Else
    ErrorMessage
   Screen.MousePointer = vbDefault
    End If
    
End Sub


Private Sub mnuEdit_Click()
EditOption = 2
End Sub



Public Sub RemoveCurrentListItem()
'On Error GoTo Err
With Me
Dim i, j, k
   j = .ListView2.ListItems.Count: i = 1
     If j = 0 Then Exit Sub
     
     For i = 1 To j
      If .ListView2.ListItems(i).Checked = True Then
         .ListView2.ListItems.Remove (i): Exit Sub
      End If
    Next i
End With
Exit Sub
Err:
   ErrorMessage
End Sub


Private Function ValidRecord() As Boolean
On Error Resume Next
With Me
    For Each i In Me
    If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
        If i.Text = Empty And i.Name <> "txtSerialNo" And i.Name <> "txtDiscountedAmnt" And i.Name <> "txtDiscountValue" And i.Name <> "txtDescription" And i.Name <> "txtTaxAmount" And i.Name <> "txtChequeNo" And i.Name <> "txtExpiryDate" And i.Name <> "cboDiscountAllowed" Then
            i.Text = "Missing Field": i.BackColor = vbRed
            MsgBox "All the Fields are Required. Please Enter the Missing Data...!", vbCritical + vbOKOnly, "Data Validation"
            i.SetFocus:  ValidRecord = False: Exit Function
        End If
               
    End If
    Next i
    ValidRecord = True
End With
End Function




Private Sub cboPaymentMethod_GotFocus()

'On Error GoTo Err
'If Not NewRecord Or Not EditRecord Then Exit Sub
With frmReceiveCustomerPayments
        Set rsCOMBO = cnCOMMON.Execute("SELECT Descriptions FROM ParamPaymentMethod;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    .cboPaymentMethod.Clear
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!Descriptions) And rsCOMBO!Descriptions <> "" Then
        .cboPaymentMethod.AddItem rsCOMBO!Descriptions
        End If
        
    rsCOMBO.MoveNext
    Loop
'    .cboPaymentMethod.AddItem "<Add New>"
    
OUTS:
    Set rsCOMBO = Nothing
End With
Exit Sub
Err:
    ErrorMessage

End Sub


Private Sub cboPaymentMethod_LostFocus()

'On Error GoTo Err
With frmReceiveCustomerPayments

    Set rsFindRecord = cnCOMMON.Execute("SELECT *  FROM ParamPaymentMethod  WHERE Descriptions ='" & Trim(.cboPaymentMethod.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        
        .cboPaymentMethod.Text = rsFindRecord!PaymentMethod & ""
        
    End If
        
End With
Exit Sub
Err:
    ErrorMessage


End Sub



Private Function GetNextInvoiceNo() As Variant
On Error GoTo Err

Dim rsLastID As ADODB.Recordset 'used to retrieve current LastId in the Table
Dim strLastID As String 'SQL statement

Dim strTemp As String 'store current record
Dim iNumPos As Integer 'store position of the first numeral
Dim strPrefix As String 'stores Id Prefix


'With frmDistributionReturns
strLastID = "SELECT MAX(InvoiceNo) AS LastID FROM ParamInvoiceNumbers;"
Set rsLastID = New ADODB.Recordset
'End With
With rsLastID
'open the recordset
    .Open strLastID, cnCOMMON, adOpenKeyset, adLockOptimistic
    If .EOF And .BOF Then 'shows empty recordset
        GetNextInvoiceNo = "0000001" 'format of desired format of the string value
    ElseIf IsNull(!lastid) = True Or !lastid = "" Then
        GetNextInvoiceNo = "0000001"
    Else
        .MoveFirst
        strTemp = !lastid
        iNumPos = 1
        Dim sChar As String
        Dim iIDLen As Integer
        iIDLen = Len(strTemp)
        sChar = Mid(strTemp, iNumPos, 1)
        While InStr("1234567890", sChar) = 0
            iNumPos = iNumPos + 1
            sChar = Mid(strTemp, iNumPos, 1)
        Wend
        'store the ID prefix eg AP
        strPrefix = Left(strTemp, iNumPos - 1)
        'store the number portion eg and the length with leading Zeros
        strTemp = Right(strTemp, Len(strTemp) + 1 - iNumPos)
        strTemp = Format(Int(strTemp) + 1, String(iIDLen + 1 - iNumPos, "0"))
        GetNextInvoiceNo = strPrefix & strTemp
    End If
End With
    Exit Function
Err:
    ErrorMessage
End Function

Public Sub ClearForNextItem()
With Me
.chkAllowDiscount.Value = 0
.cboTax.Text = ""
.cboDiscountAllowed.Text = ""
.txtAmount.Text = ""
.txtBalanceDue.Text = ""
.txtDiscountedAmnt.Text = ""
.txtDiscountValue.Text = ""
.txtDescription.Text = ""
.txtTaxAmount.Text = ""
End With
End Sub

Private Sub GetItemListStructure()
'On Error GoTo Err
With Me

.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Serial No", .ListView2.Width / 6
.ListView2.ColumnHeaders.Add , , "Advert Code ", .ListView2.Width / 5.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Advert Name", .ListView2.Width / 2.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Advert Type", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Advert Length", .ListView2.Width / 6 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Advert Width", .ListView2.Width / 5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Duration", .ListView2.Width / 4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Days", .ListView2.Width / 5.4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Advert Cost", .ListView2.Width / 5.4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contract Start Date", .ListView2.Width / 5.4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Contract End Date", .ListView2.Width / 5.4 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Approved", .ListView2.Width / 6.5 ', lvwColumnCenter
.ListView2.ColumnHeaders.Add , , "Paid", .ListView2.Width / 6.5 ', lvwColumnCenter

.ListView2.View = lvwReport

End With
Exit Sub
Err:
    ErrorMessage
End Sub



Private Sub GetCurrentInvoiceItems()


'On Error GoTo Err
With Me
Screen.MousePointer = vbHourglass

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Contract No", .ListView1.Width / 5.5  ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Order Date", .ListView1.Width / 5.5  ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Start Date", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "End Date", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Client Code", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Client Name", .ListView1.Width / 4.5
.ListView1.ColumnHeaders.Add , , "Amount Due", .ListView1.Width / 5.5

.ListView1.View = lvwReport
Screen.MousePointer = vbDefault
End With
End Sub

Public Function GetInvoiceTotalSum() As Currency
'On Error GoTo Err
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT Sum(EndingBalance) As Total FROM AccountsReceivableRegister WHERE Number = '" & .txtInvoiceNo.Text & "'")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetInvoiceTotalSum = Empty: Exit Function
    ElseIf IsNull(rsFindRecord!Total) = True Or rsFindRecord!Total = "" Then
        GetInvoiceTotalSum = Empty: Exit Function
    Else
        GetInvoiceTotalSum = CCur(rsFindRecord!Total)
    End If
          
    Set rsFindRecord = Nothing
    
End With
Exit Function
Err:
    ErrorMessage
End Function

Private Sub ShowInvoiceDetailsData()

End Sub



Private Sub cboDiscountAllowed_GotFocus()

'On Error GoTo Err
'If Not NewRecord Then Exit Sub
With frmReceiveCustomerPayments
        Set rsCOMBO = cnCOMMON.Execute("SELECT NoOfDays FROM ParamPaymentperiods Order by NoOfDays;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    .cboDiscountAllowed.Clear
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!NoOfDays) And rsCOMBO!NoOfDays <> "" Then
        .cboDiscountAllowed.AddItem rsCOMBO!NoOfDays
        End If
        
    rsCOMBO.MoveNext
    Loop
'    .cboPaymentMethod.AddItem "<Add New>"
    
OUTS:
    Set rsCOMBO = Nothing
End With
Exit Sub
Err:
    ErrorMessage

End Sub


Private Sub cboDiscountAllowed_LostFocus()

'On Error GoTo Err
With frmReceiveCustomerPayments

    Set rsFindRecord = cnCOMMON.Execute("SELECT *  FROM ParamPaymentPeriods WHERE NoOfDays ='" & Trim(.cboDiscountAllowed.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        
        .cboDiscountAllowed.Text = rsFindRecord!DiscountAllowed & ""
        .Label11.Caption = "Percentage Alwd"
    End If
    
End With
Exit Sub
Err:
    ErrorMessage


End Sub


Private Sub txtAmount_Change()
On Error Resume Next
With Me
.txtBalanceDue.Text = FormatNumber(CDbl(.txtTotalAmount.Text) - (CDbl(.txtAmount.Text) - CDbl(.txtTaxAmount.Text)), 2, vbUseDefault, vbUseDefault, vbTrue)
End With
End Sub


