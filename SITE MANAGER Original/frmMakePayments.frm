VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMakePayments 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resource Scheduling -MAKE PAYMENTS"
   ClientHeight    =   8625
   ClientLeft      =   -270
   ClientTop       =   2355
   ClientWidth     =   11910
   Icon            =   "frmMakePayments.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboPaymentMethod 
      Height          =   315
      Left            =   5880
      TabIndex        =   36
      Top             =   1320
      Width           =   1935
   End
   Begin VB.ComboBox cboReceivedFrom 
      Height          =   315
      Left            =   960
      TabIndex        =   35
      Top             =   1320
      Width           =   3495
   End
   Begin VB.TextBox txtNoOfMonths 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9720
      TabIndex        =   34
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtSiteNo 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   7920
      TabIndex        =   33
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtSiteName 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1680
      TabIndex        =   32
      Top             =   1800
      Width           =   6135
   End
   Begin VB.TextBox txtLLNo 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   9840
      TabIndex        =   30
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtAmountInWords 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1800
      TabIndex        =   29
      Top             =   3000
      Width           =   6135
   End
   Begin VB.TextBox txtTotalAmount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7920
      TabIndex        =   27
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtPaymentInterval 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9720
      TabIndex        =   25
      Top             =   2160
      Width           =   1815
   End
   Begin VB.OptionButton optFee 
      BackColor       =   &H80000009&
      Caption         =   "Council Fee"
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   23
      Top             =   720
      Width           =   1335
   End
   Begin VB.OptionButton optFee 
      BackColor       =   &H80000009&
      Caption         =   "Rent Fee"
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   22
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
      Left            =   5760
      TabIndex        =   21
      Top             =   2400
      Width           =   2055
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
            Picture         =   "frmMakePayments.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":0984
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":0EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":1408
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":194A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":1E8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":2026
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":21C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":235A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":24F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":268E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":2EC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":36F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":388C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":3A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":3D58
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":458A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":4724
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":48BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":4A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":4E0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":509C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":5196
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":5290
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":538A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":5484
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":59C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":5AD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":5BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":5CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":5E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":5F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":6032
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":6144
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":6256
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":6368
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":647A
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":658C
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":669E
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":67B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":68C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":69D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":6AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":6BF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":6D0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":6E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":6F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":7040
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":7152
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":7264
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":7376
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":7488
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":759A
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":76AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":77BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakePayments.frx":78D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "List of sites owned by selected payee"
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
      TabIndex        =   16
      Top             =   4080
      Width           =   11655
      Begin MSComctlLib.ListView ListView1 
         Height          =   4095
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   11415
         _ExtentX        =   20135
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
      TabIndex        =   15
      Top             =   2400
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   255
      Left            =   11280
      TabIndex        =   13
      Top             =   1560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Format          =   64356353
      CurrentDate     =   38210
   End
   Begin VB.TextBox txtNextDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9720
      TabIndex        =   12
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtChequeNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7920
      TabIndex        =   11
      Top             =   1560
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   9360
      TabIndex        =   8
      Top             =   960
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Format          =   64356353
      CurrentDate     =   38210
   End
   Begin VB.TextBox txtInvoiceDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7920
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtInvoiceNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9720
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   960
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
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&NEW"
            Key             =   "N"
            ImageIndex      =   26
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&EDIT"
            Key             =   "E"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&REFRESH"
            Key             =   "R"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&PRINT"
            Key             =   "P"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&SEARCH"
            Key             =   "S"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&VIEW"
            ImageIndex      =   36
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Line Line4 
      X1              =   1680
      X2              =   7800
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Site Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   1800
      X2              =   7920
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount in words"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "TOTAL AMNT"
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
      Left            =   7920
      TabIndex        =   26
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "PYMNT INTERVAL"
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
      Left            =   9720
      TabIndex        =   24
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Line Line3 
      X1              =   5640
      X2              =   7680
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      X1              =   1680
      X2              =   4320
      Y1              =   2760
      Y2              =   2760
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
      TabIndex        =   20
      Top             =   2400
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
      Left            =   4560
      TabIndex        =   19
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Payee"
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
      TabIndex        =   18
      Top             =   1320
      Width           =   1215
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
      TabIndex        =   14
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "NEXT DATE"
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
      Top             =   1320
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
      Top             =   1320
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
      Top             =   720
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
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   " Payments"
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
Attribute VB_Name = "frmMakePayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim MyCommonData As clsCommonData, EditOption, PM



Private Sub cboDepositTo_Change()

End Sub



Private Sub cboDiscountAllowed_Click()
Me.txtInvoiceDate.SetFocus
End Sub

Private Sub cboReceivedFrom_Click()
Me.cboPaymentMethod.SetFocus
End Sub


Private Sub cboReceivedFrom_GotFocus()

'On Error GoTo Err
If Not NewRecord Then Exit Sub
With frmMakePayments
        Set rsCOMBO = cnCOMMON.Execute("SELECT SurName FROM AdvertSites A,AdvertSiteLords B WHERE A.landlordNo = B.LandLordNo AND A.RenewalApprovalStatus IS NOT NULL AND A.ApprovedStatus IS NOT NULL AND A.ValidStatus IS NULL  Order by SurName;")
    
    If rsCOMBO.EOF And rsCOMBO.BOF Then GoTo OUTS
    .cboReceivedFrom.Clear
    rsCOMBO.MoveFirst
    Do While Not rsCOMBO.EOF
        If Not IsNull(rsCOMBO!Surname) And rsCOMBO!Surname <> "" Then
        .cboReceivedFrom.AddItem rsCOMBO!Surname
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
If Not NewRecord Then Exit Sub
With Me

    Set rsFindRecord = cnCOMMON.Execute("SELECT *  FROM AdvertSiteLords A ,AdvertSites B WHERE A.LandLordNo =B.LandLordNo AND A.Surname = '" & Trim(.cboReceivedFrom.Text) & "' AND B.RenewalApprovalStatus IS NOT NULL AND B.ContractFinish > '" & MyCurrentDate & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        
         .cboReceivedFrom.Text = rsFindRecord!Surname & ""
         .txtInvoiceDate = MyCurrentDate
         .txtLLNo.Text = rsFindRecord!LandLordNo & ""
         .txtSiteName = rsFindRecord!SiteName & ""
         .txtSiteNo = rsFindRecord!SiteNo
         
    End If
    
    Screen.MousePointer = vbHourglass

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Site No", .ListView1.Width / 5.5  ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Site Name", .ListView1.Width / 4.5  ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "LandLord No", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Rent Fee Amount", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Council Fee Amount", .ListView1.Width / 4.5
.ListView1.ColumnHeaders.Add , , "Date Council Payed", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Date Rent Paid", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Valid Status", .ListView1.Width / 5.5




.ListView1.View = lvwReport

If CheckRentPaymentStatus = "Y" Then
.optFee(0).Enabled = False
Else: .optFee(0).Enabled = True
End If

If CheckCouncilPaymentStatus = "Y" Then
.optFee(1).Enabled = False
Else: .optFee(1).Enabled = True
End If


Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT *  FROM AdvertSites  WHERE LandLordNo = '" & Trim(.txtLLNo.Text) & "'", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem, NCount As Double

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Screen.MousePointer = vbDefault: Exit Sub

  Else

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!SiteNo))
    
    If Not IsNull(rsLIST!SiteName) Then
        MyList.SubItems(1) = CStr(rsLIST!SiteName)
    End If
    If Not IsNull(rsLIST!sitephysicalAddress) Then
        MyList.SubItems(2) = CStr(rsLIST!sitephysicalAddress)
    End If
    
    If Not IsNull(rsLIST!LandLordNo) Then
        MyList.SubItems(3) = CStr(rsLIST!LandLordNo)
    End If
    
    If Not IsNull(rsLIST!SiteCharges) Then
        MyList.SubItems(4) = FormatNumber(rsLIST!SiteCharges, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    If Not IsNull(rsLIST!CouncilCharges) Then
        MyList.SubItems(5) = FormatNumber(rsLIST!CouncilCharges, 2, vbUseDefault, vbUseDefault, vbTrue)
    End If
    
    
    If Not IsNull(rsLIST!DateCouncilPaid) Then
        MyList.SubItems(6) = CStr(rsLIST!DateCouncilPaid)
    End If
    
   
    If Not IsNull(rsLIST!DateRentPaid) Then
        MyList.SubItems(7) = (rsLIST!DateRentPaid)
    End If
    
    
    
    If IsNull(rsLIST!Validstatus) = True Then
        MyList.SubItems(8) = CStr("No")
    ElseIf Not IsNull(rsLIST!Validstatus) Then
        MyList.SubItems(8) = CStr("Yes")
    End If
    
 
 
   
    rsLIST.MoveNext
      
Wend
 End If
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

Call GetCurrentInvoiceItems
Me.mnuEditOptions.Visible = False
Me.ListView1.Visible = True

Me.Frame1.Visible = True



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
       
        Me.txtSiteNo = Item
        Me.txtSiteName.Text = Item.SubItems(1)
        
        If CheckRentPaymentStatus = "Y" Then
        Me.optFee(0).Enabled = False
        Else: Me.optFee(0).Enabled = True
        End If

       If CheckCouncilPaymentStatus = "Y" Then
        Me.optFee(1).Enabled = False
       Else: Me.optFee(1).Enabled = True
       End If
          
    ElseIf Item.Checked = False Then
        Me.txtSiteNo.Text = Empty
        Me.txtSiteName.Text = Empty
        
    End If
    
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub optFee_Click(Index As Integer)

'On Error GoTo Err
With Me
If Not NewRecord And Not EditRecord Then Exit Sub
Select Case Index
Case 0
    Call GetRentFee
    Me.optFee(0).Value = True
    Me.txtBalanceDue.SetFocus
Case 1
    Call GetCouncilFee
    Me.optFee(1).Value = True
    Me.txtBalanceDue.SetFocus
Case Else
    Exit Sub
End Select
Exit Sub
Err:
    ErrorMessage
End With
End Sub
Private Sub GetRentFee()
'On Error GoTo Err
Dim Interval, Num, Today As Variant
With Me
    If .txtSiteNo.Text = "" Then Exit Sub
    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM AdvertSites WHERE SiteNo='" & Trim(.txtSiteNo.Text) & "' AND Discontinued = '" & "N" & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing
    Else
         .txtAmount.Text = rsFindRecord!SiteCharges & ""
         .txtPaymentInterval.Text = rsFindRecord!SiteChargesInterval & ""
         .txtNoOfMonths.Text = rsFindRecord!SiteChargeNoMonths & ""
          Interval = CStr(.txtPaymentInterval.Text)
          Num = CDbl(.txtNoOfMonths.Text)
         .txtNextDate.Text = DateAdd(Interval, Num, MyCurrentDate)
         .txtTotalAmount.Text = rsFindRecord!SiteCharges & ""
        
        
    End If
End With
Exit Sub
Err:
    ErrorMessage
End Sub
Private Sub GetCouncilFee()
'On Error GoTo Err
Dim Interval, Num, Today As Variant
With Me
    If .txtSiteNo.Text = "" Then Exit Sub
    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM AdvertSites WHERE SiteNo='" & Trim(.txtSiteNo.Text) & "' AND Discontinued = '" & "N" & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing
    Else
        .txtAmount.Text = rsFindRecord!CouncilCharges & ""
        .txtPaymentInterval.Text = rsFindRecord!CouncilChargesInterval & ""
        .txtNoOfMonths.Text = rsFindRecord!CouncilChargeNoMonths & ""
        Interval = CStr(.txtPaymentInterval.Text)
        Num = CDbl(.txtNoOfMonths.Text)
        .txtNextDate.Text = DateAdd(Interval, Num, MyCurrentDate)
        .txtTotalAmount.Text = rsFindRecord!CouncilCharges & ""
        
        
    End If
End With
Exit Sub
Err:
    ErrorMessage
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo Err
Dim SQL1, SQL2
Dim NextDate As Variant
With Me
NextDate = Format(.txtNextDate.Text, "MMMM dd,yy")
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
          If .optFee(0).Enabled = False And .optFee(1).Enabled = False Then Exit Sub
    
        If .optFee(0) Then
          PM = "Rent Fee"
           ElseIf .optFee(1) Then
          PM = "Council Fee"
           Else
          PM = Empty
        End If
    
                 
        If NewRecord Then
        If ValidRecord Then
         'save payment
            
            SQL1 = "INSERT INTO RsPayments(SiteName,PayeeNo,Payee,Amount,BalanceDue,PaymentMethod,ChequeNo,AmountInWords,TotalAmount,NextDate,PaymentInterval,LLNo,SiteNo,FeeType,CreatedBy,DateCreated,AccPeriod)"
            SQL2 = "VALUES('" & Trim(.txtSiteName.Text) & "','" & Trim(.txtInvoiceNo.Text) & "','" & Trim(.cboReceivedFrom.Text) & "'," & CCur(.txtAmount.Text) & "," & CCur(.txtBalanceDue.Text) & ",'" & Trim(.cboPaymentMethod.Text) & "','" & Trim(.txtChequeNo.Text) & "','" & Trim(.txtAmountInWords.Text) & "'," & CCur(.txtTotalAmount.Text) & ",'" & Trim(NextDate) & "','" & Trim(.txtPaymentInterval.Text) & "','" & Trim(.txtLLNo.Text) & "','" & Trim(.txtSiteNo.Text) & "','" & Trim(PM) & "','" & CurrentUserName & "','" & MyCurrentDate & "','" & MyCurrentPeriod & "');"
            
            NewSQL = SQL1 + SQL2
            Set rsNewRecord = New ADODB.Recordset
            rsNewRecord.Open NewSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsNewRecord = Nothing
            
            If .optFee(0) Then
            
            Set rsEditRecord = New ADODB.Recordset
            rsEditRecord.Open "UPDATE AdvertSites SET RentPaymentStatus = '" & "Y" & "',DateRentPaid = '" & MyCurrentDate & "',RentPaidBy = '" & CurrentUserName & "' WHERE SiteNo = '" & Trim(.txtSiteNo.Text) & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsEditRecord = Nothing
            
            Else
            
            Set rsEditRecord = New ADODB.Recordset
            rsEditRecord.Open "UPDATE AdvertSites SET CouncilPaymentStatus = '" & "Y" & "',DateCouncilPaid = '" & MyCurrentDate & "',CouncilPaidBy = '" & CurrentUserName & "' WHERE SiteNo = '" & Trim(.txtSiteNo.Text) & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsEditRecord = Nothing
            
            End If
            
            If CheckCouncilPaymentStatus = "Y" And CheckRentPaymentStatus = "Y" Then
            
            Set rsEditRecord = New ADODB.Recordset
            rsEditRecord.Open "UPDATE AdvertSites SET ValidStatus = '" & "Y" & "' WHERE SiteNo = '" & Trim(.txtSiteNo.Text) & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsEditRecord = Nothing
            
            Else
            End If
            
            
            NewRecord = False
             .Toolbar1.Buttons(1).Caption = "&NEW"
             .Toolbar1.Buttons(1).Image = 26
             .ListView1.ListItems.Clear
                                
            
           End If
             End If
                                                        
         Case Else
        Exit Sub
    End Select
Case "E"
   Select Case Button.Caption
    Case "&EDIT"
    If NewRecord Then Exit Sub
        If .txtInvoiceNo.Text = Empty Then
            MsgBox "There is NO Current Record to Edit. Please Search and Display a Record First...!", vbCritical + vbOKOnly, ""
            .txtInvoiceNo.SetFocus
        Else
               .txtInvoiceNo.Locked = True
                Button.Caption = "SAVE": Button.Image = 3
                EditRecord = True
            
        End If
    Case "SAVE"
        If EditRecord Then
          If ValidRecord Then
                     
                   
                EditSQL = "UPDATE  SET Amount=" & CCur(.txtAmount.Text) & ",PaymentMethod = '" & Trim(.cboPaymentMethod.Text) & "',BalanceDue = " & CCur(.txtBalanceDue.Text) & "',ChequeNo = '" & Trim(.txtChequeNo.Text) & "',AmountInWords = '" & Trim(.txtAmountInWords.Text) & "',NextDate = '" & Trim(NextDate) & "'  WHERE PayeeNo='" & Trim(.txtInvoiceNo.Text) & "';"
                Set rsEditRecord = New ADODB.Recordset
                rsEditRecord.Open EditSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                Set rsEditRecord = Nothing
                .txtInvoiceNo.Locked = False: EditRecord = False: Button.Caption = "&EDIT": Button.Image = 5
           
             End If
        End If
    Case Else
    Exit Sub
    End Select
    
                                                      
Case "S"
    If NewRecord Or EditRecord Then Exit Sub
    INPQRY = InputBox("Please Enter the Payment Number to Search and Display Record...!!!", "Enter Payment Number...")
    If Len(INPQRY) = 0 Then
        MsgBox "Required Search Parameter Missing or the Operation Was Cancelled...! No Work was Done!!!", vbCritical + vbOKOnly, "Missing Parameter"
        Exit Sub
    Else
        Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM RsPayments WHERE PaymentNo='" & Trim(INPQRY) & "';")
        If rsFindRecord.EOF And rsFindRecord.BOF Then
            MsgBox "Requested Record Missing or Has Been Deleted Or Check your Entries to Ensure they are Accurately Spelt...!", vbOKOnly + vbExclamation, "Record NOT Found...!"
            Set rsFindRecord = Nothing: Exit Sub
        Else
            .txtInvoiceNo.Text = Trim(rsFindRecord!PayeeNo & "")
            .cboReceivedFrom.Text = Trim(rsFindRecord!Payee & "")
            .txtAmount.Text = Trim(rsFindRecord!Amount & "")
            .txtBalanceDue.Text = Trim(rsFindRecord!BalanceDue & "")
            .cboPaymentMethod.Text = Trim(rsFindRecord!PaymentMethod & "")
            .txtChequeNo.Text = Trim(rsFindRecord!ChequeNo & "")
            .txtAmountInWords.Text = Trim(rsFindRecord!Amountinwords & "")
            .txtNextDate.Text = Trim(rsFindRecord!NextDate & "")
            .txtPaymentInterval.Text = Trim(rsFindRecord!PaymentInterval & "")
            .txtLLNo.Text = Trim(rsFindRecord!LLNo & "")
            .txtSiteNo.Text = Trim(rsFindRecord!SiteNo & "")
            .txtSiteName.Text = Trim(rsFindRecord!SiteName & "")
         
        End If
        Set rsFindRecord = Nothing
    End If
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
Private Function CheckCouncilPaymentStatus() As Variant
'On Error GoTo Err
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT CouncilPaymentStatus FROM AdvertSites  WHERE SiteNo = '" & Trim(.txtSiteNo.Text) & "' ;")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        CheckCouncilPaymentStatus = 0: Set rsFindRecord = Nothing
    ElseIf IsNull(rsFindRecord!CouncilPaymentStatus) = True Or rsFindRecord!CouncilPaymentStatus = "" Then
        CheckCouncilPaymentStatus = 0: Set rsFindRecord = Nothing
    Else
        CheckCouncilPaymentStatus = (rsFindRecord!CouncilPaymentStatus)
    End If
        
    Set rsFindRecord = Nothing
Exit Function
Err:
    ErrorMessage
    End With
End Function
Private Function CheckRentPaymentStatus() As Variant
'On Error GoTo Err
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT RentPaymentStatus FROM AdvertSites  WHERE SiteNo = '" & Trim(.txtSiteNo.Text) & "' ;")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        CheckRentPaymentStatus = 0: Set rsFindRecord = Nothing
    ElseIf IsNull(rsFindRecord!RentPaymentStatus) = True Or rsFindRecord!RentPaymentStatus = "" Then
        CheckRentPaymentStatus = 0: Set rsFindRecord = Nothing
    Else
        CheckRentPaymentStatus = (rsFindRecord!RentPaymentStatus)
    End If
        
    Set rsFindRecord = Nothing
Exit Function
Err:
    ErrorMessage
    End With
End Function
Private Sub mnuEdit_Click()
EditOption = 2
End Sub






Private Function ValidRecord() As Boolean
With Me
    For Each i In Me
    If TypeOf i Is TextBox Or TypeOf i Is ComboBox Then
        If i.Text = Empty And i.Name <> "txtAmountInWords" And i.Name <> "txtChequeNo" Then
            MsgBox "All the Fields are Required. Please Enter the Missing Data...!", vbCritical + vbOKOnly, "Data Validation"
            i.SetFocus: ValidRecord = False: Exit Function
        End If
    End If
    Next i
    ValidRecord = True
End With
End Function




Private Sub cboPaymentMethod_GotFocus()

'On Error GoTo Err
'If Not NewRecord Or Not EditRecord Then Exit Sub
With Me
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


Private Sub txtBalanceDue_GotFocus()
On Error Resume Next
With Me
If .txtTotalAmount.Text = "" Then Exit Sub
.txtBalanceDue.Text = FormatNumber(CDbl(.txtTotalAmount.Text) - CDbl(.txtAmount.Text), 2, vbUseDefault, vbUseDefault, vbTrue)
End With
End Sub

Private Sub txtBalanceDue_LostFocus()
On Error Resume Next
With Me
If .txtTotalAmount.Text = "" Then Exit Sub
.txtBalanceDue.Text = FormatNumber(CDbl(.txtTotalAmount.Text) - CDbl(.txtAmount.Text), 2, vbUseDefault, vbUseDefault, vbTrue)
End With
End Sub
