VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmALISMReceiptNew1 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A.L.I.S ENTERPRISE [PAYMENT RECEIPTS PROCESSING]"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   1560
   ClientWidth     =   11880
   Icon            =   "frmALISMReceiptNew1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   11880
   Begin TabDlg.SSTab SSTReceipt 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "RECEIPT DETAILS"
      TabPicture(0)   =   "frmALISMReceiptNew1.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame14"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "LOAN DATASHEET"
      TabPicture(1)   =   "frmALISMReceiptNew1.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "PROPOSAL DATASHEET"
      TabPicture(2)   =   "frmALISMReceiptNew1.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "POLICY DATASHEET"
      TabPicture(3)   =   "frmALISMReceiptNew1.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "RECEIPTS DATASHEET"
      TabPicture(4)   =   "frmALISMReceiptNew1.frx":04B2
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame17"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame17 
         Height          =   6615
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   11655
         Begin VB.Frame Frame2 
            Height          =   6615
            Left            =   0
            TabIndex        =   48
            Top             =   0
            Width           =   11655
            Begin VB.TextBox txtTotalAmount 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5640
               Locked          =   -1  'True
               TabIndex        =   101
               Top             =   5760
               Width           =   2415
            End
            Begin VB.TextBox txtBalance 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   99
               Top             =   5760
               Width           =   2175
            End
            Begin VB.TextBox txtReceiptAmountDetails 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5640
               Locked          =   -1  'True
               TabIndex        =   97
               Top             =   6120
               Width           =   2415
            End
            Begin VB.Frame Frame9 
               Caption         =   "Receipts"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3135
               Left            =   8280
               TabIndex        =   88
               Top             =   2880
               Width           =   3255
               Begin VB.CommandButton cmdPrintReceiptDetails 
                  Appearance      =   0  'Flat
                  Caption         =   "&Print Receipt"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Left            =   120
                  TabIndex        =   93
                  Top             =   2160
                  Width           =   3015
               End
               Begin VB.CommandButton cmdAddNewDetails 
                  Appearance      =   0  'Flat
                  Caption         =   "&Add New"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Left            =   120
                  TabIndex        =   92
                  Top             =   240
                  Width           =   3015
               End
               Begin VB.CommandButton cmdUpdateDetails 
                  Appearance      =   0  'Flat
                  Caption         =   "&Update"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Left            =   120
                  TabIndex        =   14
                  Top             =   720
                  Width           =   3015
               End
               Begin VB.CommandButton cmdSearchDetails 
                  Appearance      =   0  'Flat
                  Caption         =   "&Search"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Left            =   120
                  TabIndex        =   91
                  Top             =   1200
                  Width           =   3015
               End
               Begin VB.CommandButton cmdCancelDetails 
                  Appearance      =   0  'Flat
                  Caption         =   "&Cancel"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Left            =   120
                  TabIndex        =   90
                  Top             =   1680
                  Width           =   3015
               End
               Begin VB.CommandButton cmdPrintReportListing 
                  Appearance      =   0  'Flat
                  Caption         =   "&Print Receipt Listing"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Left            =   120
                  TabIndex        =   89
                  Top             =   2640
                  Width           =   3015
               End
            End
            Begin VB.Frame Frame8 
               Caption         =   "Premium Details"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2775
               Left            =   8280
               TabIndex        =   75
               Top             =   120
               Width           =   3255
               Begin VB.TextBox txtSuspenseAccount 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFC0&
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
                  Left            =   1080
                  TabIndex        =   81
                  Top             =   2265
                  Width           =   2055
               End
               Begin VB.TextBox txtUnitCount 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFC0&
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
                  Left            =   1080
                  TabIndex        =   80
                  Top             =   240
                  Width           =   2055
               End
               Begin VB.TextBox txtunitsPaid 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFC0&
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
                  Left            =   1080
                  TabIndex        =   79
                  Top             =   1455
                  Width           =   2055
               End
               Begin VB.TextBox txtStatusCode 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFC0&
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
                  Left            =   1080
                  TabIndex        =   78
                  Top             =   1050
                  Width           =   2055
               End
               Begin VB.TextBox txtDueDate 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFC0&
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
                  Left            =   1080
                  TabIndex        =   77
                  Top             =   645
                  Width           =   2055
               End
               Begin VB.TextBox txtUnitCountBeforePayment 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFC0&
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
                  Left            =   1080
                  TabIndex        =   76
                  Top             =   1920
                  Width           =   2055
               End
               Begin VB.Label lblSuspenseAccount 
                  Caption         =   "Suspense"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   87
                  Top             =   2340
                  Width           =   735
               End
               Begin VB.Label lblUnitCount 
                  Caption         =   "Prem Count:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   86
                  Top             =   315
                  Width           =   975
               End
               Begin VB.Label lblDueDate 
                  Caption         =   "Due Date:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   85
                  Top             =   720
                  Width           =   855
               End
               Begin VB.Label lblStatusCode 
                  Caption         =   "Status Code"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   84
                  Top             =   1125
                  Width           =   975
               End
               Begin VB.Label lblUnitssPaid 
                  Caption         =   "Prem Paid:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   83
                  Top             =   1530
                  Width           =   855
               End
               Begin VB.Label lblUnitCountBeforePayment 
                  Caption         =   "Prem Prior:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   82
                  Top             =   1935
                  Width           =   855
               End
            End
            Begin VB.Frame Frame7 
               Height          =   2175
               Left            =   120
               TabIndex        =   59
               Top             =   120
               Width           =   8055
               Begin VB.ComboBox cboEmployerCode 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFC0&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   1320
                  Locked          =   -1  'True
                  TabIndex        =   67
                  Top             =   840
                  Width           =   1695
               End
               Begin VB.TextBox txtEmployeeNo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFC0&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   1320
                  TabIndex        =   66
                  Top             =   1725
                  Width           =   1695
               End
               Begin VB.ComboBox cboEmployer 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFC0C0&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   3120
                  Sorted          =   -1  'True
                  TabIndex        =   65
                  Top             =   840
                  Width           =   4695
               End
               Begin VB.TextBox txtReferenceNo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFC0&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   5520
                  Locked          =   -1  'True
                  TabIndex        =   64
                  Top             =   1725
                  Width           =   2295
               End
               Begin VB.ComboBox cboReceiptType 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFC0C0&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   3120
                  Sorted          =   -1  'True
                  TabIndex        =   11
                  Top             =   345
                  Width           =   3015
               End
               Begin VB.TextBox txtTransactionDate 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFC0C0&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   6600
                  TabIndex        =   63
                  Top             =   345
                  Width           =   1215
               End
               Begin VB.TextBox txtReceiptNoDetails 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFC0&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   1320
                  TabIndex        =   62
                  Top             =   345
                  Width           =   1215
               End
               Begin VB.TextBox txtNames 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFC0&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   3120
                  Locked          =   -1  'True
                  TabIndex        =   61
                  Top             =   1320
                  Width           =   4695
               End
               Begin VB.TextBox txtExpectedAmount 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFC0&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   1320
                  Locked          =   -1  'True
                  TabIndex        =   60
                  Top             =   1320
                  Width           =   1695
               End
               Begin VB.Label Label18 
                  Caption         =   "Emp #"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   74
                  Top             =   1785
                  Width           =   615
               End
               Begin VB.Label Label17 
                  Caption         =   "Employer"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   73
                  Top             =   893
                  Width           =   1095
               End
               Begin VB.Label Label29 
                  Caption         =   "Type"
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   72
                  Top             =   398
                  Width           =   375
               End
               Begin VB.Label Label16 
                  Caption         =   " Date"
                  Height          =   255
                  Left            =   6120
                  TabIndex        =   71
                  Top             =   405
                  Width           =   495
               End
               Begin VB.Label Label15 
                  Caption         =   "Receipt No"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   70
                  Top             =   398
                  Width           =   1215
               End
               Begin VB.Label Label14 
                  Caption         =   "Ref No"
                  Height          =   255
                  Left            =   4320
                  TabIndex        =   69
                  Top             =   1785
                  Width           =   615
               End
               Begin VB.Label Label13 
                  Caption         =   "Expected Amt"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   68
                  Top             =   1398
                  Width           =   1095
               End
            End
            Begin VB.Frame Frame6 
               Height          =   1575
               Left            =   120
               TabIndex        =   50
               Top             =   2280
               Width           =   8055
               Begin VB.TextBox txtDateOfLastPayment 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFC0&
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
                  Left            =   1320
                  Locked          =   -1  'True
                  TabIndex        =   95
                  Top             =   1080
                  Width           =   1695
               End
               Begin VB.TextBox cboDocumentNo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFC0C0&
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
                  Left            =   1320
                  TabIndex        =   12
                  Top             =   240
                  Width           =   1695
               End
               Begin VB.TextBox txtCurrentPeriodDetails 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFC0&
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
                  Left            =   1320
                  Locked          =   -1  'True
                  TabIndex        =   53
                  Top             =   645
                  Width           =   1695
               End
               Begin VB.TextBox txtReceivedTodate 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFC0&
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
                  Left            =   5520
                  Locked          =   -1  'True
                  TabIndex        =   52
                  Top             =   645
                  Width           =   2340
               End
               Begin VB.TextBox txtTransactionAmount 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFC0C0&
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
                  Left            =   5520
                  TabIndex        =   13
                  Top             =   240
                  Width           =   2340
               End
               Begin VB.TextBox txtPaymentStatusDetails 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFC0&
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
                  Left            =   5520
                  Locked          =   -1  'True
                  TabIndex        =   51
                  Top             =   1065
                  Width           =   2340
               End
               Begin VB.Label lblDateOfLastPayment 
                  Caption         =   "Last Pay Date"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   96
                  Top             =   1140
                  Width           =   1095
               End
               Begin VB.Label Label8 
                  Caption         =   "Period"
                  Height          =   210
                  Left            =   240
                  TabIndex        =   58
                  Top             =   727
                  Width           =   495
               End
               Begin VB.Label lblTotalReceived 
                  Caption         =   "Total Received"
                  Height          =   255
                  Left            =   4200
                  TabIndex        =   57
                  Top             =   705
                  Width           =   1335
               End
               Begin VB.Label lblDocumentNo 
                  Caption         =   "Document No"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   56
                  Top             =   300
                  Width           =   1335
               End
               Begin VB.Label Label3 
                  Caption         =   "Amount"
                  Height          =   255
                  Left            =   4200
                  TabIndex        =   55
                  Top             =   300
                  Width           =   615
               End
               Begin VB.Label Label1 
                  Caption         =   "Status"
                  Height          =   210
                  Left            =   4200
                  TabIndex        =   54
                  Top             =   1147
                  Width           =   615
               End
            End
            Begin VB.TextBox txtTransactionNODetails 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   49
               Top             =   6120
               Width           =   2175
            End
            Begin MSDataGridLib.DataGrid receiptDetailsGRID 
               Height          =   1815
               Left            =   120
               TabIndex        =   106
               Top             =   3960
               Width           =   8055
               _ExtentX        =   14208
               _ExtentY        =   3201
               _Version        =   393216
               AllowUpdate     =   0   'False
               BackColor       =   16777152
               HeadLines       =   1
               RowHeight       =   19
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   2
               BeginProperty Column00 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label11 
               Caption         =   "Total Amount"
               Height          =   255
               Left            =   4440
               TabIndex        =   102
               Top             =   5820
               Width           =   1335
            End
            Begin VB.Label Label10 
               Caption         =   "Balance"
               Height          =   255
               Left            =   120
               TabIndex        =   100
               Top             =   5820
               Width           =   1335
            End
            Begin VB.Label Label9 
               Caption         =   "Receipt Amount"
               Height          =   255
               Left            =   4440
               TabIndex        =   98
               Top             =   6180
               Width           =   1335
            End
            Begin VB.Label Label19 
               Caption         =   "Entry No"
               Height          =   255
               Left            =   120
               TabIndex        =   94
               Top             =   6180
               Width           =   1335
            End
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Policies"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         TabIndex        =   29
         Top             =   360
         Width           =   11655
         Begin MSDataGridLib.DataGrid policiesGRID 
            Height          =   6255
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   11033
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777152
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "POLICY INFORMATION"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Proposals"
         Height          =   6615
         Left            =   -74880
         TabIndex        =   27
         Top             =   360
         Width           =   11655
         Begin MSDataGridLib.DataGrid proposalGRID 
            Height          =   6255
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   11033
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777152
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "PROPOSAL DETAILS"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Loans"
         Height          =   6615
         Left            =   -74880
         TabIndex        =   25
         Top             =   360
         Width           =   11655
         Begin MSDataGridLib.DataGrid loanGRID 
            Height          =   6255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   11033
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777152
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "LOAN DETAILS"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame14 
         Height          =   6615
         Left            =   -74880
         TabIndex        =   16
         Top             =   360
         Width           =   11655
         Begin VB.Frame Frame16 
            Height          =   3615
            Left            =   120
            TabIndex        =   24
            Top             =   2760
            Width           =   8055
            Begin MSDataGridLib.DataGrid ReceiptMGRID 
               Height          =   3255
               Left            =   120
               TabIndex        =   103
               Top             =   240
               Width           =   7815
               _ExtentX        =   13785
               _ExtentY        =   5741
               _Version        =   393216
               BackColor       =   16777152
               HeadLines       =   1
               RowHeight       =   19
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   2
               BeginProperty Column00 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
         End
         Begin VB.Frame Frame15 
            Height          =   2655
            Left            =   120
            TabIndex        =   21
            Top             =   120
            Width           =   11415
            Begin VB.TextBox txtRemark 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1560
               TabIndex        =   105
               Top             =   2100
               Width           =   6255
            End
            Begin MSComCtl2.DTPicker DTPickerReceiptDate 
               Height          =   375
               Left            =   7560
               TabIndex        =   3
               Top             =   240
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   661
               _Version        =   393216
               Format          =   55902209
               CurrentDate     =   37953
            End
            Begin VB.ComboBox cboCurrencyCode 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   8880
               Locked          =   -1  'True
               Sorted          =   -1  'True
               TabIndex        =   5
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txtPaymentMethod 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   47
               Top             =   1185
               Width           =   735
            End
            Begin VB.TextBox txtBankNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   46
               Top             =   1680
               Width           =   735
            End
            Begin VB.TextBox txtTransactionNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   8880
               Locked          =   -1  'True
               TabIndex        =   44
               Top             =   1680
               Width           =   2415
            End
            Begin VB.TextBox txtPaymentStatus 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
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
               Left            =   8880
               Locked          =   -1  'True
               TabIndex        =   42
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox txtLocal 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
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
               Left            =   10560
               Locked          =   -1  'True
               TabIndex        =   40
               Top             =   1200
               Width           =   735
            End
            Begin VB.TextBox txtReceiptAmount 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
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
               Left            =   9720
               TabIndex        =   6
               Top             =   720
               Width           =   1575
            End
            Begin VB.TextBox txtAccountingPeriod 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
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
               Left            =   8880
               Locked          =   -1  'True
               TabIndex        =   37
               Top             =   240
               Width           =   2415
            End
            Begin VB.ComboBox cboBankNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   2280
               Sorted          =   -1  'True
               TabIndex        =   9
               Top             =   1680
               Width           =   5535
            End
            Begin VB.TextBox txtChequeNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5880
               TabIndex        =   8
               Top             =   1200
               Width           =   1935
            End
            Begin VB.ComboBox cboPaymentMethod 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   2280
               Sorted          =   -1  'True
               TabIndex        =   7
               Top             =   1185
               Width           =   2295
            End
            Begin VB.TextBox txtPayer 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1560
               TabIndex        =   4
               Top             =   720
               Width           =   6255
            End
            Begin VB.TextBox txtReceiptNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1560
               TabIndex        =   1
               Top             =   262
               Width           =   3015
            End
            Begin VB.TextBox txtReceiptDate 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   5880
               TabIndex        =   2
               Top             =   262
               Width           =   1935
            End
            Begin VB.Label Label12 
               Caption         =   "Remark"
               Height          =   255
               Left            =   240
               TabIndex        =   104
               Top             =   2160
               Width           =   855
            End
            Begin VB.Label Label7 
               Caption         =   "Entries"
               Height          =   255
               Left            =   8040
               TabIndex        =   45
               Top             =   1740
               Width           =   1095
            End
            Begin VB.Label lblPaymentStatus 
               Caption         =   "Status"
               Height          =   210
               Left            =   8040
               TabIndex        =   43
               Top             =   1275
               Width           =   615
            End
            Begin VB.Label lblReferenceNo 
               Caption         =   "Local?"
               Height          =   255
               Left            =   9960
               TabIndex        =   41
               Top             =   1260
               Width           =   735
            End
            Begin VB.Label lblReceiptAmount 
               Caption         =   "Amount"
               Height          =   255
               Left            =   8040
               TabIndex        =   39
               Top             =   780
               Width           =   615
            End
            Begin VB.Label lblCurrentPeriod 
               Caption         =   "Period"
               Height          =   210
               Left            =   8040
               TabIndex        =   38
               Top             =   330
               Width           =   495
            End
            Begin VB.Label Label6 
               Caption         =   "Bank"
               Height          =   255
               Left            =   240
               TabIndex        =   36
               Top             =   1725
               Width           =   495
            End
            Begin VB.Label Label5 
               Caption         =   "Cheque No"
               Height          =   255
               Left            =   4800
               TabIndex        =   35
               Top             =   1245
               Width           =   1215
            End
            Begin VB.Label Label4 
               Caption         =   "Pay Method"
               Height          =   255
               Left            =   240
               TabIndex        =   34
               Top             =   1245
               Width           =   855
            End
            Begin VB.Label Label2 
               Caption         =   "Received From"
               Height          =   255
               Left            =   240
               TabIndex        =   33
               Top             =   780
               Width           =   1095
            End
            Begin VB.Label Label27 
               Caption         =   "Receipt No"
               Height          =   255
               Left            =   240
               TabIndex        =   23
               Top             =   315
               Width           =   1215
            End
            Begin VB.Label lblReceiptDate 
               Caption         =   " Receipt Date"
               Height          =   255
               Left            =   4800
               TabIndex        =   22
               Top             =   315
               Width           =   1095
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Receipts"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3135
            Left            =   8280
            TabIndex        =   17
            Top             =   2880
            Width           =   3255
            Begin VB.CommandButton cmdprintlisting 
               Appearance      =   0  'Flat
               Caption         =   "&Print Receipt Listing"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   120
               TabIndex        =   32
               Top             =   2640
               Width           =   3015
            End
            Begin VB.CommandButton cmdCancel 
               Appearance      =   0  'Flat
               Caption         =   "&Cancel"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   120
               TabIndex        =   20
               Top             =   1680
               Width           =   3015
            End
            Begin VB.CommandButton cmdSearch 
               Appearance      =   0  'Flat
               Caption         =   "&Search"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   120
               TabIndex        =   19
               Top             =   1200
               Width           =   3015
            End
            Begin VB.CommandButton cmdUpdate 
               Appearance      =   0  'Flat
               Caption         =   "&Update"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   120
               TabIndex        =   10
               Top             =   720
               Width           =   3015
            End
            Begin VB.CommandButton cmdAddNew 
               Appearance      =   0  'Flat
               Caption         =   "&Add New"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Width           =   3015
            End
            Begin VB.CommandButton cmdprintreceipt 
               Appearance      =   0  'Flat
               Caption         =   "&Print Receipt"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   120
               TabIndex        =   15
               Top             =   2160
               Width           =   3015
            End
         End
      End
   End
End
Attribute VB_Name = "frmALISMReceiptNew1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsreceipt As clsReceipting, rsReceiptDetails As clsReceiptDetails
Dim rsRCPT As ADODB.Recordset, strSQL As String, bunloadFORM As Boolean
Public rsDEP As ADODB.Recordset, strDEP As String

Private Sub Form_Unload(Cancel As Integer)
    If addpen = True Then
        Cancel = True
        MsgBox "Data entry in progress. Click Refresh to Cancel", vbCritical
    Else
        Cancel = False
    End If
End Sub

Private Sub ClearReceipt()
    Set rsreceipt = New clsReceipting
    rsreceipt.clearRECORD
    Set rsreceipt = Nothing
End Sub

Private Sub DisableControls()
        Set rsreceipt = New clsReceipting
        Dim Bval As Boolean
        rsreceipt.disableDATAENTRY
        Set rsreceipt = Nothing
End Sub

Private Sub UpdatePremium()

    Set rsreceipt = New clsReceipting
    rsreceipt.updateRECORD
    Set rsreceipt = Nothing

End Sub

Sub enableCBRECEIPT()
    frmALISMReceipt.cmdAddNew.Enabled = True
    frmALISMReceipt.cmdUpdate.Enabled = False
    frmALISMReceipt.cmdSearch.Enabled = True
    frmALISMReceipt.cmdCancel.Enabled = True
End Sub

Sub DisableCBRECEIPT()
    frmALISMReceipt.cmdAddNew.Enabled = False
    frmALISMReceipt.cmdUpdate.Enabled = True
    frmALISMReceipt.cmdSearch.Enabled = False
    frmALISMReceipt.cmdCancel.Enabled = True
End Sub

Private Sub cbobankNo_GotFocus()
    Set rsreceipt = New clsReceipting
    rsreceipt.selectBankNOGotFocus
    Set rsreceipt = Nothing
    
End Sub

Private Sub cboBankNo_KeyPress(KeyAscii As Integer)

        Set rsreceipt = New clsReceipting
        rsreceipt.selectBankNoKeyPress (KeyAscii)
        Set rsreceipt = Nothing

End Sub

Private Sub cbobankNo_LostFocus()
        Set rsreceipt = New clsReceipting
        rsreceipt.selectBankNoLostFocus
        Set rsreceipt = Nothing
End Sub

Private Sub cboCurrencyCode_GotFocus()
        Set rsreceipt = New clsReceipting
        rsreceipt.selectCURRENCYGOTFOCUS
        Set rsreceipt = Nothing
End Sub

Private Sub cboCurrencyCode_KeyPress(KeyAscii As Integer)
        Set rsreceipt = New clsReceipting
        rsreceipt.selectCURRENCYKEYPRESS (KeyAscii)
        Set rsreceipt = Nothing
End Sub
Private Sub cboCurrencyCode_LostFocus()
        Set rsreceipt = New clsReceipting
        rsreceipt.selectCURRENCYLOSTFOCUS
        Set rsreceipt = Nothing
End Sub

Private Sub cboDocumentNo_LostFocus()
            Set rsReceiptDetails = New clsReceiptDetails
            rsReceiptDetails.processDOCUMENTNO
            rsReceiptDetails.loadRECEIPTDETAILSGRID
            Set rsReceiptDetails = Nothing
End Sub
Private Sub cboEmployer_GotFocus()
        Set rsReceiptDetails = New clsReceiptDetails
        
        If bloadINVOICE = True Then
            loadInvoiceGOTFOCUS
        Else: rsReceiptDetails.SelectEmployerGotFocus
        End If
        
        Set rsReceiptDetails = Nothing
End Sub

Private Sub cboEmployer_LostFocus()
        If bloadINVOICE = True Then
                loadINVOICELOSTFOCUS
        Else
            Set rsReceiptDetails = New clsReceiptDetails
            rsReceiptDetails.selectEmployerLostFocus
            Set rsReceiptDetails = Nothing
        End If
End Sub


Private Sub cboPaymentMethod_GotFocus()
        
        Set rsreceipt = New clsReceipting
        rsreceipt.selectPaymentMethodGotFocus
        Set rsreceipt = Nothing

End Sub

Private Sub cboPaymentMethod_KeyPress(KeyAscii As Integer)
        
        Set rsreceipt = New clsReceipting
        rsreceipt.selectPaymentMethodKeyPress (KeyAscii)
        Set rsreceipt = Nothing

End Sub

Private Sub cboPaymentMethod_LostFocus()

        Set rsreceipt = New clsReceipting
        rsreceipt.selectPaymentMethodLostFocus
        Set rsreceipt = Nothing

End Sub
Private Sub cboPaymentMethodGotFocus()
        
        Set rsreceipt = New clsReceipting
        rsreceipt.selectPaymentMethodGotFocus
        Set rsreceipt = Nothing

End Sub

Private Sub cboPaymentMethodKeyPress(KeyAscii As Integer)
        
        Set rsreceipt = New clsReceipting
        rsreceipt.selectPaymentMethodKeyPress (KeyAscii)
        Set rsreceipt = Nothing
        
End Sub

Private Sub cboPaymentMethodLostFocus()
        
        Set rsreceipt = New clsReceipting
        rsreceipt.selectPaymentMethodLostFocus
        Set rsreceipt = Nothing
        
End Sub


Private Sub cboReceiptType_GotFocus()
            
            Set rsReceiptDetails = New clsReceiptDetails
            rsReceiptDetails.selectRECEIPTTYPEGOTFOCUS
            Set rsReceiptDetails = Nothing

End Sub

Private Sub cboReceiptType_KeyPress(KeyAscii As Integer)
        
        Set rsReceiptDetails = New clsReceiptDetails
        rsReceiptDetails.selectRECEIPTTYPEKEYPRESS (KeyAscii)
        Set rsReceiptDetails = Nothing

End Sub

Private Sub cboReceiptType_LostFocus()
        If frmALISMReceiptNew.cboReceiptType.Text <= "" Then Exit Sub
        Set rsReceiptDetails = New clsReceiptDetails
        rsReceiptDetails.selectRECEIPTTYPELOSTFOCUS
        Set rsReceiptDetails = Nothing
        
End Sub

Private Sub cmdAddNew_Click()
        Set rsreceipt = New clsReceipting
        clearALLRECORD
        rsreceipt.lockRECEIPTNO
        rsreceipt.addRECORD
        Set rsreceipt = Nothing
End Sub

Private Sub cmdAddNewDetails_Click()
        bloadINVOICE = False
        Set rsReceiptDetails = New clsReceiptDetails
        rsReceiptDetails.addRECORD
        Set rsReceiptDetails = Nothing

End Sub

Private Sub cmdCancel_Click()
        Set rsreceipt = New clsReceipting
        addpen = False
        rsreceipt.enableCBRECEIPT
        rsreceipt.clearRECORD
        rsreceipt.disableDATAENTRY
        Set rsreceipt = Nothing

End Sub

Private Sub cmdCancelDetails_Click()
        addpen = False
        Set rsReceiptDetails = New clsReceiptDetails
        rsReceiptDetails.clearRECORD
        rsReceiptDetails.disableDATAENTRY
        Set rsReceiptDetails = Nothing
End Sub

Private Sub cmdprintlisting_Click()
    Load frmReceiptListing
    frmReceiptListing.Show 1, Me
End Sub

Private Sub cmdprintreceipt_Click()
        Load frmNewReceipt
        frmNewReceipt.Show 1, Me

End Sub

Private Sub cmdPrintReceiptDetails_Click()
    If frmALISMReceiptNew.txtReceiptNo.Text <= "" Then
        MsgBox "Cannot Use this Form Directly, Load the Receipt on the First Tab", vbOKOnly
        
        Exit Sub
        Else: Load frmNewReceipt
        frmNewReceipt.Show 1, Me
    End If

End Sub

Private Sub cmdPrintReportListing_Click()
    If frmALISMReceiptNew.txtReceiptNo.Text <= "" Then
        MsgBox "Cannot Use this Form Directly, Load the Receipt on the First Tab", vbOKOnly
        Exit Sub
    End If

End Sub

Private Sub cmdSearch_Click()
    Set rsreceipt = New clsReceipting
        rsreceipt.searchRECORD
        If bsearchRECORD = True Then
            rsreceipt.loadRECEIPTDETAILS
            rsreceipt.loadEMPLOYER
            rsreceipt.showRECEIPTITEMS
        End If
    Set rsreceipt = Nothing
End Sub

Private Sub cmdUpdate_Click()
    Set rsreceipt = New clsReceipting
        rsreceipt.updateRECORD
    Set rsreceipt = Nothing
Exit Sub
End Sub


Private Sub cmdSearchDetails_Click()
    Set rsReceiptDetails = New clsReceiptDetails
        rsReceiptDetails.searchRECORD
        rsReceiptDetails.disableDATAENTRY
        rsReceiptDetails.showRECEIPTITEMS
    Set rsReceiptDetails = Nothing
End Sub


Private Sub cmdUpdateDetails_Click()
    Set rsReceiptDetails = New clsReceiptDetails
        rsReceiptDetails.processUPDATE
        rsReceiptDetails.UpdatePolicyLedger
        rsReceiptDetails.enableCBRECEIPT
        rsReceiptDetails.loadRECEIPTDETAILSGRID
    Set rsReceiptDetails = Nothing
End Sub

Private Sub DTPickerReceiptDate_Change()
On Error GoTo err

    With frmALISMReceiptNew
            .txtReceiptDate.Text = .DTPickerReceiptDate.Value
    End With

Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Form_Load()
        OpenConnection
End Sub

Private Sub Form_Activate()
        Call DisableControls
        Set rsreceipt = New clsReceipting
            rsreceipt.clearRECORD
            rsreceipt.enableCBRECEIPT
            rsreceipt.disableDATAENTRY
        Set rsreceipt = Nothing
        
        Set rsReceiptDetails = New clsReceiptDetails
            rsReceiptDetails.clearRECORD
            rsReceiptDetails.enableCBRECEIPT
            rsReceiptDetails.disableDATAENTRY
        Set rsReceiptDetails = Nothing
        
        frmALISMReceiptNew.SSTReceipt.Tab = 0
End Sub

Private Sub txtEmployeeNo_LostFocus()
        Set rsReceiptDetails = New clsReceiptDetails
            rsReceiptDetails.selectEmployerLostFocus
        Set rsReceiptDetails = Nothing
End Sub

Public Sub generateRECEIPTNo()
            Set rsReceiptDetails = New clsReceipting
            rsReceiptDetails.createRECEIPT
            Set rsReceiptDetails = Nothing
End Sub



Private Sub txtTransactionAmount_LostFocus()
            If frmALISMReceiptNew.txtTransactionAmount.Text <= "" Then Exit Sub
            Set rsReceiptDetails = New clsReceiptDetails
            rsReceiptDetails.processReceipt
            Set rsReceiptDetails = Nothing
End Sub

