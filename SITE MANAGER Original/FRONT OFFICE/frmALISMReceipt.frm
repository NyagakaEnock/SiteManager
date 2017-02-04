VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmALISMReceipt 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A.L.I.S ENTERPRISE [PAYMENT RECEIPTS PROCESSING]"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   1560
   ClientWidth     =   11880
   Icon            =   "frmALISMReceipt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   11880
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   5
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
      TabPicture(0)   =   "frmALISMReceipt.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame14"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "LOAN DATASHEET"
      TabPicture(1)   =   "frmALISMReceipt.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "PROPOSAL DATASHEET"
      TabPicture(2)   =   "frmALISMReceipt.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "POLICY DATASHEET"
      TabPicture(3)   =   "frmALISMReceipt.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "RECEIPTS DATASHEET"
      TabPicture(4)   =   "frmALISMReceipt.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame17"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame17 
         Height          =   6135
         Left            =   -74880
         TabIndex        =   62
         Top             =   360
         Width           =   11655
         Begin MSDataGridLib.DataGrid ReceiptGrid 
            Height          =   5775
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   10186
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
            Caption         =   "RECEIPT DATASHEET"
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
         Height          =   6135
         Left            =   -74880
         TabIndex        =   60
         Top             =   360
         Width           =   11655
         Begin MSDataGridLib.DataGrid policiesGRID 
            Height          =   5775
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   10186
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
         Height          =   6135
         Left            =   -74880
         TabIndex        =   58
         Top             =   360
         Width           =   11655
         Begin MSDataGridLib.DataGrid proposalGRID 
            Height          =   5775
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   10186
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
         Height          =   6135
         Left            =   -74880
         TabIndex        =   56
         Top             =   360
         Width           =   11655
         Begin MSDataGridLib.DataGrid loanGRID 
            Height          =   5775
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   10186
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
         Height          =   6135
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   11655
         Begin VB.TextBox txtTransactionNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   375
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   5640
            Width           =   2175
         End
         Begin VB.Frame Frame16 
            Height          =   2535
            Left            =   120
            TabIndex        =   42
            Top             =   3000
            Width           =   8055
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
               Height          =   405
               Left            =   6960
               Locked          =   -1  'True
               TabIndex        =   46
               Top             =   1965
               Width           =   975
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
               Left            =   4800
               TabIndex        =   5
               Top             =   360
               Width           =   3135
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
               Left            =   4800
               Locked          =   -1  'True
               TabIndex        =   45
               Top             =   890
               Width           =   3135
            End
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
               TabIndex        =   44
               Top             =   1965
               Width           =   2175
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
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   43
               Top             =   890
               Width           =   2175
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
               Left            =   1320
               Sorted          =   -1  'True
               TabIndex        =   6
               Top             =   1450
               Width           =   2175
            End
            Begin VB.TextBox txtChequeNo 
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
               Left            =   4800
               TabIndex        =   8
               Top             =   1965
               Width           =   1455
            End
            Begin VB.ComboBox cboBankNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   4800
               Sorted          =   -1  'True
               TabIndex        =   7
               Top             =   1450
               Width           =   3135
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
               TabIndex        =   4
               Top             =   360
               Width           =   2175
            End
            Begin VB.Label lblPaymentStatus 
               Caption         =   "Status"
               Height          =   210
               Left            =   6360
               TabIndex        =   55
               Top             =   2062
               Width           =   615
            End
            Begin VB.Label lblReceiptAmount 
               Caption         =   "Receipt Amount"
               Height          =   255
               Left            =   3600
               TabIndex        =   54
               Top             =   420
               Width           =   1335
            End
            Begin VB.Label lblDocumentNo 
               Caption         =   "Document No"
               Height          =   255
               Left            =   240
               TabIndex        =   53
               Top             =   420
               Width           =   1335
            End
            Begin VB.Label lblDateOfLastPayment 
               Caption         =   "Last Pay Date"
               Height          =   255
               Left            =   240
               TabIndex        =   52
               Top             =   2025
               Width           =   1095
            End
            Begin VB.Label lblTotalReceived 
               Caption         =   "Total Received"
               Height          =   255
               Left            =   3600
               TabIndex        =   51
               Top             =   950
               Width           =   1335
            End
            Begin VB.Label lblCurrentPeriod 
               Caption         =   "Period"
               Height          =   210
               Left            =   240
               TabIndex        =   50
               Top             =   987
               Width           =   495
            End
            Begin VB.Label Label4 
               Caption         =   "Pay Method"
               Height          =   255
               Left            =   240
               TabIndex        =   49
               Top             =   1503
               Width           =   855
            End
            Begin VB.Label Label5 
               Caption         =   "Cheque No"
               Height          =   255
               Left            =   3600
               TabIndex        =   48
               Top             =   2025
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Bank"
               Height          =   255
               Left            =   3600
               TabIndex        =   47
               Top             =   1495
               Width           =   495
            End
         End
         Begin VB.Frame Frame15 
            Height          =   2895
            Left            =   120
            TabIndex        =   29
            Top             =   120
            Width           =   8055
            Begin VB.TextBox txtPayer 
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
               TabIndex        =   70
               Top             =   2280
               Width           =   6615
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
               TabIndex        =   65
               Top             =   1320
               Width           =   1695
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
               TabIndex        =   64
               Top             =   1320
               Width           =   4815
            End
            Begin VB.TextBox txtReceiptNo 
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
               TabIndex        =   34
               Top             =   345
               Width           =   1215
            End
            Begin VB.TextBox txtReceiptDate 
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
               Left            =   6720
               TabIndex        =   33
               Top             =   345
               Width           =   1215
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
               TabIndex        =   1
               Top             =   345
               Width           =   3135
            End
            Begin VB.ComboBox cboCostCenter 
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
               Left            =   3480
               Sorted          =   -1  'True
               TabIndex        =   3
               Top             =   1845
               Width           =   2175
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
               Left            =   6480
               Locked          =   -1  'True
               TabIndex        =   32
               Top             =   1845
               Width           =   1455
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
               TabIndex        =   2
               Top             =   840
               Width           =   4815
            End
            Begin VB.TextBox txtEmployeeNo 
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
               Left            =   1320
               TabIndex        =   31
               Top             =   1845
               Width           =   1695
            End
            Begin VB.ComboBox cboEmployerCode 
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
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   30
               Top             =   840
               Width           =   1695
            End
            Begin VB.Label Label8 
               Caption         =   "Payer"
               Height          =   255
               Left            =   240
               TabIndex        =   71
               Top             =   2340
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "Expected Amt"
               Height          =   255
               Left            =   240
               TabIndex        =   66
               Top             =   1398
               Width           =   1095
            End
            Begin VB.Label lblReferenceNo 
               Caption         =   "Ref No"
               Height          =   255
               Left            =   5760
               TabIndex        =   41
               Top             =   1898
               Width           =   615
            End
            Begin VB.Label Label27 
               Caption         =   "Receipt No"
               Height          =   255
               Left            =   240
               TabIndex        =   40
               Top             =   398
               Width           =   1215
            End
            Begin VB.Label lblReceiptDate 
               Caption         =   " Date"
               Height          =   255
               Left            =   6240
               TabIndex        =   39
               Top             =   398
               Width           =   495
            End
            Begin VB.Label Label29 
               Caption         =   "Type"
               Height          =   255
               Left            =   2640
               TabIndex        =   38
               Top             =   398
               Width           =   375
            End
            Begin VB.Label lblCostCenter 
               Caption         =   "CC"
               Height          =   255
               Left            =   3120
               TabIndex        =   37
               Top             =   1898
               Width           =   735
            End
            Begin VB.Label Label2 
               Caption         =   "Employer"
               Height          =   255
               Left            =   240
               TabIndex        =   36
               Top             =   893
               Width           =   1095
            End
            Begin VB.Label Label3 
               Caption         =   "Emp #"
               Height          =   255
               Left            =   240
               TabIndex        =   35
               Top             =   1898
               Width           =   615
            End
         End
         Begin VB.Frame Frame2 
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
            TabIndex        =   16
            Top             =   120
            Width           =   3255
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
               TabIndex        =   22
               Top             =   1860
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
               TabIndex        =   21
               Top             =   645
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
               TabIndex        =   20
               Top             =   1050
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
               TabIndex        =   19
               Top             =   1455
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
               TabIndex        =   18
               Top             =   240
               Width           =   2055
            End
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
               TabIndex        =   17
               Top             =   2265
               Width           =   2055
            End
            Begin VB.Label lblUnitCountBeforePayment 
               Caption         =   "Prem Prior:"
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   1935
               Width           =   855
            End
            Begin VB.Label lblUnitssPaid 
               Caption         =   "Prem Paid:"
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Top             =   1530
               Width           =   855
            End
            Begin VB.Label lblStatusCode 
               Caption         =   "Status Code"
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   1125
               Width           =   975
            End
            Begin VB.Label lblDueDate 
               Caption         =   "Due Date:"
               Height          =   255
               Left            =   120
               TabIndex        =   25
               Top             =   720
               Width           =   855
            End
            Begin VB.Label lblUnitCount 
               Caption         =   "Prem Count:"
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   315
               Width           =   975
            End
            Begin VB.Label lblSuspenseAccount 
               Caption         =   "Suspense"
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   2340
               Width           =   735
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
            TabIndex        =   12
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
               TabIndex        =   69
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
               TabIndex        =   15
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
               TabIndex        =   14
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
               TabIndex        =   9
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
               TabIndex        =   13
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
               TabIndex        =   10
               Top             =   2160
               Width           =   3015
            End
         End
         Begin VB.Label Label7 
            Caption         =   "Transaction No"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   5700
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmALISMReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsreceipt As cReceipting
Dim rsRCPT As ADODB.Recordset, strSQL As String
Public rsDEP As ADODB.Recordset, strDEP As String

Sub ClearReceipt()
    Set rsreceipt = New cReceipting
    rsreceipt.clearRECORD
    Set rsreceipt = Nothing
End Sub

Private Sub DisableControls()
'On Error GoTo err
        Set rsreceipt = New cReceipting
        Dim Bval As Boolean
        rsreceipt.disableDATAENTRY
        Set rsreceipt = Nothing
    Exit Sub
err:
    UpdateErrorMessage
End Sub

Private Sub UpdatePremium()

    Set rsreceipt = New cReceipting
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
    Set rsreceipt = New cReceipting
    rsreceipt.selectBankNOGotFocus
    Set rsreceipt = Nothing
    
End Sub

Private Sub cboBankNo_KeyPress(KeyAscii As Integer)

        Set rsreceipt = New cReceipting
        rsreceipt.selectBankNoKeyPress (KeyAscii)
        Set rsreceipt = Nothing

End Sub

Private Sub cbobankNo_LostFocus()

        Set rsreceipt = New cReceipting
        rsreceipt.selectBankNoLostFocus
        Set rsreceipt = Nothing

End Sub

Private Sub cboCostCenter_GotFocus()

        Set rsreceipt = New cReceipting
        rsreceipt.selectCOSTCENTERGOTFOCUS
        Set rsreceipt = Nothing
End Sub

Private Sub cboCostCenter_KeyPress(KeyAscii As Integer)
        
        Set rsreceipt = New cReceipting
        rsreceipt.selectCOSTCENTERKEYPRESS (KeyAscii)
        Set rsreceipt = Nothing

End Sub
Private Sub cboCostCenter_LostFocus()
        
        Set rsreceipt = New cReceipting
        rsreceipt.selectCOSTCENTERLOSTFOCUS
        Set rsreceipt = Nothing

End Sub


Private Sub cboPaymentMethod_GotFocus()
        
        Set rsreceipt = New cReceipting
        rsreceipt.selectPaymentMethodGotFocus
        Set rsreceipt = Nothing

End Sub

Private Sub cboPaymentMethod_KeyPress(KeyAscii As Integer)
        
        Set rsreceipt = New cReceipting
        rsreceipt.selectPaymentMethodKeyPress (KeyAscii)
        Set rsreceipt = Nothing

End Sub

Private Sub cboPaymentMethod_LostFocus()

        Set rsreceipt = New cReceipting
        rsreceipt.selectPaymentMethodLostFocus
        Set rsreceipt = Nothing

End Sub
Private Sub cboPaymentMethodGotFocus()
        
        Set rsreceipt = New cReceipting
        rsreceipt.selectPaymentMethodGotFocus
        Set rsreceipt = Nothing

End Sub

Private Sub cboPaymentMethodKeyPress(KeyAscii As Integer)
        
        Set rsreceipt = New cReceipting
        rsreceipt.selectPaymentMethodKeyPress (KeyAscii)
        Set rsreceipt = Nothing
        
End Sub

Private Sub cboPaymentMethodLostFocus()
        
        Set rsreceipt = New cReceipting
        rsreceipt.selectPaymentMethodLostFocus
        Set rsreceipt = Nothing
        
End Sub

Private Sub cboEmployer_GotFocus()
'On Error GoTo err
        Set rsreceipt = New cReceipting
        cboEmployer.Clear
        rsreceipt.SelectEmployerGotFocus
        Set rsreceipt = Nothing

Exit Sub

err:
    UpdateErrorMessage
End Sub

Private Sub cboEmployer_LostFocus()
'On Error GoTo err
        
        Set rsreceipt = New cReceipting
        rsreceipt.selectEmployerLostFocus
        Set rsreceipt = Nothing
    
Exit Sub

err:
    UpdateErrorMessage
End Sub



Private Sub cboReceiptType_GotFocus()
            
            Set rsreceipt = New cReceipting
            rsreceipt.selectRECEIPTTYPEGOTFOCUS
            Set rsreceipt = Nothing

End Sub

Private Sub cboReceiptType_KeyPress(KeyAscii As Integer)
        
        Set rsreceipt = New cReceipting
        rsreceipt.selectRECEIPTTYPEKEYPRESS (KeyAscii)
        Set rsreceipt = Nothing

End Sub

Private Sub cboReceiptType_LostFocus()

        Set rsreceipt = New cReceipting
        rsreceipt.selectRECEIPTTYPELOSTFOCUS
        Set rsreceipt = Nothing
        
End Sub

Private Sub cmdAddNew_Click()

        Set rsreceipt = New cReceipting
        rsreceipt.addRECORD
        Set rsreceipt = Nothing
        Me.txtTransactionNo.Text = GetTransactionNo
End Sub

Private Sub cboDocumentNo_LostFocus()
            
            Set rsreceipt = New cReceipting
            rsreceipt.processDOCUMENTNO
            Set rsreceipt = Nothing

Exit Sub

End Sub

Private Sub cmdCancel_Click()
   
        enableCBRECEIPT
        ClearReceipt
        DisableControls
    
End Sub

Private Sub cmdprintlisting_Click()
'On Error GoTo err
Load frmReceiptListing
frmReceiptListing.Show 1, Me

Exit Sub
err:
ErrorMessage
End Sub

Private Sub cmdprintreceipt_Click()
'On Error GoTo err
        Load frmtrinityreciept
        frmtrinityreciept.Show 1, Me

Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdUpdate_Click()
    Me.txtTransactionNo.Text = GetTransactionNo
    
    Set rsreceipt = New cReceipting
    
        rsreceipt.processUPDATE
        rsreceipt.UpdatePolicyLedger
        rsreceipt.enableCBRECEIPT
        
    Set rsreceipt = Nothing
    
Exit Sub
End Sub

Private Function GetTransactionNo() As Double
'On Error GoTo err

    Dim N1 As Variant, N2 As Variant
    
    Set rsFindRecord = New ADODB.Recordset
    rsFindRecord.Open "SELECT MAX(TransactionNo) AS LNum FROM ALISMReceipt;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetTransactionNo = 1: Exit Function
    ElseIf IsNull(rsFindRecord!lnum) = True Then
        GetTransactionNo = 1: Exit Function
    Else
        N1 = rsFindRecord!lnum
        N2 = N1 + 1
        
        GetTransactionNo = N2
    End If
    
    Exit Function
err:
    ErrorMessage
End Function

Private Sub cmdSearch_Click()
    Set rsreceipt = New cReceipting
        rsreceipt.searchRECORD
    Set rsreceipt = Nothing
End Sub

Private Sub Form_Load()
'On Error GoTo err

        OpenConnection
        Call DisableControls
        Set rsreceipt = New cReceipting
            rsreceipt.enableCBRECEIPT
        Set rsreceipt = Nothing

Exit Sub

err:
        UpdateErrorMessage
End Sub

Private Sub txtEmployeeNo_LostFocus()
'On Error GoTo err:
        Set rsreceipt = New cReceipting
            rsreceipt.selectEmployeeNoLostFocus
        Set rsreceipt = Nothing
Exit Sub

err:
    UpdateErrorMessage
End Sub

Private Sub txtReceiptAmount_lostFocus()
'On Error GoTo err
            
            Set rsreceipt = New cReceipting
            rsreceipt.processReceipt
            Set rsreceipt = Nothing
Exit Sub

err:
    UpdateErrorMessage
End Sub

Public Sub generateRECEIPTNo()
'On Error GoTo err

            Set rsreceipt = New cReceipting
            rsreceipt.createRECEIPT
            Set rsreceipt = Nothing

Exit Sub

err:
    UpdateErrorMessage
End Sub

