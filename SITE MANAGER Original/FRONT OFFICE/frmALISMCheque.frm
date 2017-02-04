VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmALISMCheque 
   BackColor       =   &H80000016&
   Caption         =   "Cheque"
   ClientHeight    =   7005
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11385
   Icon            =   "frmALISMCheque.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTabCheque 
      Height          =   6975
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12303
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Assign Cheque No"
      TabPicture(0)   =   "frmALISMCheque.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label30"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CheckGRID"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtTotalAmount"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Cheque Entries"
      TabPicture(1)   =   "frmALISMCheque.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Cheque Listing"
      TabPicture(2)   =   "frmALISMCheque.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame12"
      Tab(2).Control(1)=   "Frame11"
      Tab(2).Control(2)=   "Frame10"
      Tab(2).Control(3)=   "Frame9"
      Tab(2).ControlCount=   4
      Begin VB.Frame Frame12 
         Caption         =   "Issued Cheques"
         Height          =   3255
         Left            =   -69120
         TabIndex        =   103
         Top             =   3600
         Width           =   5415
         Begin MSDataGridLib.DataGrid IssuedGrid 
            Height          =   2895
            Left            =   120
            TabIndex        =   104
            Top             =   240
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   5106
            _Version        =   393216
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
      Begin VB.Frame Frame11 
         Caption         =   "Approved Cheques"
         Height          =   3255
         Left            =   -69120
         TabIndex        =   101
         Top             =   360
         Width           =   5415
         Begin MSDataGridLib.DataGrid ApprovedGrid 
            Height          =   2895
            Left            =   120
            TabIndex        =   102
            Top             =   240
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   5106
            _Version        =   393216
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
      Begin VB.Frame Frame10 
         Caption         =   "Authorized Cheques"
         Height          =   3255
         Left            =   -74880
         TabIndex        =   99
         Top             =   3600
         Width           =   5775
         Begin MSDataGridLib.DataGrid AuthorizedGrid 
            Height          =   2895
            Left            =   120
            TabIndex        =   100
            Top             =   240
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   5106
            _Version        =   393216
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
      Begin VB.Frame Frame9 
         Caption         =   "Pending Cheques"
         Height          =   3255
         Left            =   -74880
         TabIndex        =   97
         Top             =   360
         Width           =   5775
         Begin MSDataGridLib.DataGrid PendingGrid 
            Height          =   2895
            Left            =   120
            TabIndex        =   98
            Top             =   240
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   5106
            _Version        =   393216
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
      Begin VB.TextBox txtTotalAmount 
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
         Left            =   6000
         TabIndex        =   88
         Top             =   4080
         Width           =   2775
      End
      Begin MSDataGridLib.DataGrid CheckGRID 
         Height          =   1455
         Left            =   120
         TabIndex        =   74
         Top             =   2640
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   2566
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
      Begin VB.Frame Frame7 
         Height          =   6495
         Left            =   -74880
         TabIndex        =   55
         Top             =   360
         Width           =   11175
         Begin VB.TextBox txtEntryTotalAmount 
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
            Left            =   6600
            TabIndex        =   94
            Top             =   4800
            Width           =   2055
         End
         Begin VB.TextBox txtEntryDatePrepared 
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
            Left            =   6600
            TabIndex        =   91
            Top             =   5400
            Width           =   2055
         End
         Begin VB.TextBox txtEntryPreparedBy 
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
            Left            =   1800
            TabIndex        =   90
            Top             =   5400
            Width           =   2895
         End
         Begin VB.TextBox txtEntryStatus 
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
            Height          =   400
            Left            =   9480
            Locked          =   -1  'True
            TabIndex        =   86
            Top             =   240
            Width           =   1575
         End
         Begin VB.Frame Frame8 
            Height          =   3735
            Left            =   9000
            TabIndex        =   83
            Top             =   2640
            Width           =   2175
            Begin VB.CommandButton cmdPrintEntry 
               Caption         =   "&Print"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               TabIndex        =   85
               Top             =   2100
               Width           =   1935
            End
            Begin VB.CommandButton cmdSearchEntry 
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
               Height          =   495
               Left            =   120
               TabIndex        =   84
               Top             =   1110
               Width           =   1935
            End
            Begin VB.CommandButton cmdAddNewEntry 
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
               Height          =   495
               Left            =   120
               TabIndex        =   0
               Top             =   120
               Width           =   1935
            End
            Begin VB.CommandButton cmdEntryUpdate 
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
               Height          =   495
               Left            =   120
               TabIndex        =   9
               Top             =   615
               Width           =   1935
            End
            Begin VB.CommandButton cmdCancelEntry 
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
               Height          =   495
               Left            =   120
               TabIndex        =   10
               Top             =   1605
               Width           =   1935
            End
         End
         Begin VB.TextBox txtEntryChequeNo 
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
            Left            =   1680
            TabIndex        =   15
            Top             =   240
            Width           =   3015
         End
         Begin VB.TextBox txtEntryChequeDate 
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
            Left            =   6240
            TabIndex        =   80
            Top             =   240
            Width           =   2535
         End
         Begin VB.CommandButton cmdFirstEntry 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   3
            Left            =   7125
            Picture         =   "frmALISMCheque.frx":0496
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   5880
            Width           =   1695
         End
         Begin VB.CommandButton cmdFirstEntry 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   2
            Left            =   5550
            Picture         =   "frmALISMCheque.frx":08D8
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   5880
            Width           =   1575
         End
         Begin VB.CommandButton cmdFirstEntry 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   1
            Left            =   3855
            Picture         =   "frmALISMCheque.frx":0D1A
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   5880
            Width           =   1695
         End
         Begin VB.CommandButton cmdFirstEntry 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   0
            Left            =   2160
            Picture         =   "frmALISMCheque.frx":115C
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   5880
            Width           =   1695
         End
         Begin VB.Frame Frame1 
            Height          =   1935
            Left            =   120
            TabIndex        =   56
            Top             =   720
            Width           =   10935
            Begin VB.TextBox txtAmount 
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
               TabIndex        =   64
               Top             =   1366
               Width           =   3015
            End
            Begin VB.TextBox txtReference 
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
               Left            =   6120
               TabIndex        =   63
               Top             =   1410
               Width           =   4695
            End
            Begin VB.ComboBox cboClaimCode 
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
               Sorted          =   -1  'True
               TabIndex        =   62
               Top             =   1004
               Width           =   3015
            End
            Begin VB.ComboBox Combo2 
               BackColor       =   &H00FFC0C0&
               Height          =   315
               Left            =   1440
               TabIndex        =   61
               Text            =   "Combo1"
               Top             =   5280
               Width           =   2535
            End
            Begin VB.ComboBox cboDocumentNo 
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
               Left            =   6120
               TabIndex        =   60
               Top             =   1050
               Width           =   4695
            End
            Begin VB.ComboBox cboPaymentType 
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
               Sorted          =   -1  'True
               TabIndex        =   59
               Top             =   600
               Width           =   3015
            End
            Begin VB.TextBox txtRequisitionDate 
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
               TabIndex        =   58
               Top             =   240
               Width           =   3015
            End
            Begin VB.TextBox txtPaymentDescription 
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
               Left            =   6120
               TabIndex        =   57
               Top             =   645
               Width           =   4695
            End
            Begin VB.ComboBox cboRequisitionNo 
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
               Left            =   6120
               Sorted          =   -1  'True
               TabIndex        =   1
               Top             =   240
               Width           =   4695
            End
            Begin VB.Label Label1 
               Caption         =   "Requisition No"
               Height          =   255
               Left            =   4800
               TabIndex        =   73
               Top             =   300
               Width           =   1095
            End
            Begin VB.Label Label8 
               Caption         =   "Payment Code"
               Height          =   255
               Left            =   240
               TabIndex        =   72
               Top             =   1065
               Width           =   1335
            End
            Begin VB.Label Label7 
               Caption         =   "Amount"
               Height          =   255
               Left            =   240
               TabIndex        =   71
               Top             =   1440
               Width           =   1335
            End
            Begin VB.Label Label5 
               Caption         =   "Requisition Date"
               Height          =   255
               Left            =   240
               TabIndex        =   70
               Top             =   315
               Width           =   1215
            End
            Begin VB.Label Label9 
               Caption         =   "Document No"
               Height          =   255
               Left            =   4800
               TabIndex        =   69
               Top             =   1065
               Width           =   1815
            End
            Begin VB.Label Label4 
               Caption         =   "Reference "
               Height          =   255
               Left            =   4800
               TabIndex        =   68
               Top             =   1485
               Width           =   855
            End
            Begin VB.Label Label10 
               Caption         =   "Label1"
               Height          =   255
               Left            =   600
               TabIndex        =   67
               Top             =   5280
               Width           =   1815
            End
            Begin VB.Label Label12 
               Caption         =   "Description"
               Height          =   255
               Left            =   4800
               TabIndex        =   66
               Top             =   660
               Width           =   1815
            End
            Begin VB.Label Label13 
               Caption         =   "PaymentType"
               Height          =   255
               Left            =   240
               TabIndex        =   65
               Top             =   660
               Width           =   1335
            End
         End
         Begin MSDataGridLib.DataGrid ChequeGRID 
            Height          =   2055
            Left            =   120
            TabIndex        =   75
            Top             =   2760
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   3625
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
         Begin VB.Label Label33 
            Caption         =   "Amount"
            Height          =   255
            Left            =   5040
            TabIndex        =   95
            Top             =   4875
            Width           =   1335
         End
         Begin VB.Label Label32 
            Caption         =   "Date Prepared"
            Height          =   255
            Left            =   5040
            TabIndex        =   93
            Top             =   5475
            Width           =   1095
         End
         Begin VB.Label Label31 
            Caption         =   "Prepared By"
            Height          =   255
            Left            =   240
            TabIndex        =   92
            Top             =   5475
            Width           =   975
         End
         Begin VB.Label Label29 
            Caption         =   "Status"
            Height          =   255
            Left            =   8880
            TabIndex        =   87
            Top             =   315
            Width           =   615
         End
         Begin VB.Label Label28 
            Caption         =   "Cheque No"
            Height          =   255
            Left            =   360
            TabIndex        =   82
            Top             =   315
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "Cheque Date"
            Height          =   255
            Left            =   4920
            TabIndex        =   81
            Top             =   315
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2415
         Left            =   120
         TabIndex        =   38
         Top             =   4440
         Width           =   8775
         Begin VB.TextBox txtDateScheduled 
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
            Left            =   5160
            TabIndex        =   115
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtScheduledBy 
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
            Left            =   1440
            TabIndex        =   114
            Top             =   1320
            Width           =   2295
         End
         Begin VB.TextBox txtScheduled 
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
            Left            =   8280
            TabIndex        =   113
            Top             =   1320
            Width           =   255
         End
         Begin VB.TextBox txtIssued 
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
            Left            =   8280
            TabIndex        =   108
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox txtAuthorized 
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
            Left            =   8280
            TabIndex        =   107
            Top             =   930
            Width           =   255
         End
         Begin VB.TextBox txtApproved 
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
            Left            =   8280
            TabIndex        =   106
            Top             =   525
            Width           =   255
         End
         Begin VB.TextBox txtPrepared 
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
            Left            =   8280
            TabIndex        =   105
            Top             =   120
            Width           =   255
         End
         Begin VB.TextBox txtPreparedBy 
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
            Left            =   1440
            TabIndex        =   46
            Top             =   120
            Width           =   2295
         End
         Begin VB.TextBox txtDatePrepared 
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
            Left            =   5160
            TabIndex        =   45
            Top             =   120
            Width           =   1695
         End
         Begin VB.TextBox txtDateApproved 
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
            Left            =   5160
            TabIndex        =   44
            Top             =   525
            Width           =   1695
         End
         Begin VB.TextBox txtApprovedBy 
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
            Left            =   1440
            TabIndex        =   43
            Top             =   525
            Width           =   2295
         End
         Begin VB.TextBox txtDateAuthorized 
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
            Left            =   5160
            TabIndex        =   42
            Top             =   930
            Width           =   1695
         End
         Begin VB.TextBox txtAuthorizedBy 
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
            Left            =   1440
            TabIndex        =   41
            Top             =   930
            Width           =   2295
         End
         Begin VB.TextBox txtissuedBy 
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
            Left            =   1440
            TabIndex        =   40
            Top             =   1680
            Width           =   2295
         End
         Begin VB.TextBox txtDateIssued 
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
            Left            =   5160
            TabIndex        =   39
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label40 
            Caption         =   "Date Issued"
            Height          =   255
            Left            =   3840
            TabIndex        =   118
            Top             =   1733
            Width           =   1335
         End
         Begin VB.Label Label39 
            Caption         =   "Issued By"
            Height          =   255
            Left            =   120
            TabIndex        =   117
            Top             =   1733
            Width           =   975
         End
         Begin VB.Label Label38 
            Caption         =   "Issued ?"
            Height          =   255
            Left            =   7080
            TabIndex        =   116
            Top             =   1733
            Width           =   975
         End
         Begin VB.Label Label37 
            Caption         =   "Scheduled ?"
            Height          =   255
            Left            =   7080
            TabIndex        =   112
            Top             =   1380
            Width           =   1095
         End
         Begin VB.Label Label36 
            Caption         =   "Authorized?"
            Height          =   255
            Left            =   7080
            TabIndex        =   111
            Top             =   1005
            Width           =   1335
         End
         Begin VB.Label Label35 
            Caption         =   "Approved? "
            Height          =   255
            Left            =   7080
            TabIndex        =   110
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label34 
            Caption         =   "Prepared ?"
            Height          =   255
            Left            =   7080
            TabIndex        =   109
            Top             =   195
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Prepared By"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   195
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Date Prepared"
            Height          =   255
            Left            =   3840
            TabIndex        =   53
            Top             =   195
            Width           =   1095
         End
         Begin VB.Label Label23 
            Caption         =   "Date Approved "
            Height          =   255
            Left            =   3840
            TabIndex        =   52
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label24 
            Caption         =   "Approved By"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label26 
            Caption         =   "Date Authorized"
            Height          =   255
            Left            =   3840
            TabIndex        =   50
            Top             =   1005
            Width           =   1335
         End
         Begin VB.Label Label27 
            Caption         =   "Authorized By"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   1005
            Width           =   975
         End
         Begin VB.Label Label22 
            Caption         =   "Scheduled By"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   1373
            Width           =   1215
         End
         Begin VB.Label Label25 
            Caption         =   "Date Scheduled"
            Height          =   255
            Left            =   3840
            TabIndex        =   47
            Top             =   1373
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3135
         Left            =   9000
         TabIndex        =   33
         Top             =   1440
         Width           =   2295
         Begin VB.CommandButton cmdCancel 
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
            Height          =   495
            Left            =   120
            TabIndex        =   37
            Top             =   1635
            Width           =   2055
         End
         Begin VB.CommandButton cmdUpdate 
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
            Height          =   495
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   2055
         End
         Begin VB.CommandButton cmdAdd 
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
            Height          =   495
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   2055
         End
         Begin VB.CommandButton cmdSearch 
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
            Height          =   495
            Left            =   120
            TabIndex        =   36
            Top             =   1080
            Width           =   2055
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   35
            Top             =   2130
            Width           =   2055
         End
         Begin VB.CommandButton cmdvoucher 
            Caption         =   "&Print Voucher"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   34
            Top             =   2640
            Width           =   2055
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Payee Details"
         Height          =   1095
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   8775
         Begin VB.TextBox cboTownCode 
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
            Height          =   400
            Left            =   6960
            TabIndex        =   96
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtPayeeAddress 
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
            Height          =   405
            Left            =   1440
            TabIndex        =   28
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtPostalCode 
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
            Height          =   400
            Left            =   4320
            TabIndex        =   27
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtPayeeDetails 
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
            Height          =   405
            Left            =   1440
            TabIndex        =   6
            Top             =   120
            Width           =   7095
         End
         Begin VB.Label Label14 
            Caption         =   " Address"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   675
            Width           =   1095
         End
         Begin VB.Label Label19 
            Caption         =   "Postal Code"
            Height          =   255
            Left            =   3240
            TabIndex        =   31
            Top             =   675
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Payee "
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label20 
            Caption         =   "Town "
            Height          =   255
            Left            =   6120
            TabIndex        =   29
            Top             =   675
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1695
         Left            =   9000
         TabIndex        =   22
         Top             =   4560
         Width           =   2295
         Begin VB.CommandButton cmdAuthorize 
            BackColor       =   &H00808000&
            Caption         =   "&Authorize"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   660
            Width           =   2055
         End
         Begin VB.CommandButton cmdApprove 
            BackColor       =   &H00808000&
            Caption         =   "&Approve"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   120
            Width           =   2055
         End
         Begin VB.CommandButton cmdIssueCheck 
            BackColor       =   &H00808000&
            Caption         =   "&Issue Check"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   1155
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1095
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   11175
         Begin MSComCtl2.DTPicker DTPickerChequeDate 
            Height          =   375
            Left            =   5760
            TabIndex        =   4
            Top             =   120
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Format          =   63373313
            CurrentDate     =   37945
         End
         Begin VB.TextBox txtChequeNo 
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
            Height          =   400
            Left            =   1440
            TabIndex        =   3
            Top             =   100
            Width           =   1575
         End
         Begin VB.TextBox txtChequeDate 
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
            Height          =   400
            Left            =   4320
            TabIndex        =   11
            Top             =   100
            Width           =   1695
         End
         Begin VB.TextBox txtBankAccountNo 
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
            Height          =   400
            Left            =   1440
            TabIndex        =   12
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtBankName 
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
            Height          =   400
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   600
            Width           =   3015
         End
         Begin VB.ComboBox cboBankNo 
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
            Left            =   6960
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   120
            Width           =   4095
         End
         Begin VB.TextBox txtStatus 
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
            Height          =   400
            Left            =   6960
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   600
            Width           =   4095
         End
         Begin VB.Label Label15 
            Caption         =   "Cheque No"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   173
            Width           =   1215
         End
         Begin VB.Label Label21 
            Caption         =   "Cheque Date"
            Height          =   255
            Left            =   3240
            TabIndex        =   20
            Top             =   173
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "AccountNo"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   673
            Width           =   1335
         End
         Begin VB.Label Label16 
            Caption         =   "Bank No"
            Height          =   255
            Left            =   6240
            TabIndex        =   18
            Top             =   173
            Width           =   735
         End
         Begin VB.Label Label18 
            Caption         =   "Status"
            Height          =   255
            Left            =   6240
            TabIndex        =   17
            Top             =   673
            Width           =   615
         End
      End
      Begin VB.Label Label30 
         Caption         =   "TOTAL AMOUNT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   89
         Top             =   4155
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmALISMCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCode As ADODB.Recordset, strcode As String, BSave, bsaveCHEQUE, bcheckREQENTRY, bcheckREQUISITION As Boolean, bExitSub As Boolean
Dim rsLOANVALUE As cLoanApprovers
Dim rsPAYMENTTYPELF As ADODB.Recordset, strPAYMENTTYPElf As String, bLifeAssured As Boolean
Dim rsClaimApproval As clsALISApproval

Private Sub clearENTRY()
On Error GoTo err

        With frmALISMCheque
            .txtAmount.Text = ""
            .txtEntryPreparedBy.Text = ""
            .txtRequisitionDate.Text = ""
            .cboRequisitionNo.Text = ""
            .cboClaimCode.Text = ""
            .cboDocumentNo.Text = ""
            .txtReference.Text = ""
            .txtPayeeDetails.Text = ""
            .txtDatePrepared.Text = ""
            .txtPostalCode.Text = ""
            .cboTownCode.Text = ""
            .cboPaymentType.Text = ""
            .txtPaymentDescription.Text = ""
            .txtPayeeAddress.Text = ""
        End With

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub clearRECORD()
On Error GoTo err

        With frmALISMCheque
            .txtTotalAmount.Text = ""
            .txtpreparedby.Text = ""
            .txtRequisitionDate.Text = ""
            .cboRequisitionNo.Text = ""
            .txtStatus.Text = ""
            .cboClaimCode.Text = ""
            .cboDocumentNo.Text = ""
            .txtReference.Text = ""
            .txtPayeeDetails.Text = ""
            .txtDatePrepared.Text = ""
            .txtPostalCode.Text = ""
            .cboTownCode.Text = ""
            .cboPaymentType.Text = ""
            .txtApprovedBy.Text = ""
            .txtAuthorizedBy.Text = ""
            .txtBankAccountNo.Text = ""
            .txtBankName.Text = ""
            .txtChequeDate.Text = ""
            .txtChequeNo.Text = ""
            .txtDateApproved.Text = ""
            .cboBankNo.Text = ""
            .txtPaymentDescription.Text = ""
            .txtPayeeAddress.Text = ""
            .txtIssuedBy.Text = ""
            .txtDateIssued.Text = ""
        End With

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub enableENTRY()
On Error GoTo err

        With frmALISMCheque
            .txtAmount.Locked = False
            .txtEntryPreparedBy.Locked = True
            .txtRequisitionDate.Locked = True
            .cboRequisitionNo.Locked = False
            .txtEntryStatus.Locked = True
            .cboClaimCode.Locked = True
            .cboDocumentNo.Locked = True
            .txtReference.Locked = True
            .txtEntryDatePrepared.Locked = True
            .cboPaymentType.Locked = True
            .txtPaymentDescription.Locked = True
            .txtPayeeAddress.Locked = True
            .txtEntryChequeDate.Locked = True
            .txtEntryChequeNo.Locked = False
    
        End With

Exit Sub
err:
    ErrorMessage

End Sub
Private Sub disableENTRY()
On Error GoTo err

        With frmALISMCheque
            .txtAmount.Locked = True
            .txtEntryPreparedBy.Locked = True
            .txtRequisitionDate.Locked = True
            .cboRequisitionNo.Locked = True
            .txtEntryStatus.Locked = True
            .cboClaimCode.Locked = True
            .cboDocumentNo.Locked = True
            .txtReference.Locked = True
            .txtEntryDatePrepared.Locked = True
            .cboPaymentType.Locked = True
            .txtPaymentDescription.Locked = True
            .txtPayeeAddress.Locked = True
            .txtEntryChequeDate.Locked = True
            .txtEntryChequeNo.Locked = False
    
        End With

Exit Sub
err:
    ErrorMessage

End Sub

Private Sub EnableCONTROL()

On Error GoTo err

        With frmALISMCheque
            .txtTotalAmount.Locked = False
            .txtpreparedby.Locked = True
            .txtStatus.Locked = True
            .txtReference.Locked = True
            .txtPayeeDetails.Locked = False
            .txtDatePrepared.Locked = True
            .txtPayeeAddress.Locked = True
            .txtApprovedBy.Locked = True
            .txtAuthorizedBy.Locked = True
            .txtBankAccountNo.Locked = True
            .txtBankName.Locked = True
            .txtChequeDate.Locked = True
            .txtChequeNo.Locked = False
            .txtDateApproved.Locked = True
            .cboBankNo.Locked = False
            .txtDateIssued.Locked = True
            .txtIssuedBy.Locked = True
            .DTPickerChequeDate.Enabled = True
            .cboBankNo.Enabled = True
        End With

Exit Sub
err:
    ErrorMessage

End Sub

Private Sub disableCONTROL()
On Error GoTo err

        With frmALISMCheque
            .txtTotalAmount.Locked = True
            .txtpreparedby.Locked = True
            .txtStatus.Locked = True
            .txtReference.Locked = True
            .txtPayeeDetails.Locked = True
            .txtDatePrepared.Locked = True
            .txtPayeeAddress.Locked = True
            .txtApprovedBy.Locked = True
            .txtAuthorizedBy.Locked = True
            .txtBankAccountNo.Locked = True
            .txtBankName.Locked = True
            .txtChequeDate.Locked = True
            .txtChequeNo.Locked = True
            .txtDateApproved.Locked = True
            .cboBankNo.Enabled = False
            .txtDateIssued.Locked = True
            .txtIssuedBy.Locked = True
            .DTPickerChequeDate.Enabled = False
        End With

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub loadRecord()
'On Error GoTo err

        With RsCode
            frmALISMCheque.txtTotalAmount.Text = !Amount & ""
            frmALISMCheque.txtpreparedby.Text = !Preparedby
            frmALISMCheque.txtDatePrepared.Text = !dateprepared
            
            If IsNull(!Status) = True Or !Status <= "" Then
                        frmALISMCheque.txtStatus.Text = "CHQ-PREPARED"
                Else: frmALISMCheque.txtStatus.Text = !Status
            End If

            frmALISMCheque.txtApprovedBy = !ApprovedBy & ""
            frmALISMCheque.txtAuthorizedBy = !AuthorizedBy & ""
            frmALISMCheque.txtChequeDate = !ChequeDate
            frmALISMCheque.txtChequeNo = !ChequeNo
            frmALISMCheque.txtDateApproved = !DateApproved & ""
            frmALISMCheque.cboBankNo = !BankNo
            frmALISMCheque.txtPayeeDetails = !PayeeDetails & ""
            frmALISMCheque.txtDateAuthorized = !DateAuthorized & ""
            frmALISMCheque.txtIssuedBy = !IssuedBy & ""
            frmALISMCheque.txtDateIssued = !DateIssued & ""
        End With
    
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub loadBankName()

On Error GoTo err
            
          Dim rsbank As ADODB.Recordset, strBANK As String
          Set rsbank = New ADODB.Recordset
          
          strBANK = "SELECT * FROM ALISPbankAccount where bankno = '" & frmALISMCheque.cboBankNo.Text & "' ;"
          rsbank.Open strBANK, cnALIS, adOpenKeyset, adLockOptimistic

            With rsbank
                    If .EOF Or .BOF Then
                        MsgBox "The Bank Selected is invalid", vbOKOnly
                        Exit Sub
                    End If
                
                    frmALISMCheque.txtBankName = !details
                    frmALISMCheque.txtBankAccountNo = !AccountNo

            End With
            
        rsbank.Close
        strBANK = ""
Exit Sub

err:
    UpdateErrorMessage
End Sub


Private Sub loadCHEQUEENTRIES()
On Error GoTo err
            
          Dim rsREQNO As ADODB.Recordset, strREQNO As String
          Set rsREQNO = New ADODB.Recordset
          
          strREQNO = "SELECT * FROM ALISMPaymentRequisition where requisitionno = '" & frmALISMCheque.cboRequisitionNo.Text & "' ;"
          rsREQNO.Open strREQNO, cnALIS, adOpenKeyset, adLockOptimistic

            With rsREQNO
                    If .EOF Or .BOF Then
                        MsgBox "The Requisition Selected is invalid", vbOKOnly
                        Exit Sub
                    End If
                    
                    frmALISMCheque.txtRequisitionDate.Text = !RequisitionDate
                    
                    frmALISMCheque.cboPaymentType = !PaymentType
                    frmALISMCheque.cboClaimCode.Text = !claimcode
                    frmALISMCheque.cboDocumentNo.Text = !DocumentNo
                    frmALISMCheque.txtReference.Text = !reference
                    frmALISMCheque.txtPayeeDetails.Text = !PayeeDetails
                    frmALISMCheque.txtPostalCode.Text = !PostalCode
                    frmALISMCheque.cboTownCode.Text = !TownCode
                    frmALISMCheque.txtPayeeAddress = !Address
                    frmALISMCheque.cboPaymentType.Text = !PaymentType
                    
            End With
            
            
            loadPaymentDescription
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub loadRECORD1()
'On Error GoTo err
            
          Dim rsREQNO As ADODB.Recordset, strREQNO As String
          Set rsREQNO = New ADODB.Recordset
          
          strREQNO = "SELECT * FROM ALISMPaymentRequisition where requisitionno = '" & frmALISMCheque.cboRequisitionNo.Text & "' ;"
          rsREQNO.Open strREQNO, cnALIS, adOpenKeyset, adLockOptimistic

            With rsREQNO
                    If .EOF Or .BOF Then
                        MsgBox "The Requisition Selected is invalid", vbOKOnly
                        Exit Sub
                    End If
                    
                    frmALISMCheque.txtRequisitionDate.Text = !RequisitionDate
                    
                    frmALISMCheque.cboPaymentType = !PaymentType
                    frmALISMCheque.cboClaimCode.Text = !claimcode
                    frmALISMCheque.cboDocumentNo.Text = !DocumentNo
                    frmALISMCheque.txtReference.Text = !reference
                    frmALISMCheque.txtPayeeDetails.Text = !PayeeDetails
                    frmALISMCheque.txtPostalCode.Text = !PostalCode
                    frmALISMCheque.cboTownCode.Text = !TownCode
                    frmALISMCheque.txtPayeeAddress = !Address
                    frmALISMCheque.cboPaymentType.Text = !PaymentType
                    
            End With
            
            
            loadPaymentDescription
Exit Sub

err:
    ErrorMessage
End Sub


Private Sub SaveRECORD()
On Error GoTo err
    
    Dim rsSAVE As ADODB.Recordset, strSAVE As String
    Set rsSAVE = New ADODB.Recordset
    
    strSAVE = "SELECT * from ALISMCheque where chequeNo = '" & frmALISMCheque.txtChequeNo.Text & "';"
    rsSAVE.Open strSAVE, cnALIS, adOpenKeyset, adLockOptimistic

        With rsSAVE
            .AddNew
                !ChequeNo = frmALISMCheque.txtChequeNo.Text
                !Preparedby = frmALISMCheque.txtpreparedby.Text
                !Status = frmALISMCheque.txtStatus.Text
                !dateprepared = frmALISMCheque.txtDatePrepared.Text
                !BankNo = frmALISMCheque.cboBankNo.Text
                !ChequeDate = frmALISMCheque.txtChequeDate.Text
                !PayeeDetails = frmALISMCheque.txtPayeeDetails
                !Prepared = "Y"
                
            .Update
            .Requery
        End With

rsSAVE.Close
strSAVE = ""

Exit Sub

err:
If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
    rsSAVE.CancelUpdate
    rsSAVE.Requery
End If
    UpdateErrorMessage
End Sub
Private Sub loadMGRID()
    loadCHKENTRIESGRID
    loadApprovedGRID
    loadAuthorizedGRID
    loadIssuedGRID
    loadPendingGRID
End Sub

Private Sub loadCHKENTRIESGRID()

On Error GoTo err

    Dim rsMGRID As ADODB.Recordset, strMGRID As String
    Set rsMGRID = New Recordset
    
    strMGRID = "SELECT * from ALISMChequeEntry where chequeNo = '" & frmALISMCheque.txtChequeNo.Text & "' ;"
    rsMGRID.Open strMGRID, cnALIS, adOpenKeyset, adLockOptimistic
    Set frmALISMCheque.CheckGRID.DataSource = rsMGRID

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub loadPendingGRID()
On Error GoTo err

    Dim rsMGRID As ADODB.Recordset, strMGRID As String
    Set rsMGRID = New Recordset
    
    strMGRID = "SELECT * from ALISMCheque where Status = 'CHK APPROVAL' ;"
    rsMGRID.Open strMGRID, cnALIS, adOpenKeyset, adLockOptimistic
    Set frmALISMCheque.PendingGRID.DataSource = rsMGRID

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub loadApprovedGRID()
On Error GoTo err

    Dim rsMGRID As ADODB.Recordset, strMGRID As String
    Set rsMGRID = New Recordset
    
    strMGRID = "SELECT * from ALISMCheque where Status = 'CHK APPROVAL' ;"
    rsMGRID.Open strMGRID, cnALIS, adOpenKeyset, adLockOptimistic
    Set frmALISMCheque.ApprovedGrid.DataSource = rsMGRID

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub loadIssuedGRID()
On Error GoTo err

    Dim rsMGRID As ADODB.Recordset, strMGRID As String
    Set rsMGRID = New Recordset
    
    strMGRID = "SELECT * from ALISMCheque where Status = 'CHECK ISSUANCE' ;"
    rsMGRID.Open strMGRID, cnALIS, adOpenKeyset, adLockOptimistic
    Set frmALISMCheque.IssuedGrid.DataSource = rsMGRID

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub loadAuthorizedGRID()
On Error GoTo err

    Dim rsMGRID As ADODB.Recordset, strMGRID As String
    Set rsMGRID = New Recordset
    
    strMGRID = "SELECT * from ALISMCheque where Status = 'CHK AUTHORIZATION' ;"
    rsMGRID.Open strMGRID, cnALIS, adOpenKeyset, adLockOptimistic
    Set frmALISMCheque.AuthorizedGrid.DataSource = rsMGRID

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub loadGRID()
On Error GoTo err

    Dim rsGRID As ADODB.Recordset, StrGRID As String
    Set rsGRID = New Recordset
    
    StrGRID = "SELECT * from ALISMChequeEntry where chequeNo = '" & frmALISMCheque.txtChequeNo.Text & "' ;"
    rsGRID.Open StrGRID, cnALIS, adOpenKeyset, adLockOptimistic
    Set frmALISMCheque.ChequeGRID.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub updateTOTAL()
'On Error GoTo err
    
    Dim rsTOTAL, rsUPDATE As ADODB.Recordset, strTOTAL, strUPDATE As String
    Set rsTOTAL = New Recordset
    
    strTOTAL = "SELECT sum(ALISMChequeEntry.amount) as totals from ALISMChequeEntry where chequeNo = '" & frmALISMCheque.txtEntryChequeNo.Text & "' ;"
    rsTOTAL.Open strTOTAL, cnALIS, adOpenKeyset, adLockOptimistic
                
                With rsTOTAL
                            If .EOF And .BOF Then frmALISMCheque.txtTotalAmount = 0
                        
                                          
                            If IsNull(!Totals) = True Then
                                    frmALISMCheque.txtEntryTotalAmount.Text = 0
                            Else
                                    frmALISMCheque.txtTotalAmount = !Totals
                                    frmALISMCheque.txtEntryTotalAmount = !Totals
                            End If

                End With
    
rsTOTAL.Close
strTOTAL = ""

    Set rsUPDATE = New ADODB.Recordset
    
    strUPDATE = "SELECT * from ALISMCheque where chequeNo = '" & frmALISMCheque.txtEntryChequeNo.Text & "' ;"
    rsUPDATE.Open strUPDATE, cnALIS, adOpenKeyset, adLockOptimistic
                
                With rsUPDATE
                            If .EOF And .BOF Then Exit Sub
                            
                            !Amount = frmALISMCheque.txtEntryTotalAmount
                            .Update
                            .Requery
                    
                End With
    
rsUPDATE.Close
strUPDATE = ""

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub SaveChequeENTRIES()
On Error GoTo err
    Dim rsSAVE As ADODB.Recordset, strSAVE As String
    Set rsSAVE = New Recordset
    
    strSAVE = "SELECT * from ALISMChequeEntry where chequeNo = '" & frmALISMCheque.txtChequeNo.Text & "' and requisitionNo = '" & frmALISMCheque.cboRequisitionNo.Text & "';"
    rsSAVE.Open strSAVE, cnALIS, adOpenKeyset, adLockOptimistic

        With rsSAVE
            .AddNew
            !Amount = frmALISMCheque.txtAmount
            !Preparedby = frmALISMCheque.txtpreparedby
            !RequisitionNo = frmALISMCheque.cboRequisitionNo
            !Status = frmALISMCheque.txtStatus
            !dateprepared = Date
            !ChequeNo = frmALISMCheque.txtChequeNo
            .Update
            .Requery
        End With

strSAVE = ""
rsSAVE.Close

Exit Sub

err:
        If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
            rsSAVE.CancelUpdate
            rsSAVE.Requery
        Else
            UpdateErrorMessage
        End If
End Sub
 
Private Sub updateREQUISITION()
          
        Dim rsREQ As ADODB.Recordset
        Set rsREQ = New ADODB.Recordset
        
        rsREQ.Open "SELECT * FROM ALISMPaymentRequisition WHERE RequisitionNo = '" & frmALISMCheque.cboRequisitionNo.Text & "' ", cnALIS, adOpenKeyset, adLockOptimistic
  
        With rsREQ
            If .BOF Or .EOF = True Then Exit Sub
            !Status = frmALISMCheque.txtStatus.Text
            !ChequeNo = frmALISMCheque.txtEntryChequeNo.Text
            !ChequeDate = frmALISMCheque.txtChequeDate.Text
            !ChequePrepared = "Y"
            .Update
            .Requery
        End With

rsREQ.Close

Exit Sub

err:
    
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then
        rsREQ.CancelUpdate
        rsREQ.Requery
    Else
        UpdateErrorMessage
    End If
    
End Sub


Private Sub loadPaymentDescription()
On Error GoTo err
        
        Dim rsPAYMENTTYPE  As ADODB.Recordset
        
        Set rsPAYMENTTYPE = New ADODB.Recordset
        
        rsPAYMENTTYPE.Open "SELECT * FROM ALISPPaymentType WHERE PaymentType = '" & frmALISMCheque.cboPaymentType.Text & "' ", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsPAYMENTTYPE
                If .EOF And .BOF Then Exit Sub
                frmALISMCheque.txtPaymentDescription = !Description
        End With
        
rsPAYMENTTYPE.Close

Exit Sub

err:
        ErrorMessage
End Sub

Private Sub cbobankNo_GotFocus()
On Error GoTo err

          Dim rsBANKGF As ADODB.Recordset, strBANKGF As String
          Set rsBANKGF = New ADODB.Recordset
          
          strBANKGF = "SELECT * FROM ALISPbankAccount;"
          rsBANKGF.Open strBANKGF, cnALIS, adOpenKeyset, adLockOptimistic

          frmALISMCheque.cboBankNo.Clear

          With rsBANKGF
                  If .EOF Or .BOF Then Exit Sub
                  Do Until .EOF
                      Me.cboBankNo.AddItem !details & ""
                      .MoveNext
                  Loop
          End With
        
          GlobalClaimNo = frmALISMCheque.cboDocumentNo
 
          rsBANKGF.Close
          strBANKGF = ""
         
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cboBankNo_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub

Private Sub cbobankNo_LostFocus()
    loadBANK
    Set rsClaimApproval = New clsALISApproval
'    rsClaimApproval.loadApprovalDetails
    rsClaimApproval.switchCOMMANDBUTTONS
    Set rsClaimApproval = Nothing
End Sub

Private Sub loadBANK()
On Error GoTo err
        
        Dim rsBANKLF As ADODB.Recordset
        
        Set rsBANKLF = New ADODB.Recordset
        
        rsBANKLF.Open "SELECT * FROM ALISPBankAccount WHERE Details = '" & cboBankNo.Text & "' ", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsBANKLF
                If .EOF And .BOF Then Exit Sub
                frmALISMCheque.cboBankNo.Text = !BankNo
                frmALISMCheque.txtBankName.Text = !details
                frmALISMCheque.txtBankAccountNo = !AccountNo
                frmALISMCheque.txtChequeDate = Date
                frmALISMCheque.txtpreparedby = CurrentUserName
                frmALISMCheque.txtDatePrepared = Date
        End With
        
        
rsBANKLF.Close

Exit Sub

err:
        ErrorMessage

End Sub
Private Sub loadBANKDETAILS()
On Error GoTo err
        
        Dim rsBANKLF As ADODB.Recordset
        
        Set rsBANKLF = New ADODB.Recordset
        
        rsBANKLF.Open "SELECT * FROM ALISPBankAccount WHERE BankNo = '" & frmALISMCheque.cboBankNo.Text & "' ", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsBANKLF
                If .EOF And .BOF Then Exit Sub
                frmALISMCheque.cboBankNo.Text = !BankNo
                frmALISMCheque.txtBankName.Text = !details
                frmALISMCheque.txtBankAccountNo = !AccountNo
                frmALISMCheque.txtChequeDate = Date
                frmALISMCheque.txtpreparedby = CurrentUserName
                frmALISMCheque.txtDatePrepared = Date
        End With
        
        
rsBANKLF.Close

Exit Sub

err:
        ErrorMessage

End Sub

Private Sub verifyRequisition()
On Error GoTo err

          Dim rsVERIFY As ADODB.Recordset, strVERIFY As String
          Set rsVERIFY = New ADODB.Recordset
          
          strVERIFY = "SELECT * FROM ALISMChequeEntry where requisitionno = '" & frmALISMCheque.cboRequisitionNo & "';"
          rsVERIFY.Open strVERIFY, cnALIS, adOpenKeyset, adLockOptimistic

          With rsVERIFY
                If .BOF Or .EOF Then Exit Sub
                
                MsgBox "This requisition Has already been processed", vbOKOnly
                        clearENTRY
                        frmALISMCheque.cboRequisitionNo.SetFocus
                Exit Sub
          
          End With

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub checkREQUISITION()
On Error GoTo err
        
          Dim rsREQNO As ADODB.Recordset, strREQNO As String
          Set rsREQNO = New ADODB.Recordset
          
          strREQNO = "SELECT * FROM ALISMPaymentRequisition where ChequePrepared = 'N' ;"
          rsREQNO.Open strREQNO, cnALIS, adOpenKeyset, adLockOptimistic

          With rsREQNO
                  If .EOF Or .BOF Then
                        MsgBox "There are no Requisitions to be Processed, All Payment have been Made", vbOKOnly
                        Exit Sub
                  End If
                    
                  bcheckREQUISITION = True
                  bcheckREQENTRY = True
          End With
            
          rsREQNO.Close
          strREQNO = ""
         
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cboRequisitionNo_GotFocus()
On Error GoTo err
        
          If frmALISMCheque.txtEntryChequeNo.Text <= "" Then Exit Sub
          
          Dim rsREQNOGF As ADODB.Recordset, strREQNOGF As String
          Set rsREQNOGF = New ADODB.Recordset
          
          strREQNOGF = "SELECT * FROM ALISMPaymentRequisition where (ChequePrepared = 'N' or chequePrepared is Null) AND AUTHORIZED = 'Y' ;"
          rsREQNOGF.Open strREQNOGF, cnALIS, adOpenKeyset, adLockOptimistic

          frmALISMCheque.cboRequisitionNo.Clear

          With rsREQNOGF
                  If .EOF Or .BOF Then Exit Sub
                  Do Until .EOF
                      frmALISMCheque.cboRequisitionNo.AddItem !reference
                      .MoveNext
                  Loop
          End With
            
          rsREQNOGF.Close
          strREQNOGF = ""
         
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cboRequisitionNo_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub

Private Sub cboRequisitionNo_LostFocus()
'On Error GoTo err
        
        Dim rsREQNOLF As ADODB.Recordset
        
        Set rsREQNOLF = New ADODB.Recordset
        
        rsREQNOLF.Open "SELECT * FROM ALISMPaymentRequisition WHERE Reference = '" & cboRequisitionNo.Text & "' ", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsREQNOLF
                If .EOF And .BOF Then Exit Sub
                frmALISMCheque.cboRequisitionNo.Text = !RequisitionNo
                frmALISMCheque.txtAmount = !Amount
                frmALISMCheque.txtStatus = "CHQ-PREPARED"
        End With
        
        loadCHEQUEENTRIES
        'verifyRequisition
        
'rsREQNOLF.Close
'strREQNOlf = ""

Exit Sub

err:
        ErrorMessage

End Sub

Private Sub cmdAddNewEntry_Click()
        If frmALISMCheque.txtEntryChequeNo.Text <= "" Then
            MsgBox "MUST LOAD The Cheque Before Loading the Entries", vbOKOnly
            Exit Sub
        End If
 
        bcheckREQENTRY = False
        checkREQUISITION
        If bcheckREQENTRY = True Then
                enableENTRY
                clearENTRY
                DisableCBEntry
                Exit Sub
        End If
End Sub

Private Sub cmdApprove_Click()
    bApproveCheque = True
    bAuthorizeCheque = False
    Set rsClaimApproval = New clsALISApproval
    rsClaimApproval.checkAPPROVEDDISCHARGE
    If bApproveCheque = False Then Exit Sub
    rsClaimApproval.approveCLAIM
    rsClaimApproval.loadAPPROVALDETAILS
    rsClaimApproval.switchCOMMANDBUTTONS
    Set rsClaimApproval = Nothing
    bapproveREQUISITION = False
End Sub

Private Sub checkAPPROVALSTATUS()
On Error GoTo err
                
                Dim rsAUTHORIZATION As ADODB.Recordset, strAuthorization As String
                Set rsAUTHORIZATION = New Recordset
                    
                strAuthorization = "SELECT * FROM ALISPLoanOperationType  WHERE PaymentApproval = '1' ;"
                rsAUTHORIZATION.Open strAuthorization, cnALIS, adOpenKeyset, adLockOptimistic

                With rsAUTHORIZATION
                        If .EOF Or .BOF Then Exit Sub
                        
                        Dim rsAPPROVED As ADODB.Recordset, strAPPROVED As String
                        Set rsAPPROVED = New Recordset
                            
                        strAPPROVED = "SELECT * FROM ALISMLoanOperation  WHERE ApplicationNo = '" & frmALISMCheque.cboRequisitionNo & "' and operationType = '" & !OperationType & "' ;"
                        rsAPPROVED.Open strAPPROVED, cnALIS, adOpenKeyset, adLockOptimistic
        
                        With rsAPPROVED
                                If .BOF Or .EOF Then
                                        MsgBox "Authorization can only take place immediately after Approval", vbOKOnly
                                                bExitSub = True
                                                Exit Sub
                                ElseIf !Accept = "N" Then
                                        MsgBox "Cannot Approve payment That has been Rejected", vbOKOnly
                                                bExitSub = True
                                                Exit Sub
                                End If
                        
                        End With

            End With
                    '/ Authorization
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub checkAUTHORIZATIONSTATUS()
On Error GoTo err
                
                Dim rsAUTHORIZATION As ADODB.Recordset, strAuthorization As String
                Set rsAUTHORIZATION = New Recordset
                    
                strAuthorization = "SELECT * FROM ALISPLoanOperationType  WHERE ChequeAuthorization = '1' ;"
                rsAUTHORIZATION.Open strAuthorization, cnALIS, adOpenKeyset, adLockOptimistic

                With rsAUTHORIZATION
                        If .EOF Or .BOF Then Exit Sub
                        
                        Dim rsAPPROVED As ADODB.Recordset, strAPPROVED As String
                        Set rsAPPROVED = New Recordset
                            
                        strAPPROVED = "SELECT * FROM ALISMLoanOperation  WHERE ApplicationNo = '" & frmALISMCheque.cboRequisitionNo & "' and operationType = '" & !OperationType & "' ;"
                        rsAPPROVED.Open strAPPROVED, cnALIS, adOpenKeyset, adLockOptimistic
        
                        With rsAPPROVED
                                If .BOF Or .EOF Then
                                        Exit Sub
                                Else
                                        MsgBox "This Record has Already been Authorized", vbOKOnly
                                                bExitSub = True
                                End If
                        
                        End With

                        Dim rsCHKAPPROVAL As ADODB.Recordset, strCHKAPPROVAL As String
                        Set rsCHKAPPROVAL = New Recordset
                            
                        strCHKAPPROVAL = "SELECT * FROM ALISPLoanOperationType  WHERE ChequeApproval = '1' ;"
                        rsCHKAPPROVAL.Open strCHKAPPROVAL, cnALIS, adOpenKeyset, adLockOptimistic
                        
                        With rsCHKAPPROVAL
                                If .BOF Or .EOF Then Exit Sub
                                
                                    Dim rsAPPROVE As ADODB.Recordset, strAPPROVE As String
                                    Set rsAPPROVE = New Recordset
                                        
                                    strAPPROVE = "SELECT * FROM ALISMLoanOperation  WHERE ApplicationNo = '" & frmALISMCheque.txtChequeNo & "' and operationType = '" & !OperationType & "' ;"
                                    rsAPPROVE.Open strAPPROVE, cnALIS, adOpenKeyset, adLockOptimistic
                    
                                    With rsAPPROVE
                                            If .BOF Or .EOF Then
                                                    Exit Sub
                                            Else
                                                    MsgBox "This Record has Already been Approved", vbOKOnly
                                                            bExitSub = True
                                            End If
                                    
                                    End With
                        
                        End With

            End With
                    '/ Authorization
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub checkISSUANCESTATUS()
On Error GoTo err
                
                Dim rsAUTHORIZATION As ADODB.Recordset, strAuthorization As String
                Set rsAUTHORIZATION = New Recordset
                    
                strAuthorization = "SELECT * FROM ALISPLoanOperationType  WHERE ChequeAuthorization = '1' ;"
                rsAUTHORIZATION.Open strAuthorization, cnALIS, adOpenKeyset, adLockOptimistic

                With rsAUTHORIZATION
                        If .EOF Or .BOF Then Exit Sub
                        
                        Dim rsAPPROVED As ADODB.Recordset, strAPPROVED As String
                        Set rsAPPROVED = New Recordset
                            
                        strAPPROVED = "SELECT * FROM ALISMLoanOperation  WHERE ApplicationNo = '" & frmALISMCheque.cboRequisitionNo & "' and operationType = '" & !OperationType & "' ;"
                        rsAPPROVED.Open strAPPROVED, cnALIS, adOpenKeyset, adLockOptimistic
        
                        With rsAPPROVED
                                If .BOF Or .EOF Then
                                        Exit Sub
                                Else
                                        MsgBox "This Record has Already been Authorized", vbOKOnly
                                                bExitSub = True
                                End If
                        
                        End With

                        Dim rsCHKAPPROVAL As ADODB.Recordset, strCHKAPPROVAL As String
                        Set rsCHKAPPROVAL = New Recordset
                            
                        strCHKAPPROVAL = "SELECT * FROM ALISPLoanOperationType  WHERE ChequeApproval = '1' ;"
                        rsCHKAPPROVAL.Open strCHKAPPROVAL, cnALIS, adOpenKeyset, adLockOptimistic
                        
                        With rsCHKAPPROVAL
                                If .BOF Or .EOF Then Exit Sub
                                
                                    Dim rsAPPROVE As ADODB.Recordset, strAPPROVE As String
                                    Set rsAPPROVE = New Recordset
                                        
                                    strAPPROVE = "SELECT * FROM ALISMLoanOperation  WHERE ApplicationNo = '" & frmALISMCheque.txtChequeNo & "' and operationType = '" & !OperationType & "' ;"
                                    rsAPPROVE.Open strAPPROVE, cnALIS, adOpenKeyset, adLockOptimistic
                    
                                    With rsAPPROVE
                                            If .BOF Or .EOF Then
                                                    Exit Sub
                                            Else
                                                    MsgBox "This Record has Already been Approved", vbOKOnly
                                                            bExitSub = True
                                            End If
                                    
                                    End With
                        
                        End With

            End With
                    '/ Authorization
Exit Sub

err:
    ErrorMessage
End Sub


Private Sub cmdAuthorize_Click()
    bAuthorizeCheque = True
    Set rsClaimApproval = New clsALISApproval
    rsClaimApproval.checkAPPROVEDDISCHARGE
    If bAuthorizeCheque = False Then Exit Sub
    rsClaimApproval.approveCLAIM
    rsClaimApproval.loadAPPROVALDETAILS
    rsClaimApproval.switchCOMMANDBUTTONS
    Set rsClaimApproval = Nothing
    bAuthorizeCheque = False
End Sub

Private Sub cmdCancelEntry_Click()
On Error GoTo Myerr
        enableCBEntry
        clearENTRY
        disableENTRY
        Exit Sub

Myerr:
        ErrorMessage
End Sub




Private Sub cmdFirstEntry_Click(Index As Integer)
    If frmALISMCheque.txtEntryChequeNo.Text <= "" Then Exit Sub
    browseCHEQUEENTRY (Index)
    loadCHECKENTRIES
    loadGRID
    loadMGRID
End Sub

Private Sub cmdIssueCheck_Click()
    With frmALISMCheque
            If .txtAuthorized.Text = "Y" And .txtIssued.Text <> "Y" Then
                processISSUANCE
'                .txtIssued.Text = !Issued & ""
'                .txtIssuedBy.Text = !issuedBy & ""
'                .txtDateIssued.Text = !DateIssued & ""
            End If
    End With
End Sub

Private Sub processISSUANCE()
On Error GoTo err
        GlobalApplicationNo = Screen.ActiveForm.txtChequeNo.Text
        GlobalClaimNo = Screen.ActiveForm.txtChequeNo
        
        Dim rsAUTHORIZATION As ADODB.Recordset, strAPPROVED As String
        Set rsAUTHORIZATION = New Recordset
            
        strAPPROVED = "SELECT * FROM ALISPLoanOperationType  WHERE ChequeIssuance = '1' ;"
        rsAUTHORIZATION.Open strAPPROVED, cnALIS, adOpenKeyset, adLockOptimistic

        With rsAUTHORIZATION
                If .EOF Or .BOF Then Exit Sub
                
                GlobalOperationType = rsAUTHORIZATION!OperationType
                GlobalOperationDescription = rsAUTHORIZATION!Description
        End With
        
        '/Check Whether this payment has been approved
        

        Set rsLOANVALUE = New cLoanApprovers
        Call rsLOANVALUE.operationAPPROVED
        Set rsLOANVALUE = Nothing
        
        GlobalOperationType = ""
        GlobalOperationDescription = ""
        
        GlobalApplicationNo = ""
Exit Sub

err:
    ErrorMessage
End Sub


Private Sub cmdPrint_Click()
        If frmALISMCheque.txtAuthorized.Text = "Y" Then
                If frmALISMCheque.txtChequeNo.Text <= "" Then
                        MsgBox "Cannot Execute this Function Before Loading the Record", vbOKOnly
                        Exit Sub
                End If
                payvoucher = True
                Load frmclaimprocessing
                frmclaimprocessing.Show 1, Me
        End If
End Sub
Private Sub cmdEntryUpdate_Click()
        If frmALISMCheque.txtEntryChequeNo.Text <= "" Then
            MsgBox "MUST LOAD The Cheque Before Loading the Entries", vbOKOnly
            Exit Sub
        End If
 
        validateENTRY
        If bsaveCHEQUE = True Then
                    SaveChequeENTRIES
                    updateREQUISITION
                    updateTOTAL
                    loadGRID
                    loadMGRID
                    'saveCLAIM
                    bsaveCHEQUE = False
                    enableCBEntry
                    disableENTRY
                    cmdAddNewEntry.SetFocus
        End If

End Sub




Private Sub cmdPrintEntry_Click()
    If frmALISMCheque.txtEntryChequeNo.Text <= "" Then
        MsgBox "MUST LOAD The Cheque Before Loading the Entries", vbOKOnly
        Exit Sub
    End If
 
End Sub

Private Sub cmdSearchEntry_Click()
    If frmALISMCheque.txtEntryChequeNo.Text <= "" Then
        MsgBox "MUST LOAD The Cheque Before Loading the Entries", vbOKOnly
        Exit Sub
    End If
        searchCHECKENTRIES
        loadPaymentDescription
        loadMGRID
        loadGRID
End Sub

Private Sub cmdvoucher_Click()
On Error GoTo err
        If frmALISMCheque.txtChequeNo.Text <= "" Then
                MsgBox "Cannot Execute this Function Before Loading the Record", vbOKOnly
                Exit Sub
        End If

        Load frmDischargeVoucher
        frmDischargeVoucher.Show 1, ALISENTPMAIN

Exit Sub
err:
ErrorMessage

End Sub



Private Sub DTPickerChequeDate_Change()
On Error GoTo err
        
        frmALISMCheque.txtChequeDate.Text = Date
        frmALISMCheque.txtChequeDate.Text = frmALISMCheque.DTPickerChequeDate.Value

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub Form_Activate()
    disableENTRY
    enableCBEntry
    disableCONTROL
    enableCButtons
    
    frmALISMCheque.SSTabCheque.Tab = 0

End Sub

Private Sub Form_Load()
    OpenConnection
End Sub

Private Sub cmdAdd_Click()

        bcheckREQUISITION = False
        checkREQUISITION
        If bcheckREQUISITION = True Then
                bcheckREQUISITION = False
                EnableCONTROL
                clearRECORD
                DisableCButtons
                initializeRECORD
                Exit Sub
        End If
End Sub

Private Sub initializeRECORD()

        With frmALISMCheque
                  .SSTabCheque.TabEnabled(1) = False
                  .SSTabCheque.TabEnabled(2) = False
                  .txtChequeDate.Text = Date
                  .DTPickerChequeDate.Value = Date
                  .txtStatus.Text = "CHK-PREPARED"
        End With

End Sub

Private Sub cmdCancel_Click()
On Error GoTo Myerr
        enableCButtons
        clearRECORD
        disableCONTROL
        Exit Sub

Myerr:
        ErrorMessage
End Sub

Private Sub cmdDelete_Click()
On Error GoTo Myerr

If frmALISMCheque.cboRequisitionNo.Text = "" Then
            MsgBox "There is no current record to delete", vbInformation, "Delete Information"
'ElseIf frmALISMCheque.txtRequisitionDate = "" Then
 '           MsgBox "There is no current record", vbInformation, "Delete Information"
Else
        If MsgBox("Are you sure you want to completely delete the current record?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            With RsCode
                
                If .EOF And .BOF Then Exit Sub
                .Delete
                .Requery
                clearRECORD
                                
                End With
    End If
        '/* End if Msg Box
        
End If
        '/* If txt = ""
        
Exit Sub

Myerr:
    ErrorMessage


End Sub

Private Sub browseCHEQUEENTRY(Index As Integer)
On Error GoTo err
        Dim rsCHK As ADODB.Recordset
        
        Set rsCHK = New ADODB.Recordset
        rsCHK.Open "SELECT * FROM ALISMChequeEntry WHERE ALISMChequeEntry.ChequeNo = '" & frmALISMCheque.txtEntryChequeNo.Text & "'", cnALIS, adOpenKeyset, adLockOptimistic

        cmdEntryUpdate.Enabled = False

        With rsCHK
        If .EOF And .BOF Then Exit Sub
    
                    Select Case Index
                                Case 0
                                    .MoveFirst
                                Case 1
                                    .MovePrevious
                                    If .BOF Then .MoveFirst
                                Case 2
                                    .MoveNext
                                    If .EOF Then .MoveLast
                                Case 3
                                    .MoveLast
                    End Select
        End With

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub browseRECORD(Index As Integer)
On Error GoTo err

        cmdUpdate.Enabled = False

        With RsCode
        If .EOF And .BOF Then Exit Sub
    
                    Select Case Index
                                Case 0
                                    .MoveFirst
                                Case 1
                                    .MovePrevious
                                    If .BOF Then .MoveFirst
                                Case 2
                                    .MoveNext
                                    If .EOF Then .MoveLast
                                Case 3
                                    .MoveLast
                    End Select
        End With

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub cmdFirstCode_Click(Index As Integer)
        browseRECORD (Index)
        loadRecord
        loadBankName
        'loadRequisition
        loadPaymentDescription
        updateCHEQUEENTRIES
        loadCHECKENTRIES
        loadGRID
        loadMGRID

Exit Sub

Myerr:
    ErrorMessage

End Sub
Private Sub searchCHECKENTRIES()
On Error GoTo err

        Dim strQRE As Variant
        Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset
        strQRE = InputBox("Enter The Requisition No to search.", "Search Value")
        
        rsFind.Open "SELECT ALISMPaymentRequisition.*, ALISMChequeEntry.* FROM ALISMChequeEntry,ALISMPaymentRequisition  WHERE ALISMChequeEntry.ChequeNo = '" & frmALISMCheque.txtEntryChequeNo.Text & "' and ALISMPaymentRequisition.requisitionNo = '" & strQRE & "' and ALISMPaymentRequisition.RequisitionNo = ALISMChequeEntry.RequisitionNo; ", cnALIS, adOpenKeyset, adLockOptimistic

        With rsFind
                If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                Else
                            frmALISMCheque.cboRequisitionNo = !RequisitionNo
                            frmALISMCheque.txtRequisitionDate = !RequisitionDate
                            frmALISMCheque.cboDocumentNo = !DocumentNo
                            frmALISMCheque.cboPaymentType = !PaymentType
                            frmALISMCheque.cboClaimCode = !claimcode
                            frmALISMCheque.txtAmount = !Amount
                            frmALISMCheque.txtReference = !reference
                            Edit = True
                End If

            End With

        Exit Sub

err:
            ErrorMessage

End Sub

Private Sub loadCHECKENTRIES()
On Error GoTo err
        Dim rsLOAD As ADODB.Recordset

        Set rsLOAD = New ADODB.Recordset
        
        rsLOAD.Open "SELECT ALISMchequeEntry.*, ALISMPaymentRequisition.* FROM ALISMChequeEntry, ALISMPaymentRequisition WHERE ALISMChequeEntry.ChequeNo = '" & frmALISMCheque.txtChequeNo.Text & "' and ALISMchequeEntry.requisitionNo = ALISMPaymentRequisition.requisitionNo ; ", cnALIS, adOpenKeyset, adLockOptimistic

        With rsLOAD
                If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly
                Else
                            frmALISMCheque.cboRequisitionNo = !RequisitionNo
                            frmALISMCheque.txtRequisitionDate = !RequisitionDate
                            frmALISMCheque.cboDocumentNo = !DocumentNo
                            frmALISMCheque.cboPaymentType = !PaymentType
                            frmALISMCheque.cboClaimCode = !claimcode
                            frmALISMCheque.txtAmount = !Amount
                            frmALISMCheque.txtReference = !reference
                            
                End If

            End With

        Exit Sub

err:
            ErrorMessage

End Sub

Private Sub cmdSearch_Click()
   bApproveCheque = True
   searchRECORD
   loadCHECKENTRIES
   loadCHKENTRIESGRID
   loadGRID
   loadBANKDETAILS
   
   Set rsClaimApproval = New clsALISApproval
   rsClaimApproval.loadAPPROVALDETAILS
   rsClaimApproval.switchCOMMANDBUTTONS
   Set rsClaimApproval = Nothing
   bApproveCheque = False
End Sub

Private Sub searchRECORD()
On Error GoTo err

        Dim strQRE As Variant
        Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset
        strQRE = InputBox("Enter The Cheque No to search.", "Search Value")
        
        rsFind.Open "SELECT * FROM ALISMCheque WHERE ChequeNo = '" & strQRE & "';", cnALIS, adOpenKeyset, adLockOptimistic

        With rsFind
                If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                Else
                            frmALISMCheque.txtTotalAmount.Text = !Amount & ""
                            frmALISMCheque.txtpreparedby.Text = !Preparedby
                            frmALISMCheque.txtDatePrepared.Text = !dateprepared
                            frmALISMCheque.txtStatus.Text = !Status
                            frmALISMCheque.txtPayeeDetails = !PayeeDetails & ""
                            frmALISMCheque.txtApprovedBy = !ApprovedBy & ""
                            frmALISMCheque.txtPrepared.Text = !Prepared & ""
                            frmALISMCheque.txtApproved.Text = !Approved & ""
                            frmALISMCheque.txtAuthorized.Text = !Authorized & ""
                            frmALISMCheque.txtAuthorizedBy = !AuthorizedBy & ""
                            frmALISMCheque.txtChequeDate = !ChequeDate
                            frmALISMCheque.txtChequeNo = !ChequeNo
                            frmALISMCheque.txtDateApproved = !DateApproved & ""
                            frmALISMCheque.cboBankNo = !BankNo
                            frmALISMCheque.txtDateAuthorized = !DateAuthorized & ""
                            frmALISMCheque.txtScheduled.Text = !sCHEDULED & ""
                            frmALISMCheque.txtScheduledBy.Text = !ScheduledBy & ""
                            frmALISMCheque.txtDateScheduled.Text = !DateScheduled & ""
                            frmALISMCheque.txtIssued.Text = !Issued & ""
                            frmALISMCheque.txtIssuedBy.Text = !IssuedBy & ""
                            frmALISMCheque.txtDateIssued.Text = !DateIssued & ""

                            updateCHEQUEENTRIES
                            Edit = True
                End If

            End With

        Exit Sub

err:
            ErrorMessage

End Sub
Private Sub DisableCBEntry()
        With frmALISMCheque
            .cmdEntryUpdate.Enabled = True
            .cmdAddNewEntry.Enabled = False
            .cmdSearchEntry.Enabled = False
            .cmdCancelEntry.Enabled = True
            .cmdPrintEntry.Enabled = False
        End With
End Sub


Private Sub enableCBEntry()
        With frmALISMCheque
            .cmdEntryUpdate.Enabled = False
            .cmdAddNewEntry.Enabled = True
            .cmdSearchEntry.Enabled = True
            .cmdCancelEntry.Enabled = True
            .cmdPrintEntry.Enabled = True
        End With

End Sub


Private Sub DisableCButtons()
        With frmALISMCheque
            .cmdUpdate.Enabled = True
            .cmdAdd.Enabled = False
            .cmdSearch.Enabled = False
            .cmdcancel.Enabled = True
            .cmdApprove.Enabled = False
            .cmdAuthorize.Enabled = False
            .cmdPrint.Enabled = False
            .cmdIssueCheck.Enabled = False
            
        End With
End Sub

Private Sub enableCButtons()
        With frmALISMCheque
            .cmdUpdate.Enabled = False
            .cmdAdd.Enabled = True
            .cmdSearch.Enabled = True
            .cmdcancel.Enabled = True
            .cmdApprove.Enabled = True
            .cmdAuthorize.Enabled = True
            .cmdPrint.Enabled = True
            .cmdIssueCheck.Enabled = True
        End With

End Sub
Private Sub validateENTRY()
On Error GoTo err
            
            If frmALISMCheque.txtAmount <= "" Then
                    MsgBox "The Requisition Amount cannot be left Blank", vbOKOnly
                    frmALISMCheque.txtAmount.SetFocus
                        
            ElseIf frmALISMCheque.txtChequeNo.Text <= "" Then
                    MsgBox "The Cheque No cannot be left Blank", vbOKOnly
                    frmALISMCheque.txtChequeNo.SetFocus
            
            ElseIf frmALISMCheque.txtChequeDate.Text <= "" Then
                    MsgBox "The Cheque Date cannot be left Blank", vbOKOnly
                    frmALISMCheque.txtChequeDate.SetFocus
            
            ElseIf frmALISMCheque.cboBankNo.Text <= "" Then
                    MsgBox "The bank No cannot be left Blank", vbOKOnly
                    frmALISMCheque.cboBankNo.SetFocus
            

            ElseIf frmALISMCheque.cboRequisitionNo <= "" Then
                    MsgBox "The Requisition date cannot be Left Blank", vbOKOnly
                    frmALISMCheque.cboRequisitionNo.SetFocus

            ElseIf frmALISMCheque.txtRequisitionDate <= "" Then
                    MsgBox "The Requisition date cannot be Left Blank", vbOKOnly
                    frmALISMCheque.txtRequisitionDate.SetFocus
                    
            ElseIf frmALISMCheque.cboClaimCode <= "" Then
                    MsgBox "The Claim Code is Required for all the Transaction", vbOKOnly
                    frmALISMCheque.cboClaimCode.SetFocus

            ElseIf frmALISMCheque.cboDocumentNo <= "" Then
                    MsgBox "The Document No is Used to Determine the Payees details", vbOKOnly
                    frmALISMCheque.cboDocumentNo.SetFocus
                    
            ElseIf frmALISMCheque.txtPayeeDetails <= "" Then
                    MsgBox "you must cobfirm the Payee Details prior to Approval", vbOKOnly
                    frmALISMCheque.txtPayeeDetails.SetFocus
            
            Else: bsaveCHEQUE = True
            
            End If

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub validateRECORD()
On Error GoTo err
            
            If frmALISMCheque.txtChequeDate.Text <= "" Then
                    MsgBox "The Cheque Date cannot in the Future", vbOKOnly
                    frmALISMCheque.txtChequeDate.SetFocus
            
            ElseIf DateDiff("D", frmALISMCheque.txtChequeDate.Text, Date) <= 0 Then
                    MsgBox "The Cheque Date cannot be left Blank", vbOKOnly
                    frmALISMCheque.txtChequeDate.SetFocus

            ElseIf frmALISMCheque.cboBankNo.Text <= "" Then
                    MsgBox "The bank No cannot be left Blank", vbOKOnly
                    frmALISMCheque.cboBankNo.SetFocus
            
            ElseIf frmALISMCheque.txtBankAccountNo.Text <= "" Then
                    MsgBox "The Bank Account No cannot be left Blank", vbOKOnly
                    frmALISMCheque.txtBankAccountNo.SetFocus
            
            Else: BSave = True
            
            End If

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub saveCLAIM()
On Error GoTo err
    
    Dim rsCLAIM As ADODB.Recordset, strCLAIM As String
    
    Set rsCLAIM = New Recordset
    
    strCLAIM = "SELECT * from ALISMClaimStatus where claimNo = '" & frmALISMCheque.cboDocumentNo.Text & " '; "
    rsCLAIM.Open strCLAIM, cnALIS, adOpenKeyset, adLockOptimistic

    With rsCLAIM
            !claimstatus = "CHQ-PREPARED"
            !InstallmentNo = 0
            !ClaimSequence = 8
            !StatusDate = Date
            .Update
            .Requery
    End With
Exit Sub

rsCLAIM.Close
strCLAIM = ""

err:
    UpdateErrorMessage
End Sub
Private Sub updateCHEQUEENTRIES()
On Error GoTo err

        With frmALISMCheque
                .txtEntryChequeNo.Text = .txtChequeNo.Text
                .txtEntryChequeDate.Text = .txtChequeDate.Text
                .txtEntryStatus.Text = .txtStatus.Text
        End With

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub GenerateChequeNo()
On Error GoTo err
        Dim rsLAST As ADODB.Recordset, strLAST As String
        
        Set rsLAST = New Recordset
      
        strLAST = "SELECT * FROM ALISPBankAccount Where AccountNo = '" & frmALISMCheque.txtBankAccountNo & "';"
        rsLAST.Open strLAST, cnALIS, adOpenKeyset, adLockOptimistic

                 With rsLAST
                    txtChequeNo = !ChequeNo
                    !ChequeNo = !ChequeNo + 1
                    .Update
            End With

Exit Sub

err:
    UpdateErrorMessage
End Sub

Private Sub cmdUpdate_Click()
        validateRECORD
        If BSave = True Then
            'GenerateChequeNo
            updateCHEQUEENTRIES
            SaveRECORD
            'saveCLAIM
            loadGRID
            loadMGRID
            BSave = False
            enableCButtons
            disableCONTROL
            frmALISMCheque.SSTabCheque.TabEnabled(1) = True
            frmALISMCheque.SSTabCheque.TabEnabled(2) = False
            frmALISMCheque.SSTabCheque.Tab = 1
            disableENTRY
            frmALISMCheque.cmdAddNewEntry.SetFocus
        End If
        
        Set rsClaimApproval = New clsALISApproval
        rsClaimApproval.loadAPPROVALDETAILS
        rsClaimApproval.switchCOMMANDBUTTONS
        Set rsClaimApproval = Nothing

End Sub

