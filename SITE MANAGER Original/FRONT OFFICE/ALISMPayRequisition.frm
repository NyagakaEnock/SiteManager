VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmALISMPayRequisition 
   BackColor       =   &H80000016&
   Caption         =   "Payment Requisition"
   ClientHeight    =   7665
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   Icon            =   "ALISMPayRequisition.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SStPaymentRequisition 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   13361
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Registration"
      TabPicture(0)   =   "ALISMPayRequisition.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label21"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label22"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtTotalDeductions"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtTotalPayments"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Approval/Authorization"
      TabPicture(1)   =   "ALISMPayRequisition.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame13"
      Tab(1).Control(1)=   "Frame12"
      Tab(1).Control(2)=   "Frame7"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Pending/Rejected Requisition"
      TabPicture(2)   =   "ALISMPayRequisition.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame10"
      Tab(2).Control(1)=   "Frame9"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame13 
         Caption         =   "Authorized"
         Height          =   2415
         Left            =   -74640
         TabIndex        =   69
         Top             =   4560
         Width           =   11175
         Begin MSDataGridLib.DataGrid AuthorizedGrid 
            Height          =   2055
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   3625
            _Version        =   393216
            BackColor       =   16777215
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
      Begin VB.Frame Frame6 
         Height          =   1335
         Left            =   9360
         TabIndex        =   66
         Top             =   6120
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
            TabIndex        =   68
            Top             =   720
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
            TabIndex        =   67
            Top             =   120
            Width           =   2055
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Approved"
         Height          =   2295
         Left            =   -74640
         TabIndex        =   64
         Top             =   2280
         Width           =   11175
         Begin MSDataGridLib.DataGrid ApprovedGrid 
            Height          =   1935
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   3413
            _Version        =   393216
            BackColor       =   16777215
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
         Caption         =   "Rejected Requisitions"
         Height          =   3135
         Left            =   -74640
         TabIndex        =   60
         Top             =   3720
         Width           =   11175
         Begin MSDataGridLib.DataGrid RejectionGrid 
            Height          =   2775
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   4895
            _Version        =   393216
            BackColor       =   16777215
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
         Caption         =   "Pending Requisition"
         Height          =   3015
         Left            =   -74640
         TabIndex        =   58
         Top             =   720
         Width           =   11175
         Begin MSDataGridLib.DataGrid PendingGrid 
            Height          =   2655
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   4683
            _Version        =   393216
            BackColor       =   16777215
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
      Begin VB.TextBox txtTotalPayments 
         BackColor       =   &H00FFFFC0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """ ""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
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
         TabIndex        =   56
         Top             =   6960
         Width           =   2535
      End
      Begin VB.TextBox txtTotalDeductions 
         BackColor       =   &H00FFFFC0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """ ""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
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
         TabIndex        =   54
         Top             =   6960
         Width           =   2535
      End
      Begin VB.Frame Frame8 
         Caption         =   "Deduction Grid"
         Height          =   2415
         Left            =   4800
         TabIndex        =   52
         Top             =   4440
         Width           =   4455
         Begin MSDataGridLib.DataGrid DeductionGrid 
            Height          =   2055
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   3625
            _Version        =   393216
            BackColor       =   16777215
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
      Begin VB.Frame Frame7 
         Height          =   1575
         Left            =   -74640
         TabIndex        =   38
         Top             =   720
         Width           =   11175
         Begin VB.Frame Frame11 
            Caption         =   "Rejected Requisitions"
            Height          =   3135
            Left            =   120
            TabIndex        =   62
            Top             =   6120
            Width           =   11175
            Begin MSDataGridLib.DataGrid DataGrid3 
               Height          =   2775
               Left            =   120
               TabIndex        =   63
               Top             =   240
               Width           =   10935
               _ExtentX        =   19288
               _ExtentY        =   4895
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
                  Weight          =   400
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
         Begin VB.Frame Frame2 
            Height          =   1335
            Left            =   120
            TabIndex        =   39
            Top             =   120
            Width           =   10935
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
               Left            =   10320
               TabIndex        =   74
               Top             =   840
               Width           =   240
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
               Left            =   10320
               TabIndex        =   73
               Top             =   480
               Width           =   240
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
               Left            =   10320
               TabIndex        =   72
               Top             =   120
               Width           =   240
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
               Left            =   1680
               TabIndex        =   45
               Top             =   120
               Width           =   2535
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
               Left            =   6000
               TabIndex        =   44
               Top             =   120
               Width           =   2535
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
               Left            =   6000
               TabIndex        =   43
               Top             =   525
               Width           =   2535
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
               Left            =   1680
               TabIndex        =   42
               Top             =   525
               Width           =   2535
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
               Left            =   1680
               TabIndex        =   41
               Top             =   930
               Width           =   2535
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
               Left            =   6000
               TabIndex        =   40
               Top             =   930
               Width           =   2535
            End
            Begin VB.Label Label25 
               Caption         =   "Authorized"
               Height          =   255
               Left            =   9000
               TabIndex        =   77
               Top             =   900
               Width           =   855
            End
            Begin VB.Label Label24 
               Caption         =   "Approved"
               Height          =   255
               Left            =   9000
               TabIndex        =   76
               Top             =   540
               Width           =   1095
            End
            Begin VB.Label Label23 
               Caption         =   "Prepared"
               Height          =   255
               Left            =   9000
               TabIndex        =   75
               Top             =   180
               Width           =   1215
            End
            Begin VB.Label Label3 
               Caption         =   "Prepared By"
               Height          =   255
               Left            =   120
               TabIndex        =   51
               Top             =   195
               Width           =   975
            End
            Begin VB.Label Label2 
               Caption         =   "Date "
               Height          =   255
               Left            =   4560
               TabIndex        =   50
               Top             =   195
               Width           =   495
            End
            Begin VB.Label Label11 
               Caption         =   "Approved By"
               Height          =   255
               Left            =   120
               TabIndex        =   49
               Top             =   555
               Width           =   975
            End
            Begin VB.Label Label15 
               Caption         =   "Date "
               Height          =   255
               Left            =   4560
               TabIndex        =   48
               Top             =   598
               Width           =   495
            End
            Begin VB.Label Label16 
               Caption         =   "Authorized By"
               Height          =   255
               Left            =   120
               TabIndex        =   47
               Top             =   915
               Width           =   975
            End
            Begin VB.Label Label17 
               Caption         =   "Date "
               Height          =   255
               Left            =   4560
               TabIndex        =   46
               Top             =   1003
               Width           =   495
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1935
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   11535
         Begin VB.TextBox txtRequisitionNo 
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
            TabIndex        =   27
            Top             =   240
            Width           =   4575
         End
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
            Left            =   1680
            TabIndex        =   26
            Top             =   1366
            Width           =   1695
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
            Left            =   7680
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   1410
            Width           =   3735
         End
         Begin VB.ComboBox cboClaimCode 
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
            Left            =   1680
            TabIndex        =   24
            Top             =   1004
            Width           =   4575
         End
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1440
            TabIndex        =   23
            Text            =   "Combo1"
            Top             =   5280
            Width           =   2535
         End
         Begin VB.ComboBox cboDocumentNo 
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
            Left            =   7680
            TabIndex        =   22
            Top             =   1024
            Width           =   3735
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
            Height          =   360
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1376
            Width           =   2295
         End
         Begin VB.ComboBox cboPaymentType 
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
            Left            =   1680
            TabIndex        =   20
            Top             =   642
            Width           =   4575
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
            Left            =   7680
            TabIndex        =   19
            Top             =   240
            Width           =   3735
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
            Left            =   7680
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   632
            Width           =   3735
         End
         Begin VB.Label Label1 
            Caption         =   "Requisition No"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   293
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Payment Code"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   1057
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Amount"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   1419
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   " Date"
            Height          =   255
            Left            =   6480
            TabIndex        =   34
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "Document No"
            Height          =   255
            Left            =   6480
            TabIndex        =   33
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Reference"
            Height          =   255
            Left            =   6480
            TabIndex        =   32
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Label1"
            Height          =   255
            Left            =   600
            TabIndex        =   31
            Top             =   5280
            Width           =   1815
         End
         Begin VB.Label Label18 
            Caption         =   "Status"
            Height          =   255
            Left            =   3480
            TabIndex        =   30
            Top             =   1425
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Description"
            Height          =   255
            Left            =   6480
            TabIndex        =   29
            Top             =   690
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "PaymentType"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   695
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3855
         Left            =   9360
         TabIndex        =   8
         Top             =   2280
         Width           =   2295
         Begin VB.CommandButton cmdSearchClaimNo 
            Caption         =   "&Search - Claim No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   120
            TabIndex        =   71
            Top             =   3330
            Width           =   2055
         End
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
            Height          =   420
            Left            =   120
            TabIndex        =   16
            Top             =   2160
            Width           =   2055
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   1800
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
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   2055
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   1410
            Width           =   2055
         End
         Begin VB.CommandButton cmdAddNew 
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
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   240
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
            Height          =   420
            Left            =   120
            TabIndex        =   11
            Top             =   960
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
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   2580
            Width           =   2055
         End
         Begin VB.CommandButton cmdPrintListing 
            Caption         =   "&Print Listing"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   2955
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Payment Grid"
         Height          =   4575
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   4575
         Begin MSDataGridLib.DataGrid PaymentGRID 
            Height          =   4215
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   7435
            _Version        =   393216
            BackColor       =   16777215
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
      Begin VB.Frame Frame5 
         Caption         =   "Payee Details"
         Height          =   2175
         Left            =   4800
         TabIndex        =   1
         Top             =   2280
         Width           =   4575
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
            Height          =   360
            Left            =   360
            TabIndex        =   5
            Top             =   615
            Width           =   3975
         End
         Begin VB.ComboBox cboTownCode 
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
            Left            =   360
            TabIndex        =   4
            Top             =   1335
            Width           =   3975
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
            Height          =   360
            Left            =   360
            TabIndex        =   3
            Top             =   975
            Width           =   3975
         End
         Begin VB.TextBox txtPayeeDetails 
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
            Left            =   360
            TabIndex        =   2
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Label Label22 
         Caption         =   "Payments"
         Height          =   255
         Left            =   480
         TabIndex        =   57
         Top             =   7020
         Width           =   1335
      End
      Begin VB.Label Label21 
         Caption         =   "Deductions"
         Height          =   255
         Left            =   4920
         TabIndex        =   55
         Top             =   7020
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmALISMPayRequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLOANVALUE As cLoanApprovers
Dim rsLOADGRID As clsALISGRID
Dim rsPAYREQ As clsPaymentRequisition
Dim bPaymentType, bClaimCode, bDocumentCode As Boolean
Public rsClaimApproval As clsALISApproval

Private Sub cboClaimCode_GotFocus()
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.ClaimCodeGotFocus
    Set rsPAYREQ = Nothing
End Sub

Private Sub cboClaimCode_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub

Private Sub cboClaimCode_LostFocus()
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.ClaimCodeLostFocus
    Set rsPAYREQ = Nothing
End Sub

Private Sub cboDocumentNo_GotFocus()
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.DocumentNoGotFocus
    Set rsPAYREQ = Nothing
End Sub

Private Sub cboDocumentNo_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboDocumentNo_LostFocus()
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.DocumentNoLostFocus
    Set rsPAYREQ = Nothing
End Sub

Private Sub loadPaymentGRID()
On Error GoTo err

    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset

    If frmALISMPaymentRequisition.cboDocumentNo.Text = Empty Then Exit Sub
    rsGRID.Open "SELECT ALISPClaimType.ClaimTypeDescription, ALISMClaim.Amount FROM ALISMClaim, ALISPClaimType WHERE ALISMClaim.claimNo =  '" & frmALISMPaymentRequisition.cboDocumentNo & "' and ALISMClaim.type = 'A' and ALISMClaim.ClaimType = ALISPClaimType.ClaimType;", cnALIS, adOpenKeyset, adLockOptimistic
 
    Set PaymentGRID.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub LoadDeductionGRID()
On Error GoTo err
    Dim rsGRID1 As ADODB.Recordset
    Set rsGRID1 = New ADODB.Recordset
    
    If frmALISMPaymentRequisition.cboDocumentNo.Text = Empty Then Exit Sub

    rsGRID1.Open "SELECT ALISPClaimType.ClaimTypeDescription, ALISMClaim.Amount FROM ALISMClaim, ALISPClaimType WHERE ALISMClaim.claimNo =  '" & frmALISMPaymentRequisition.cboDocumentNo & "' and ALISMClaim.type = 'D' and ALISMClaim.ClaimType = ALISPClaimType.ClaimType;", cnALIS, adOpenKeyset, adLockOptimistic
    Set DeductionGrid.DataSource = rsGRID1
    
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub LoadRejectedGRID()
On Error GoTo err
    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset
    
    If frmALISMPaymentRequisition.cboDocumentNo.Text = Empty Then Exit Sub

    rsGRID.Open "SELECT * FROM ALISMPaymentRequisition WHERE Status =  'REQ-PREP' ;", cnALIS, adOpenKeyset, adLockOptimistic
    Set RejectionGrid.DataSource = rsGRID
    
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub loadPendingGRID()
On Error GoTo err
    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset
    
    If frmALISMPaymentRequisition.cboDocumentNo.Text = Empty Then Exit Sub

    rsGRID.Open "SELECT * FROM ALISMPaymentRequisition WHERE Status =  'REQ-PREP' ;", cnALIS, adOpenKeyset, adLockOptimistic
    Set PendingGrid.DataSource = rsGRID
    
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub loadApprovedGRID()
On Error GoTo err
    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset
    
    If frmALISMPaymentRequisition.cboDocumentNo.Text = Empty Then Exit Sub

    rsGRID.Open "SELECT * FROM ALISMPaymentRequisition WHERE Status =  'REQ APPROVAL' ;", cnALIS, adOpenKeyset, adLockOptimistic
    Set ApprovedGrid.DataSource = rsGRID
    
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub loadAuthorizedGRID()
On Error GoTo err
    Dim rsGRID As ADODB.Recordset
    Set rsGRID = New ADODB.Recordset
    
    If frmALISMPaymentRequisition.cboDocumentNo.Text = Empty Then Exit Sub

    rsGRID.Open "SELECT * FROM ALISMPaymentRequisition WHERE Status =  'REQ AUTHORIZATION' ;", cnALIS, adOpenKeyset, adLockOptimistic
    Set AuthorizedGrid.DataSource = rsGRID
    
    Exit Sub
err:
    ErrorMessage
End Sub


Private Sub loadTOTALS()
On Error GoTo err
    Dim rsTOTALS As ADODB.Recordset
    Set rsTOTALS = New ADODB.Recordset
    
    rsTOTALS.Open "SELECT * FROM ALISMClaimTotal WHERE claimNo =  '" & frmALISMPaymentRequisition.cboDocumentNo & "';", cnALIS, adOpenKeyset, adLockOptimistic
    
    With rsTOTALS
            If .BOF Or .EOF = True Then Exit Sub
            frmALISMPaymentRequisition.txtTotalPayments.Text = !proceeds
            frmALISMPaymentRequisition.txtTotalDeductions.Text = !deductions

    End With

rsTOTALS.Close

Exit Sub
err:
    ErrorMessage
End Sub


Private Sub LoadGrid()
    loadPaymentGRID
    LoadDeductionGRID
    loadPendingGRID
    LoadRejectedGRID
    loadApprovedGRID
    loadAuthorizedGRID
    loadTOTALS
End Sub


Private Sub cboPaymentType_GotFocus()
    bPaymentType = True
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.selectPaymentTypeGotFocus
    Set rsPAYREQ = Nothing
End Sub

Private Sub cboPaymentType_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub

Private Sub cboPaymentType_LostFocus()
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.selectPaymentTypeLostFocus
    Set rsPAYREQ = Nothing
    bPaymentType = False
End Sub

Private Sub loadPaymentDescription()
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.loadPaymentDescription
    Set rsPAYREQ = Nothing
End Sub

Private Sub cmdApprove_Click()
        bapproveREQUISITION = True
        bAuthorizeREQUISITION = False
        Set rsClaimApproval = New clsALISApproval
        rsClaimApproval.checkAPPROVEDDISCHARGE
        If bapproveREQUISITION = False Then Exit Sub
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
                            
                        strAPPROVED = "SELECT * FROM ALISMLoanOperation  WHERE ApplicationNo = '" & frmALISMPaymentRequisition.txtRequisitionNo & "' and operationType = '" & !OperationType & "' ;"
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
                    
                strAuthorization = "SELECT * FROM ALISPLoanOperationType  WHERE PaymentAuthorization = '1' ;"
                rsAUTHORIZATION.Open strAuthorization, cnALIS, adOpenKeyset, adLockOptimistic

                With rsAUTHORIZATION
                        If .EOF Or .BOF Then Exit Sub
                        
                        Dim rsAPPROVED As ADODB.Recordset, strAPPROVED As String
                        Set rsAPPROVED = New Recordset
                            
                        strAPPROVED = "SELECT * FROM ALISMLoanOperation  WHERE ApplicationNo = '" & frmALISMPaymentRequisition.txtRequisitionNo & "' and operationType = '" & !OperationType & "' ;"
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
                            
                        strCHKAPPROVAL = "SELECT * FROM ALISPLoanOperationType  WHERE PaymentApproval = '1' ;"
                        rsCHKAPPROVAL.Open strCHKAPPROVAL, cnALIS, adOpenKeyset, adLockOptimistic
                        
                        With rsCHKAPPROVAL
                                If .BOF Or .EOF Then Exit Sub
                                
                                    Dim rsAPPROVE As ADODB.Recordset, strAPPROVE As String
                                    Set rsAPPROVE = New Recordset
                                        
                                    strAPPROVE = "SELECT * FROM ALISMLoanOperation  WHERE ApplicationNo = '" & frmALISMPaymentRequisition.txtRequisitionNo & "' and operationType = '" & !OperationType & "' ;"
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
    If frmALISMPaymentRequisition.txtApproved = "Y" Then
            bAuthorizeREQUISITION = True
            Set rsClaimApproval = New clsALISApproval
            rsClaimApproval.checkAPPROVEDDISCHARGE
            If bAuthorizeREQUISITION = False Then Exit Sub
            rsClaimApproval.approveCLAIM
            rsClaimApproval.loadAPPROVALDETAILS
            rsClaimApproval.switchCOMMANDBUTTONS
            Set rsClaimApproval = Nothing
            bAuthorizeREQUISITION = False
    End If
End Sub

Private Sub cmdPrint_Click()
    If frmALISMPaymentRequisition.txtAuthorized.Text = "Y" Then
            Load frmPayRequisition
            frmPayRequisition.Show 1, Me
    End If
End Sub

Private Sub cmdprintlisting_Click()
    Load frmPayRequisitionListing
    frmPayRequisitionListing.Show 1, Me
End Sub



Private Sub cmdSearch_Click()
    bapproveREQUISITION = True
    
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.searchRECORD
    rsPAYREQ.loadPaymentDescription
    Set rsPAYREQ = Nothing
    
    Set rsClaimApproval = New clsALISApproval
    rsClaimApproval.switchCOMMANDBUTTONS
    rsClaimApproval.loadAPPROVALDETAILS
    Set rsClaimApproval = Nothing
    bapproveREQUISITION = False

End Sub

Private Sub cmdSearchClaimNo_Click()
    bapproveREQUISITION = True
    bUseClaimNo = True
    
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.searchRECORD
    Set rsPAYREQ = Nothing
    
    Set rsClaimApproval = New clsALISApproval
    rsClaimApproval.switchCOMMANDBUTTONS
    rsClaimApproval.loadAPPROVALDETAILS
    Set rsClaimApproval = Nothing
    bUseClaimNo = False
    bapproveREQUISITION = False
End Sub

Private Sub Form_Activate()
    Set rsPAYREQ = New clsPaymentRequisition
    disableALLRECORD
    enableButtons
    rsPAYREQ.enableXButtons
    Set rsPAYREQ = Nothing
  
   frmALISMPaymentRequisition.SStPaymentRequisition.Tab = 0
End Sub

Private Sub Form_Load()
    Call OpenConnection
    
'    Set RsCode = New Recordset
'            strcode = "SELECT * from ALISMPaymentRequisition;"
'
'    RsCode.Open strcode, cnALIS, adOpenKeyset, adLockOptimistic

End Sub

Private Sub cmdAddNew_Click()
        Set rsPAYREQ = New clsPaymentRequisition
        rsPAYREQ.EnableCONTROL
        clearALLRECORD
        disableButtons
        rsPAYREQ.DisableXButtons
        Set rsPAYREQ = Nothing
End Sub

Private Sub cmdCancel_Click()
        Set rsPAYREQ = New clsPaymentRequisition
        rsPAYREQ.cancelRECORD
        Set rsPAYREQ = Nothing
End Sub

Private Sub cmdDelete_Click()
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.DeleteRecord
    Set rsPAYREQ = Nothing
End Sub

Private Sub cmdEdit_Click()
    Set rsPAYREQ = New clsPaymentRequisition
    rsPAYREQ.EditRECORD
    Set rsPAYREQ = Nothing
    
    Set rsClaimApproval = New clsALISApproval
    rsClaimApproval.switchCOMMANDBUTTONS
    rsClaimApproval.loadAPPROVALDETAILS
    Set rsClaimApproval = Nothing

End Sub


Private Sub cmdUpdate_Click()
        Set rsPAYREQ = New clsPaymentRequisition
        rsPAYREQ.updateRECORD
        Set rsPAYREQ = Nothing
End Sub
Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
        
        Dim strCategoryCode, StrLOAD As String
        Dim rsLOAD As ADODB.Recordset
        Set rsLOAD = New ADODB.Recordset
        
                
        If bPaymentType = True Then
                
                    With frmALISMPaymentRequisition
                            .cboPaymentType.Text = ""
                    End With
                    
                    strCategoryCode = Item.Text
                    frmALISMPaymentRequisition.cboPaymentType.Text = Item.Text
            
                    StrLOAD = "Select * from ALISPPaymentType where PaymentType = '" & frmALISMPaymentRequisition.cboPaymentType & "' ;"
                    rsLOAD.Open StrLOAD, cnALIS, adOpenKeyset, adLockOptimistic
                    
                    With rsLOAD
                            If .BOF Or .EOF Then Exit Sub
                                    frmALISMPaymentRequisition.txtPaymentDescription = !Description
                    End With

        End If
        
End Sub

