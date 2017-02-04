VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmALISMOrders 
   Caption         =   "Order Processing"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   12091
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Order Processing"
      TabPicture(0)   =   "frmALISMOrders.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Bankers Order Per Policy"
      TabPicture(1)   =   "frmALISMOrders.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "All Bankers Orders"
      TabPicture(2)   =   "frmALISMOrders.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   14
         Top             =   360
         Width           =   11535
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   5535
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   9763
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
         Height          =   5775
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   11535
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   5535
            Left            =   120
            TabIndex        =   62
            Top             =   120
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   9763
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
         Height          =   6375
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   11535
         Begin VB.TextBox txtDatePrepared 
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
            Left            =   5520
            TabIndex        =   56
            Top             =   5760
            Width           =   3255
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
            Left            =   1560
            TabIndex        =   54
            Top             =   5760
            Width           =   2775
         End
         Begin VB.Frame Frame6 
            Caption         =   "Company Details"
            Height          =   1215
            Left            =   120
            TabIndex        =   48
            Top             =   4440
            Width           =   8895
            Begin VB.TextBox txtIssuedBy 
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
               Left            =   5400
               TabIndex        =   8
               Top             =   720
               Width           =   3255
            End
            Begin VB.TextBox txtCoyAccountNo 
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
               TabIndex        =   51
               Top             =   720
               Width           =   2775
            End
            Begin VB.TextBox txtCoyBankName 
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
               Left            =   4200
               TabIndex        =   49
               Top             =   240
               Width           =   4455
            End
            Begin VB.ComboBox cboCoyBankNO 
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
               Left            =   1440
               TabIndex        =   10
               Top             =   240
               Width           =   2775
            End
            Begin VB.Label Label18 
               Caption         =   "Issued By"
               Height          =   255
               Left            =   4320
               TabIndex        =   53
               Top             =   780
               Width           =   1215
            End
            Begin VB.Label Label17 
               Caption         =   "Account No"
               Height          =   255
               Left            =   120
               TabIndex        =   52
               Top             =   780
               Width           =   1215
            End
            Begin VB.Label Label16 
               Caption         =   "Bank No"
               Height          =   255
               Left            =   120
               TabIndex        =   50
               Top             =   300
               Width           =   1335
            End
         End
         Begin VB.Frame Frame5 
            Height          =   2295
            Left            =   120
            TabIndex        =   33
            Top             =   2160
            Width           =   8895
            Begin VB.TextBox txtDateofFirstPayment 
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
               Left            =   6240
               TabIndex        =   7
               Top             =   1200
               Width           =   2175
            End
            Begin VB.TextBox txtPaymentModeDescription 
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
               Left            =   4200
               TabIndex        =   45
               Top             =   1680
               Width           =   4455
            End
            Begin VB.ComboBox cboPaymentMode 
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
               Left            =   1440
               TabIndex        =   9
               Top             =   1680
               Width           =   2775
            End
            Begin VB.TextBox txtAccountNo 
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
               Left            =   1440
               TabIndex        =   6
               Top             =   1200
               Width           =   2775
            End
            Begin VB.TextBox txtBankNo 
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
               TabIndex        =   42
               Top             =   720
               Width           =   1215
            End
            Begin VB.ComboBox cboBankName 
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
               Left            =   2760
               TabIndex        =   5
               Top             =   720
               Width           =   5895
            End
            Begin VB.TextBox txtPaymentMethodDescription 
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
               Left            =   4200
               TabIndex        =   34
               Top             =   240
               Width           =   4455
            End
            Begin VB.ComboBox cboPaymentMethod 
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
               Left            =   1440
               TabIndex        =   4
               Top             =   240
               Width           =   2775
            End
            Begin MSComCtl2.DTPicker DTPickerDateofFirstPayment 
               Height          =   375
               Left            =   8400
               TabIndex        =   58
               Top             =   1200
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   661
               _Version        =   393216
               Format          =   56623105
               CurrentDate     =   38015
            End
            Begin VB.Label Label15 
               Caption         =   "Date of First Payment"
               Height          =   255
               Left            =   4320
               TabIndex        =   47
               Top             =   1260
               Width           =   1575
            End
            Begin VB.Label Label14 
               Caption         =   "Payment Mode"
               Height          =   255
               Left            =   120
               TabIndex        =   46
               Top             =   1740
               Width           =   1095
            End
            Begin VB.Label Label13 
               Caption         =   "Account No"
               Height          =   255
               Left            =   120
               TabIndex        =   44
               Top             =   1260
               Width           =   1215
            End
            Begin VB.Label Label12 
               Caption         =   "Bank No"
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   780
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "Payment Method"
               Height          =   255
               Left            =   120
               TabIndex        =   35
               Top             =   300
               Width           =   1335
            End
         End
         Begin VB.Frame fraCButtons 
            Height          =   3495
            Index           =   0
            Left            =   9120
            TabIndex        =   16
            Top             =   2160
            Width           =   2295
            Begin VB.CommandButton cmdSearchPolicy 
               Caption         =   "&Search Pol"
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
               TabIndex        =   61
               Top             =   1806
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
               TabIndex        =   19
               Top             =   2850
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
               Height          =   495
               Left            =   120
               TabIndex        =   18
               Top             =   2328
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
               TabIndex        =   11
               Top             =   762
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
               Height          =   495
               Left            =   120
               TabIndex        =   0
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
               Height          =   495
               Left            =   120
               TabIndex        =   17
               Top             =   1284
               Width           =   2055
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Policy Details"
            Height          =   2055
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   11295
            Begin VB.TextBox txtOrderNo 
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
               TabIndex        =   59
               Top             =   240
               Width           =   2055
            End
            Begin VB.TextBox txtMaturityDate 
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
               Left            =   8760
               TabIndex        =   38
               Top             =   680
               Width           =   2295
            End
            Begin VB.TextBox txtTermOfPolicy 
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
               Left            =   5040
               TabIndex        =   37
               Top             =   680
               Width           =   2055
            End
            Begin VB.TextBox txtDateOFCommencement 
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
               TabIndex        =   36
               Top             =   660
               Width           =   2055
            End
            Begin VB.TextBox txtNames 
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
               Left            =   6720
               TabIndex        =   25
               Top             =   240
               Width           =   4335
            End
            Begin VB.TextBox txtPolicyNo 
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
               Left            =   5040
               TabIndex        =   2
               Top             =   240
               Width           =   2055
            End
            Begin VB.TextBox txtExpectedPremium 
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
               TabIndex        =   24
               Top             =   1080
               Width           =   2055
            End
            Begin VB.TextBox txtPremiumCount 
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
               TabIndex        =   23
               Top             =   1560
               Width           =   2055
            End
            Begin VB.TextBox txtSurrenderValue 
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
               Left            =   5040
               TabIndex        =   22
               Top             =   1120
               Width           =   2055
            End
            Begin VB.TextBox txtLoanAmount 
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
               Left            =   8760
               TabIndex        =   21
               Top             =   1120
               Width           =   2295
            End
            Begin VB.TextBox txtOrderAmount 
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
               Left            =   5040
               TabIndex        =   3
               Top             =   1560
               Width           =   2055
            End
            Begin VB.TextBox txtOrderDate 
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
               Left            =   8760
               TabIndex        =   20
               Top             =   1560
               Width           =   2295
            End
            Begin VB.Label Label21 
               Caption         =   "OrderNo"
               Height          =   255
               Left            =   120
               TabIndex        =   60
               Top             =   300
               Width           =   1215
            End
            Begin VB.Label Label11 
               Caption         =   "Maturity Date"
               Height          =   255
               Left            =   7440
               TabIndex        =   41
               Top             =   780
               Width           =   1215
            End
            Begin VB.Label Label10 
               Caption         =   "Term of Policy"
               Height          =   255
               Left            =   3840
               TabIndex        =   40
               Top             =   780
               Width           =   1215
            End
            Begin VB.Label Label9 
               Caption         =   "DOC"
               Height          =   255
               Left            =   120
               TabIndex        =   39
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label2 
               Caption         =   "Policy No"
               Height          =   255
               Left            =   3840
               TabIndex        =   32
               Top             =   300
               Width           =   1215
            End
            Begin VB.Label Label3 
               Caption         =   "Premium "
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   1140
               Width           =   1215
            End
            Begin VB.Label Label4 
               Caption         =   "Premium Count"
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   1620
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Surrender Value "
               Height          =   255
               Left            =   3840
               TabIndex        =   29
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Loan Amount"
               Height          =   255
               Left            =   7440
               TabIndex        =   28
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "Order Amount"
               Height          =   255
               Left            =   3840
               TabIndex        =   27
               Top             =   1620
               Width           =   1215
            End
            Begin VB.Label Label8 
               Caption         =   "Order Date"
               Height          =   255
               Left            =   7440
               TabIndex        =   26
               Top             =   1620
               Width           =   1215
            End
         End
         Begin VB.Label Label20 
            Caption         =   "Date"
            Height          =   255
            Left            =   4440
            TabIndex        =   57
            Top             =   5820
            Width           =   1215
         End
         Begin VB.Label Label19 
            Caption         =   "Prepared By"
            Height          =   255
            Left            =   240
            TabIndex        =   55
            Top             =   5820
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmALISMOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim translater As New cMoneyConverter

Private Sub DTPickerDateofFirstPayment_Change()
On Error GoTo err
    frmALISMOrders.DTPickerDateofFirstPayment.MinDate = Date
    frmALISMOrders.DTPickerDateofFirstPayment.MousePointer = cc2IBeam
    frmALISMOrders.txtDateofFirstPayment.Text = frmALISMOrders.DTPickerDateofFirstPayment.Value
Exit Sub
err:
    ErrorMessage
End Sub



Private Sub cboCoybankNo_GotFocus()
On Error GoTo err

        Dim rsBANKGF As ADODB.Recordset, strBANKGF As String
        Set rsBANKGF = New Recordset
        
        strBANKGF = "SELECT * FROM ALISPBankAccount;"
        rsBANKGF.Open strBANKGF, cnALIS, adOpenKeyset, adLockOptimistic
        
        frmALISMOrders.cboCoyBankNO.Clear

        With rsBANKGF
            Do Until .EOF
            frmALISMOrders.cboCoyBankNO.AddItem !details
                    .MoveNext
            Loop
    
        End With

rsBANKGF.Close
strBANKGF = ""

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cboCoyBankNo_KeyPress(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub

Private Sub cboCoybankNo_LostFocus()
On Error GoTo err

        Dim rsBANKLF As ADODB.Recordset, strBANKLF As String
        Set rsBANKLF = New Recordset
        
        rsBANKLF.Open "SELECT * FROM ALISPBankAccount WHERE Details = '" & frmALISMOrders.cboCoyBankNO.Text & "'", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsBANKLF
                If .EOF And .BOF Then Exit Sub
                frmALISMOrders.cboCoyBankNO.Text = !BankNo
                frmALISMOrders.txtCoyBankName.Text = !details
                frmALISMOrders.txtCoyAccountNo.Text = !AccountNo
        End With
        
rsBANKLF.Close
strBANKLF = ""

Exit Sub

err:
        ErrorMessage

End Sub

Private Sub clearRECORD()
On Error GoTo err
    With frmALISMOrders
        .txtOrderNo.Text = ""
        .txtAccountNo.Text = ""
        .txtBankNo.Text = ""
        .txtCoyAccountNo.Text = ""
        .txtDateofCommencement.Text = ""
        .txtDateofFirstPayment.Text = ""
        .txtDatePrepared.Text = ""
        .txtexpectedpremium.Text = ""
        .txtIssuedBy.Text = ""
        .txtLoanAmount.Text = ""
        .txtMaturityDate.Text = ""
        .txtNames.Text = ""
        .txtOrderAmount.Text = ""
        .txtOrderDate.Text = ""
        .txtPaymentMethodDescription.Text = ""
        .txtPaymentModeDescription.Text = ""
        .txtPolicyNo.Text = ""
        .txtPremiumcount.Text = ""
        .txtPreparedBy.Text = ""
        .txtSurrenderValue.Text = ""
        .txtTermofPolicy.Text = ""
        .cboBankName.Text = ""
        .cboCoyBankNO.Text = ""
        .cboPaymentMethod.Text = ""
        .cboPaymentMode.Text = ""
        .txtCoyBankName.Text = ""
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub testRECORD()
On Error GoTo err
    With frmALISMOrders
        .txtOrderNo.Locked = True
        .txtAccountNo.Locked = Bval
        .txtBankNo.Locked = True
        .txtCoyAccountNo.Locked = True
        .txtDateofCommencement.Locked = True
        .txtDateofFirstPayment.Locked = Bval
        .txtDatePrepared.Locked = True
        .txtexpectedpremium.Locked = True
        .txtIssuedBy.Locked = Bval
        .txtLoanAmount.Locked = True
        .txtMaturityDate.Locked = True
        .txtNames.Locked = True
        .txtOrderAmount.Locked = Bval
        .txtOrderDate.Locked = True
        .txtPaymentMethodDescription.Locked = True
        .txtPaymentModeDescription.Locked = True
        .txtPolicyNo.Locked = Bval
        .txtPremiumcount.Locked = True
        .txtPreparedBy.Locked = True
        .txtSurrenderValue.Locked = True
        .txtTermofPolicy.Locked = Bval
        .cboBankName.Locked = Bval
        .cboCoyBankNO.Locked = Bval
        .cboPaymentMethod.Locked = Bval
        .cboPaymentMode.Locked = Bval
        .txtCoyBankName.Locked = True
    End With

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub cmdPrint_Click()
On Error GoTo err
    If frmALISMOrders.txtOrderNo.Text <= "" Then
        MsgBox "Cannot Print The Details, Kindly Load the Bankers Other", vbOKOnly
        frmALISMOrders.txtOrderNo.SetFocus
    ElseIf frmALISMOrders.txtPolicyNo.Text <= "" Then
        MsgBox "The Policy No Does not Exist", vbOKOnly
    Else
        Load frmALISRBankersOrder
        frmALISRBankersOrder.Show 1, Me
    End If
    
Exit Sub
err:
    ErrorMessage
End Sub




Private Sub Form_Load()
On Error GoTo err
    OpenConnection
    
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub cmdAddNew_Click()
        enableRECORD
        clearRECORD
        DisableCommandButtons
End Sub

Private Sub cmdUpdate_Click()

        validateRECORD
        If bsaveRECORD = True Then
            saveRECORD
            savePOLICY
            LoadGrid
            loadMGRID
            bsaveRECORD = False
            EnableCommandButtons
            disableRECORD
        End If
End Sub


Private Sub enableRECORD()
On Error GoTo err

    Bval = False
    testRECORD
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub disableRECORD()
On Error GoTo err

    Bval = True
    testRECORD
Exit Sub
err:
    ErrorMessage
End Sub


Private Sub validateRECORD()
On Error GoTo err
    
    With frmALISMOrders
        If .txtPolicyNo.Text <= "" Then
                MsgBox " The Policy No cannot be Left Blank", vbOKOnly
                .txtPolicyNo.SetFocus
        
        ElseIf .txtOrderAmount.Text <= "" Then
                MsgBox "The Amount cannot be Left Blank", vbOKOnly
                .txtOrderAmount.SetFocus
                
        ElseIf .txtOrderDate.Text <= "" Then
                MsgBox "The Order date cannot be Left Blank", vbOKOnly
                .txtOrderDate.SetFocus
        
        ElseIf .cboPaymentMethod.Text <= "" Then
                MsgBox "The Payment Method Cannot be Left Blank", vbOKOnly
                .cboPaymentMethod.SetFocus
        
        ElseIf .cboBankName.Text <= "" Then
                MsgBox "The Policyholders Bank No is Needed for Processing", vbOKOnly
                .cboBankName.SetFocus
                
        ElseIf .txtDateofFirstPayment.Text <= "" Then
                MsgBox "The First Payment Date is Needed", vbOKOnly
                .txtDateofFirstPayment.SetFocus
        
        ElseIf .cboPaymentMode.Text <= "" Then
                MsgBox "The Payment Mode is Required "
                .cboPaymentMode.SetFocus
        
        ElseIf .cboCoyBankNO.Text <= "" Then
                MsgBox "The Companys Bank Account Number is Needed", vbOKOnly
                .cboCoyBankNO.SetFocus
        
        ElseIf .txtCoyAccountNo.Text <= "" Then
                MsgBox "The Companys Bank Account Number is Needed", vbOKOnly
                .txtCoyAccountNo.SetFocus
        
        ElseIf .txtIssuedBy.Text <= "" Then
                MsgBox "The Details of the Person Issuing the Instructions is Required", vbOKOnly
                .txtIssuedBy.SetFocus
        Else
                bsaveRECORD = True
        
        End If
    End With

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub saveRECORD()
On Error GoTo err
      Dim rsSAVE As ADODB.Recordset
      Set rsSAVE = New ADODB.Recordset
    
      rsSAVE.Open "SELECT * FROM ALISMOrders ; ", cnALIS, adOpenKeyset, adLockOptimistic
    
    With rsSAVE
            .AddNew
            !DocumentNo = frmALISMOrders.txtPolicyNo.Text
            !OrderAmount = Val(frmALISMOrders.txtOrderAmount.Text)
            !OrderDate = frmALISMOrders.txtOrderDate.Text
            !PaymentMethod = frmALISMOrders.cboPaymentMethod
            !BankNo = frmALISMOrders.txtBankNo.Text
            !AccountNo = frmALISMOrders.txtAccountNo.Text
            
            figures = frmALISMOrders.txtOrderAmount
            Set translater = New cMoneyConverter
            translater.CallConverter
            !inwords = inwords
            Set translater = Nothing
            
            !DateofFirstPayment = frmALISMOrders.txtDateofFirstPayment.Text
            !PaymentMode = frmALISMOrders.cboPaymentMode.Text
            !CoyBankNo = frmALISMOrders.cboCoyBankNO.Text
            !CoyAccountNo = frmALISMOrders.txtCoyAccountNo.Text
            !issuedBy = frmALISMOrders.txtIssuedBy.Text
            !DocumentType = "POL"
            !Preparedby = frmALISMOrders.txtPreparedBy.Text
            !dateprepared = frmALISMOrders.txtDatePrepared.Text
            .Update
            .Requery
            
            frmALISMOrders.txtOrderNo.Text = !OrderNo
    End With

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub cmdSearchPolicy_Click()
On Error GoTo Myerr

        Dim strQRE As Variant
        Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset
        strQRE = InputBox("Enter The Policy No to search.", "Search Value")
        
        rsFind.Open "SELECT * FROM ALISMOrders, ALISMPolicy, ALISMReference WHERE ALISMOrders.DocumentNo = '" & strQRE & "' and ALISMPolicy.PolicyNo = ALISMOrders.DocumentNo and ALISMReference.ReferenceNo = ALISMPolicy.ReferenceNo ;", cnALIS, adOpenKeyset, adLockOptimistic

        With rsFind
                If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                Else
                        frmALISMOrders.txtOrderNo.Text = !OrderNo
                        frmALISMOrders.txtPolicyNo.Text = !PolicyNo
                        frmALISMOrders.txtNames.Text = Trim(!Othernames) + " " + Trim(!Surname)
                        frmALISMOrders.txtexpectedpremium.Text = !ExpectedPremium
                        frmALISMOrders.txtPremiumcount.Text = !UnitCount
                        frmALISMOrders.cboPaymentMethod.Text = !PaymentMethod
                        frmALISMOrders.txtSurrenderValue.Text = !SurrenderValue & ""
                        frmALISMOrders.txtOrderAmount.Text = !ExpectedPremium
                        frmALISMOrders.txtDateofCommencement.Text = !DateofCommencement
                        frmALISMOrders.txtTermofPolicy.Text = !TermOfPolicy
                        frmALISMOrders.txtMaturityDate.Text = !MaturityDate
                        frmALISMOrders.txtDateofFirstPayment.Text = !DateofCommencement
                        frmALISMOrders.txtAccountNo.Text = !AccountNo & ""
                        frmALISMOrders.txtIssuedBy.Text = !issuedBy
                        frmALISMOrders.cboPaymentMode.Text = !PaymentMode
                        frmALISMOrders.txtBankNo.Text = !BankNo & ""
                        frmALISMOrders.txtPreparedBy.Text = UserName
                        frmALISMOrders.txtDatePrepared.Text = !dateprepared
                        frmALISMOrders.txtOrderDate.Text = !OrderDate
                        frmALISMOrders.cboCoyBankNO.Text = !CoyBankNo
                     Edit = True
                End If
            
           ' LoadGrid
           loadBANK
           loadCOYBANK
           loadPAYMENTMETHOD
           loadPaymentModeDESCRIPTION
            LoadGrid
            loadMGRID

            End With

        Exit Sub

Myerr:
            ErrorMessage

End Sub
Private Sub cmdSearch_Click()
On Error GoTo Myerr

        Dim strQRE As Variant
        Dim rsFind As ADODB.Recordset, Edit As Boolean

        Set rsFind = New ADODB.Recordset
        strQRE = InputBox("Enter The Order No to search.", "Search Value")
        
        rsFind.Open "SELECT * FROM ALISMOrders, ALISMPolicy, ALISMReference WHERE ALISMOrders.OrderNo LIKE '" & strQRE & "' and ALISMPolicy.PolicyNo = ALISMOrders.DocumentNo and ALISMReference.ReferenceNo = ALISMPolicy.ReferenceNo ;", cnALIS, adOpenKeyset, adLockOptimistic

        With rsFind
                If .EOF And .BOF Then
                        MsgBox "The requested record does not exist in the database. Check your search statement.", vbOKOnly + vbExclamation, "Searching"
                Else
                        frmALISMOrders.txtOrderNo.Text = !OrderNo
                        frmALISMOrders.txtPolicyNo.Text = !PolicyNo
                        frmALISMOrders.txtNames.Text = Trim(!Othernames) + " " + Trim(!Surname)
                        frmALISMOrders.txtexpectedpremium.Text = !ExpectedPremium
                        frmALISMOrders.txtPremiumcount.Text = !UnitCount
                        frmALISMOrders.cboPaymentMethod.Text = !PaymentMethod
                        frmALISMOrders.txtSurrenderValue.Text = !SurrenderValue & ""
                        frmALISMOrders.txtOrderAmount.Text = !ExpectedPremium
                        frmALISMOrders.txtDateofCommencement.Text = !DateofCommencement
                        frmALISMOrders.txtTermofPolicy.Text = !TermOfPolicy
                        frmALISMOrders.txtMaturityDate.Text = !MaturityDate
                        frmALISMOrders.txtDateofFirstPayment.Text = !DateofCommencement
                        frmALISMOrders.txtAccountNo.Text = !AccountNo & ""
                        frmALISMOrders.txtIssuedBy.Text = !issuedBy
                        frmALISMOrders.cboPaymentMode.Text = !PaymentMode
                        frmALISMOrders.txtBankNo.Text = !BankNo & ""
                        frmALISMOrders.txtPreparedBy.Text = UserName
                        frmALISMOrders.txtDatePrepared.Text = !dateprepared
                        frmALISMOrders.txtOrderDate.Text = !OrderDate
                        frmALISMOrders.cboCoyBankNO.Text = !CoyBankNo
                     Edit = True
                End If
            
           ' LoadGrid
           loadBANK
           loadCOYBANK
           loadPAYMENTMETHOD
           loadPaymentModeDESCRIPTION
            LoadGrid
            loadMGRID

            End With

        Exit Sub

Myerr:
            ErrorMessage

End Sub



Private Sub cmdCancel_Click()
On Error GoTo Myerr
        EnableCommandButtons
        clearRECORD
        disableRECORD
        Exit Sub

Myerr:
        ErrorMessage
End Sub
Private Sub savePOLICY()
'On Error GoTo Err

        Dim rsPOLICY As ADODB.Recordset
        Set rsPOLICY = New ADODB.Recordset
      
        rsPOLICY.Open "SELECT * FROM ALISMPolicy WHERE ALISMPolicy.PolicyNo= '" & frmALISMOrders.txtPolicyNo.Text & "' ; ", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsPOLICY
                If .EOF And .BOF Then Exit Sub
                !PaymentMethod = frmALISMOrders.cboPaymentMethod.Text
                !AccountNo = frmALISMOrders.txtAccountNo.Text
                !BankNo = frmALISMOrders.txtBankNo.Text
        End With
        
Exit Sub

err:
    ErrorMessage
End Sub



Private Sub LoadGrid()
On Error GoTo err
        
        Dim rsGRID As ADODB.Recordset, StrGRID As String
        Set rsGRID = New ADODB.Recordset
        
        StrGRID = "SELECT * FROM ALISMOrders WHERE DocumentNo= '" & frmALISMOrders.txtPolicyNo.Text & "'; "
        rsGRID.Open StrGRID, cnALIS, adOpenKeyset, adLockOptimistic
        
        If rsGRID.BOF Or rsGRID.BOF Then Exit Sub
        Set frmALISMOrders.DataGrid1.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub loadMGRID()
On Error GoTo err
        
        Dim rsGRID As ADODB.Recordset, StrGRID As String
        Set rsGRID = New ADODB.Recordset
        
        StrGRID = "SELECT * FROM ALISMOrders; "
        rsGRID.Open StrGRID, cnALIS, adOpenKeyset, adLockOptimistic
        
        Set frmALISMOrders.DataGrid2.DataSource = rsGRID

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub txtPolicyNo_LostFocus()
'On Error GoTo Err

        Dim rsPOLICY As ADODB.Recordset
        Set rsPOLICY = New ADODB.Recordset
      
        rsPOLICY.Open "SELECT * FROM ALISMPolicy, ALISMReference WHERE ALISMPolicy.PolicyNo= '" & frmALISMOrders.txtPolicyNo.Text & "' and ALISMPolicy.ReferenceNo LIKE ALISMReference.ReferenceNo ; ", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsPOLICY
                If .EOF And .BOF Then Exit Sub
                frmALISMOrders.txtPolicyNo.Text = !PolicyNo
                frmALISMOrders.txtNames.Text = Trim(!Othernames) + " " + Trim(!Surname)
                frmALISMOrders.txtexpectedpremium.Text = !ExpectedPremium
                frmALISMOrders.txtPremiumcount.Text = !UnitCount
                frmALISMOrders.cboPaymentMethod.Text = !PaymentMethod
                frmALISMOrders.txtSurrenderValue.Text = !SurrenderValue & ""
                frmALISMOrders.txtOrderAmount.Text = !ExpectedPremium
                frmALISMOrders.txtDateofCommencement.Text = !DateofCommencement
                frmALISMOrders.txtTermofPolicy.Text = !TermOfPolicy
                frmALISMOrders.txtMaturityDate.Text = !MaturityDate
                frmALISMOrders.txtDateofFirstPayment.Text = !DateofCommencement
                frmALISMOrders.txtAccountNo.Text = !AccountNo & ""
                frmALISMOrders.txtIssuedBy.Text = UserName
                frmALISMOrders.cboPaymentMode.Text = !PaymentMode
                frmALISMOrders.cboBankName.Text = !BankNo & ""
                frmALISMOrders.txtPreparedBy.Text = UserName
                frmALISMOrders.txtDatePrepared.Text = Date
                frmALISMOrders.txtOrderDate.Text = Date
       End With
loadPAYMENTMETHOD
loadPaymentModeDESCRIPTION
loadBANK
LoadGrid
loadMGRID

rsPOLICY.Close

Exit Sub

err:
        ErrorMessage

End Sub
Private Sub DisableCommandButtons()
On Error GoTo err

    With frmALISMOrders
        .cmdUpdate.Enabled = True
        .cmdAddNew.Enabled = False
        .cmdSearch.Enabled = False
        .cmdCancel.Enabled = True
        .cmdSearchPolicy.Enabled = False
    End With
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub EnableCommandButtons()
On Error GoTo err

    With frmALISMOrders
        .cmdUpdate.Enabled = False
        .cmdAddNew.Enabled = True
        .cmdSearch.Enabled = True
        .cmdCancel.Enabled = True
        .cmdSearchPolicy.Enabled = True
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cboPaymentMode_gotFocus()
On Error GoTo err

    Dim rsPAYMENTMODEGF As ADODB.Recordset, strPAYMENTMODEgf As String
    Set rsPAYMENTMODEGF = New Recordset
    
    strPAYMENTMODEgf = "SELECT * FROM ALISPPaymentMode;"
    rsPAYMENTMODEGF.Open strPAYMENTMODEgf, cnALIS, adOpenKeyset, adLockOptimistic
    
    frmALISMOrders.cboPaymentMode.Clear

    With rsPAYMENTMODEGF
            Do Until .EOF
                    frmALISMOrders.cboPaymentMode.AddItem !Description
                    .MoveNext
            Loop
    
    End With
        
rsPAYMENTMODEGF.Close
strPAYMENTMODEgf = ""

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cboPaymentMode_KeyPress(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub

Private Sub cboPaymentMode_LostFocus()
On Error GoTo err

        Dim rsPAYMENTMODELF As ADODB.Recordset
        Set rsPAYMENTMODELF = New Recordset
        
        rsPAYMENTMODELF.Open "SELECT * FROM ALISPPaymentMode WHERE Description = '" & frmALISMOrders.cboPaymentMode.Text & "'", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsPAYMENTMODELF
                If .EOF And .BOF Then Exit Sub
                frmALISMOrders.cboPaymentMode.Text = !PaymentMode
                frmALISMOrders.txtPaymentModeDescription.Text = !Description
        End With
  
rsPAYMENTMODELF.Close

Exit Sub

err:
        ErrorMessage

End Sub


Private Sub loadPaymentModeDESCRIPTION()
On Error GoTo err

        Dim rsPAYMENTMODELF As ADODB.Recordset
        Set rsPAYMENTMODELF = New Recordset
        
        rsPAYMENTMODELF.Open "SELECT * FROM ALISPPaymentMode WHERE PaymentMode = '" & frmALISMOrders.cboPaymentMode.Text & "'", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsPAYMENTMODELF
                If .EOF And .BOF Then Exit Sub
                frmALISMOrders.txtPaymentModeDescription.Text = !Description
        End With
  
rsPAYMENTMODELF.Close

Exit Sub

err:
        ErrorMessage

End Sub

Private Sub cboPaymentMethod_GotFocus()
On Error GoTo err

    Dim rsPAYMENTMETHODGF As ADODB.Recordset, strPAYMENTMETHODgf As String
    Set rsPAYMENTMETHODGF = New Recordset
    
    strPAYMENTMETHODgf = "SELECT * FROM ALISPPaymentMethod;"
    rsPAYMENTMETHODGF.Open strPAYMENTMETHODgf, cnALIS, adOpenKeyset, adLockOptimistic
    
    frmALISMOrders.cboPaymentMethod.Clear

    With rsPAYMENTMETHODGF
            Do Until .EOF
                    frmALISMOrders.cboPaymentMethod.AddItem !PaymentMethodDescription
                    .MoveNext
            Loop
    
    End With
        
rsPAYMENTMETHODGF.Close
strPAYMENTMETHODgf = ""

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub cboPaymentMethod_LostFocus()
On Error GoTo err

        Dim rsPAYMENTMETHODLF As ADODB.Recordset, strPAYMENTMETHODlf As String
        Set rsPAYMENTMETHODLF = New Recordset
        
        rsPAYMENTMETHODLF.Open "SELECT * FROM ALISPPaymentMethod WHERE PaymentMethodDescription= '" & frmALISMOrders.cboPaymentMethod.Text & "'", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsPAYMENTMETHODLF
                If .EOF And .BOF Then Exit Sub
                frmALISMOrders.cboPaymentMethod.Text = !PaymentMethod
                frmALISMOrders.txtPaymentMethodDescription.Text = !PaymentMethodDescription
         End With
                
rsPAYMENTMETHODLF.Close
strPAYMENTMETHODlf = ""
                
Exit Sub

err:
        ErrorMessage

End Sub

Private Sub loadPAYMENTMETHOD()
On Error GoTo err

        Dim rsPAYMENTMETHODLF As ADODB.Recordset, strPAYMENTMETHODlf As String
        Set rsPAYMENTMETHODLF = New Recordset
        
        rsPAYMENTMETHODLF.Open "SELECT * FROM ALISPPaymentMethod WHERE PaymentMethod = '" & frmALISMOrders.cboPaymentMethod.Text & "'", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsPAYMENTMETHODLF
                If .EOF And .BOF Then Exit Sub
                frmALISMOrders.txtPaymentMethodDescription.Text = !PaymentMethodDescription
        End With
                
rsPAYMENTMETHODLF.Close
strPAYMENTMETHODlf = ""
                
Exit Sub

err:
        ErrorMessage

End Sub

Private Sub cboBankName_GotFocus()
On Error GoTo err

        Dim rsBANKGF As ADODB.Recordset, strBANKGF As String
        Set rsBANKGF = New Recordset
        
        strBANKGF = "SELECT * FROM ALISPBank;"
        rsBANKGF.Open strBANKGF, cnALIS, adOpenKeyset, adLockOptimistic
        
        frmALISMOrders.cboBankName.Clear

        With rsBANKGF
            Do Until .EOF
            frmALISMOrders.cboBankName.AddItem !CompanyName
                    .MoveNext
            Loop
    
        End With

rsBANKGF.Close
strBANKGF = ""

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cboBankName_KeyPress(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub

Private Sub cboBankName_LostFocus()
On Error GoTo err

        Dim rsBANKLF As ADODB.Recordset, strBANKLF As String
        Set rsBANKLF = New Recordset
        
        rsBANKLF.Open "SELECT * FROM ALISPBank WHERE CompanyName= '" & frmALISMOrders.cboBankName.Text & "'", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsBANKLF
                If .EOF And .BOF Then Exit Sub
                frmALISMOrders.txtBankNo.Text = !BankNo
                frmALISMOrders.cboBankName.Text = !CompanyName
        End With
        
rsBANKLF.Close
strBANKLF = ""

Exit Sub

err:
        ErrorMessage

End Sub

Private Sub loadBANK()
On Error GoTo err

        Dim rsBANKLF As ADODB.Recordset, strBANKLF As String
        Set rsBANKLF = New Recordset
        
        rsBANKLF.Open "SELECT * FROM ALISPBank WHERE BankNo = '" & frmALISMOrders.txtBankNo.Text & "'", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsBANKLF
                If .EOF And .BOF Then Exit Sub
                 frmALISMOrders.cboBankName.Text = !CompanyName
        End With
        
rsBANKLF.Close
strBANKLF = ""

Exit Sub

err:
        ErrorMessage

End Sub


Private Sub loadCOYBANK()
On Error GoTo err

        Dim rsBANKLF As ADODB.Recordset, strBANKLF As String
        Set rsBANKLF = New Recordset
        
        rsBANKLF.Open "SELECT * FROM ALISPBank WHERE BankNo = '" & frmALISMOrders.cboCoyBankNO.Text & "'", cnALIS, adOpenKeyset, adLockOptimistic
        
        With rsBANKLF
                If .EOF And .BOF Then Exit Sub
                 frmALISMOrders.txtCoyBankName.Text = !CompanyName
        End With
        
rsBANKLF.Close
strBANKLF = ""

Exit Sub

err:
        ErrorMessage

End Sub



