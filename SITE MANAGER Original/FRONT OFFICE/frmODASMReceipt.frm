VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmODASMReceipt 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receipt"
   ClientHeight    =   7860
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11670
   Icon            =   "frmODASMReceipt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   11670
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   7095
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   11655
      Begin VB.TextBox txtInvoiceTotal 
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
         Height          =   315
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   6720
         Width           =   1215
      End
      Begin VB.TextBox txtJobBriefTotal 
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
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   6720
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Caption         =   "Previous Receipts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   57
         Top             =   5040
         Width           =   11415
         Begin MSComctlLib.ListView ListView1 
            Height          =   1215
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   2143
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame15 
         Height          =   1695
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   8295
         Begin VB.ComboBox cboCurrencyCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   6000
            Locked          =   -1  'True
            Sorted          =   -1  'True
            TabIndex        =   63
            Top             =   520
            Width           =   735
         End
         Begin VB.TextBox txtReceiptAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   6840
            TabIndex        =   62
            Top             =   520
            Width           =   1335
         End
         Begin VB.TextBox txtPayer 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   1560
            TabIndex        =   61
            Top             =   520
            Width           =   3615
         End
         Begin VB.TextBox txtPaymentMethod 
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
            Height          =   315
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   920
            Width           =   735
         End
         Begin VB.TextBox txtReceiptDate 
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
            Height          =   315
            Left            =   3960
            TabIndex        =   47
            Top             =   165
            Width           =   1215
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
            Height          =   315
            Left            =   1560
            TabIndex        =   46
            Top             =   165
            Width           =   1455
         End
         Begin VB.TextBox txtChequeNo 
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
            Height          =   315
            Left            =   1560
            TabIndex        =   45
            Top             =   1320
            Width           =   1455
         End
         Begin VB.ComboBox cboBankNo 
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
            Height          =   315
            Left            =   3360
            Sorted          =   -1  'True
            TabIndex        =   44
            Top             =   920
            Width           =   4815
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
            Height          =   315
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   165
            Width           =   2175
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
            Height          =   315
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   1320
            Width           =   735
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
            Height          =   315
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtTransactionNo 
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
            Height          =   315
            Left            =   7320
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox txtBankNo 
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
            Height          =   315
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   920
            Width           =   975
         End
         Begin VB.Label lblReceiptAmount 
            Caption         =   "Amount"
            Height          =   255
            Left            =   5400
            TabIndex        =   65
            Top             =   550
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Received From"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   540
            Width           =   1095
         End
         Begin VB.Label lblReceiptDate 
            Caption         =   "Date"
            Height          =   255
            Left            =   3240
            TabIndex        =   56
            Top             =   195
            Width           =   615
         End
         Begin VB.Label Label27 
            Caption         =   "Receipt No"
            Height          =   255
            Left            =   240
            TabIndex        =   55
            Top             =   195
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Cheque No"
            Height          =   255
            Left            =   240
            TabIndex        =   54
            Top             =   1350
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Bank"
            Height          =   255
            Left            =   240
            TabIndex        =   53
            Top             =   943
            Width           =   495
         End
         Begin VB.Label lblCurrentPeriod 
            Caption         =   "Period"
            Height          =   210
            Left            =   5400
            TabIndex        =   52
            Top             =   217
            Width           =   495
         End
         Begin VB.Label lblReferenceNo 
            Caption         =   "Local?"
            Height          =   255
            Left            =   5400
            TabIndex        =   51
            Top             =   1350
            Width           =   735
         End
         Begin VB.Label lblPaymentStatus 
            Caption         =   "Status"
            Height          =   210
            Left            =   3240
            TabIndex        =   50
            Top             =   1372
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Entries"
            Height          =   255
            Left            =   6840
            TabIndex        =   49
            Top             =   1350
            Width           =   1095
         End
      End
      Begin VB.TextBox txtJobBriefNo 
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
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   6720
         Width           =   1215
      End
      Begin VB.Frame Frame8 
         Caption         =   "More Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   8520
         TabIndex        =   26
         Top             =   120
         Width           =   3015
         Begin VB.TextBox txtSuspenseAccount 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   1200
            TabIndex        =   31
            Top             =   1590
            Width           =   1575
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
            Height          =   315
            Left            =   1200
            TabIndex        =   30
            Top             =   1140
            Width           =   1575
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
            Height          =   315
            Left            =   1200
            TabIndex        =   29
            Top             =   690
            Width           =   1575
         End
         Begin VB.TextBox txtAccountNo 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   1200
            TabIndex        =   28
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtAccountBalance 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   1200
            TabIndex        =   27
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label lblSuspenseAccount 
            Caption         =   "Suspense"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1620
            Width           =   735
         End
         Begin VB.Label lblDueDate 
            Caption         =   "Due Date:"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblStatusCode 
            Caption         =   "Status Code"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1170
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Account No"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   270
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Balance"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   2055
            Width           =   735
         End
      End
      Begin VB.Frame Frame7 
         Height          =   735
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   8295
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
            Height          =   315
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   240
            Width           =   4695
         End
         Begin VB.TextBox txtExpectedAmount 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label13 
            Caption         =   "Expected Amt"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   270
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Height          =   2535
         Left            =   3720
         TabIndex        =   9
         Top             =   2520
         Width           =   3495
         Begin VB.TextBox txtDateOfLastPayment 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   1002
            Width           =   1500
         End
         Begin VB.TextBox cboDocumentNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1800
            TabIndex        =   14
            Top             =   621
            Width           =   1500
         End
         Begin VB.TextBox txtReceivedTodate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   1764
            Width           =   1500
         End
         Begin VB.TextBox txtTransactionAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1800
            TabIndex        =   12
            Top             =   1383
            Width           =   1500
         End
         Begin VB.TextBox txtPaymentStatusDetails 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   2145
            Width           =   1500
         End
         Begin VB.ComboBox cboReceiptType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   600
            Sorted          =   -1  'True
            TabIndex        =   10
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label lblDateOfLastPayment 
            Caption         =   "Last Pay Date"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1032
            Width           =   1095
         End
         Begin VB.Label lblTotalReceived 
            Caption         =   "Total Received"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1794
            Width           =   1335
         End
         Begin VB.Label lblDocumentNo 
            Caption         =   "Document No"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   651
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Amount"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1413
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Status"
            Height          =   210
            Left            =   120
            TabIndex        =   17
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label29 
            Caption         =   "Type"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Receipts Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   7320
         TabIndex        =   5
         Top             =   2520
         Width           =   4215
         Begin VB.TextBox txtTotalAmount 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   2160
            Width           =   1335
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   1815
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   3201
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin VB.Label Label11 
            Caption         =   "Total Amount"
            Height          =   255
            Left            =   1440
            TabIndex        =   8
            Top             =   2175
            Width           =   1335
         End
      End
      Begin VB.TextBox txtRemark 
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
         Height          =   315
         Left            =   7200
         TabIndex        =   4
         Top             =   6720
         Width           =   4215
      End
      Begin VB.Frame FramePayee 
         Caption         =   "Payees"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         TabIndex        =   2
         Top             =   2520
         Width           =   3495
         Begin MSComctlLib.ListView ListView3 
            Height          =   2175
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   3836
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Label Label14 
         Caption         =   "INV Total"
         Height          =   255
         Left            =   4440
         TabIndex        =   69
         Top             =   6750
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "JB Total"
         Height          =   255
         Left            =   2280
         TabIndex        =   67
         Top             =   6750
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "JobBriefNo"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   6750
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Remark"
         Height          =   255
         Left            =   6600
         TabIndex        =   59
         Top             =   6750
         Width           =   855
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5880
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMReceipt.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMReceipt.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMReceipt.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMReceipt.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMReceipt.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMReceipt.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMReceipt.frx":236E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   1164
      ButtonWidth     =   3069
      ButtonHeight    =   1005
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New &Record "
            Key             =   "N"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Edit/Change "
            Key             =   "E"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Search/Find "
            Key             =   "S"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Refresh "
            Key             =   "R"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print &Preview "
            Key             =   "P"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Help"
            Key             =   "F"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComDlg.CommonDialog HelpCommonDialog 
         Left            =   10800
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         HelpCommand     =   11
         HelpContext     =   72
         HelpFile        =   "REGHELP.HLP"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClear 
         Caption         =   "Clear the &Screen"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnumm 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuShow 
      Caption         =   "&Show/View"
      Begin VB.Menu mnuClosedJobs 
         Caption         =   "Closed Jobs"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuKJHGFDGFVHJ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFullInventory 
         Caption         =   "Full Inventory"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help System"
      Begin VB.Menu mnuHow 
         Caption         =   "How to use this System"
         Shortcut        =   ^{F1}
      End
   End
End
Attribute VB_Name = "frmODASMReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsreceipt As clsReceipting1, rsReceiptDetails As clsReceiptDetails1
Dim bunloadFORM As Boolean
Public rsDEP As ADODB.Recordset, strDEP As String
Private Sub cboReceiptType_Click()
    Me.cboDocumentNo.SetFocus
End Sub

Private Sub Form_Unload(cancel As Integer)
    If addpen = True Then
        cancel = True
        MsgBox "Data entry in progress. Click Refresh to Cancel", vbCritical
    Else
        cancel = False
    End If
End Sub

Private Sub ClearReceipt()
    Set rsreceipt = New clsReceipting1
    rsreceipt.clearRECORD
    Set rsreceipt = Nothing
End Sub

Private Sub disableControls()
        Set rsreceipt = New clsReceipting1
        Dim bVAL As Boolean
        rsreceipt.disableDATAENTRY
        Set rsreceipt = Nothing
End Sub

Private Sub UpdatePremium()

    Set rsreceipt = New clsReceipting1
    rsreceipt.updateRECORD
    Set rsreceipt = Nothing

End Sub
Private Sub cbobankNo_GotFocus()
    Set rsreceipt = New clsReceipting1
    rsreceipt.selectBankNOGotFocus
    Set rsreceipt = Nothing
    
End Sub
Private Sub cboBankNo_KeyPress(KeyAscii As Integer)

        Set rsreceipt = New clsReceipting1
        rsreceipt.selectBankNoKeyPress (KeyAscii)
        Set rsreceipt = Nothing

End Sub
Private Sub cbobankNo_LostFocus()
        Set rsreceipt = New clsReceipting1
        rsreceipt.selectBankNoLostFocus
        Set rsreceipt = Nothing
End Sub

Private Sub cboCurrencyCode_GotFocus()
        Set rsreceipt = New clsReceipting1
        rsreceipt.selectCURRENCYGOTFOCUS
        Set rsreceipt = Nothing
End Sub
Private Sub cboCurrencyCode_LostFocus()
        Set rsreceipt = New clsReceipting1
        rsreceipt.selectCURRENCYLOSTFOCUS
        Set rsreceipt = Nothing
End Sub

Private Sub cboDocumentNo_LostFocus()
    Set rsReceiptDetails = New clsReceiptDetails1
    rsReceiptDetails.processDOCUMENTNO
'    Me.txtTransactionAmount.SetFocus
    showRECEIPTITEMS
    showALLPreviousRECEIPTS
    Set rsReceiptDetails = Nothing
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        
        If bsearchRECORD <> True And beditRECORD <> True Then
                Item.Checked = False
                Exit Sub
        End If
        Dim i, j As Double
        
        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView1.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                    If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                                Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                    End If
            Next i
            

            With frmODASMReceipt
                    If .ListView1.Checkboxes = False Then
                            Exit Sub
                    End If

                    frmODASMReceipt.txtReceiptNo = Item.Text
                    frmODASMReceipt.txtTransactionNo = Item.SubItems(1)
            End With
        
        
        End If
  
Exit Sub

err:
    ErrorMessage

End Sub

Private Sub cboReceiptType_GotFocus()
        Set rsReceiptDetails = New clsReceiptDetails1
        rsReceiptDetails.selectRECEIPTTYPEGOTFOCUS
        frmODASMReceipt.FramePayee.Enabled = True
        Set rsReceiptDetails = Nothing
End Sub

Private Sub cboReceiptType_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub cboReceiptType_LostFocus()
        If Screen.ActiveForm.cboReceiptType.Text <= "" Then Exit Sub
        Set rsReceiptDetails = New clsReceiptDetails1
        rsReceiptDetails.selectRECEIPTTYPELOSTFOCUS
        Set rsReceiptDetails = Nothing
        
End Sub


Private Sub AddNewRecord()
        
        Set rsreceipt = New clsReceipting1
        rsreceipt.lockRECEIPTNO
        rsreceipt.loadDEFAULTS
        Set rsreceipt = Nothing
        
        If Screen.ActiveForm.txtPaymentMethod.Text = Empty Then
                MsgBox "Close the Form and select the Payment Method relaunching you Process"
        End If
        
        baddRECORD = True
        bloadINVOICE = False
        Set rsReceiptDetails = New clsReceiptDetails1
        rsReceiptDetails.addRECORD
        frmODASMReceipt.FramePayee.Enabled = False
        Set rsReceiptDetails = Nothing

End Sub


Private Sub cancelALL()
        addpen = False
        clearALLRECORD
        disableALLRECORD
        With frmODASMReceipt
            .ListView1.ListItems.Clear
            .ListView2.ListItems.Clear
            .ListView3.ListItems.Clear
        End With
        baddRECORD = False
        NewRecord = False
        
End Sub



Private Sub cmdSearch_Click()
    Set rsreceipt = New clsReceipting1
        rsreceipt.searchRECORD
        If bsearchRECORD = True Then
            rsreceipt.loadRECEIPTDETAILS
            rsreceipt.loadEMPLOYER
            showRECEIPTITEMS
            showALLPreviousRECEIPTS
        End If
    Set rsreceipt = Nothing

End Sub


Private Sub SaveNewRecord()
    Set rsreceipt = New clsReceipting1
    Set rsReceiptDetails = New clsReceiptDetails1

    rsreceipt.updateRECORD
    If bsaveRECORD = True Then
            rsReceiptDetails.processUPDATE
            disableALLRECORD
            showRECEIPTITEMS
            showALLPreviousRECEIPTS
            frmODASMReceipt.FramePayee.Enabled = False
            NewRecord = False
            beditRECORD = False
    End If
            
    Set rsReceiptDetails = Nothing
    Set rsreceipt = Nothing

End Sub


Private Sub Form_Load()
        OpenConnection
End Sub

Private Sub Form_Activate()
        disableALLRECORD

        Set rsreceipt = New clsReceipting1
            rsreceipt.lockRECEIPTNO
            rsreceipt.loadDEFAULTS
            rsreceipt.checkPaymentMethod
        Set rsreceipt = Nothing
        
        Set rsReceiptDetails = New clsReceiptDetails1
            rsReceiptDetails.clearRECORD
        Set rsReceiptDetails = Nothing
        showRECEIPTITEMS
End Sub

Private Sub txtAccountNo_LostFocus()
        Set rsReceiptDetails = New clsReceiptDetails1
            rsReceiptDetails.selectEmployerLostFocus
        Set rsReceiptDetails = Nothing
End Sub

Public Sub generateRECEIPTNo()
            Set rsReceiptDetails = New clsReceipting1
            rsReceiptDetails.createRECEIPT
            Set rsReceiptDetails = Nothing
End Sub





Private Sub ListView3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    
    ListView3.SortKey = ColumnHeader.Index - 1
    ListView3.Sorted = True
    
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView3_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        
    Dim i, j As Double
    
    If Item.Checked = True Then
        
        j = Screen.ActiveForm.ListView3.ListItems.Count
        
        If j = 0 Then Exit Sub
        
        For i = 1 To j
            If Screen.ActiveForm.ListView3.ListItems(i) <> Item Then
                Screen.ActiveForm.ListView3.ListItems(i).Checked = False
            End If
        Next i
        
        With frmODASMReceipt
            If .ListView3.Checkboxes = False Then
                    Exit Sub
            End If

            .cboDocumentNo = Item.Text
            
'            If .txtPayer.Text = Empty Then
                .txtPayer.Text = Item.SubItems(1)
'            End If
            
            .cboDocumentNo.SetFocus
        End With
    
    
    End If
  
Exit Sub

err:
    ErrorMessage

End Sub
Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
    
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
    Dim i, j As Double
    
    If Item.Checked = True Then
        
        j = Screen.ActiveForm.ListView2.ListItems.Count
        
        If j = 0 Then Exit Sub
        
        For i = 1 To j
                If Screen.ActiveForm.ListView2.ListItems(i) <> Item Then
                            Screen.ActiveForm.ListView2.ListItems(i).Checked = False
                End If
        Next i
        

        With frmODASMReceipt
                If .ListView2.Checkboxes = False Then
                        Exit Sub
                End If

                .cboDocumentNo = Item.Text
                .txtTransactionNo = Item.SubItems(1)
                
                Set rsReceiptDetails = New clsReceiptDetails1
                    rsReceiptDetails.searchRECORD
                Set rsReceiptDetails = Nothing
        
        End With
    
    
    End If

Exit Sub

err:
    ErrorMessage

End Sub

Private Sub txtTransactionAmount_LostFocus()
    frmODASMReceipt.txtTransactionAmount.Text = FormatNumber(frmODASMReceipt.txtTransactionAmount)
    If frmODASMReceipt.txtTransactionAmount.Text <= Empty Then Exit Sub
    Set rsReceiptDetails = New clsReceiptDetails1
    
    Me.txtAccountBalance.Text = CDbl(Me.txtAccountBalance.Text) - CDbl(Me.txtTransactionAmount)
    rsReceiptDetails.processReceipt
    Set rsReceiptDetails = Nothing
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
        
    With frmODASMReceipt
    
    Set rsreceipt = New clsReceipting1
    Set rsReceiptDetails = New clsReceiptDetails1

    Select Case Button.Key
            Case "N"
                Select Case Button.Caption
                
                Case "New &Record "
                        If editRECORD Then Exit Sub
                        NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4:
                        AddNewRecord
                Case "&Save Record "
                
                        bsaveRECORD = False
                        SaveNewRecord
                                
                        If bsaveRECORD = True Then
                            bsaveRECORD = False
                            Button.Caption = "&NEXT RECEIPT": Button.Image = 2
                            .Toolbar1.Buttons(4).Caption = "FINISH"
                              disableALLRECORD
                        End If
                
                Case "&NEXT RECEIPT"
                        NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4:
                        AddNewRecord
                Case Else
                    Exit Sub
                End Select

    Case "E"
        Select Case Button.Caption
            Case "&Edit/Change "
            
            Case "&Save Record "

                    bsaveRECORD = False
                    rsreceipt.validateRECORD
                    
                    If bsaveRECORD = True Then
                            rsreceipt.updateRECORD
                            bsaveRECORD = False
                            .Toolbar1.Buttons(2).Caption = "New &Record "
                            .Toolbar1.Buttons(3).Caption = "&NEXT RECEIPT"
                            .Toolbar1.Buttons(4).Caption = "FINISH"
                            disableALLRECORD
                    End If
            
            Case Else
        End Select
    
    Case "S"
        Select Case Button.Caption
            Case "FINISH"
                .Toolbar1.Buttons(2).Caption = "New &Record "
                .Toolbar1.Buttons(2).Image = 2
                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                .Toolbar1.Buttons(3).Image = 5
                NewRecord = False: editRECORD = False: clearALLRECORD
            End Select
    Case "R"
        cancelALL
        
        If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
            .Toolbar1.Buttons(2).Caption = "New &Record "
            .Toolbar1.Buttons(2).Image = 2
            .Toolbar1.Buttons(3).Caption = "&Edit/Change "
            .Toolbar1.Buttons(3).Image = 5
            NewRecord = False: editRECORD = False: clearALLRECORD
    Case "P"
        If Screen.ActiveForm.txtReceiptNo.Text <= "" Then
            MsgBox "Cannot Use this Form Directly, Load the Receipt on the First Tab", vbOKOnly
            Exit Sub
            Else: Load frmNewReceipt
            frmNewReceipt.Show 1, Me
        End If

    Case "F"
        Me.HelpCommonDialog.DialogTitle = "Using the Main System"
        Me.HelpCommonDialog.HelpFile = App.HelpFile
        Me.HelpCommonDialog.HelpContext = 71
        Me.HelpCommonDialog.HelpCommand = cdlHelpContext
        Me.HelpCommonDialog.ShowHelp

 
    Case Else
        Exit Sub
    End Select
    
    Set rsreceipt = Nothing
    Set rsReceiptDetails = Nothing

End With
Exit Sub
err:
    ErrorMessage

End Sub

