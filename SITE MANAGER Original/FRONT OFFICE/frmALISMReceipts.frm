VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmALISMReceipts 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "[PAYMENT RECEIPTS PROCESSING]"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   1560
   ClientWidth     =   11880
   Icon            =   "frmALISMReceipts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   11880
   Begin VB.Frame Frame2 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11655
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
         TabIndex        =   64
         Top             =   2520
         Width           =   3375
         Begin MSComctlLib.ListView ListView3 
            Height          =   2175
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   3135
            _ExtentX        =   5530
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
         Left            =   5160
         TabIndex        =   62
         Top             =   6720
         Width           =   6255
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
         Left            =   6840
         TabIndex        =   58
         Top             =   2520
         Width           =   3495
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   60
            Top             =   2160
            Width           =   1695
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   1815
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
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
            Left            =   360
            TabIndex        =   61
            Top             =   2175
            Width           =   1335
         End
      End
      Begin VB.Frame Frame6 
         Height          =   2535
         Left            =   3600
         TabIndex        =   46
         Top             =   2520
         Width           =   3135
         Begin VB.ComboBox cboReceiptType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   720
            Sorted          =   -1  'True
            TabIndex        =   66
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox txtPaymentStatusDetails 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   1905
            Width           =   1500
         End
         Begin VB.TextBox txtTransactionAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1440
            TabIndex        =   50
            Top             =   1245
            Width           =   1500
         End
         Begin VB.TextBox txtReceivedTodate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   1575
            Width           =   1500
         End
         Begin VB.TextBox cboDocumentNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   1440
            TabIndex        =   48
            Top             =   600
            Width           =   1500
         End
         Begin VB.TextBox txtDateOfLastPayment 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   900
            Width           =   1500
         End
         Begin VB.Label Label29 
            Caption         =   "Type"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Status"
            Height          =   210
            Left            =   240
            TabIndex        =   56
            Top             =   1950
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Amount"
            Height          =   255
            Left            =   240
            TabIndex        =   55
            Top             =   1275
            Width           =   615
         End
         Begin VB.Label lblDocumentNo 
            Caption         =   "Document No"
            Height          =   255
            Left            =   240
            TabIndex        =   54
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lblTotalReceived 
            Caption         =   "Total Received"
            Height          =   255
            Left            =   240
            TabIndex        =   53
            Top             =   1605
            Width           =   1335
         End
         Begin VB.Label lblDateOfLastPayment 
            Caption         =   "Last Pay Date"
            Height          =   255
            Left            =   240
            TabIndex        =   52
            Top             =   930
            Width           =   1095
         End
      End
      Begin VB.Frame Frame7 
         Height          =   735
         Left            =   120
         TabIndex        =   42
         Top             =   1800
         Width           =   8295
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
            Height          =   315
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   240
            Width           =   1455
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
            Height          =   315
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   240
            Width           =   4695
         End
         Begin VB.Label Label13 
            Caption         =   "Expected Amt"
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   270
            Width           =   1095
         End
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
         TabIndex        =   35
         Top             =   120
         Width           =   3015
         Begin VB.TextBox txtAccounBalance 
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
            TabIndex        =   70
            Top             =   1800
            Width           =   1335
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
            TabIndex        =   68
            Top             =   240
            Width           =   1335
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
            TabIndex        =   38
            Top             =   600
            Width           =   1335
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
            TabIndex        =   37
            Top             =   960
            Width           =   1335
         End
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
            TabIndex        =   36
            Top             =   1425
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Balance"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   1815
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Account No"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   270
            Width           =   855
         End
         Begin VB.Label lblStatusCode 
            Caption         =   "Status Code"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   990
            Width           =   975
         End
         Begin VB.Label lblDueDate 
            Caption         =   "Due Date:"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   630
            Width           =   855
         End
         Begin VB.Label lblSuspenseAccount 
            Caption         =   "Suspense"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   1440
            Width           =   735
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Actions"
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
         Left            =   10440
         TabIndex        =   28
         Top             =   2520
         Width           =   1095
         Begin VB.CommandButton cmdEdit 
            Appearance      =   0  'Flat
            Caption         =   "&Listing"
            Height          =   315
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   1935
            Width           =   855
         End
         Begin VB.CommandButton cmdCancel 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Picture         =   "frmALISMReceipts.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   1305
            Width           =   855
         End
         Begin VB.CommandButton cmdSearch 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Picture         =   "frmALISMReceipts.frx":0544
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   990
            Width           =   855
         End
         Begin VB.CommandButton cmdUpdate 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Picture         =   "frmALISMReceipts.frx":0646
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   675
            Width           =   855
         End
         Begin VB.CommandButton cmdAddNew 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Picture         =   "frmALISMReceipts.frx":0748
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton cmdDelete 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Picture         =   "frmALISMReceipts.frx":084A
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   1620
            Width           =   855
         End
      End
      Begin VB.TextBox txtBalance 
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   6720
         Width           =   1815
      End
      Begin VB.Frame Frame15 
         Height          =   1695
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   8295
         Begin VB.ComboBox cboCurrencyCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   6000
            Locked          =   -1  'True
            Sorted          =   -1  'True
            TabIndex        =   16
            Top             =   505
            Width           =   735
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
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   890
            Width           =   975
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
            TabIndex        =   14
            Top             =   1320
            Width           =   855
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
            TabIndex        =   13
            Top             =   1320
            Width           =   1215
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
            TabIndex        =   12
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtReceiptAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   6720
            TabIndex        =   11
            Top             =   505
            Width           =   1455
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
            TabIndex        =   10
            Top             =   120
            Width           =   2175
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
            Left            =   3000
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   890
            Width           =   5175
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
            TabIndex        =   8
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtPayer 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   1560
            TabIndex        =   7
            Top             =   505
            Width           =   3615
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
            TabIndex        =   6
            Top             =   120
            Width           =   1455
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
            TabIndex        =   5
            Top             =   135
            Width           =   1215
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
            TabIndex        =   4
            Top             =   890
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "Entries"
            Height          =   255
            Left            =   6840
            TabIndex        =   26
            Top             =   1350
            Width           =   1095
         End
         Begin VB.Label lblPaymentStatus 
            Caption         =   "Status"
            Height          =   210
            Left            =   3240
            TabIndex        =   25
            Top             =   1372
            Width           =   615
         End
         Begin VB.Label lblReferenceNo 
            Caption         =   "Local?"
            Height          =   255
            Left            =   5400
            TabIndex        =   24
            Top             =   1350
            Width           =   735
         End
         Begin VB.Label lblReceiptAmount 
            Caption         =   "Amount"
            Height          =   255
            Left            =   5400
            TabIndex        =   23
            Top             =   535
            Width           =   615
         End
         Begin VB.Label lblCurrentPeriod 
            Caption         =   "Period"
            Height          =   210
            Left            =   5400
            TabIndex        =   22
            Top             =   210
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Bank"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   943
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "Cheque No"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   1350
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Received From"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   540
            Width           =   1095
         End
         Begin VB.Label Label27 
            Caption         =   "Receipt No"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   195
            Width           =   1215
         End
         Begin VB.Label lblReceiptDate 
            Caption         =   "Date"
            Height          =   255
            Left            =   3240
            TabIndex        =   17
            Top             =   195
            Width           =   615
         End
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
         TabIndex        =   1
         Top             =   5040
         Width           =   11415
         Begin MSComctlLib.ListView ListView1 
            Height          =   1215
            Left            =   120
            TabIndex        =   2
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
      Begin VB.Label Label12 
         Caption         =   "Remark"
         Height          =   255
         Left            =   4080
         TabIndex        =   63
         Top             =   6750
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Balance"
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   6750
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmALISMReceipts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsreceipt As clsReceipting1, rsReceiptDetails As clsReceiptDetails1
Dim bunloadFORM As Boolean
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

Private Sub cboCurrencyCode_KeyPress(KeyAscii As Integer)
        Set rsreceipt = New clsReceipting1
        rsreceipt.selectCURRENCYKEYPRESS (KeyAscii)
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
            showRECEIPTITEMS
            showALLPreviousRECEIPTS
            Set rsReceiptDetails = Nothing
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'''On Error GoTo err
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'''On Error GoTo err
        
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


Private Sub cmdAddNew_Click()
        
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


Private Sub cmdCancel_Click()
        addpen = False
        clearALLRECORD
        disableALLRECORD
        enableButtons
        With frmODASMReceipt
            .ListView1.ListItems.Clear
            .ListView2.ListItems.Clear
            .ListView3.ListItems.Clear
        End With
        baddRECORD = False
        
End Sub


Private Sub cmdDelete_Click()
    If Screen.ActiveForm.txtReceiptNo.Text <= "" Then
        MsgBox "Cannot Use this Form Directly, Load the Receipt on the First Tab", vbOKOnly
        
        Exit Sub
        Else: Load frmNewReceipt
        frmNewReceipt.Show 1, Me
    End If

End Sub

Private Sub cmdEdit_Click()
    If Screen.ActiveForm.txtReceiptNo.Text <= "" Then
        MsgBox "Cannot Use this Form Directly, Load the Receipt on the First Tab", vbOKOnly
        Exit Sub
    Else
        Load frmReceiptListing
        frmReceiptListing.Show 1, Me

    End If

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


Private Sub cmdUpdate_Click()
    Set rsreceipt = New clsReceipting1
    Set rsReceiptDetails = New clsReceiptDetails1

    rsreceipt.updateRECORD
    If bsaveRECORD = True Then
            rsReceiptDetails.processUPDATE
            enableButtons
            disableALLRECORD
            showRECEIPTITEMS
            showALLPreviousRECEIPTS
            frmODASMReceipt.FramePayee.Enabled = False
    End If
            
    Set rsReceiptDetails = Nothing
    Set rsreceipt = Nothing

End Sub


Private Sub Form_Load()
        OpenConnection
End Sub

Private Sub Form_Activate()
        Set rsreceipt = New clsReceipting1
            rsreceipt.lockRECEIPTNO
            rsreceipt.loadDEFAULTS
            rsreceipt.disableDATAENTRY
            rsreceipt.checkPaymentMethod
        Set rsreceipt = Nothing
        
        Set rsReceiptDetails = New clsReceiptDetails1
            rsReceiptDetails.clearRECORD
            enableButtons
            rsReceiptDetails.disableDATAENTRY
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
'''On Error GoTo err
    
    ListView3.SortKey = ColumnHeader.Index - 1
    ListView3.Sorted = True
    
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ListView3_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'''On Error GoTo err
        
        'If bsearchRECORD = True Or beditRECORD = True And baddRECORD = True Then
        'Else
                'Item.Checked = False
        'End If
        
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
                    
                    If .txtPayer.Text = Empty Then
                            .txtPayer.Text = Item.SubItems(1)
                    End If
                    
                    .cboDocumentNo.SetFocus
            End With
        
        
        End If
  
Exit Sub

err:
    ErrorMessage

End Sub
Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'''On Error GoTo err
    
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
    
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'''On Error GoTo err
        
        'If bsearchRECORD = True Or beditRECORD = True And baddRECORD = True Then
        'Else
                'Item.Checked = False
        'End If
        
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
            If Screen.ActiveForm.txtTransactionAmount.Text <= Empty Then Exit Sub
            Set rsReceiptDetails = New clsReceiptDetails1
            rsReceiptDetails.processReceipt
            Set rsReceiptDetails = Nothing
End Sub

