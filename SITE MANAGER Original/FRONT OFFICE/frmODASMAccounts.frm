VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmODASMAccounts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process Invoice"
   ClientHeight    =   7920
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11565
   Icon            =   "frmODASMAccounts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      Caption         =   "Invoice Items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1215
      Left            =   5400
      TabIndex        =   35
      Top             =   6600
      Width           =   6135
      Begin MSComctlLib.ListView ListView7 
         Height          =   855
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1508
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
   Begin VB.Frame FrameSites 
      Caption         =   "Invoice Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1695
      Left            =   5400
      TabIndex        =   29
      Top             =   4920
      Width           =   6135
      Begin VB.TextBox txtInstallmentNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1440
         TabIndex        =   64
         Text            =   " "
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox cboVATRate 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   4320
         TabIndex        =   61
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtInvoiceReference 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1440
         TabIndex        =   50
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtDueDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1440
         TabIndex        =   48
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtJBPriceInclusive 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4320
         TabIndex        =   45
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtJBVATAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4320
         TabIndex        =   44
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtJBPriceExclusive 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4320
         TabIndex        =   32
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtPaymentMethod 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1440
         TabIndex        =   31
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtItemNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1680
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Installment #"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "VAT Rate"
         Height          =   255
         Left            =   3000
         TabIndex        =   62
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "VAT"
         Height          =   255
         Left            =   3000
         TabIndex        =   52
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Reference"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Due Date"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   " Amount Incl"
         Height          =   255
         Left            =   3000
         TabIndex        =   46
         Top             =   1350
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "Form of Payment"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label Label27 
         Caption         =   "Amount Excl"
         Height          =   255
         Left            =   3000
         TabIndex        =   33
         Top             =   270
         Width           =   1215
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Related Invoices"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1695
      Left            =   120
      TabIndex        =   27
      Top             =   6120
      Width           =   5175
      Begin MSComctlLib.ListView ListView6 
         Height          =   1335
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2355
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
   Begin VB.Frame Frame8 
      Caption         =   "Invoice Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1935
      Left            =   5400
      TabIndex        =   16
      Top             =   1320
      Width           =   6135
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   4320
         MaxLength       =   120
         TabIndex        =   42
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtInvoiceDescription 
         BackColor       =   &H00FFC0C0&
         Height          =   555
         Left            =   1440
         MaxLength       =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox txtPriceInclusive 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4320
         TabIndex        =   25
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtVATAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1440
         TabIndex        =   23
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtPriceExclusive 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1440
         TabIndex        =   21
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtInvoiceDate 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4320
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtInvoiceNo 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label30 
         Caption         =   "Remark"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   3000
         TabIndex        =   43
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Description"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Amount Incl"
         Height          =   255
         Left            =   3000
         TabIndex        =   26
         Top             =   1230
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "VAT Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "InVoice Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1230
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Invoice Date"
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "InVoice No"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Invoice Sent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2055
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   5175
      Begin MSComctlLib.ListView ListView1 
         Height          =   1695
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2990
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
   Begin VB.Frame Frame6 
      Caption         =   "Contract Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1695
      Left            =   5400
      TabIndex        =   8
      Top             =   3240
      Width           =   6135
      Begin VB.TextBox txtJobBriefNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1440
         TabIndex        =   59
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtDeposit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4320
         TabIndex        =   57
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtAmountQuoted 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1440
         TabIndex        =   54
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtTotalCost 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4320
         TabIndex        =   53
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtDescriptionOfOrder 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   2880
         MaxLength       =   120
         TabIndex        =   38
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtProductCode 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   37
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtCommencementDate 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtExpiryDate 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4320
         TabIndex        =   11
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   " Brief No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "Deposit "
         Height          =   255
         Left            =   3120
         TabIndex        =   58
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Total Cost"
         Height          =   255
         Left            =   3120
         TabIndex        =   56
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Total Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Product"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Comm Date"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "Expiry Date"
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Invoices Received"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   5175
      Begin MSComctlLib.ListView ListView3 
         Height          =   1455
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2566
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
   Begin VB.Frame Frame2 
      Caption         =   "Client Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   5400
      TabIndex        =   3
      Top             =   720
      Width           =   6015
      Begin VB.TextBox txtCurrentPeriod 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4920
         TabIndex        =   47
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtAccountNo 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   720
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtCompanyName 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   " Name"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Receipts for Job Brief"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5175
      Begin MSComctlLib.ListView ListView2 
         Height          =   1215
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2143
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7080
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
            Picture         =   "frmODASMAccounts.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAccounts.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAccounts.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAccounts.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAccounts.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAccounts.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMAccounts.frx":236E
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
      Width           =   11565
      _ExtentX        =   20399
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
      Begin VB.TextBox txtNoOfItems 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   11040
         TabIndex        =   65
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComDlg.CommonDialog HelpCommonDialog 
         Left            =   10560
         Top             =   0
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
      Begin VB.Menu mnuRegisteredClients 
         Caption         =   "Registered Clients"
      End
      Begin VB.Menu mnuKHJGGFDHJ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowQuotations 
         Caption         =   "Show Quotations"
      End
      Begin VB.Menu mnuExtraInfo 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExtraInformation 
         Caption         =   "Extra Inform"
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
Attribute VB_Name = "frmODASMAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsINVOICE As clsODASMAccounts

Private Sub cboVATRate_Change()
  With Me
     calculateVAT
  End With
End Sub

Private Sub cboVATRate_GotFocus()
        selectVATRATE_GotFocus
End Sub

Private Sub cboVATRate_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub cboVATRate_LostFocus()
        selectVATRate_LostFocus
End Sub

Private Sub Form_Activate()
        
        Set rsINVOICE = New clsODASMAccounts
        rsINVOICE.loadRECORD
        rsINVOICE.LoadDEFAULT
        rsINVOICE.loadINSTALLMENT
        rsINVOICE.calculateVAT
        rsINVOICE.updatePRICEExclusive
        rsINVOICE.updateVAT
        rsINVOICE.updatePRICEinclusive
        CalculateNoOfItems
        disableALLRECORD
        
        showBRIEFRECEIPTS
        showBRIEFINVOICESsenT
        showINVOICEDETAILS
        rsINVOICE.calculateTOTALRECEIPTS
        rsINVOICE.calculateInvoicesSend
        Set rsINVOICE = Nothing
        showBRIEFINACCOUNT
        showINVOICEitems
    Set rsINVOICE = Nothing
End Sub
Private Sub Form_Initialize()
        Set rsINVOICE = New clsODASMAccounts
End Sub
Private Sub Form_Unload(cancel As Integer)
        showRECEIPTSCHEDULE
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
        Dim i, j As Double
        
        'If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then

        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView1.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView1.ListItems(i).Checked = False
                End If
            Next i
            
            frmODASMAccounts.txtInvoiceNo.Text = Item.Text

        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub Form_Load()
        OpenConnection
End Sub

Private Sub Form_Terminate()
       Set rsINVOICE = Nothing
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
        
        'If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then

        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView2.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView2.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView2.ListItems(i).Checked = False
                End If
            Next i
            
            
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
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
        
        'If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then

        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView3.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView3.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView3.ListItems(i).Checked = False
                End If
            Next i
            
            
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ListView4_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView4.SortKey = ColumnHeader.Index - 1
    ListView4.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView4_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        
        'If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then

        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView4.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView4.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView4.ListItems(i).Checked = False
                End If
            Next i
            
            
        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub ListView5_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    ListView5.SortKey = ColumnHeader.Index - 1
    ListView5.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView5_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo err
        Dim i, j As Double
        
        'If Item.Checked = True And baddRECORD = True Or beditRECORD = True Or bsearchRECORD = True Then

        If Item.Checked = True Then
            
            j = Screen.ActiveForm.ListView5.ListItems.Count
            
            If j = 0 Then Exit Sub
            
            For i = 1 To j
                If Screen.ActiveForm.ListView5.ListItems(i) <> Item Then
                   Screen.ActiveForm.ListView5.ListItems(i).Checked = False
                End If
            Next i
            
            frmODASMAccounts.txtPaymentMethod.Text = Item.Text

        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub mnuExtraInformation_Click()
        Load frmODASMInformation
        frmODASMInformation.Show 1, Me
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err

Set rsINVOICE = New clsODASMAccounts
   
 With frmODASMAccounts
    Select Case Button.Key
    Case "N"
    
                Select Case Button.Caption
                    Case "New &Record "
                        If editRECORD Then Exit Sub
                        enableALLRECORD
                        rsINVOICE.enableRECORD
                        NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4
                        
                    Case "&Save Record "
                        If NewRecord Then
                            rsINVOICE.updateRECORD
                            disableALLRECORD
                            Button.Caption = "NE&XT ITEM": Button.Image = 2

                            .Toolbar1.Buttons(3).Caption = "FINISH"
                                
                        End If
                        
                     Case "NE&XT ITEM"
                            Button.Caption = "&Save Record ": Button.Image = 4
                            
                            rsINVOICE.clearRECORD
                            rsINVOICE.enableRECORD

                            NewRecord = False
                    Case Else
                            Exit Sub
                End Select
    
      Case "E"
        Select Case Button.Caption
            Case "FINISH"
                .Toolbar1.Buttons(2).Caption = "New &Record "
                .Toolbar1.Buttons(2).Image = 2
                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                .Toolbar1.Buttons(3).Image = 5
                NewRecord = False: editRECORD = False:
            End Select
      Case "S"
                
        Case "R"
            If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                .Toolbar1.Buttons(2).Caption = "New &Record "
                .Toolbar1.Buttons(2).Image = 2
                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                .Toolbar1.Buttons(3).Image = 5
               NewRecord = False: editRECORD = False:
        
        Case "P"
                CurrentRecord = .txtInvoiceNo
                Load frmODASRInvoice
                frmODASRInvoice.Show 1, Me
        Case "F"
            Me.HelpCommonDialog.DialogTitle = "Using the Main System"
            Me.HelpCommonDialog.HelpFile = App.HelpFile
            Me.HelpCommonDialog.HelpContext = 71
            Me.HelpCommonDialog.HelpCommand = cdlHelpContext
            Me.HelpCommonDialog.ShowHelp

        Case Else
            Exit Sub
        End Select

    End With
Set rsINVOICE = Nothing

Exit Sub
err:
    ErrorMessage

End Sub
Public Sub calculateVAT()
On Error GoTo err
        With frmODASMAccounts
                .txtJBVATAmount.Text = FormatNumber(CDbl(.cboVATRate.Text) / 100 * CDbl(.txtJBPriceExclusive))
                .txtJBPriceInclusive.Text = FormatNumber(CDbl(.txtJBVATAmount.Text) + CDbl(.txtJBPriceExclusive.Text))
        End With
        
Exit Sub
err:
    ErrorMessage
End Sub
Public Sub CalculateNoOfItems()
On Error GoTo err

    Set rsSAVE = New ADODB.Recordset
    
    strSQL = "SELECT * FROM ODASMJobBriefItems WHERE JobBriefNo = '" & frmODASMAccounts.txtJobBriefNo.Text & "'"
    rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
    With rsSAVE
            Me.txtNoOfItems.Text = .RecordCount
            
    End With
        
rsSAVE.Close
strTRANS = ""

Exit Sub
err:
ErrorMessage
End Sub
