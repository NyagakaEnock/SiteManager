VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmODASMOpenJobCard 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7515
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "frmODASMOpenJobCard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtReqPriceExclusive 
      Alignment       =   2  'Center
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
      Left            =   3600
      TabIndex        =   83
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox txtItemNo 
      Alignment       =   2  'Center
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
      Left            =   1200
      TabIndex        =   81
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox txtTotalVATAmount 
      Alignment       =   2  'Center
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
      Left            =   1200
      TabIndex        =   78
      Top             =   7065
      Width           =   1215
   End
   Begin VB.TextBox txtTotalPriceExcl 
      Alignment       =   2  'Center
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
      Left            =   3600
      TabIndex        =   77
      Top             =   7065
      Width           =   1455
   End
   Begin VB.Frame Frame6 
      Caption         =   "List of Suppliers"
      Height          =   1575
      Left            =   5280
      TabIndex        =   67
      Top             =   2280
      Width           =   6495
      Begin MSComctlLib.ListView ListView3 
         Height          =   1215
         Left            =   120
         TabIndex        =   68
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   2143
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
   Begin VB.Frame Cumulative 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   5280
      TabIndex        =   47
      Top             =   6360
      Width           =   6495
      Begin VB.TextBox txtRequisitionApproved 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   2520
         TabIndex        =   52
         Top             =   375
         Width           =   1575
      End
      Begin VB.TextBox txtQuantityApproved 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   2520
         TabIndex        =   51
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtRequisitionPrepared 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   960
         TabIndex        =   49
         Top             =   375
         Width           =   1575
      End
      Begin VB.TextBox txtQuantityPrepared 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   960
         TabIndex        =   48
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Cost"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "Approval"
         Height          =   255
         Left            =   2640
         TabIndex        =   53
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Order"
         Height          =   255
         Left            =   1440
         TabIndex        =   50
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.TextBox txtRemarks 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   42
      Top             =   6720
      Width           =   3855
   End
   Begin VB.Frame Frame4 
      Caption         =   "Requisitions"
      Height          =   2535
      Left            =   5280
      TabIndex        =   26
      Top             =   3840
      Width           =   6495
      Begin VB.TextBox txtUnitCode 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         TabIndex        =   76
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtTotalUnitPriceExcl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4200
         TabIndex        =   73
         Top             =   1520
         Width           =   1695
      End
      Begin VB.TextBox txtVATRate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   5400
         TabIndex        =   70
         Top             =   880
         Width           =   375
      End
      Begin VB.TextBox txtCostCenter 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4200
         TabIndex        =   69
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtItemQuantity 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   4200
         TabIndex        =   64
         Top             =   880
         Width           =   495
      End
      Begin VB.ComboBox cboCurrencyCode 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1320
         TabIndex        =   62
         Top             =   1500
         Width           =   1215
      End
      Begin VB.TextBox txtAccountNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         TabIndex        =   60
         Top             =   870
         Width           =   1215
      End
      Begin VB.TextBox txtExchangeRate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1320
         TabIndex        =   46
         Top             =   1845
         Width           =   1215
      End
      Begin VB.TextBox txtRequisitionDate 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4200
         TabIndex        =   41
         Top             =   560
         Width           =   1695
      End
      Begin VB.TextBox txtRequisitionNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         TabIndex        =   40
         Top             =   555
         Width           =   1215
      End
      Begin VB.TextBox txtTotalUnitPriceIncl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4200
         TabIndex        =   35
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtVATAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4200
         TabIndex        =   33
         Top             =   1840
         Width           =   1695
      End
      Begin VB.TextBox txtUnitPrice 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   4200
         TabIndex        =   31
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtItemSize 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1320
         TabIndex        =   29
         Top             =   1185
         Width           =   1215
      End
      Begin VB.TextBox txtProductCode 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin MSComCtl2.UpDown UpDownQuantity 
         Height          =   255
         Left            =   4680
         TabIndex        =   63
         Top             =   895
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Max             =   1000
         Enabled         =   -1  'True
      End
      Begin VB.Label Label29 
         Caption         =   "Unit Code"
         Height          =   255
         Left            =   360
         TabIndex        =   75
         Top             =   2175
         Width           =   855
      End
      Begin VB.Label Label33 
         Caption         =   "Price Excl"
         Height          =   255
         Left            =   3240
         TabIndex        =   74
         Top             =   1535
         Width           =   735
      End
      Begin VB.Label Label32 
         Caption         =   "%"
         Height          =   255
         Left            =   5760
         TabIndex        =   72
         Top             =   895
         Width           =   135
      End
      Begin VB.Label Label31 
         Caption         =   "VAT"
         Height          =   255
         Left            =   5040
         TabIndex        =   71
         Top             =   895
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   3240
         TabIndex        =   66
         Top             =   895
         Width           =   855
      End
      Begin VB.Label Label27 
         Caption         =   "Curr Code"
         Height          =   255
         Left            =   360
         TabIndex        =   65
         Top             =   1530
         Width           =   855
      End
      Begin VB.Label Label30 
         Caption         =   "Supplier"
         Height          =   255
         Left            =   360
         TabIndex        =   61
         Top             =   885
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Cost Center"
         Height          =   255
         Left            =   3240
         TabIndex        =   59
         Top             =   225
         Width           =   1335
      End
      Begin VB.Label Label28 
         Caption         =   "Exch Rate"
         Height          =   255
         Left            =   360
         TabIndex        =   58
         Top             =   1860
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "Req Date"
         Height          =   255
         Left            =   3240
         TabIndex        =   57
         Top             =   575
         Width           =   735
      End
      Begin VB.Label Label24 
         Caption         =   "Req No"
         Height          =   255
         Left            =   360
         TabIndex        =   56
         Top             =   570
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Total Cost"
         Height          =   255
         Left            =   3240
         TabIndex        =   36
         Top             =   2175
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "VAT"
         Height          =   255
         Left            =   3240
         TabIndex        =   34
         Top             =   1855
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Unit Price"
         Height          =   255
         Left            =   3240
         TabIndex        =   32
         Top             =   1215
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Size"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Product"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   255
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Inventory Items"
      Height          =   1695
      Left            =   5280
      TabIndex        =   25
      Top             =   600
      Width           =   6495
      Begin MSComctlLib.ListView ListView2 
         Height          =   1335
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   2355
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
   Begin VB.Frame Frame5 
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   5055
      Begin VB.TextBox txtDateOfCompletion 
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
         Left            =   3720
         TabIndex        =   38
         Top             =   615
         Width           =   1215
      End
      Begin VB.TextBox txtDateOfCommencement 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1320
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtSupervisedBy 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3720
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtDoneBy 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Date of Completion"
         Height          =   255
         Left            =   2520
         TabIndex        =   39
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "DOC"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "Supervised By"
         Height          =   255
         Left            =   2520
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label21 
         Caption         =   "Done By"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   255
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   5055
      Begin VB.TextBox txtTotalCost 
         Alignment       =   1  'Right Justify
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
         Left            =   3720
         TabIndex        =   44
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtJobBriefDate 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3720
         TabIndex        =   23
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtDepartmentCode 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtLpono 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtDescriptionOfOrder 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox txtJobCardNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3720
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtDeadlineDate 
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
         Left            =   1320
         TabIndex        =   5
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtCustomerName 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label Label16 
         Caption         =   "Total Cost (Incl)"
         Height          =   255
         Left            =   2520
         TabIndex        =   45
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Brief Date"
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Department"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Job Card No"
         Height          =   255
         Left            =   2640
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "L.P.O No"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Desc of Order"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "Deadline Date"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label25 
         Caption         =   "Customer"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Requisition Raised By The Department"
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
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   5055
      Begin MSComctlLib.ListView ListView1 
         Height          =   1935
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   3413
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
            Picture         =   "frmODASMOpenJobCard.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMOpenJobCard.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMOpenJobCard.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMOpenJobCard.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMOpenJobCard.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMOpenJobCard.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMOpenJobCard.frx":236E
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
      Width           =   11910
      _ExtentX        =   21008
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
         Left            =   10560
         Top             =   -120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         HelpCommand     =   11
         HelpContext     =   72
         HelpFile        =   "REGHELP.HLP"
      End
   End
   Begin VB.Label Label37 
      Caption         =   "Req Totals"
      Height          =   255
      Left            =   2760
      TabIndex        =   84
      Top             =   6375
      Width           =   975
   End
   Begin VB.Label Label36 
      Caption         =   "ItemNo"
      Height          =   255
      Left            =   120
      TabIndex        =   82
      Top             =   6375
      Width           =   975
   End
   Begin VB.Label Label35 
      Caption         =   "Total VAT:"
      Height          =   255
      Left            =   120
      TabIndex        =   80
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label34 
      Caption         =   "Price Excl"
      Height          =   255
      Left            =   2760
      TabIndex        =   79
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label Label22 
      Caption         =   "Remarks"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   6720
      Width           =   1095
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
Attribute VB_Name = "frmODASMOpenJobCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsREQ As clsODASRequisition
Dim rsOPERATION As clsODASOperation

Private Sub cboCurrencyCode_GotFocus()
        selectCURRENCYGOTFOCUS
End Sub

Private Sub cboCurrencyCode_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub cboCurrencyCode_LostFocus()
        selectCURRENCYLOSTFOCUS
End Sub
Private Sub txtCostCenter_GotFocus()
        SelectCostingGotFocus
End Sub

Private Sub txtCostCenter_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub txtCostCenter_LostFocus()
        selectCostingLostFocus
End Sub


Private Sub Form_Activate()
        disableALLRECORD
        Set rsREQ = New clsODASRequisition
        rsREQ.loadDEFAULT
        rsREQ.loadJOBBrief
        rsREQ.obtainBASECURRENCY
        rsREQ.calculateTOTALS
        showALLREQUISITIONSRAISED
        frmODASMOpenJobCard.ListView1.Enabled = False
        frmODASMOpenJobCard.ListView2.Enabled = False
        frmODASMOpenJobCard.ListView3.Enabled = False
        
        If bRequisitionAPPROVAL = True Or bRequisitionAUTHORIZATION = True Then
                frmODASMOpenJobCard.ListView1.Enabled = True
                frmODASMOpenJobCard.ListView2.Enabled = True
                frmODASMOpenJobCard.ListView3.Enabled = True
        End If
        
        showDeptPRODUCT
'        disableFRAME
End Sub



Private Sub Form_Unload(Cancel As Integer)
        Set rsREQ = Nothing
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'On Error GoTo err
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView1_DblClick()
'On Error GoTo err
            With SchedulingMain
                    
                    CurrentRecord = Trim(Me.ListView1.SelectedItem.Text)
                    GlobalDepartmentCode = ""
                    frmODASMOperation.txtApplicationNo.Text = CurrentRecord
                    Set rsOPERATION = New clsODASOperation
                    GlobalDepartmentCode = .ListView1.SelectedItem.SubItems(1)
                    rsOPERATION.checkAPPROVEDDISCHARGE
                    If bRequisitionAPPROVAL = False And bRequisitionAUTHORIZATION = False Then Exit Sub
                    
                    rsOPERATION.approveOPERATION
                    Set rsOPERATION = Nothing
            End With

Exit Sub

err:
    ErrorMessage
End Sub


Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo err
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
            
            Screen.ActiveForm.txtAccountNo.Text = Item.Text
            'Screen.ActiveForm.txtNames.Text = Item.SubItems(1)
            showALLLandLORDSites

        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'On Error GoTo err
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo err
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
            
            Screen.ActiveForm.txtProductCode.Text = Item.Text
            Screen.ActiveForm.ListView3.Enabled = True

            ShowSupplierCOST3
            rsREQ.loadCostCenter
            Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub
Private Sub ListView3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'On Error GoTo err
    ListView3.SortKey = ColumnHeader.Index - 1
    ListView3.Sorted = True
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ListView3_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo err
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
            
            Screen.ActiveForm.txtAccountNo.Text = Item.Text
            rsREQ.LoadItemCOST
            calculatePRICE
            Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub disableFRAME()
'On Error GoTo err
    
    With Screen.ActiveForm
        .Frame1.Enabled = False
        .Frame2.Enabled = False
        .Frame3.Enabled = False
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub enableFRAME()
'On Error GoTo err
    
    With Screen.ActiveForm
        .Frame1.Enabled = False
        .Frame2.Enabled = True
        .Frame3.Enabled = True
    End With
Exit Sub

err:
    ErrorMessage
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
''On Error GoTo Err
        
        With frmODASMOpenJobCard
        
        Select Case Button.Key
                Case "N"
                    Select Case Button.Caption
                    
                    Case "New &Record "
                            If editRECORD Then Exit Sub
                            .ListView2.Enabled = True
                            .ListView3.Enabled = False

                            NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4:
                            enableFRAME
                            rsREQ.enableRECORD
                            rsREQ.clearRECORD
                    Case "&Save Record "
                    
                            If NewRecord Then
                                    rsREQ.GenerateRequisitionNo
                            End If
    
                            bsaveRECORD = False
                            rsREQ.validateRECORD

                            If bsaveRECORD = True Then
                            
                                    rsREQ.updateRECORD
                                    
                                    If bsaveRECORD = False Then
                                              .Toolbar1.Buttons(2).Caption = "New &Record "
                                              .Toolbar1.Buttons(3).Caption = "&NEXT ITEM "
                                              .Toolbar1.Buttons(4).Caption = "FINISH"
                                                disableALLRECORD
                                    End If
                            End If
                    
                    Case "&NEXT ITEM "
                            
                            .Toolbar1.Buttons(1).Caption = "&Save Record"
                            rsREQ.enableRECORD
                            rsREQ.clearRECORDPartially
                    Case Else
                        Exit Sub
                    End Select
    
        Case "E"
            Select Case Button.Caption
                Case "&Edit/Change "
                
                Case "&Save Record "

                        bsaveRECORD = False
                        rsREQ.validateRECORD
                        
                        If bsaveRECORD = True Then
                                rsREQ.updateRECORD
                                If bsaveRECORD = False Then
                                          .Toolbar1.Buttons(2).Caption = "New &Record "
                                          .Toolbar1.Buttons(3).Caption = "&NEXT ITEM "
                                          .Toolbar1.Buttons(4).Caption = "FINISH"
                                          disableALLRECORD
                                End If
                        End If
                
                Case "&NEXT ITEM "
                            .Toolbar1.Buttons(3).Caption = "&Save Record "
                            rsREQ.enableRECORD
                            rsREQ.clearRECORDPartially
                Case Else
            End Select
        
        Case "S"
                .Toolbar1.Buttons(2).Caption = "New &Record "
                .Toolbar1.Buttons(2).Image = 2
                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                .Toolbar1.Buttons(3).Image = 5
                NewRecord = False: editRECORD = False: clearALLRECORD

        Case "R"
            If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                .Toolbar1.Buttons(2).Caption = "New &Record "
                .Toolbar1.Buttons(2).Image = 2
                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                .Toolbar1.Buttons(3).Image = 5
                NewRecord = False: editRECORD = False: clearALLRECORD
        Case "P"
                If .txtJobCardNo.Text <> Empty Then
                        Load frmODASRJobCard
                        frmODASRJobCard.Show 1, Me
                End If
        Case "F"
     
     
        Case Else
            Exit Sub
        End Select
End With
Exit Sub
err:
    ErrorMessage

End Sub
Private Sub calculatePRICE()
On Error GoTo err
        
        With frmODASMOpenJobCard
            If .txtExchangeRate.Text <= Empty Or CDbl(.txtUnitPrice.Text) <= 0 Then Exit Sub
            .txtItemQuantity.Text = .UpDownQuantity.Value
            .txtTotalUnitPriceExcl.Text = FormatNumber(CDbl(.txtUnitPrice.Text) * CDbl(.txtItemQuantity.Text) * CDbl(.txtExchangeRate.Text))
            .txtVATAmount.Text = FormatNumber(CDbl(.txtTotalUnitPriceExcl.Text) * (CDbl(.txtVATRate) / 100))
            .txtTotalUnitPriceIncl.Text = FormatNumber(CDbl(.txtTotalUnitPriceExcl) + CDbl(.txtVATAmount.Text))
        End With

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub UpDownQuantity_Change()
'On Error GoTo err
        
        With frmODASMOpenJobCard
            If .txtExchangeRate.Text <= Empty Or CDbl(.txtUnitPrice.Text) <= 0 Then Exit Sub
            calculatePRICE
        End With

Exit Sub

err:
    ErrorMessage
End Sub
