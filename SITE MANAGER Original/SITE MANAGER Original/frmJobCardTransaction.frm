VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmJobCardTransaction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Job Card Transactions"
   ClientHeight    =   7800
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "frmJobCardTransaction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "List of Items Under Selected Job Card No"
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
      Left            =   5760
      TabIndex        =   63
      Top             =   3360
      Width           =   6135
      Begin MSComctlLib.ListView ListView3 
         Height          =   1575
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2778
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777152
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   0
      TabIndex        =   53
      Top             =   2520
      Width           =   11895
      Begin VB.TextBox txtDateOfCompletion 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   10440
         TabIndex        =   61
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtSupervisor 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   7080
         TabIndex        =   59
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtDateOfCommence 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4800
         TabIndex        =   57
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtJobDoneBy 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1200
         TabIndex        =   55
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label30 
         Caption         =   "Date Of Completion"
         Height          =   255
         Left            =   9000
         TabIndex        =   60
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label29 
         Caption         =   "Supervisor"
         Height          =   255
         Left            =   6240
         TabIndex        =   58
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label28 
         Caption         =   "Date Of Commencement"
         Height          =   375
         Left            =   2880
         TabIndex        =   56
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label27 
         Caption         =   "Job Done By"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   0
      TabIndex        =   22
      Top             =   720
      Width           =   11895
      Begin MSComCtl2.DTPicker dtpRequisitionDate 
         Height          =   255
         Left            =   9960
         TabIndex        =   81
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Format          =   63176705
         CurrentDate     =   38292
      End
      Begin VB.TextBox txtSiding 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtSideCode 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   11280
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtOtherSite 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtDeptCode 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chkIlluminated 
         Caption         =   "Illuminated"
         Height          =   255
         Left            =   5760
         TabIndex        =   62
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtBorderWidth 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   10920
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtBorderLength 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   9600
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox chkBorder 
         Caption         =   "Check3"
         Height          =   195
         Left            =   7920
         TabIndex        =   45
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox txtWidth 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   10920
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtLength 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   9600
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtLPONo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtOrderQuantity 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtLocation 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtRequistionDate 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtJobCardNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtDeadLineDate 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtOrderDesc 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtClientName 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   600
         Width           =   5655
      End
      Begin VB.TextBox txtDepartment 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label31 
         Caption         =   "Siding"
         Height          =   255
         Left            =   3120
         TabIndex        =   71
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "OtherSite"
         Height          =   255
         Left            =   4200
         TabIndex        =   68
         Top             =   960
         Width           =   735
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   120
         X2              =   11760
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label26 
         Caption         =   "Deadline Date"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "W"
         Height          =   255
         Left            =   10560
         TabIndex        =   50
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label24 
         Caption         =   "L"
         Height          =   255
         Left            =   9360
         TabIndex        =   48
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label Label23 
         Caption         =   "Size"
         Height          =   255
         Left            =   9000
         TabIndex        =   47
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label22 
         Caption         =   "Border"
         Height          =   255
         Left            =   7080
         TabIndex        =   46
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "W"
         Height          =   255
         Left            =   10560
         TabIndex        =   43
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label17 
         Caption         =   "L"
         Height          =   255
         Left            =   9360
         TabIndex        =   41
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label5 
         Caption         =   "Size"
         Height          =   255
         Left            =   9000
         TabIndex        =   40
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "L.P.O No"
         Height          =   255
         Left            =   7080
         TabIndex        =   38
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   7080
         TabIndex        =   33
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Location"
         Height          =   255
         Left            =   7080
         TabIndex        =   32
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Date"
         Height          =   255
         Left            =   9600
         TabIndex        =   31
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Job  Card No"
         Height          =   255
         Left            =   4680
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Desc Of Order"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Name Of Client"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Department"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Item(s) Requisitioned"
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
      Left            =   5760
      TabIndex        =   5
      Top             =   5400
      Width           =   6135
      Begin VB.TextBox txtRequisitionNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtTaxTotalPrice 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4440
         TabIndex        =   77
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Frame Frame7 
         Height          =   375
         Left            =   0
         TabIndex        =   74
         Top             =   1920
         Width           =   3135
         Begin VB.OptionButton optPrices 
            Caption         =   "Wholesale Qty"
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
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   76
            Top             =   120
            Width           =   1575
         End
         Begin VB.OptionButton optPrices 
            Caption         =   "Retail Qty"
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
            Height          =   195
            Index           =   1
            Left            =   1800
            TabIndex        =   75
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.ComboBox cboCategory 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1080
         TabIndex        =   73
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtTaxAmount 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4440
         TabIndex        =   66
         Top             =   1920
         Width           =   1575
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Left            =   4440
         Max             =   0
         Min             =   32767
         TabIndex        =   21
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox txtQuantity 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   4680
         TabIndex        =   20
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtTotalPrice 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4440
         TabIndex        =   18
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Top             =   1320
         Width           =   2415
      End
      Begin VB.ComboBox cboTax 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1080
         TabIndex        =   13
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtItemName 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtItemCode 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   4440
         TabIndex        =   9
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtCategoryCode 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3720
         TabIndex        =   7
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label33 
         Caption         =   "Req No."
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label32 
         Caption         =   "Tax Cost"
         Height          =   255
         Left            =   3600
         TabIndex        =   78
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Tax Amount"
         Height          =   255
         Left            =   3480
         TabIndex        =   65
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   3600
         TabIndex        =   19
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label20 
         Caption         =   "Total Cost"
         Height          =   255
         Left            =   3600
         TabIndex        =   17
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "Cost"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Category Name"
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Tax"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Item Name"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Item Code"
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Cat Code "
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "List Of Inventory Items"
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
      Height          =   2415
      Left            =   0
      TabIndex        =   3
      Top             =   5280
      Width           =   5655
      Begin MSComctlLib.ListView ListView2 
         Height          =   2055
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3625
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777152
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
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
            Picture         =   "frmJobCardTransaction.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobCardTransaction.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobCardTransaction.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobCardTransaction.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobCardTransaction.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobCardTransaction.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobCardTransaction.frx":236E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
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
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         HelpCommand     =   11
         HelpContext     =   72
         HelpFile        =   "REGHELP.HLP"
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "List Of Open Job Cards"
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
      Left            =   0
      TabIndex        =   0
      Top             =   3360
      Width           =   5655
      Begin MSComctlLib.ListView ListView1 
         Height          =   1575
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2778
         View            =   3
         MultiSelect     =   -1  'True
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
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
      Begin VB.Menu mnuOpenedJobs 
         Caption         =   "Opened Jobs"
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
Attribute VB_Name = "frmJobCardTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public billbno, QuantityType, MediaItemCode, DepCode, CustomerCode, PhysicalAddress, JoboNumber, PurchaseOrderNo, QuantityUnits




Private Sub cboBBNo_GotFocus()

End Sub

Private Sub cboBBNo_LostFocus()

End Sub

Private Sub cboCouncilPeriod_Click()
Me.ListView1.SetFocus
End Sub



Private Sub cboCountry_Click()
Me.ListView1.SetFocus
End Sub

Private Sub cboElecPeriod_Click()
Me.ListView1.SetFocus
End Sub


Private Sub cboLandLord_Click()
Me.ListView1.SetFocus
End Sub

Private Sub cboRentPeriod_Click()
Me.ListView1.SetFocus
End Sub



Private Sub cboTown_Click()
Me.ListView1.SetFocus
End Sub



Private Sub Combo1_Click()
Me.ListView1.SetFocus
End Sub

Private Sub cboCategory_GotFocus()
'On Error GoTo Err
If Not NewRecord Then Exit Sub
With Me
   
    If .cboCategory.ListCount <> 0 Then Exit Sub
    
     AttachSQL = "SELECT A.CategoryCode AS SelectField FROM ParamDrugCategories A ,GenProductsInventory B WHERE A.CategoryCode = B.CategoryCode ORDER BY A.CategoryCode;"
    .cboCategory.Clear
    MyCommonData.AttachInventDropDown
    
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub cboCity_GotFocus()
If Not NewRecord Then Exit Sub
With Me
    AttachSQL = "SELECT Town AS SelectField FROM ParamTownS ORDER BY Town;"
    .cboCity.Clear
    MyCommonData.AttachDropDown
End With

End Sub

Private Sub cboSiding_GotFocus()
If Not NewRecord Then Exit Sub
With Me
    AttachSQL = "SELECT SidingDescription AS SelectField FROM Advertsiding ORDER BY sidingdescription;"
    .cboSiding.Clear
    MyCommonData.AttachDropDown
End With

End Sub

Private Sub cboCategory_Click()
Me.ListView3.SetFocus
End Sub
Private Sub cboCategory_LostFocus()
'On Error GoTo Err
With Me

    Set rsFindRecord = cnINVENT.Execute("SELECT * FROM ParamDrugCategories A,GenproductsInventory B WHERE A.CategoryCode = B.CategoryCode AND B.CategoryCODE='" & Trim(.cboCategory.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .txtCategoryCode.Text = rsFindRecord!CategoryName & ""
        .cboCategory.Text = rsFindRecord!CategoryCode & ""
        
        .ListView2.SetFocus
        
        If .cboCategory.Text = "AAA" Then
            Call ShowAllInventoryItems
        Else
            Call ShowInventoryItemsPerCategory
        End If
        
    End If
    
End With
Exit Sub
Err:
    ErrorMessage
End Sub
Private Sub ShowInventoryItemsPerCategory()
'On Error GoTo Err
With Me
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Product Code", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Purchase Order No", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Product Name", .ListView2.Width / 4.5
.ListView2.ColumnHeaders.Add , , "Current Quantity", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Quantity Units", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Current Total Pieces", .ListView2.Width / 5.5

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM GenProductsInventory WHERE Discontinued = '" & "N" & "' AND CategoryCode = '" & Trim(.cboCategory.Text) & "' ORDER BY DrugName;", cnINVENT, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!DrugCode))


    If Not IsNull(rsLIST!PurchaseOrderNo) Then
        MyList.SubItems(1) = CStr(rsLIST!PurchaseOrderNo)
    End If
     
    If Not IsNull(rsLIST!DrugName) Then
        MyList.SubItems(2) = CStr(rsLIST!DrugName)
    End If
    
    If Not IsNull(rsLIST!CurrentQuantity) Then
        MyList.SubItems(3) = CStr(rsLIST!CurrentQuantity)
    End If
    
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(4) = CStr(rsLIST!QuantityUnits)
    End If
    
    If Not IsNull(rsLIST!TotalPieces) Then
        MyList.SubItems(5) = CStr(rsLIST!TotalPieces)
    End If
    
    rsLIST.MoveNext
    
Wend



Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub ShowAllInventoryProductsStructure()
'On Error GoTo Err
With Me
.ListView3.ListItems.Clear
.ListView3.ColumnHeaders.Clear

.ListView3.ColumnHeaders.Add , , "Product Code", .ListView3.Width / 6#
.ListView3.ColumnHeaders.Add , , "Purchase Order No", .ListView3.Width / 6
.ListView3.ColumnHeaders.Add , , "Product Name", .ListView3.Width / 5
.ListView3.ColumnHeaders.Add , , "Current Quantity", .ListView3.Width / 6
.ListView3.ColumnHeaders.Add , , "Quantity Units", .ListView3.Width / 6
.ListView3.ColumnHeaders.Add , , "Current Total Pieces", .ListView3.Width / 6

.ListView3.View = lvwReport
End With
Exit Sub
Err:
     ErrorMessage
End Sub

Private Sub ShowAllInventoryItems()
'On Error GoTo Err
With Me
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Product Code", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Purchase Order No", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Product Name", .ListView2.Width / 4.5
.ListView2.ColumnHeaders.Add , , "Current Quantity", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Quantity Units", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Current Total Pieces", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Category Code", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Category Name", .ListView2.Width / 5.5

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM GenProductsInventory A ,ParamDrugCategories B WHERE A.CategoryCode = B.Categorycode AND A.Discontinued = '" & "N" & "' ORDER BY A.DrugName;", cnINVENT, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!DrugCode))


    If Not IsNull(rsLIST!PurchaseOrderNo) Then
        MyList.SubItems(1) = CStr(rsLIST!PurchaseOrderNo)
    End If
     
    If Not IsNull(rsLIST!DrugName) Then
        MyList.SubItems(2) = CStr(rsLIST!DrugName)
    End If
    
    If Not IsNull(rsLIST!CurrentQuantity) Then
        MyList.SubItems(3) = CStr(rsLIST!CurrentQuantity)
    End If
    
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(4) = CStr(rsLIST!QuantityUnits)
    End If
    
    If Not IsNull(rsLIST!TotalPieces) Then
        MyList.SubItems(5) = CStr(rsLIST!TotalPieces)
    End If
    
     If Not IsNull(rsLIST!CategoryCode) Then
        MyList.SubItems(6) = CStr(rsLIST!CategoryCode)
    End If
    
     If Not IsNull(rsLIST!CategoryName) Then
        MyList.SubItems(7) = CStr(rsLIST!CategoryName)
    End If
    
    rsLIST.MoveNext
    
Wend



Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage

End Sub

Private Function GetAdvertPrice()
'On Error GoTo Err
With Me
    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM AdvertPricing WHERE BBNo='" & Trim(.txtItemCode.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing
    Else
        GetAdvertPrice = rsFindRecord!BBCharges & ""
     
        
    End If
End With
Exit Function
Err:
    ErrorMessage
End Function

Private Sub ShowBillBoardsPerCategory()
''On Error GoTo Err
With Me
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Category Code", .ListView2.Width / 6#
.ListView2.ColumnHeaders.Add , , "Category Name", .ListView2.Width / 4
.ListView2.ColumnHeaders.Add , , "Item Code", .ListView2.Width / 6
.ListView2.ColumnHeaders.Add , , "Item Name", .ListView2.Width / 4
.ListView2.ColumnHeaders.Add , , "Length", .ListView2.Width / 6.5
.ListView2.ColumnHeaders.Add , , "Width", .ListView2.Width / 6.5

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertBBDetails A ,AdvertCategories B WHERE A.CategoryCode = B.CategoryCode AND A.CategoryCode = '" & Trim(.txtCategoryCode.Text) & "' ORDER BY A.Name;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing:  Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!CategoryCode))


    If Not IsNull(rsLIST!CategoryName) Then
        MyList.SubItems(1) = CStr(rsLIST!CategoryName)
    End If
     
    If Not IsNull(rsLIST!BillBoardNo) Then
        MyList.SubItems(2) = CStr(rsLIST!BillBoardNo)
    End If
    
    If Not IsNull(rsLIST!Name) Then
        MyList.SubItems(3) = CStr(rsLIST!Name)
    End If
    
    If Not IsNull(rsLIST!Length) Then
        MyList.SubItems(4) = CStr(rsLIST!Length)
    End If
    
    If Not IsNull(rsLIST!Width) Then
        MyList.SubItems(5) = CStr(rsLIST!Width)
    End If
    
    rsLIST.MoveNext
    
Wend


Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Private Sub ShowAllClientsStructure()
''On Error GoTo Err
With Me
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Client Code", .ListView1.Width / 6#
.ListView1.ColumnHeaders.Add , , "Company Name", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Address", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "City", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Phone", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Mobile Phone", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Fax", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Contact Title", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Contact Name", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Physical Adress", .ListView1.Width / 5.5

.ListView1.View = lvwReport
End With
Exit Sub
Err:
    ErrorMessage
 End Sub
Private Sub ShowAllClients()
''On Error GoTo Err
With Me
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Client Code", .ListView1.Width / 6#
.ListView1.ColumnHeaders.Add , , "Company Name", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Address", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "City", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Phone", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Mobile Phone", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Fax", .ListView1.Width / 5.5
.ListView1.ColumnHeaders.Add , , "Contact Title", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Contact Name", .ListView1.Width / 4
.ListView1.ColumnHeaders.Add , , "Physical Adress", .ListView1.Width / 1.5
.ListView1.ColumnHeaders.Add , , "Client Code", .ListView1.Width / 6.5

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertClients ORDER BY CompanyName;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!CustomerId))


    If Not IsNull(rsLIST!CompanyName) Then
        MyList.SubItems(1) = CStr(rsLIST!CompanyName)
    End If
     
    If Not IsNull(rsLIST!Address) Then
        MyList.SubItems(2) = CStr(rsLIST!Address)
    End If
    
    If Not IsNull(rsLIST!City) Then
        MyList.SubItems(3) = CStr(rsLIST!City)
    End If
    
    If Not IsNull(rsLIST!Phone) Then
        MyList.SubItems(4) = CStr(rsLIST!Phone)
    End If
    
    If Not IsNull(rsLIST!MobilePhone) Then
        MyList.SubItems(5) = CStr(rsLIST!MobilePhone)
    End If
    
    If Not IsNull(rsLIST!Fax) Then
        MyList.SubItems(6) = CStr(rsLIST!Fax)
    End If
    
    If Not IsNull(rsLIST!ContactTitle) Then
        MyList.SubItems(7) = CStr(rsLIST!ContactTitle)
    End If
    
    If Not IsNull(rsLIST!ContactName) Then
        MyList.SubItems(8) = CStr(rsLIST!ContactName)
    End If
    
    If Not IsNull(rsLIST!PhysicalAddress) Then
        MyList.SubItems(9) = CStr(rsLIST!PhysicalAddress)
    End If
    
    If Not IsNull(rsLIST!CustomerId) Then
        MyList.SubItems(10) = CStr(rsLIST!CustomerId)
    End If
    
    rsLIST.MoveNext
    
Wend

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub
Private Sub ShowAllBillBoardCategoriesStructure()
''On Error GoTo Err
With Me
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Category Code", .ListView2.Width / 6#
.ListView2.ColumnHeaders.Add , , "Category Name", .ListView2.Width / 5
.ListView2.ColumnHeaders.Add , , "Item Code", .ListView2.Width / 6
.ListView2.ColumnHeaders.Add , , "Item Name", .ListView2.Width / 4
.ListView2.ColumnHeaders.Add , , "Length", .ListView2.Width / 6.5
.ListView2.ColumnHeaders.Add , , "Width", .ListView2.Width / 6.5

.ListView2.View = lvwReport
End With
Exit Sub
Err:
    ErrorMessage
End Sub


Private Sub ShowAllBillBoardCategories()
'On Error GoTo Err
With Me
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Category Code", .ListView2.Width / 6#
.ListView2.ColumnHeaders.Add , , "Category Name", .ListView2.Width / 4
.ListView2.ColumnHeaders.Add , , "Item Code", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Item Name", .ListView2.Width / 4
.ListView2.ColumnHeaders.Add , , "Length", .ListView2.Width / 6.5
.ListView2.ColumnHeaders.Add , , "Width", .ListView2.Width / 6.5

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertBBDetails A ,AdvertCategories B WHERE A.CategoryCode = B.CategoryCode ORDER BY A.Name;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!CategoryCode))


    If Not IsNull(rsLIST!CategoryName) Then
        MyList.SubItems(1) = CStr(rsLIST!CategoryName)
    End If
     
    If Not IsNull(rsLIST!BillBoardNo) Then
        MyList.SubItems(2) = CStr(rsLIST!BillBoardNo)
    End If
    
    If Not IsNull(rsLIST!Name) Then
        MyList.SubItems(3) = CStr(rsLIST!Name)
    End If
    
    If Not IsNull(rsLIST!Length) Then
        MyList.SubItems(4) = CStr(rsLIST!Length)
    End If
    
    If Not IsNull(rsLIST!Width) Then
        MyList.SubItems(5) = CStr(rsLIST!Width)
    End If
    
    rsLIST.MoveNext
    
Wend



Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage
End Sub

Private Sub cboCity_Click()
Me.ListView1.SetFocus
End Sub

Private Sub cboSiding_Click()
Me.ListView2.SetFocus
End Sub

Private Sub cboSiding_LostFocus()
''On Error GoTo Err
With Me

    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM AdvertSiding WHERE SidingDescription ='" & Trim(.cboSiding.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .cboSiding.Text = rsFindRecord!SidingType & ""
        .txtQuantity.SetFocus
    End If
    End With
  Exit Sub
Err:
     ErrorMessage
End Sub

Private Sub cboTax_Click()
Me.ListView2.SetFocus
End Sub

Private Sub cboTax_GotFocus()
'On Error GoTo Err
If Not NewRecord Then Exit Sub
With Me
   
    If .cboTax.ListCount <> 0 Then Exit Sub
    
     AttachSQL = "SELECT Description AS SelectField FROM ParamTaxes ORDER BY Description;"
    .cboTax.Clear
    MyCommonData.AttachDropDown
    
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub cboTax_LostFocus()
'On Error GoTo Err
With Me

    Set rsFindRecord = cnCOMMON.Execute("SELECT * FROM ParamTaxes WHERE Description ='" & Trim(.cboTax.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing: Exit Sub
    Else
        .cboTax.Text = rsFindRecord!TaxRate & ""
        If .txtPrice.Text = "" And .txtTotalPrice.Text = "" Then Exit Sub
          .txtTaxAmount.Text = Val(CSng(.cboTax.Text) / 100) * Val(CSng(.txtTotalPrice.Text))
          .txtTaxTotalPrice.Text = Val(CSng(.txtTotalPrice.Text)) + Val(.txtTaxAmount.Text)
          .txtTaxAmount.Text = FormatNumber(.txtTaxAmount.Text, 2, vbUseDefault, vbUseDefault, vbTrue)
'          .txtTaxPrice.Text = FormatNumber(.txtTaxAmount.Text, 2, vbUseDefault, vbUseDefault, vbTrue)
         .ListView2.SetFocus
    End If
    
End With
Exit Sub
Err:
    ErrorMessage

End Sub

Private Sub dtpRequisitionDate_CloseUp()
On Error GoTo Err
With Me
  If .dtpRequisitionDate < Date Then
     MsgBox "The requisition date can not be in the past", vbExclamation, "Invalid Requisition Date"
     .txtRequistionDate.SetFocus
     Else
     .txtRequistionDate.Text = .dtpRequisitionDate.Value
   End If
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub Form_Initialize()
Set MyCommonData = New clsCommonData
'Set mycommondata New clsCommonData

End Sub

Private Sub Form_Load()
ShowAllClientsStructure
ShowAllBillBoardCategoriesStructure
ShowAllInventoryProductsStructure
DisableCostControls
End Sub
Private Sub DisableCostControls()
On Error GoTo Err
With Me
   .cboTax.Enabled = False
   .txtPrice.Enabled = False
   .VScroll1.Enabled = False
   .txtQuantity.Enabled = False
   .txtTotalPrice.Enabled = False
   .txtTaxTotalPrice.Enabled = False
   .txtTaxAmount.Enabled = False
End With
Exit Sub
Err:
    ErrorMessage
End Sub
Private Sub EnableCostControls()
On Error GoTo Err
With Me
   .cboTax.Enabled = True
   .txtPrice.Enabled = True
   .VScroll1.Enabled = True
   .txtQuantity.Enabled = True
   .txtTotalPrice.Enabled = True
   .txtTaxTotalPrice.Enabled = True
   .txtTaxAmount.Enabled = True
End With
Exit Sub
Err:
    ErrorMessage

End Sub
Public Function AutoPurchaseOrderNo() As String
'On Error GoTo Err
With Me

Dim rsLastID As ADODB.Recordset 'used to retrieve current LastId in the Table
Dim strLastID As String 'SQL statement

Dim strTemp As String 'store current record
Dim iNumPos As Integer 'store position of the first numeral
Dim strPrefix As String 'stores Id Prefix

'Retrieve the last record in the recrdset where order is ascending

'strLastID = "SELECT max(cnCliniclDetails.cnCliniclID) as lastid from cnCliniclDetails"
strLastID = "SELECT MAX(RequisitionNo) AS LastID FROM AdvertJobCardTransactions;"
Set rsLastID = New ADODB.Recordset

With rsLastID
'open the recordset
    .Open strLastID, cnCOMMON, adOpenKeyset, adLockOptimistic
    If .EOF And .BOF Then 'shows empty recordset
        AutoPurchaseOrderNo = "RQ00001" 'format of desired format of the string value
    ElseIf IsNull(!lastid) = True Or !lastid = "" Then
        AutoPurchaseOrderNo = "RQ00001"
    Else
       ' If .EOF And .BOF Then .MoveFirst
        '.MoveLast
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
        AutoPurchaseOrderNo = strPrefix & strTemp
    End If
End With
End With
    Exit Function
Err:
    ErrorMessage
End Function

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo Err
'If Not NewRecord And Not EditRecord Then Item.Checked = False: Exit Sub
If Me.ListView1.ListItems.Count = 0 Or Me.ListView1.View <> lvwReport Then Exit Sub
    
    Dim i, j, k
    j = Me.ListView1.ListItems.Count
    
    If j = 0 Then Exit Sub
    
    For i = 1 To j
        If Me.ListView1.ListItems(i).Text <> Item Then
            Me.ListView1.ListItems(i).Checked = False
        End If
    Next i
    
      If Item.Checked = True Then
        
        Me.txtJobCardNo.Text = Item
        Me.txtDeptCode.Text = Item.SubItems(1)
        Me.txtDepartment.Text = Item.SubItems(2)
        Me.txtSupervisor.Text = Item.SubItems(3)
        Me.txtDateOfCompletion.Text = Item.SubItems(4)
        Me.txtDateOfCommence.Text = Item.SubItems(5)
        Me.txtClientName.Text = Item.SubItems(6)
        Me.txtDeadLineDate.Text = Item.SubItems(7)
        Me.txtLPONo.Text = Item.SubItems(8)
        JoboNumber = Item
        DepCode = Item.SubItems(1)
        
        Call ShowAllItemsUnderSelectedJob
        
             
    ElseIf Item.Checked = False Then

    End If
Exit Sub
Err:
    ErrorMessage
End Sub
Private Sub ShowAllItemsUnderSelectedJob()
'On Error GoTo Err
With Me
.ListView3.ListItems.Clear
.ListView3.ColumnHeaders.Clear

.ListView3.ColumnHeaders.Add , , "Job Card No", .ListView3.Width / 5
.ListView3.ColumnHeaders.Add , , "Media Name", .ListView3.Width / 4
.ListView3.ColumnHeaders.Add , , "Site Code", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Site Name", .ListView3.Width / 4.5
.ListView3.ColumnHeaders.Add , , "Quantity", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Item Code", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Length", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Width", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Illuminated", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Siding", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Bordered", .ListView3.Width / 6.5
.ListView3.ColumnHeaders.Add , , "Registered Site", .ListView3.Width / 4.5
.ListView3.ColumnHeaders.Add , , "Other Site", .ListView3.Width / 4.5

.ListView3.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertJobBrief A ,AdvertJobBriefItems B,AdvertBBDetails C,AdvertSites D WHERE  A.JobBriefno = B.JobBriefNo AND C.BillBoardNo = B.ItemCode AND B.SiteCode = D.SiteNo AND B.JoBbriefNo = '" & JoboNumber & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView3.View = lvwList
    Set MyList = .ListView3.ListItems.Add(, , "Requisitions have been made for all items")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView3.ListItems.Add(, , CStr(rsLIST!JobBriefNo))


    If Not IsNull(rsLIST!CategoryName) Then
        MyList.SubItems(1) = CStr(rsLIST!CategoryName) + " " + CStr(rsLIST!ItemName)
    End If
     
    If Not IsNull(rsLIST!SiteCode) Then
        MyList.SubItems(2) = CStr(rsLIST!SiteCode)
    End If
    
    If Not IsNull(rsLIST!SiteName) Then
        MyList.SubItems(3) = CStr(rsLIST!SiteName)
    End If
        
    If Not IsNull(rsLIST!Quantity) Then
        MyList.SubItems(4) = CStr(rsLIST!Quantity)
    End If
    
     If Not IsNull(rsLIST!ItemCode) Then
        MyList.SubItems(5) = CStr(rsLIST!ItemCode)
    End If
    
     If Not IsNull(rsLIST!Length) Then
        MyList.SubItems(6) = CStr(rsLIST!Length)
    End If
    
     If Not IsNull(rsLIST!Width) Then
        MyList.SubItems(7) = CStr(rsLIST!Width)
    End If
    
     If Not IsNull(rsLIST!Illuminate) And (rsLIST!Illuminate) = 0 Then
        MyList.SubItems(8) = CStr("NO")
      ElseIf Not IsNull(rsLIST!Illuminate) And (rsLIST!Illuminate) = 1 Then
        MyList.SubItems(8) = CStr("YES")
    End If
    
    If Not IsNull(rsLIST!Siding) Then
        MyList.SubItems(9) = CStr(rsLIST!Siding)
    End If
    
     If Not IsNull(rsLIST!Border) And (rsLIST!Border) = 0 Then
        MyList.SubItems(10) = CStr("NO")
      ElseIf Not IsNull(rsLIST!Border) And (rsLIST!Border) = 1 Then
        MyList.SubItems(10) = CStr("YES")
    End If
    
    If Not IsNull(rsLIST!SiteName) Then
        MyList.SubItems(11) = CStr(rsLIST!SiteName)
    End If
    
    If Not IsNull(rsLIST!OtherSite) Then
        MyList.SubItems(12) = CStr(rsLIST!OtherSite)
    End If
     
       
    rsLIST.MoveNext
    
Wend

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage

End Sub
Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo Err
'If Not NewRecord And Not EditRecord Then Item.Checked = False: Exit Sub
If Me.ListView2.ListItems.Count = 0 Or Me.ListView2.View <> lvwReport Then Exit Sub
    
    Dim i, j, k
    j = Me.ListView2.ListItems.Count
    
    If j = 0 Then Exit Sub
    
    For i = 1 To j
        If Me.ListView2.ListItems(i).Text <> Item Then
            Me.ListView2.ListItems(i).Checked = False
        End If
    Next i
    
      If Item.Checked = True Then
       DontChange = True

        Me.txtItemCode.Text = Item
        Me.txtItemName.Text = Item.SubItems(2)
        PurchaseOrderNo = Item.SubItems(1)
        QuantityUnits = Item.SubItems(4)
        Me.cboCategory.Text = Item.SubItems(6)
        Me.txtCategoryCode.Text = Item.SubItems(7)
        
    ElseIf Item.Checked = False Then
    
       
    End If
    DontChange = False
Exit Sub
Err:
    On Error Resume Next
End Sub
Public Sub ClearTextFields()
For Each i In Screen.ActiveForm
    If TypeOf i Is TextBox And i.Name <> "txtTitle" Then
        i.Text = Empty
    End If
    If TypeOf i Is ComboBox Then
        i.Clear
    End If
    If TypeOf i Is Image Then
        i.Picture = LoadPicture("")
    End If
Next i
End Sub

Private Sub ListView3_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo Err
'If Not NewRecord And Not EditRecord Then Item.Checked = False: Exit Sub
If Me.ListView3.ListItems.Count = 0 Or Me.ListView3.View <> lvwReport Then Exit Sub
    
    Dim i, j, k
    j = Me.ListView3.ListItems.Count
    
    If j = 0 Then Exit Sub
    
    For i = 1 To j
        If Me.ListView3.ListItems(i).Text <> Item Then
            Me.ListView3.ListItems(i).Checked = False
        End If
    Next i
    
      If Item.Checked = True Then
    
        
'        Me.txtItemCode.Text = Item
        Me.txtOrderDesc.Text = Item.SubItems(1)
        Me.txtSideCode.Text = Item.SubItems(2)
        Me.txtLocation.Text = Item.SubItems(3)
        Me.txtOrderQuantity.Text = Item.SubItems(4)
        Me.txtLength.Text = Item.SubItems(6)
        Me.txtWidth.Text = Item.SubItems(7)
        MediaItemCode = Item.SubItems(5)
        Me.txtOtherSite.Text = Item.SubItems(12)
        Me.txtSiding.Text = Item.SubItems(9)
                
        If Item.SubItems(8) = "YES" Then
        Me.chkIlluminated.Value = 1
        Else
        Me.chkIlluminated.Value = 0
        End If
        
        If Item.SubItems(10) = "YES" Then
        Me.chkBorder.Value = 1
        Me.txtBorderLength.Text = Item.SubItems(6)
        Me.txtBorderWidth.Text = Item.SubItems(7)
        Else
        Me.chkBorder.Value = 0
        End If
        
        
    ElseIf Item.Checked = False Then
    
        
    End If
    
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub mnuClear_Click()
    MyCommonData.ClearTextFields
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuCurrent_Click()
    Call ShowCurrentSettings

End Sub
Private Sub ShowCurrentSettings()
''On Error GoTo Err
With Me
Screen.MousePointer = vbHourglass

.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Brokers Code", .ListView1.Width / 8 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Brokers Name", .ListView1.Width / 2.5 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Physical Address", .ListView1.Width / 4 ', lvwColumnCenter
.ListView1.ColumnHeaders.Add , , "Postal Address", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Branch", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Town/City", .ListView1.Width / 10
.ListView1.ColumnHeaders.Add , , "Country", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Contact Person", .ListView1.Width / 6
.ListView1.ColumnHeaders.Add , , "Contact Title", .ListView1.Width / 3.5
.ListView1.ColumnHeaders.Add , , "Phone Number", .ListView1.Width / 8
.ListView1.ColumnHeaders.Add , , "Email", .ListView1.Width / 5

.ListView1.View = lvwReport: .ListView1.Visible = True

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT ParamInsuranceBrokers.* FROM ParamInsuranceBrokers ORDER BY ParamInsuranceBrokers.BrokersCode;", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem, NCount As Double

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Screen.MousePointer = vbDefault: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!BrokersCode))
    
    If Not IsNull(rsLIST!BrokersCode) Then
        MyList.SubItems(1) = CStr(rsLIST!BrokersName)
        MyList.SubItems(2) = CStr(rsLIST!PhysicalAddress)
        MyList.SubItems(3) = CStr(rsLIST!PostalAddress)
        MyList.SubItems(4) = CStr(rsLIST!Branch)
        MyList.SubItems(5) = CStr(rsLIST!TownCity)
        MyList.SubItems(6) = CStr(rsLIST!Country)
        MyList.SubItems(7) = CStr(rsLIST!ContactPerson)
        MyList.SubItems(8) = CStr(rsLIST!ContactTitle)
        MyList.SubItems(9) = CStr(rsLIST!TelephoneNo)
        MyList.SubItems(10) = CStr(rsLIST!Email)

    End If

    rsLIST.MoveNext
    
Wend

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
Private Function ValidRecord() As Boolean
On Error Resume Next
With Me
    If .txtJobCardNo.Text = Empty Then
        strMessage = "Jobcard Number required...!"
        .txtJobCardNo.SetFocus
    ElseIf .txtRequisitionNo.Text = Empty Then
        strMessage = "Requisition Number required...!"
        .txtRequisitionNo.SetFocus
    ElseIf .txtItemCode.Text = Empty Then
        strMessage = "Inventory Product Code required...!"
        .txtItemCode.SetFocus
    ElseIf .txtDeptCode.Text = Empty Then
        strMessage = "Department Code Required...!"
        .txtDeptCode.SetFocus
    ElseIf .txtRequistionDate.Text = Empty Then
        strMessage = "Date Of Requisition Required...!"
        .txtRequistionDate.SetFocus
    ElseIf .txtPrice.Text = Empty Then
        strMessage = "Cost Of Item Required...!"
        .txtPrice.SetFocus
    ElseIf .txtTotalPrice.Text = Empty Then
        strMessage = "Total Cost Of Item(s)..!"
        .txtTotalPrice.SetFocus
        
    
    Else
        ValidRecord = True
    End If
    If Not ValidRecord Then
        MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
    End If
End With
End Function
Private Function ValidMainRecord() As Boolean
With Me
    If .txtRequisitionNo.Text = Empty Then
        strMessage = "Requistion Number required...!"
        .txtRequisitionNo.SetFocus
    ElseIf .txtJobCardNo.Text = Empty Then
        strMessage = "Job Card Number required...!"
        .txtJobCardNo.SetFocus
    ElseIf .txtItemCode.Text = Empty Then
        strMessage = "Inventory Product Code required...!"
        .txtItemCode.SetFocus
     
    Else
        ValidMainRecord = True
    End If
'    If Not ValidRecord Then
'        MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
'    End If
End With
End Function
Private Sub RemoveCurrentList3Item()
'On Error GoTo Err
With Me
Dim i, j, k
   j = .ListView3.ListItems.Count: i = 1
     If j = 0 Then Exit Sub
     
     For i = 1 To j
      If .ListView3.ListItems(i).Checked = True Then
         .ListView3.ListItems.Remove (i): Exit Sub
      End If
    Next i
End With
Exit Sub
Err:
   ErrorMessage
End Sub
Private Sub RemoveCurrentList1Item()
'On Error GoTo Err
With Me
Dim i, j, k
   j = .ListView1.ListItems.Count: i = 1
     If j = 0 Then Exit Sub
     
     For i = 1 To j
      If .ListView1.ListItems(i).Checked = True Then
         .ListView1.ListItems.Remove (i): Exit Sub
      End If
    Next i
End With
Exit Sub
Err:
   ErrorMessage
End Sub
Private Sub RemoveCurrentList2Item()
'On Error GoTo Err
With Me
Dim i, j, k
   j = .ListView2.ListItems.Count: i = 1
     If j = 0 Then Exit Sub
     
     For i = 1 To j
      If .ListView2.ListItems(i).Checked = True Then
         .ListView2.ListItems.Remove (i): Exit Sub
      End If
    Next i
End With
Exit Sub
Err:
   ErrorMessage
End Sub

Private Sub mnuRegisteredClients_Click()
If Not NewRecord Then Exit Sub
Call ShowAllClients
End Sub

Private Sub mnuFullInventory_Click()
'On Error GoTo Err
If Not NewRecord Then Exit Sub
DontChange = True
Call ShowAllInventoryItems

Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub mnuOpenedJobs_Click()
If Not NewRecord Then Exit Sub
Call ShowOpenedJobs
End Sub
Private Sub ShowOpenedJobs()
'On Error GoTo Err
With Me
.ListView1.ListItems.Clear
.ListView1.ColumnHeaders.Clear

.ListView1.ColumnHeaders.Add , , "Job Card No", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "Dept Code", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Department Name", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "Supervised By", .ListView1.Width / 5
.ListView1.ColumnHeaders.Add , , "Envisiaged D.O.C", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Commence Date", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "Client Name", .ListView1.Width / 4.5
.ListView1.ColumnHeaders.Add , , "DeadLine Date", .ListView1.Width / 6.5
.ListView1.ColumnHeaders.Add , , "L.P.O No", .ListView1.Width / 6.5

.ListView1.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM AdvertJobCard A ,AdvertParamDepartments B,AdvertJobBrief C WHERE A.JobCardNo = C.JobBriefNo AND A.DeptCode = B.DepartmentCode AND A.Opened = '" & "Y" & "';", cnCOMMON, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView1.View = lvwList
    Set MyList = .ListView1.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView1.ListItems.Add(, , CStr(rsLIST!JobCardNo))


    If Not IsNull(rsLIST!DeptCode) Then
        MyList.SubItems(1) = CStr(rsLIST!DeptCode)
    End If
     
    If Not IsNull(rsLIST!DepartmentDescription) Then
        MyList.SubItems(2) = CStr(rsLIST!DepartmentDescription)
    End If
    
    If Not IsNull(rsLIST!SupervisedBy) Then
        MyList.SubItems(3) = CStr(rsLIST!SupervisedBy)
    End If
    
    If Not IsNull(rsLIST!EnvisiagedDateOfCompletion) Then
        MyList.SubItems(4) = CStr(rsLIST!EnvisiagedDateOfCompletion)
    End If
    
    If Not IsNull(rsLIST!StartDate) Then
        MyList.SubItems(5) = CStr(rsLIST!StartDate)
    End If
    
     If Not IsNull(rsLIST!CustomerName) Then
        MyList.SubItems(6) = CStr(rsLIST!CustomerName)
    End If
    
     If Not IsNull(rsLIST!DeadLineDate) Then
        MyList.SubItems(7) = CStr(rsLIST!DeadLineDate)
    End If
    
     If Not IsNull(rsLIST!lpono) Then
        MyList.SubItems(8) = CStr(rsLIST!lpono)
    End If
     
       
    rsLIST.MoveNext
    
Wend

Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage

End Sub

Private Sub optPrices_Click(Index As Integer)
On Error Resume Next
With Me
'If Not NewRecord And Not EditRecord Then Exit Sub
If .ListView2.ListItems.Count = 0 Then Exit Sub
Select Case Index
Case 0
    Call EnableCostControls
    Call GetWholeSaleCost
    Me.optPrices(0).Value = True
    .txtTotalPrice.Text = FormatNumber(CDbl(.txtPrice.Text), 2, vbUseDefault, vbUseDefault, vbTrue)
   
Case 1
    Call EnableCostControls
    Call GetRetailCost
    Me.optPrices(1).Value = True
   .txtTotalPrice.Text = FormatNumber(CDbl(.txtPrice.Text), 2, vbUseDefault, vbUseDefault, vbTrue)
   
Case Else
    Exit Sub
End Select
Exit Sub
Err:
    ErrorMessage
End With
End Sub
Private Sub GetRetailCost()
'On Error GoTo Err
With Me
    Set rsFindRecord = cnINVENT.Execute("SELECT * FROM ProductsCostPriceSetup WHERE DrugCode='" & Trim(.txtItemCode.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing
    Else
        .txtPrice.Text = rsFindRecord!RetailCost & ""
        .txtQuantity.Text = 1
        
    End If
End With
Exit Sub
Err:
    ErrorMessage
End Sub

Private Sub GetWholeSaleCost()
'On Error GoTo Err
With Me
    Set rsFindRecord = cnINVENT.Execute("SELECT * FROM ProductsCostPriceSetup WHERE DrugCode='" & Trim(.txtItemCode.Text) & "';")
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        Set rsFindRecord = Nothing
    Else
        .txtPrice.Text = rsFindRecord!DosageCost & ""
        .txtQuantity.Text = 1
        
    End If
End With
Exit Sub
Err:
    ErrorMessage
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo Err
Dim TotalSum, QuoteDate, RequisitionDate As Variant
With Me
RequisitionDate = Format(.txtRequistionDate.Text, "MMMM dd,yyyy")
Select Case Button.Key
Case "N"
    Select Case Button.Caption
    Case "New &Record "
        If EditRecord Then Exit Sub
        MyCommonData.ClearTextFields: .ListView1.ListItems.Clear: .ListView2.ListItems.Clear: .ListView3.ListItems.Clear: .txtJobDoneBy.SetFocus
        NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4
        .txtJobDoneBy.SetFocus
        .txtRequisitionNo.Text = AutoPurchaseOrderNo
        .txtRequistionDate.Text = MyCurrentDate
        .txtSupervisor.Text = CurrentUserName
        
    Case "&Save Record "
         If .optPrices(0).Value = True Then
         QuantityType = "WholeSale"
         ElseIf .optPrices(1).Value = True Then
         QuantityType = "Retail"
         End If
          
        If NewRecord Then
           If ItemsSelected Then
             If ItemsSelected2 Then
               If ItemsSelected3 Then
                  If ValidRecord Then
            NewSQL = "INSERT INTO AdvertJobCardTransactions(CategoryName,ProductName,QuantityType,TaxTotalCost,CategoryCode,PurchaseOrderNo,JobCardNo,RequisitionNo,ProductCode,DepartmentsCode,RequisitionDate,Qty,QuantityUnits,Cost,TotalCost,TaxAmount,CreatedBy,DateCreated,AccPeriod)VALUES('" & Trim(.txtCategoryCode.Text) & "','" & Trim(.txtItemName.Text) & "','" & QuantityType & "'," & CCur(.txtTaxTotalPrice.Text) & ",'" & .cboCategory.Text & "','" & PurchaseOrderNo & "','" & Trim(.txtJobCardNo.Text) & "','" & Trim(.txtRequisitionNo.Text) & "','" & Trim(.txtItemCode.Text) & "','" & Trim(.txtDeptCode.Text) & "','" & RequisitionDate & _
            "'," & Trim(.txtQuantity.Text) & ",'" & QuantityUnits & "'," & CCur(.txtPrice.Text) & "," & CCur(.txtTotalPrice.Text) & "," & CCur(.txtTaxAmount.Text) & ",'" & Trim(CurrentUserName) & "','" & Trim(MyCurrentDate) & "','" & Trim(MyCurrentPeriod) & "');"
            Set rsNewRecord = New ADODB.Recordset
            rsNewRecord.Open NewSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsNewRecord = Nothing
             Call RemoveCurrentList2Item
             Button.Caption = "NE&XT ITEM"
            .Toolbar1.Buttons(3).Caption = "FINISH"
            
                  End If
                End If
              End If
           End If
        End If
        
     Case "NE&XT ITEM"
          .txtItemCode.Text = ""
          .txtItemName.Text = ""
          .cboCategory.Text = ""
          .txtCategoryCode.Text = ""
          .cboTax.Text = ""
          .txtQuantity.Text = ""
          .txtPrice.Text = ""
          .txtTotalPrice.Text = ""
          .txtTaxAmount.Text = ""
          .txtTotalPrice.Text = ""
          .txtTaxTotalPrice.Text = ""
          .optPrices(0).Value = False
          .optPrices(1).Value = False
           
           Call DisableCostControls

'          RemoveCurrentList2Item
          Button.Caption = "&Save Record ": Button.Image = 4
    Case Else
        Exit Sub
    End Select
    
Case "E"
    Select Case Button.Caption
'    Case "&Edit/Change "
'    If NewRecord Then Exit Sub
'        If .txtCode.Text = Empty Then
'            MsgBox "There is NO Current Record to Edit. Please Search and Display a Record First...!", vbCritical + vbOKOnly, ""
'           .txtCode.SetFocus
'        Else
'           .txtCode.Locked = True
'            Button.Caption = "Save &Changes ": Button.Image = 4
'            EditRecord = True
'        End If
'    Case "Save &Changes "
'        If EditRecord Then
'        If ValidRecord Then
'        EditSQL = "Update ParamInsuranceBrokers SET BrokersName = '" & Trim(txtName.Text) & "'" & _
'        " ,Branch = '" & Trim(txtBranch.Text) & "'" & _
'        " ,PhysicalAddress = '" & Trim(txtAddress1.Text) & "'" & _
'        " ,PostalAddress = '" & Trim(txtAddress2.Text) & "'" & _
'        " ,TownCity = '" & Trim(cboTown.Text) & "'" & _
'        " ,Country = '" & Trim(cboCountry.Text) & "'" & _
'        " ,ContactPerson = '" & Trim(txtPerson.Text) & "'" & _
'        " ,ContactTitle = '" & Trim(cboTitle.Text) & "'" & _
'        " ,TelephoneNo = '" & Trim(txtPhone.Text) & "'" & _
'        " ,Email = '" & Trim(txtEmail.Text) & "' WHERE BrokersCode='" & Trim(txtCode.Text) & "';"
'
'            Set rsEditRecord = New ADODB.Recordset
'            rsEditRecord.Open EditSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
'            Set rsEditRecord = Nothing
'            .txtCode.Locked = False: EditRecord = False: Button.Caption = "&Edit/Change ": Button.Image = 5
'        End If
'        End If
     Case "FINISH"
         Dim NewOverallCost, NewTaxOverallCost As Variant
         If ValidMainRecord Then
         
            Dim RsFINDTOTALSUM As ADODB.Recordset
            Set RsFINDTOTALSUM = New ADODB.Recordset
              RsFINDTOTALSUM.Open "SELECT SUM(TotalCost) As OverallCost FROM AdvertJobCardTransactions WHERE JobCardNo = '" & Trim(.txtJobCardNo.Text) & "' AND RequisitionNo = '" & Trim(.txtRequisitionNo.Text) & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
              If RsFINDTOTALSUM.EOF And RsFINDTOTALSUM.BOF Then Set RsFINDTOTALSUM = Nothing: Exit Sub
              If Not IsNull(RsFINDTOTALSUM!overallCost) Then
                 NewOverallCost = RsFINDTOTALSUM!overallCost
                 
               End If
             End If
           Set RsFINDTOTALSUM = Nothing
           
         
            Set RsFINDTOTALSUM = New ADODB.Recordset
              RsFINDTOTALSUM.Open "SELECT SUM(TaxTotalCost) As TaxOverallCost FROM AdvertJobCardTransactions WHERE JobCardNo = '" & Trim(.txtJobCardNo.Text) & "' AND RequisitionNo = '" & Trim(.txtRequisitionNo.Text) & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
              If RsFINDTOTALSUM.EOF And RsFINDTOTALSUM.BOF Then Set RsFINDTOTALSUM = Nothing: Exit Sub
              If Not IsNull(RsFINDTOTALSUM!TaxOverallCost) Then
                 NewTaxOverallCost = RsFINDTOTALSUM!TaxOverallCost
                 
               End If
             
           Set RsFINDTOTALSUM = Nothing
              
            Set rsLineUpdate = New ADODB.Recordset
               rsLineUpdate.Open "UPDATE AdvertJobCardTransactions SET OverallTotalCost = " & CCur(NewOverallCost) & " WHERE JobCardNo = '" & Trim(.txtJobCardNo.Text) & "' AND RequisitionNo = '" & Trim(.txtRequisitionNo.Text) & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsLineUpdate = Nothing
            
            Set rsLineUpdate = New ADODB.Recordset
               rsLineUpdate.Open "UPDATE AdvertJobCardTransactions SET OverallTaxTotalCost = " & CCur(NewTaxOverallCost) & " WHERE JobCardNo = '" & Trim(.txtJobCardNo.Text) & "' AND RequisitionNo = '" & Trim(.txtRequisitionNo.Text) & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
            Set rsLineUpdate = Nothing
            
            
            Call RemoveCurrentList3Item
            Call RemoveCurrentList1Item
            
               
       .Toolbar1.Buttons(2).Caption = "New &Record "
       .Toolbar1.Buttons(2).Image = 2
       .Toolbar1.Buttons(3).Image = 5
       .Toolbar1.Buttons(3).Caption = "&Edit/Change "
             
'     End If
    Case Else
       
        Exit Sub
    End Select
Case "S"
'    If NewRecord Or EditRecord Then Exit Sub
'    INPQRY = InputBox("Please Enter the Brokers Code for the Record to Search and Display...!!!", "Enter Brokers Code...")
'    If Len(INPQRY) = 0 Then
'        MsgBox "Required Search Parameter Missing or the Operation Was Cancelled...! No Work was Done!!!", vbCritical + vbOKOnly, "Missing Parameter"
'        Exit Sub
'    Else
'        Set rsFindRecord = cnCOMMON.Execute("SELECT ParamInsuranceBrokers.* FROM ParamInsuranceBrokers WHERE ParamInsuranceBrokers.BrokersCode='" & Trim(INPQRY) & "';")
'        If rsFindRecord.EOF And rsFindRecord.BOF Then
'            MsgBox "Requested Record Missing or Has Been Deleted. Check your Entries to Ensure they are Accurately Spelt...!", vbOKOnly + vbExclamation, "Record NOT Found...!"
'            Set rsFindRecord = Nothing: Exit Sub
'        Else
'            .txtCode.Text = Trim(rsFindRecord!BrokersCode & "")
'            .txtName.Text = Trim(rsFindRecord!BrokersName & "")
'            .txtBranch.Text = Trim(rsFindRecord!Branch & "")
'            .txtAddress1.Text = Trim(rsFindRecord!PhysicalAddress & "")
'            .txtAddress2.Text = Trim(rsFindRecord!PostalAddress & "")
'            .cboTown.Text = Trim(rsFindRecord!TownCity & "")
'            .cboCountry.Text = Trim(rsFindRecord!Country & "")
'            .txtPerson.Text = Trim(rsFindRecord!ContactPerson & "")
'            .cboTitle.Text = Trim(rsFindRecord!ContactTitle & "")
'            .txtPhone.Text = Trim(rsFindRecord!TelephoneNo & "")
'            .txtEmail.Text = Trim(rsFindRecord!Email & "")
'
'        End If
'        Set rsFindRecord = Nothing
'    End If
Case "R"
    If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
        .Toolbar1.Buttons(2).Caption = "New &Record "
        .Toolbar1.Buttons(2).Image = 2
        .Toolbar1.Buttons(3).Caption = "&Edit/Change "
        .Toolbar1.Buttons(3).Image = 5
        NewRecord = False: EditRecord = False: MyCommonData.ClearTheScreen
Case "P"
'    Load frmRptAdvertPrintOut
'    frmRptAdvertPrintOut.Show 1, Me
Case "F"
     
     
Case Else
    Exit Sub
End Select
End With
Exit Sub
Err:
    If Err.Number = -2147217900 Then
    MsgBox "You have already requested the Inventory product under the current requisition you can either edit or place another requisiton for the product.....", vbExclamation, "Requisition Error"
    Else
    ErrorMessage
    End If
End Sub
Private Function ItemsSelected() As Boolean
'On Error GoTo Err
With Me
Dim i, j, k
j = .ListView1.ListItems.Count: k = 0

If j = 0 Or .ListView1.View <> lvwReport Then

    ItemsSelected = False
    
    strMessage = "There are no open job cards to Perform this Operation!!!"
    
Else

    For i = 1 To j
        If .ListView1.ListItems(i).Checked = True Then
            k = k + 1
        End If
    Next i
    
    If k = 0 Then
        ItemsSelected = False
        strMessage = "Please Select at Least ONE item from the open job card(s) datasheet!!"
    ElseIf k >= 1 Then
        ItemsSelected = True
    End If
    
End If
If Not ItemsSelected Then
    MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
End If
End With
Exit Function
Err:
    ErrorMessage
End Function
Private Function ItemsSelected2() As Boolean
'On Error GoTo Err
With Me
Dim i, j, k
j = .ListView2.ListItems.Count: k = 0

If j = 0 Or .ListView2.View <> lvwReport Then

    ItemsSelected2 = False
    
    strMessage = "There are no inventory items to display !!!"
    
Else

    For i = 1 To j
        If .ListView2.ListItems(i).Checked = True Then
            k = k + 1
        End If
    Next i
    
    If k = 0 Then
        ItemsSelected2 = False
        strMessage = "Please Select ONE item from list of Inventory Items Listed!!"
    ElseIf k >= 1 Then
        ItemsSelected2 = True
    End If
    
End If
If Not ItemsSelected2 Then
    MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
End If
End With
Exit Function
Err:
    ErrorMessage

End Function
Private Function ItemsSelected3()
'On Error GoTo Err
With Me
Dim i, j, k
j = .ListView3.ListItems.Count: k = 0

If j = 0 Or .ListView3.View <> lvwReport Then

    ItemsSelected3 = False
    
    strMessage = "There are no items under selected job card !!!"
    
Else

    For i = 1 To j
        If .ListView3.ListItems(i).Checked = True Then
            k = k + 1
        End If
    Next i
    
    If k = 0 Then
        ItemsSelected3 = False
        strMessage = "Please Select ONE item from list of items under selected job card!!"
    ElseIf k >= 1 Then
        ItemsSelected3 = True
    End If
    
End If
If Not ItemsSelected3 Then
    MsgBox strMessage, vbCritical + vbOKOnly, "Data Validation"
End If
End With
Exit Function
Err:
    ErrorMessage
End Function

Private Sub txtSidingCost_Change()

End Sub

Private Sub txtItemName_Change()
On Error GoTo Err
With Me
    If DontChange = True Then Exit Sub
   If .txtItemName.Text = Empty Then
         .ListView1.ListItems.Clear
    Else
     SearchByProductName
    End If

End With
Exit Sub
Err:
   ErrorMessage
End Sub
Private Sub SearchByProductName()
'On Error GoTo Err
With Me
.ListView2.ListItems.Clear
.ListView2.ColumnHeaders.Clear

.ListView2.ColumnHeaders.Add , , "Product Code", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Purchase Order No", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Product Name", .ListView2.Width / 4.5
.ListView2.ColumnHeaders.Add , , "Current Quantity", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Quantity Units", .ListView2.Width / 5.5
.ListView2.ColumnHeaders.Add , , "Current Total Pieces", .ListView2.Width / 5.5

.ListView2.View = lvwReport

Dim rsLIST As ADODB.Recordset
Set rsLIST = New ADODB.Recordset

rsLIST.Open "SELECT * FROM GenProductsInventory WHERE Discontinued = '" & "N" & "' AND DrugName LIKE '" & Trim(.txtItemName.Text) & "%' ORDER BY DrugName;", cnINVENT, adOpenKeyset, adLockOptimistic

Dim MyList As ListItem

If rsLIST.EOF And rsLIST.BOF Then
    .ListView2.View = lvwList
    Set MyList = .ListView2.ListItems.Add(, , "Search is Complete. There are No Records to Display in this View")
    Set rsLIST = Nothing: Set MyList = Nothing: Exit Sub
End If

While Not rsLIST.EOF

Set MyList = .ListView2.ListItems.Add(, , CStr(rsLIST!DrugCode))


    If Not IsNull(rsLIST!PurchaseOrderNo) Then
        MyList.SubItems(1) = CStr(rsLIST!PurchaseOrderNo)
    End If
     
    If Not IsNull(rsLIST!DrugName) Then
        MyList.SubItems(2) = CStr(rsLIST!DrugName)
    End If
    
    If Not IsNull(rsLIST!CurrentQuantity) Then
        MyList.SubItems(3) = CStr(rsLIST!CurrentQuantity)
    End If
    
    If Not IsNull(rsLIST!QuantityUnits) Then
        MyList.SubItems(4) = CStr(rsLIST!QuantityUnits)
    End If
    
    If Not IsNull(rsLIST!TotalPieces) Then
        MyList.SubItems(5) = CStr(rsLIST!TotalPieces)
    End If
    
    rsLIST.MoveNext
    
Wend



Set MyList = Nothing: Set rsLIST = Nothing

End With
Exit Sub
Err:
If Err.Number = 3265 Then Resume Next
    ErrorMessage


End Sub
Private Sub txtQuantity_Change()
On Error GoTo Err
With Me
.txtTotalPrice.Text = Val(.txtQuantity) * Val(.txtPrice.Text)
End With
Exit Sub
Err:
   ErrorMessage
End Sub

Private Sub txtQuotationNo_Change()

End Sub

Private Sub VScroll1_Change()
With Me
.txtQuantity.Text = .VScroll1.Value
End With
End Sub

Private Sub VScroll1_GotFocus()
With Me
.VScroll1.Value = 1
End With
End Sub
