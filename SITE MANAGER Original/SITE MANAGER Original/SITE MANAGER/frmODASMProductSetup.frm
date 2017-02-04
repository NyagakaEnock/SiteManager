VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmODASMProductSetup 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7065
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11235
   Icon            =   "frmODASMProductSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   11235
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Purchase Order Items"
      Height          =   2535
      Left            =   5640
      TabIndex        =   42
      Top             =   720
      Width           =   5415
      Begin VB.TextBox txtOrderDescription 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1080
         TabIndex        =   66
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtOrderDate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   3600
         TabIndex        =   59
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtOrderNo 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1080
         TabIndex        =   58
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtProductCode 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1080
         TabIndex        =   50
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtProductDescription 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   2160
         TabIndex        =   49
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtReceivedTotalUnitCostInclusive 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   3600
         TabIndex        =   48
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtReceivedTotalUnitCost 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1080
         TabIndex        =   47
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtReceivedVATAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   3600
         TabIndex        =   46
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtReceivedUnitCost 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1080
         TabIndex        =   45
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtReceivedUnitQuantity 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1080
         TabIndex        =   44
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtReceivedVATRate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   3600
         TabIndex        =   43
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Order Desc"
         Height          =   255
         Left            =   240
         TabIndex        =   67
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Order Date"
         Height          =   255
         Left            =   2520
         TabIndex        =   61
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Order No"
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   1695
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Cost"
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   1350
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "VAT"
         Height          =   255
         Left            =   2520
         TabIndex        =   55
         Top             =   1710
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Total Cost"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   2055
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Product"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Total (Inc)"
         Height          =   255
         Left            =   2520
         TabIndex        =   52
         Top             =   2070
         Width           =   1335
      End
      Begin VB.Label Label28 
         Caption         =   "VAT Rate"
         Height          =   255
         Left            =   2520
         TabIndex        =   51
         Top             =   1350
         Width           =   1335
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Mark up Details"
      Height          =   1815
      Left            =   5640
      TabIndex        =   29
      Top             =   3240
      Width           =   5415
      Begin VB.TextBox txtProductCost 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1200
         TabIndex        =   63
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtUnitType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   3600
         TabIndex        =   62
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtMarkUp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   3600
         TabIndex        =   35
         Top             =   1365
         Width           =   1575
      End
      Begin VB.TextBox txtProductPrice 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1200
         TabIndex        =   34
         Top             =   1365
         Width           =   1095
      End
      Begin VB.TextBox txtTotalQuantity 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   3600
         TabIndex        =   33
         Top             =   615
         Width           =   1575
      End
      Begin VB.TextBox txtMarkupType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1200
         TabIndex        =   32
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtQuantity 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1200
         TabIndex        =   31
         Top             =   615
         Width           =   1095
      End
      Begin VB.TextBox txtMarkupAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   3600
         TabIndex        =   30
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Unit Type"
         Height          =   255
         Left            =   2400
         TabIndex        =   65
         Top             =   990
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Product Cost"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label32 
         Caption         =   "Quantity/Set"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Label31 
         Caption         =   "Mark up Type"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label30 
         Caption         =   "Total Quantity"
         Height          =   255
         Left            =   2400
         TabIndex        =   39
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label29 
         Caption         =   "Product Price"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1395
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "Mark Up"
         Height          =   255
         Left            =   2400
         TabIndex        =   37
         Top             =   1395
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Mark Up Amt"
         Height          =   255
         Left            =   2400
         TabIndex        =   36
         Top             =   270
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Delivery Details"
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   5415
      Begin VB.TextBox txtGRNQuantity 
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
         Height          =   315
         Left            =   3960
         TabIndex        =   27
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtGRNTotalCostInclusive 
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
         Height          =   315
         Left            =   1440
         TabIndex        =   25
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtGRNTotalVATAmount 
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
         Height          =   315
         Left            =   3960
         TabIndex        =   23
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtGRNTotalCost 
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
         Height          =   315
         Left            =   1440
         TabIndex        =   21
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtInvoiceDate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   3960
         TabIndex        =   17
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtInvoiceNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1440
         TabIndex        =   16
         Top             =   1080
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPickerDeliveryDate 
         Height          =   315
         Left            =   5040
         TabIndex        =   14
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57278465
         CurrentDate     =   38339
      End
      Begin VB.TextBox txtDeliveryNoteNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtDeliveryDate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   3960
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtGRNDate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   3960
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtGRNNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPickerGRNDate 
         Height          =   315
         Left            =   5040
         TabIndex        =   15
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57278465
         CurrentDate     =   38339
      End
      Begin MSComCtl2.DTPicker DTPickerInvoiceDate 
         Height          =   315
         Left            =   5040
         TabIndex        =   20
         Top             =   1080
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57278465
         CurrentDate     =   38339
      End
      Begin VB.Label Label27 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   2880
         TabIndex        =   28
         Top             =   1815
         Width           =   1335
      End
      Begin VB.Label Label26 
         Caption         =   "Total Cost Inc"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1815
         Width           =   1575
      End
      Begin VB.Label Label24 
         Caption         =   "VAT Amt"
         Height          =   255
         Left            =   2880
         TabIndex        =   24
         Top             =   1455
         Width           =   1335
      End
      Begin VB.Label Label21 
         Caption         =   "Total Cost"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "Invoice Date"
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   1095
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "Invoice No"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1095
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Delivery Note No"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   390
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Delivery Date"
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   390
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "GRN Date"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   735
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "GRN No"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   735
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Inventory Transactions"
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   10935
      Begin MSComctlLib.ListView ListView3 
         Height          =   1335
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   10695
         _ExtentX        =   18865
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
   Begin VB.Frame Frame1 
      Caption         =   "Purchase Order Entries"
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
      TabIndex        =   1
      Top             =   3360
      Width           =   5415
      Begin MSComctlLib.ListView ListView2 
         Height          =   1335
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
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
            Picture         =   "frmODASMProductSetup.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMProductSetup.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMProductSetup.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMProductSetup.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMProductSetup.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMProductSetup.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMProductSetup.frx":236E
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
      Width           =   11235
      _ExtentX        =   19817
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
Attribute VB_Name = "frmODASMProductSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRECEIVE As clsODASReceiveOrder



Private Sub DTPickerDeliveryDate_Change()
'On Error GoTo err
    With frmODASMReceiveOrder
        .txtDeliveryDate.Text = .DTPickerDeliveryDate.Value
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub DTPickerGRNDate_Change()
'On Error GoTo err
    With frmODASMReceiveOrder
        .txtGRNDate.Text = .DTPickerGRNDate.Value
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub DTPickerInvoiceDate_Change()
'On Error GoTo err
    With frmODASMReceiveOrder
        .txtInvoiceDate.Text = .DTPickerInvoiceDate.Value
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Form_Activate()
        
        showORDERITEMSRECEIVED
        showORDERITEMSWITHBALANCE
        disableALLRECORD
        
        Set rsRECEIVE = New clsODASReceiveOrder
        rsRECEIVE.loadDELIVERYDETAILS
        rsRECEIVE.loadOrder
        Set rsRECEIVE = Nothing
        disableFRAME
End Sub


Private Sub Form_Load()
        OpenODBCConnection
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
                    bRequisitionAPPROVAL = False
                    frmODASMOperation.txtApplicationNo.Text = CurrentRecord
                    Set rsOPERATION = New clsODASOperation
                    bRequisitionAPPROVAL = True
                    GlobalDepartmentCode = .ListView1.SelectedItem.SubItems(1)
                    rsOPERATION.checkAPPROVEDDISCHARGE
                    If bRequisitionAPPROVAL = False Then Exit Sub
                    
                    rsOPERATION.approveOPERATION
                    Set rsOPERATION = Nothing
                    bRequisitionAPPROVAL = False
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
            
            Screen.ActiveForm.txtProductCode.Text = Item.Text
            Screen.ActiveForm.txtProductDescription.Text = Item.SubItems(2)
            Screen.ActiveForm.txtVATRate.Text = Item.SubItems(3)

            If Screen.ActiveForm.txtOrderDescription.Text = Empty Then
                    Screen.ActiveForm.txtOrderDescription.Text = "Supply of " + Trim(Screen.ActiveForm.txtProductDescription.Text)
            Else
                    Screen.ActiveForm.txtOrderDescription.Text = Trim(Screen.ActiveForm.txtOrderDescription.Text) + "," + Trim(Screen.ActiveForm.txtProductDescription.Text)
            End If
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
            
            Set rsRECEIVE = New clsODASReceiveOrder
            rsRECEIVE.clearGRNItems
            Screen.ActiveForm.txtProductCode.Text = Item.SubItems(1)
            rsRECEIVE.loadPuchaseOrderRECEIVED
            loadPRODUCT
            rsRECEIVE.CalculateMARKUP
            Set rsRECEIVE = Nothing
            
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
        .Frame3.Enabled = False
        .Frame5.Enabled = False
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub enableFRAME()
'On Error GoTo err
    
    With Screen.ActiveForm
        .Frame1.Enabled = True
        .Frame3.Enabled = True
        .Frame5.Enabled = True
    End With
Exit Sub

err:
    ErrorMessage
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo err
        
        With Screen.ActiveForm
        
        Select Case Button.Key
                Case "N"
                    Select Case Button.Caption
                    
                    Case "New &Record "
                        If EditRecord Then Exit Sub
                        NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4
                        enableALLRECORD
                        enableFRAME
                    
                    Case "&Save Record "
                        If NewRecord Then
                            Set rsRECEIVE = New clsODASReceiveOrder
                            
                            rsRECEIVE.ValidateRECORD
                            
                            If bSaveRECORD = True Then
                                    
                                    rsRECEIVE.updateRECORD
                            
                                    If bSaveRECORD = False Then
                                            Button.Caption = "&Next Item"
                                            rsRECEIVE.ClearPurchaseOrderItems
                                            frmODASMReceiveOrder.ListView2.SetFocus
                                            .Toolbar1.Buttons(3).Caption = "FINISH"
                                    End If
                            End If
                            
                            Set rsRECEIVE = Nothing

                        End If
                    
                    Case "&Next Item"
                            Button.Caption = "&Save Record ": Button.Image = 4
                            enableALLRECORD
                    Case Else
                        Exit Sub
                    End Select
    
        Case "E"
            Select Case Button.Caption
                Case "&Edit/Change "
                        editMYRECORD
                        Button.Caption = "&Save Record ": Button.Image = 4
                        enableALLRECORD

                Case "FINISH"
                        .Toolbar1.Buttons(2).Caption = "New &Record "
                        .Toolbar1.Buttons(2).Image = 2
                        .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                        .Toolbar1.Buttons(3).Image = 5
                        NewRecord = False: EditRecord = False: clearALLRECORD
                Case Else
            End Select
        
        Case "S"
            If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                .Toolbar1.Buttons(2).Caption = "New &Record "
                .Toolbar1.Buttons(2).Image = 2
                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                .Toolbar1.Buttons(3).Image = 5
                NewRecord = False: EditRecord = False: clearALLRECORD
  
        Case "R"
            If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                .Toolbar1.Buttons(2).Caption = "New &Record "
                .Toolbar1.Buttons(2).Image = 2
                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                .Toolbar1.Buttons(3).Image = 5
                NewRecord = False: EditRecord = False: clearALLRECORD
        Case "P"
        Case "F"
     
     
        Case Else
            Exit Sub
        End Select
End With
Exit Sub
err:
    ErrorMessage

End Sub


Private Sub txtExchangeRate_LostFocus()
        calculateTAXES
End Sub

Private Sub txtQuantity_LostFocus()
        If frmODASMProductSetup.txtQuantity.Text = Empty Then Exit Sub

        Set rsRECEIVE = New clsODASReceiveOrder
        rsRECEIVE.CalculateMARKUP
        Set rsRECEIVE = Nothing
End Sub


Private Sub txtReceivedUnitQuantity_LostFocus()
    computeTOTALCOST
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub txtReceivedVATRate_LostFocus()
'On Error GoTo err
    computeTOTALCOST
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub txtUnitCost_LostFocus()
'On Error GoTo err
        With frmODASMOrder
            If CDbl(.txtUnitQuantity.Text) > 0 And CDbl(.txtUnitCost.Text) > 0 Then
                        .txtVATAmount.Text = FormatCurrency(CDbl(.txtVATRate) / 100) * CDbl(.txtUnitCost) * CDbl(.txtUnitQuantity.Text)
            End If
        End With
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub computeTOTALCOST()
'On Error GoTo err
    With frmODASMReceiveOrder
        .txtReceivedTotalUnitCost.Text = CDbl(.txtReceivedUnitCost) * CDbl(.txtReceivedUnitQuantity)
        .txtReceivedVATAmount.Text = CDbl(.txtTotalUnitCost.Text) * CDbl(.txtReceivedVATRate) / 100
        .txtReceivedTotalUnitCostInclusive.Text = CDbl(.txtReceivedTotalUnitCost) + CDbl(.txtReceivedVATAmount)
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub UpDownQuantity_Change()
'On Error GoTo err
    With frmODASMReceiveOrder
        .txtReceivedUnitQuantity.Text = .UpDownQuantity.Value
    End With
    computeTOTALCOST
Exit Sub
err:
    ErrorMessage
End Sub
