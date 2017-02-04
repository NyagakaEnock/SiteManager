VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmODASMOpenJobCard 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7200
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "frmODASMOpenJobCard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Cumulative 
      Enabled         =   0   'False
      Height          =   1215
      Left            =   7080
      TabIndex        =   53
      Top             =   5880
      Width           =   4815
      Begin VB.TextBox txtRequisitionApproved 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   960
         TabIndex        =   58
         Top             =   735
         Width           =   1575
      End
      Begin VB.TextBox txtQuantityApproved 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3240
         TabIndex        =   57
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtRequisitionPrepared 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   960
         TabIndex        =   55
         Top             =   375
         Width           =   1575
      End
      Begin VB.TextBox txtQuantityPrepared 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3240
         TabIndex        =   54
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   3600
         TabIndex        =   61
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Cost"
         Height          =   255
         Left            =   1440
         TabIndex        =   60
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "Approval"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Order"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   390
         Width           =   1335
      End
   End
   Begin VB.TextBox txtRemarks 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   46
      Top             =   6720
      Width           =   6015
   End
   Begin VB.Frame Frame4 
      Caption         =   "Requisitions"
      Height          =   3135
      Left            =   7080
      TabIndex        =   26
      Top             =   2760
      Width           =   4815
      Begin VB.ComboBox cboCostingCode 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   960
         TabIndex        =   67
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox txtExchangeRate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   3720
         TabIndex        =   52
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtCurrencySymbol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3360
         TabIndex        =   51
         Top             =   1920
         Width           =   375
      End
      Begin VB.ComboBox cboCurrencyCode 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   3360
         TabIndex        =   50
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtRequisitionDate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3360
         TabIndex        =   45
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtRequisitionNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3360
         TabIndex        =   44
         Top             =   686
         Width           =   1335
      End
      Begin VB.TextBox txtTotalVATAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   3360
         TabIndex        =   41
         Top             =   2350
         Width           =   1335
      End
      Begin MSComCtl2.UpDown UpDownQuantity 
         Height          =   255
         Left            =   1920
         TabIndex        =   40
         Top             =   1530
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Max             =   1000
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtTotalItemCost 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   960
         TabIndex        =   37
         Top             =   2766
         Width           =   3735
      End
      Begin VB.TextBox txtVATAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   960
         TabIndex        =   35
         Top             =   2350
         Width           =   1215
      End
      Begin VB.TextBox txtItemCost 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   960
         TabIndex        =   33
         Top             =   1934
         Width           =   1215
      End
      Begin VB.TextBox txtItemQuantity 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   960
         TabIndex        =   31
         Top             =   1518
         Width           =   975
      End
      Begin VB.TextBox txtItemSize 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   960
         TabIndex        =   29
         Top             =   1102
         Width           =   1215
      End
      Begin VB.TextBox txtItemCode 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   960
         TabIndex        =   27
         Top             =   686
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Costing"
         Height          =   255
         Left            =   240
         TabIndex        =   68
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label29 
         Caption         =   "Total VAT"
         Height          =   255
         Left            =   2400
         TabIndex        =   66
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label28 
         Caption         =   "Currency"
         Height          =   255
         Left            =   2400
         TabIndex        =   65
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   "Curr Code"
         Height          =   255
         Left            =   2400
         TabIndex        =   64
         Top             =   1470
         Width           =   735
      End
      Begin VB.Label Label26 
         Caption         =   "Req Date"
         Height          =   255
         Left            =   2400
         TabIndex        =   63
         Top             =   1095
         Width           =   735
      End
      Begin VB.Label Label24 
         Caption         =   "Req No"
         Height          =   255
         Left            =   2400
         TabIndex        =   62
         Top             =   701
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Total Cost"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   2781
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "VAT"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   2365
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Cost"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1949
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1533
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Size"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1117
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Item"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   701
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Inventory Items"
      Height          =   2175
      Left            =   7080
      TabIndex        =   25
      Top             =   600
      Width           =   4815
      Begin MSComctlLib.ListView ListView2 
         Height          =   1815
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   3201
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
      Width           =   6855
      Begin VB.TextBox txtDateOfCompletion 
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
         Left            =   5040
         TabIndex        =   42
         Top             =   615
         Width           =   1695
      End
      Begin VB.TextBox txtDateOfCommencement 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1320
         TabIndex        =   19
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtSupervisedBy 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   5040
         TabIndex        =   18
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtDoneBy 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label15 
         Caption         =   "Date of Completion"
         Height          =   255
         Left            =   3600
         TabIndex        =   43
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
         Left            =   3600
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
      Width           =   6855
      Begin VB.TextBox txtTotalCost 
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
         Left            =   5040
         TabIndex        =   48
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtJobBriefDate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   5040
         TabIndex        =   23
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtDepartmentCode 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtLpono 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtDescriptionOfOrder 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   1320
         Width           =   5415
      End
      Begin VB.TextBox txtJobCardNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   5040
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtDeadlineDate 
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
         Left            =   1320
         TabIndex        =   5
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox txtCustomerName 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   960
         Width           =   5415
      End
      Begin VB.Label Label16 
         Caption         =   "Total Cost"
         Height          =   255
         Left            =   3960
         TabIndex        =   49
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Brief Date"
         Height          =   255
         Left            =   3960
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
         Left            =   3960
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
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   6855
      Begin MSComctlLib.ListView ListView1 
         Height          =   2415
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   4260
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
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         HelpCommand     =   11
         HelpContext     =   72
         HelpFile        =   "REGHELP.HLP"
      End
   End
   Begin VB.Label Label22 
      Caption         =   "Remarks"
      Height          =   255
      Left            =   120
      TabIndex        =   47
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
        SelectCurrencyGotFocus
End Sub

Private Sub cboCurrencyCode_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub cboCurrencyCode_LostFocus()
        selectCurrencyLostFocus
End Sub
Private Sub cboCostingCode_GotFocus()
        SelectCostingGotFocus
End Sub

Private Sub cboCostingCode_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub cboCostingCode_LostFocus()
        selectCostingLostFocus
End Sub


Private Sub Form_Activate()
        disableALLRECORD
        Set rsREQ = New clsODASRequisition
        rsREQ.LoadDEFAULT
        rsREQ.obtainBASECURRENCY
        rsREQ.calculateTOTALS
        showALLREQUISITIONSRAISED
'        disableFRAME
End Sub


Private Sub Form_Load()
        OpenODBCConnection
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
            
            Screen.ActiveForm.txtLandLordNo.Text = Item.Text
            Screen.ActiveForm.txtNames.Text = Item.SubItems(1)
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
            
            Screen.ActiveForm.txtItemCode.Text = Item.Text
            rsREQ.LoadItemCOST
            
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
                            If EditRecord Then Exit Sub
                            .ListView2.ListItems.Clear:
                            NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4:
                            enableFRAME
                            rsREQ.enableRECORD
                            rsREQ.clearRECORD
                    Case "&Save Record "
                    
                            If NewRecord Then
                                    rsREQ.generateRequisitionNO
                            End If
    
                            bSaveRECORD = False
                            rsREQ.ValidateRECORD
                            
                            If bSaveRECORD = True Then
                                    rsREQ.updateRECORD
                                    If bSaveRECORD = False Then
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

                        bSaveRECORD = False
                        rsREQ.ValidateRECORD
                        
                        If bSaveRECORD = True Then
                                rsREQ.updateRECORD
                                If bSaveRECORD = False Then
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
                NewRecord = False: EditRecord = False: clearALLRECORD

        Case "R"
            If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                .Toolbar1.Buttons(2).Caption = "New &Record "
                .Toolbar1.Buttons(2).Image = 2
                .Toolbar1.Buttons(3).Caption = "&Edit/Change "
                .Toolbar1.Buttons(3).Image = 5
                NewRecord = False: EditRecord = False: clearALLRECORD
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


Private Sub UpDownQuantity_Change()
'On Error GoTo err
        
        
        With frmODASMOpenJobCard
            If .txtExchangeRate.Text <= Empty Or CDbl(.txtItemCost.Text) <= 0 Then Exit Sub
            .txtItemQuantity.Text = .UpDownQuantity.Value
            .txtTotalItemCost.Text = CCur(.txtItemCost.Text) * CDbl(.txtItemQuantity.Text) * CDbl(.txtExchangeRate.Text)
            .txtTotalVATAmount.Text = CCur(.txtVATAmount.Text) * CDbl(.txtItemQuantity.Text) & CDbl(.txtExchangeRate.Text)
        
        End With

Exit Sub

err:
    ErrorMessage
End Sub
