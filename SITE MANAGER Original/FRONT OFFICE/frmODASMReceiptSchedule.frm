VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmODASMReceiptSchedule 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receipt Schedule"
   ClientHeight    =   6675
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12390
   Icon            =   "frmODASMReceiptSchedule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   12390
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Installments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   38
      Top             =   5520
      Width           =   6015
      Begin MSComCtl2.DTPicker DTPickerDueDate 
         Height          =   285
         Left            =   5520
         TabIndex        =   52
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Format          =   15794177
         CurrentDate     =   38412
      End
      Begin VB.TextBox txtPercentage 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   2280
         TabIndex        =   50
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1680
         TabIndex        =   42
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtInvoiceReference 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4080
         TabIndex        =   41
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtInstallmentNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1680
         TabIndex        =   40
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtPaymentDueDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4080
         TabIndex        =   39
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label22 
         Caption         =   "%"
         Height          =   255
         Left            =   2760
         TabIndex        =   51
         Top             =   255
         Width           =   135
      End
      Begin VB.Label Label10 
         Caption         =   "Installment Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   615
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Reference"
         Height          =   255
         Left            =   3120
         TabIndex        =   45
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "Installment No"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "Due Date"
         Height          =   255
         Left            =   3120
         TabIndex        =   43
         Top             =   255
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Receipt Schedule"
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
      Height          =   5895
      Left            =   6240
      TabIndex        =   28
      Top             =   720
      Width           =   6015
      Begin MSComctlLib.ListView ListView1 
         Height          =   5535
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9763
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
   Begin VB.Frame Frame3 
      Caption         =   "Schedule Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   21
      Top             =   3240
      Width           =   6015
      Begin VB.ComboBox cboInstallmentType 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   4080
         TabIndex        =   53
         Top             =   585
         Width           =   1815
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   48
         Top             =   1680
         Width           =   4575
      End
      Begin VB.TextBox txtTransactionType 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4080
         TabIndex        =   47
         Top             =   960
         Width           =   1815
      End
      Begin MSComCtl2.UpDown UpDownDuration 
         Height          =   255
         Left            =   1800
         TabIndex        =   36
         Top             =   1320
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         Max             =   9999
         Min             =   1
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox cboDurationMode 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   4080
         TabIndex        =   35
         Top             =   1290
         Width           =   1815
      End
      Begin VB.TextBox txtCurrentPeriod 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4080
         TabIndex        =   31
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox cboPaymentMode 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1320
         TabIndex        =   30
         Top             =   585
         Width           =   1335
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtAmountCreated 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   23
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtDuration 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1320
         TabIndex        =   22
         Top             =   1305
         Width           =   495
      End
      Begin VB.Label Label26 
         Caption         =   "Inst Type"
         Height          =   255
         Left            =   3000
         TabIndex        =   55
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "Type"
         Height          =   255
         Left            =   3000
         TabIndex        =   54
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Remark"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Duration"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Duration Mode "
         Height          =   255
         Left            =   3000
         TabIndex        =   34
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Created"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Current Period"
         Height          =   255
         Left            =   3000
         TabIndex        =   32
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Status"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   255
         Width           =   975
      End
      Begin VB.Label Label12 
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1575
         Width           =   1215
      End
      Begin VB.Label Label31 
         Caption         =   "Payment Mode"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   615
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   6015
      Begin VB.CheckBox ChkRestore 
         Caption         =   "Restore"
         Height          =   255
         Left            =   4920
         TabIndex        =   56
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   3480
         TabIndex        =   19
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtExpiryDate 
         BackColor       =   &H00FFFFFF&
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
         Left            =   4200
         TabIndex        =   16
         Top             =   1005
         Width           =   1695
      End
      Begin VB.TextBox txtProductCode 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   630
         Width           =   4455
      End
      Begin VB.TextBox txtPriceInclusive 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtJobBriefDate 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4200
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtDescriptionOfOrder 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   1770
         Width           =   4455
      End
      Begin VB.TextBox txtJobBriefNo 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtCommencementDate 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   3
         Top             =   1005
         Width           =   1335
      End
      Begin VB.TextBox txtCompanyName 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   1395
         Width           =   4455
      End
      Begin VB.Label Label19 
         Caption         =   "Balance"
         Height          =   255
         Left            =   2880
         TabIndex        =   20
         Top             =   2175
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Total Cost (Incl)"
         Height          =   255
         Left            =   2880
         TabIndex        =   18
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Expires"
         Height          =   255
         Left            =   2880
         TabIndex        =   17
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Product"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Price Inclusive"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2175
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Brief Date"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Job Brief No"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Desc of Order"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1785
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "Commencement"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label Label25 
         Caption         =   "Customer"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1410
         Width           =   855
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6840
      Top             =   240
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
            Picture         =   "frmODASMReceiptSchedule.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMReceiptSchedule.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMReceiptSchedule.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMReceiptSchedule.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMReceiptSchedule.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMReceiptSchedule.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMReceiptSchedule.frx":236E
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
      Width           =   12390
      _ExtentX        =   21855
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
   Begin VB.TextBox txtNoOfMonths 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   1440
      TabIndex        =   57
      Top             =   3720
      Width           =   975
   End
   Begin MSComCtl2.UpDown UpDownInstallment 
      Height          =   255
      Left            =   1920
      TabIndex        =   62
      Top             =   4095
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Value           =   1
      Max             =   999
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtInstallments 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   1560
      TabIndex        =   59
      Top             =   4080
      Width           =   375
   End
   Begin VB.CheckBox chkEqualInstallment 
      Caption         =   "Equal Installment?"
      Height          =   255
      Left            =   2520
      TabIndex        =   60
      Top             =   4095
      Width           =   1695
   End
   Begin VB.CheckBox chkAfterCommencementDate 
      Caption         =   "After DOC?"
      Height          =   255
      Left            =   4320
      TabIndex        =   61
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label20 
      Caption         =   "# of Months"
      Height          =   255
      Left            =   480
      TabIndex        =   58
      Top             =   3735
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "# of Installments"
      Height          =   255
      Left            =   360
      TabIndex        =   63
      Top             =   4080
      Width           =   1215
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
Attribute VB_Name = "frmODASMReceiptSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsreceipt As clsODASReceiptSchedule
Dim rsINVOICE As clsODASMAccounts
Private Sub cboDurationMode_GotFocus()
        SelectDurationCodeGotFocus
End Sub
Private Sub cboDurationMode_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub
Private Sub cboDurationMode_LostFocus()
        SelectDurationCodeLostFocus
End Sub
Private Sub cboInstallmentType_GotFocus()
        Screen.ActiveForm.cboInstallmentType.AddItem "AMOUNT"
        Screen.ActiveForm.cboInstallmentType.AddItem "PERCENT"
End Sub
Private Sub cboInstallmentType_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub
Private Sub cboInstallmentType_LostFocus()
On Error GoTo err
    With frmODASMReceiptSchedule
        
        If .cboInstallmentType = "AMOUNT" Then
            .cboInstallmentType = "A"
            .txtPercentage.Locked = False
            .txtPercentage.Text = 0
            .txtAmount.Text = FormatNumber(.txtBalance.Text)
            .txtAmount.Locked = False
            
        ElseIf .cboInstallmentType = "PERCENT" Then
            .cboInstallmentType = "P"
            .txtPercentage.Locked = True
            .txtAmount.Locked = True
        End If
    
    End With
    
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub cboPaymentMode_GotFocus()
        selectPaymentModeGotFocus
End Sub

Private Sub cboPaymentMode_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub cboPaymentMode_LostFocus()
        selectPaymentModeLostFocus
End Sub

Private Sub chkEqualInstallment_Click()
        Set rsreceipt = New clsODASReceiptSchedule
        rsreceipt.calculateAMOUNT
        Set rsreceipt = Nothing
        
        With frmODASMReceiptSchedule
            If .chkEqualInstallment.Value = 0 Then
                    .txtPercentage.Locked = False
                    .txtAmount.Locked = False
                    .DTPickerDueDate.Enabled = True
            Else
                    .txtPercentage.Locked = True
                    .txtAmount.Locked = True
                    .DTPickerDueDate.Enabled = False
            End If
        End With
End Sub


Private Sub ChkRestore_Click()
  With Me
   Set rsreceipt = New clsODASReceiptSchedule
    rsreceipt.loadRECORD
  End With
End Sub

Private Sub DTPickerDueDate_Change()
On Error GoTo err
    With frmODASMReceiptSchedule
        .txtPaymentDueDate.Text = .DTPickerDueDate.Value
    End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Form_Activate()
        Set rsreceipt = New clsODASReceiptSchedule
        Set rsINVOICE = New clsODASMAccounts
        disableALLRECORD
        rsreceipt.loadAMOUNTCREATED
        rsreceipt.loadRECORD
        rsreceipt.LoadDEFAULT
        showALLInstallments
End Sub

Private Sub Form_Initialize()
        Set rsreceipt = New clsODASReceiptSchedule
End Sub

Private Sub Form_Load()
        OpenConnection
End Sub

Private Sub Form_Terminate()
        Set rsreceipt = Nothing
End Sub
Private Sub Form_Unload(cancel As Integer)
On Error GoTo err
        If NewRecord Or beditRECORD Then MsgBox "Data Entry or Edit in Progress! No Work was Done!", vbInformation + vbOKOnly, "Screen Unload": cancel = 1: Exit Sub
        Set rsreceipt = Nothing
        
        If NewRecord = True Then
            cancel = True
            MsgBox "Data entry in progress. Click Refresh to Cancel", vbCritical
        Else
            cancel = False
        End If

Exit Sub

err:
    ErrorMessage
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
            
            Screen.ActiveForm.txtRequisitionNo.Text = Item.Text
            Set rsreceipt = New clsODASReceiptSchedule
            rsreceipt.loadRECORD
            Set rsreceipt = Nothing

        Else
            Item.Checked = False
        End If
        

Exit Sub

err:
    ErrorMessage
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
        With frmODASMReceiptSchedule
        Set rsreceipt = New clsODASReceiptSchedule
        Select Case Button.Key
                Case "N"
                    Select Case Button.Caption
                    Case "New &Record "
                            If editRECORD Then Exit Sub
                            NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4:
                            rsreceipt.enableRECORD
                    Case "&Save Record "
                            bsaveRECORD = False
                            rsreceipt.updateRECORD
                            If bsaveRECORD = True Then
                                        bsaveRECORD = False
                                        .Toolbar1.Buttons(2).Caption = "New &Record ": Button.Image = 2
                                        .Toolbar1.Buttons(3).Caption = "&NEXT INSTALLMENT"
                                        .Toolbar1.Buttons(4).Caption = "FINISH"
                                          disableALLRECORD
                            End If
                    Case "&NEXT INSTALLMENT"
                            .Toolbar1.Buttons(1).Caption = "&Save Record"
                            rsreceipt.loadAMOUNTCREATED
                            rsreceipt.clearINSTALLMENT
                            rsreceipt.enableINSTALLMENT
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
                                .Toolbar1.Buttons(3).Caption = "&NEXT INSTALLMENT"
                                .Toolbar1.Buttons(4).Caption = "FINISH"
                                disableALLRECORD
                        End If
                
                Case "&NEXT INSTALLMENT"
                            .Toolbar1.Buttons(3).Caption = "&Save Record "
                            rsreceipt.loadAMOUNTCREATED
                            rsreceipt.clearINSTALLMENT
                            rsreceipt.enableINSTALLMENT
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
            CurrentRecord = .txtJobBriefNo.Text
            Load frmODASRPaymentInstallments
            frmODASRPaymentInstallments.Show vbModal
        Case "F"
            Me.HelpCommonDialog.DialogTitle = "Using the Main System"
            Me.HelpCommonDialog.HelpFile = App.HelpFile
            Me.HelpCommonDialog.HelpContext = 35
            Me.HelpCommonDialog.HelpCommand = cdlHelpContext
            Me.HelpCommonDialog.ShowHelp

        Case Else
            Exit Sub
        End Select
        
        Set rsreceipt = Nothing
        
End With
Exit Sub
err:
    ErrorMessage

End Sub

Private Sub txtPercentage_LostFocus()
On Error GoTo err
        With frmODASMReceiptSchedule
                
                If .cboInstallmentType.Text = "P" And CDbl(.txtAmountCreated.Text) <= CDbl(.txtBalance.Text) Then
                        .txtAmount.Text = FormatNumber(CDbl(.txtBalance) * CDbl(.txtPercentage.Text) / 100)
                Else: .txtAmount.Text = 0
                End If
                .txtRemark.SetFocus
        End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub txtPriceInclusive_Change()
 With Me
   .txtBalance.Text = .txtPriceInclusive.Text
 End With
End Sub

Private Sub UpDownDuration_Change()
On Error GoTo err
        With frmODASMReceiptSchedule
            .txtDuration.Text = .UpDownDuration.Value
            Set rsreceipt = New clsODASReceiptSchedule
            If .UpDownDuration.Value <= 0 Then Exit Sub
                rsreceipt.calculateAMOUNT
            Set rsreceipt = Nothing

        End With
Exit Sub

err:
    ErrorMessage
End Sub

Private Sub UpDownDuration_DownClick()
On Error GoTo err
        With frmODASMReceiptSchedule
            Set rsreceipt = New clsODASReceiptSchedule
            .txtInstallments.Text = .UpDownInstallment.Value
            rsreceipt.calculateAMOUNT
            Set rsreceipt = Nothing
            
        End With
Exit Sub

err:
    ErrorMessage

End Sub

Private Sub UpDownDuration_GotFocus()
    With Me
        .UpDownDuration.Min = 1
    End With
End Sub

Private Sub UpDownDuration_LostFocus()
On Error GoTo err
        With frmODASMReceiptSchedule
            Set rsreceipt = New clsODASReceiptSchedule
            .txtInstallments.Text = .UpDownInstallment.Value
            rsreceipt.calculateAMOUNT
            Set rsreceipt = Nothing
            
        End With
Exit Sub

err:
    ErrorMessage

End Sub



Private Sub UpDownInstallment_Change()
On Error GoTo err
        With frmODASMReceiptSchedule
            Set rsreceipt = New clsODASReceiptSchedule
            
            If .UpDownInstallment.Value <= 0 Then Exit Sub
                .txtInstallments.Text = .UpDownInstallment.Value
                rsreceipt.calculateAMOUNT
            Set rsreceipt = Nothing
            
        End With
Exit Sub

err:
    ErrorMessage

End Sub