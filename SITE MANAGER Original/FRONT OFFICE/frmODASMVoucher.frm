VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmODASMVoucher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prepare Voucher"
   ClientHeight    =   9570
   ClientLeft      =   150
   ClientTop       =   135
   ClientWidth     =   11595
   Icon            =   "frmODASMVoucher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   11595
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Invoices within This Requisitions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2055
      Left            =   120
      TabIndex        =   40
      Top             =   6720
      Width           =   10455
      Begin VB.TextBox txtVoucherAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8760
         TabIndex        =   48
         Top             =   1680
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1335
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   2355
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
      Begin VB.Label Label6 
         Caption         =   "Voucher Amount "
         Height          =   255
         Left            =   7440
         TabIndex        =   49
         Top             =   1680
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   10455
      Begin VB.TextBox txtContractNo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8880
         TabIndex        =   65
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtContractYear 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8880
         TabIndex        =   64
         Top             =   1800
         Width           =   1455
      End
      Begin VB.PictureBox DTPStartDate 
         Height          =   375
         Left            =   1080
         ScaleHeight     =   315
         ScaleWidth      =   1275
         TabIndex        =   58
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtCurrentPeriod 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtVoucherNo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1080
         TabIndex        =   28
         Top             =   405
         Width           =   1695
      End
      Begin VB.TextBox txtReference 
         BackColor       =   &H00FFC0C0&
         Height          =   555
         Left            =   4920
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   1800
         Width           =   3855
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1440
         TabIndex        =   26
         Text            =   "Combo1"
         Top             =   5280
         Width           =   2535
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   405
         Width           =   3135
      End
      Begin VB.TextBox txtCostCenter 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1080
         TabIndex        =   24
         Top             =   750
         Width           =   1695
      End
      Begin VB.TextBox txtPaymentCodeDescription 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6000
         TabIndex        =   23
         Top             =   750
         Width           =   4335
      End
      Begin VB.TextBox txtAccountNo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1095
         Width           =   1215
      End
      Begin VB.TextBox txtProductCode 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7200
         TabIndex        =   21
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtPayeeDetails 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1440
         Width           =   4815
      End
      Begin VB.TextBox txtPaymentDescription 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1080
         TabIndex        =   19
         Top             =   1080
         Width           =   4815
      End
      Begin VB.TextBox cboPaymentCode 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4080
         TabIndex        =   18
         Top             =   750
         Width           =   1815
      End
      Begin VB.TextBox txtRequisitionDate 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4080
         TabIndex        =   17
         Top             =   405
         Width           =   1815
      End
      Begin VB.TextBox txtItems 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1080
         Width           =   1335
      End
      Begin VB.PictureBox DTPLastDate 
         Height          =   375
         Left            =   2520
         ScaleHeight     =   315
         ScaleWidth      =   1275
         TabIndex        =   59
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label24 
         Caption         =   "Narration:"
         Height          =   255
         Left            =   4080
         TabIndex        =   66
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "Last Date:"
         Height          =   255
         Left            =   2520
         TabIndex        =   62
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "Start Date:"
         Height          =   255
         Left            =   1080
         TabIndex        =   61
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Filter By Due Dates:"
         Height          =   495
         Left            =   120
         TabIndex        =   60
         Top             =   1980
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Voucher No"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   435
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Payment Code"
         Height          =   255
         Left            =   2880
         TabIndex        =   38
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   " Date"
         Height          =   255
         Left            =   2880
         TabIndex        =   37
         Top             =   435
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Reference"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1125
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Label1"
         Height          =   255
         Left            =   600
         TabIndex        =   35
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "Status"
         Height          =   255
         Left            =   6120
         TabIndex        =   34
         Top             =   435
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Cost Center"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Payee"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1470
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Account No"
         Height          =   255
         Left            =   6120
         TabIndex        =   31
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Product Code"
         Height          =   255
         Left            =   6120
         TabIndex        =   30
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Items"
         Height          =   255
         Left            =   8520
         TabIndex        =   29
         Top             =   1110
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   8760
      Width           =   10455
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Visible         =   0   'False
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   240
         Width           =   9015
      End
      Begin VB.Label Label15 
         Caption         =   "Remark"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
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
      ForeColor       =   &H00C00000&
      Height          =   3255
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   7215
      Begin VB.CheckBox chkSelectAll 
         Caption         =   "Select All"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   4683
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
   Begin VB.Frame Frame1 
      Caption         =   "Invoice Details"
      Height          =   3255
      Left            =   7440
      TabIndex        =   1
      Top             =   3360
      Width           =   3135
      Begin VB.TextBox txtAmountPaid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1440
         TabIndex        =   50
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtLPONo 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   1096
         Width           =   1575
      End
      Begin VB.TextBox txtInvoiceBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         TabIndex        =   42
         Top             =   2380
         Width           =   1575
      End
      Begin VB.TextBox cboDocumentNo 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   668
         Width           =   1575
      End
      Begin VB.TextBox txtInvoiceAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   1524
         Width           =   1575
      End
      Begin VB.TextBox txtTotalVoucherAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox txtVoucherItemNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   " Amount Paid"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1950
         Width           =   1335
      End
      Begin VB.Label Label20 
         Caption         =   "LPO No"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1125
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Invoice Balance  "
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   2415
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Document No"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Invoice Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Total Requestion  "
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2820
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Item No"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   1335
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
            Picture         =   "frmODASMVoucher.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMVoucher.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMVoucher.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMVoucher.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMVoucher.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMVoucher.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMVoucher.frx":236E
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
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   1164
      ButtonWidth     =   3307
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
      Begin VB.TextBox txtInstallmentNo 
         Height          =   285
         Left            =   9840
         TabIndex        =   57
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtExpiryDate 
         Height          =   285
         Left            =   9840
         TabIndex        =   56
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtReqNo 
         Height          =   285
         Left            =   9840
         TabIndex        =   55
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtJobDetails 
         Height          =   285
         Left            =   9840
         TabIndex        =   54
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtJobCardNo 
         Height          =   285
         Left            =   9840
         TabIndex        =   53
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCouncilCode 
         Height          =   285
         Left            =   9840
         TabIndex        =   47
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComDlg.CommonDialog HelpCommonDialog 
         Left            =   6720
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         HelpCommand     =   11
         HelpContext     =   72
         HelpFile        =   "REGHELP.HLP"
      End
   End
End
Attribute VB_Name = "frmODASMVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsPAYREQ As clsPaymentRequisition
Dim strPrevPendingContracts As String

Private Sub chkSelectAll_Click()
If Me.chkSelectAll.Value = vbChecked Then
        checkAll Me.ListView1
Else
        UnCheckAll Me.ListView1
End If
End Sub

Private Sub DTPLastDate_CloseUp()
 rsPAYREQ.SelectCostCenter
End Sub

Private Sub DTPStartDate_CloseUp()
 rsPAYREQ.SelectCostCenter
End Sub

Private Sub DTPLastDate_Click()

End Sub

Private Sub Form_Activate()
    NewRecord = False

    disableALLRECORD
    rsPAYREQ.clearRECORD
    
    If Me.cboPaymentCode.Text = "OTP" Then
         frmODASMVoucher.txtPayeeDetails.BackColor = &HFFC0C0: frmODASMVoucher.txtPayeeDetails.Locked = False
    End If
    
    If bapproveREQUISITION = True Or bAuthorizeREQUISITION = True Then
        rsPAYREQ.LoadRequisition
        rsPAYREQ.loadClaimDescription
    Else
        rsPAYREQ.loadPAYMENTRECORD
        rsPAYREQ.SelectCostCenter
        
        GetInvoicesREQUISITIONED
    End If
    
    CurrentPeriod
    Me.txtCurrentPeriod.Text = CurrentPeriod
    
    rsPAYREQ.loadPaymentDescription
    computeVOUCHERTOTAL
    countVOUCHERITEMS
    Me.txtCurrentPeriod = CurrentPeriod
End Sub

Private Function getprevPendingContracts()
getprevPendingContracts = ""
strSQL = "SELECT * FROM ODASMInstallment WHERE  ContractNo='" & Me.txtContractNo.Text & "' AND ContractYear<" & CInt(Me.txtContractYear.Text) & " AND (Requisitioned='N' OR Requisitioned IS NULL)"
Debug.Print strSQL
Set rsFindRecord1 = cnCOMMON.Execute(strSQL)
While Not rsFindRecord1.EOF
        getprevPendingContracts = getprevPendingContracts & rsFindRecord1!ContractYear & "- [" & rsFindRecord1!PaymentDueDate & "], "
rsFindRecord1.MoveNext
Wend
End Function

Private Sub Form_Initialize()
        Set rsPAYREQ = New clsPaymentRequisition
End Sub

Private Sub Form_Load()
'Me.DTPLastDate.Value = DateAdd("M", 2, Date)
'Me.DTPStartDate.Value = DateAdd("M", -2, Date)
Me.DTPLastDate.Value = Date
Me.DTPStartDate.Value = Date
End Sub

Private Sub Form_Terminate()
    Set rsPAYREQ = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If NewRecord = True Then
        Cancel = True
        MsgBox "Data entry in progress. Click Refresh to Cancel", vbCritical
    Else
        Cancel = False
    End If
End Sub

Private Sub ListView3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    
    ListView3.SortKey = ColumnHeader.Index - 1
    ListView3.Sorted = True
    
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i, j As Double
   
    If Item.Checked = True Then
        
    j = Screen.ActiveForm.ListView1.ListItems.Count
        
    If j = 0 Then Exit Sub
        
'    For i = 1 To j
'        If Screen.ActiveForm.ListView1.ListItems(i) <> Item Then
'           Screen.ActiveForm.ListView1.ListItems(i).Checked = False
'        End If
'    Next i
        
            If Me.txtCostCenter = "OTP" Then
                   frmODASMVoucher.cboDocumentNo.Text = Item.Text
                   frmODASMVoucher.txtAmountPaid.Text = Item.SubItems(3)
                   frmODASMVoucher.txtInvoiceAmount.Text = Item.SubItems(3)
                   frmODASMVoucher.txtJobCardNo.Text = Item.SubItems(2)
                   frmODASMVoucher.txtRemark.Text = Item.SubItems(5)
                   frmODASMVoucher.txtCouncilCode.Text = Item.SubItems(1)
                   frmODASMVoucher.txtReference.Text = Item.SubItems(1)
                   frmODASMVoucher.txtReqNo.Text = Item.SubItems(4)
                 
                     loadAllowance
                     loadJobBriefAccountNo
                     loadAccountName
            Else
                Me.txtAccountNo.Text = Item.SubItems(4)
                frmODASMVoucher.cboDocumentNo.Text = Item.Text
                frmODASMVoucher.txtLPONo.Text = Item.SubItems(1)
                frmODASMVoucher.txtExpiryDate.Text = Item.SubItems(2)
                frmODASMVoucher.txtCouncilCode = ""
                frmODASMVoucher.txtAmountPaid.Text = Item.SubItems(3)
                frmODASMVoucher.txtInvoiceAmount.Text = Item.SubItems(3)
                frmODASMVoucher.txtVoucherAmount = 0
                
                frmODASMVoucher.txtInvoiceBalance.Text = 0
                If bLoadRecord = True Then
                    txtReference.Text = Item.SubItems(6)
                End If
                txtInstallmentNo.Text = Item.SubItems(7)
            
                strSQL = "SELECT * FROM ODASMInstallment WHERE InstallmentNo LIKE '" & txtInstallmentNo.Text & "' "
                Set rsFindRecord = cnCOMMON.Execute(strSQL)
                If rsFindRecord.EOF Or rsFindRecord.BOF Then
                Else
                        Me.txtContractNo.Text = rsFindRecord!ContractNo & ""
                        Me.txtContractYear.Text = rsFindRecord!ContractYear & ""
                End If
                strPrevPendingContracts = getprevPendingContracts
                If Len(strPrevPendingContracts) > 0 Then
                        MsgBox "There are pending installments (" & strPrevPendingContracts & ") for contract no [" & txtContractNo.Text & "] that need to be updated before you proceed ", vbExclamation
                        Exit Sub
                End If
            
            
            End If
            If NewRecord = True Then
'                    rsPAYREQ.SelectParticularCostCenter
                    NewRecord = False
            Else
                    Set rsCONTROL = New ADODB.Recordset
                    rsPAYREQ.DocumentNoLostFocus
            End If
    Else
        Item.Checked = False
    End If
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo err
    
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err

   
 With frmODASMAccounts
    Select Case Button.Key
    Case "N"
    
        Select Case Button.Caption
            Case "New &Record "
                NewRecord = True
                If editRECORD Then Exit Sub
                enableALLRECORD
                Me.txtVoucherNo.Text = ""
                bLoadRecord = True
                NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4
                
            Case "&Save Record "
                    If Me.txtInvoiceAmount > Me.txtAmountPaid Then
                               PartialPaid = Me.txtAmountPaid
                    End If
                    Dim q As Integer
                    Dim valid As Boolean
                    valid = False
                    For j = 1 To Me.ListView1.ListItems.Count
                            If Me.ListView1.ListItems(j).Checked = True Then
                                    valid = True
                                    Exit For
                            End If
                    Next j
                    If valid = False Then
                            MsgBox "Select at least one Record before you proceed", vbExclamation
                            Exit Sub
                    End If
                    For q = 1 To Me.ListView1.ListItems.Count
                        If Me.ListView1.ListItems(q).Checked = True Then
                            Me.txtVoucherNo.Text = ""
                            bLoadRecord = False
                            ListView1_ItemCheck Me.ListView1.ListItems(q)
                            If Len(strPrevPendingContracts) > 0 Then
                                    Exit Sub
                            End If
                            
                            rsPAYREQ.updateRECORD
                        End If
                    Next q
                    rsPAYREQ.SelectCostCenter
                    'disableALLRECORD
'                    Button.Caption = "NE&XT ITEM": Button.Image = 2
                    Button.Caption = "New &Record ": Button.Image = 2
    
                    Toolbar1.Buttons(3).Caption = "FINISH"
                
             Case "NE&XT ITEM"
                bLoadRecord = True
                    Button.Caption = "&Save Record ": Button.Image = 4
                    rsPAYREQ.clearRECORD
                    rsPAYREQ.enableRECORD
                     
                    NewRecord = False
            Case Else
                    Exit Sub
        End Select
    
      Case "E"
        Select Case Button.Caption
            Case "FINISH"
                Toolbar1.Buttons(2).Caption = "New &Record "
                Toolbar1.Buttons(2).Image = 2
                Toolbar1.Buttons(3).Caption = "&Edit/Change "
                Toolbar1.Buttons(3).Image = 5
                NewRecord = False: editRECORD = False:
            End Select
      Case "S"
                
      Case "R"
            If MsgBox(RefreshMessage, vbExclamation + vbOKCancel + vbDefaultButton2, "Refresh The Screen") = vbCancel Then Exit Sub
                Me.Toolbar1.Buttons(2).Caption = "New &Record "
                Me.Toolbar1.Buttons(2).Image = 2
                Me.Toolbar1.Buttons(3).Caption = "&Edit/Change "
                Me.Toolbar1.Buttons(3).Image = 5
               NewRecord = False: editRECORD = False:
        Case "P"
                If frmODASMVoucher.txtVoucherNo.Text >= "" Then
                    CurrentRecord = Me.txtVoucherNo
                    
                    If frmODASMVoucher.txtCouncilCode <> Empty Then
                        Load frmRatesVoucher
                        frmRatesVoucher.Show 1, Me
                    Else
                        Load frmPayRequisition
                        frmPayRequisition.Show 1, Me
                    End If
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

    End With

Exit Sub
err:
    ErrorMessage

End Sub
Private Sub txtAmountPaid_Change()
        If NewRecord = True Then Exit Sub
        With frmODASMVoucher
                computeVOUCHERTOTAL
                If .txtInvoiceAmount.Text <= Empty Then Exit Sub
                If .txtAmountPaid.Text <= Empty Then Exit Sub
                If frmODASMVoucher.txtCostCenter = "OTP" Then
                  .txtInvoiceBalance.Text = 0
                  .txtInvoiceBalance.Locked = True
                Else:
                         If .txtVoucherAmount = "" Then
                            txtVoucherAmount = 0
                            
                         End If
                .txtInvoiceBalance.Text = CDbl(.txtInvoiceAmount.Text) - CDbl(.txtAmountPaid.Text)
                .txtVoucherAmount = CDbl(.txtAmountPaid.Text)
                'CDbl(.txtAmountPaid.Text)
                '.txtVoucherAmount.Text = CDbl(.txtAmountPaid.Text) - CDbl(.txtVoucherAmount)
                '.txtVoucherAmount.Text = CDbl(.txtAmountPaid.Text) - CDbl(.txtVoucherAmount)
                End If
        End With
End Sub

Private Sub txtVoucherAmount_LostFocus()
    rsPAYREQ.calculateVOUCHERSUM
End Sub
Public Sub loadAllowance()
On Error GoTo err
    
    Set rsCONTROL = New ADODB.Recordset
    rsCONTROL.Open "SELECT * FROM ODASPAdminCosting WHERE CostItem = '" & frmODASMVoucher.txtReference.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With frmODASMVoucher
                
            If rsCONTROL.EOF Or rsCONTROL.BOF = True Then Exit Sub
            frmODASMVoucher.txtPaymentDescription.Text = rsCONTROL!CostingItemName
    
    End With
    
    rsCONTROL.Close

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadJobBriefAccountNo()
On Error GoTo err
    
    Set rsCONTROL = New ADODB.Recordset
    rsCONTROL.Open "SELECT * FROM ODASMJobBrief WHERE JobBriefNo = '" & frmODASMVoucher.txtJobCardNo.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With frmODASMVoucher
                
            If rsCONTROL.EOF Or rsCONTROL.BOF = True Then Exit Sub
            frmODASMVoucher.txtAccountNo.Text = rsCONTROL!AccountNo
    
    End With
    
    rsCONTROL.Close

Exit Sub

err:
    ErrorMessage
End Sub
Public Sub loadAccountName()
On Error GoTo err
    
    Set rsCONTROL = New ADODB.Recordset
    rsCONTROL.Open "SELECT * FROM ODASPAccount WHERE AccountNo = '" & frmODASMVoucher.txtAccountNo.Text & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With frmODASMVoucher
                
            If rsCONTROL.EOF Or rsCONTROL.BOF = True Then Exit Sub
            frmODASMVoucher.txtJobDetails.Text = rsCONTROL!CompanyName
    
    End With
    
    rsCONTROL.Close

Exit Sub

err:
    ErrorMessage
End Sub
