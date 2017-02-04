VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmODASMPaymentConfirmation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Confirmation"
   ClientHeight    =   9570
   ClientLeft      =   150
   ClientTop       =   135
   ClientWidth     =   10695
   Icon            =   "frmODASMPaymentConfirmation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   10695
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
      TabIndex        =   38
      Top             =   6720
      Width           =   10455
      Begin VB.TextBox txtVoucherAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   8760
         TabIndex        =   44
         Top             =   1680
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1335
         Left            =   120
         TabIndex        =   39
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
         TabIndex        =   45
         Top             =   1680
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   10455
      Begin VB.TextBox txtContractYear 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7200
         TabIndex        =   61
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox txtContractNo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7200
         TabIndex        =   60
         Top             =   1800
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker DTPStartDate 
         Height          =   375
         Left            =   2760
         TabIndex        =   52
         Top             =   2040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   122748929
         CurrentDate     =   40961
      End
      Begin VB.TextBox txtCurrentPeriod 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtVoucherNo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1080
         TabIndex        =   26
         Top             =   405
         Width           =   1695
      End
      Begin VB.TextBox txtReference 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1440
         TabIndex        =   24
         Text            =   "Combo1"
         Top             =   5280
         Width           =   2535
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   405
         Width           =   3135
      End
      Begin VB.TextBox txtCostCenter 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1080
         TabIndex        =   22
         Top             =   750
         Width           =   1695
      End
      Begin VB.TextBox txtPaymentCodeDescription 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6000
         TabIndex        =   21
         Top             =   750
         Width           =   4335
      End
      Begin VB.TextBox txtAccountNo 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1095
         Width           =   1215
      End
      Begin VB.TextBox txtProductCode 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7200
         TabIndex        =   19
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtPayeeDetails 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1440
         Width           =   4815
      End
      Begin VB.TextBox txtPaymentDescription 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2280
         TabIndex        =   17
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox cboPaymentCode 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4080
         TabIndex        =   16
         Top             =   750
         Width           =   1815
      End
      Begin VB.TextBox txtRequisitionDate 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4080
         TabIndex        =   15
         Top             =   405
         Width           =   1815
      End
      Begin VB.TextBox txtItems 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1080
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPLastDate 
         Height          =   375
         Left            =   4680
         TabIndex        =   53
         Top             =   2040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   122748929
         CurrentDate     =   40961
      End
      Begin VB.Label Label23 
         Caption         =   "Last Date:"
         Height          =   255
         Left            =   4680
         TabIndex        =   56
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "Start Date:"
         Height          =   255
         Left            =   2760
         TabIndex        =   55
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Filter By Payment Requisition Dates:"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   2100
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Voucher No"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   435
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Payment Code"
         Height          =   255
         Left            =   2880
         TabIndex        =   36
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   " Date"
         Height          =   255
         Left            =   2880
         TabIndex        =   35
         Top             =   435
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Reference"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1125
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Label1"
         Height          =   255
         Left            =   600
         TabIndex        =   33
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "Status"
         Height          =   255
         Left            =   6120
         TabIndex        =   32
         Top             =   435
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Cost Center"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Payee"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1470
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Account No"
         Height          =   255
         Left            =   6120
         TabIndex        =   29
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Product Code"
         Height          =   255
         Left            =   6120
         TabIndex        =   28
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Items"
         Height          =   255
         Left            =   8520
         TabIndex        =   27
         Top             =   1110
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   8760
      Width           =   10455
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   42
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
         TabIndex        =   11
         Top             =   240
         Width           =   9015
      End
      Begin VB.Label Label15 
         Caption         =   "Remark"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Requisitions"
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
      TabIndex        =   8
      Top             =   3360
      Width           =   7215
      Begin VB.CheckBox chkSelectAll 
         Caption         =   "Select All"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   120
         TabIndex        =   9
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
      Caption         =   " Details"
      Height          =   3255
      Left            =   7440
      TabIndex        =   1
      Top             =   3360
      Width           =   3135
      Begin MSComCtl2.DTPicker DTPChequeDate 
         Height          =   375
         Left            =   1440
         TabIndex        =   59
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   122748931
         CurrentDate     =   40931
      End
      Begin VB.TextBox txtInvoiceBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         TabIndex        =   40
         Top             =   1905
         Width           =   1575
      End
      Begin VB.TextBox txtChequeNo 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   795
         Width           =   1575
      End
      Begin VB.TextBox txtInvoiceAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   1170
         Width           =   1575
      End
      Begin VB.TextBox txtTotalVoucherAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Cheque Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Invoice Balance  "
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1935
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Cheque No"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Invoice Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Total Requistion  "
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1620
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9600
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
            Picture         =   "frmODASMPaymentConfirmation.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPaymentConfirmation.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPaymentConfirmation.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPaymentConfirmation.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPaymentConfirmation.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPaymentConfirmation.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPaymentConfirmation.frx":236E
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
      Width           =   10695
      _ExtentX        =   18865
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
         TabIndex        =   51
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtExpiryDate 
         Height          =   285
         Left            =   9840
         TabIndex        =   50
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtReqNo 
         Height          =   285
         Left            =   9840
         TabIndex        =   49
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtJobDetails 
         Height          =   285
         Left            =   9840
         TabIndex        =   48
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtJobCardNo 
         Height          =   285
         Left            =   9840
         TabIndex        =   47
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCouncilCode 
         Height          =   285
         Left            =   9840
         TabIndex        =   43
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
Attribute VB_Name = "frmODASMPaymentConfirmation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsPAYREQ As clsPaymentRequisition


Private Sub chkSelectAll_Click()
If Me.chkSelectAll.Value = vbChecked Then
        checkAll Me.ListView1
Else
        UnCheckAll Me.ListView1
End If
End Sub

Private Sub DTPLastDate_CloseUp()
GetRentRequisitioned
End Sub

Private Sub DTPStartDate_CloseUp()
 GetRentRequisitioned
End Sub

Private Sub Form_Activate()
    NewRecord = False

    disableALLRECORD
    clearALLRECORD
    Me.DTPStartDate.Value = DateAdd("D", -7, Date)
    Me.DTPLastDate.Value = Date
    
    GetRentRequisitioned
        
    GetInvoicesREQUISITIONED
    
    CurrentPeriod
    Me.txtCurrentPeriod.Text = CurrentPeriod
    
    rsPAYREQ.loadPaymentDescription
    computeVOUCHERTOTAL
    countVOUCHERITEMS
    Me.txtCurrentPeriod = CurrentPeriod
End Sub

Private Sub Form_Initialize()
        Set rsPAYREQ = New clsPaymentRequisition
End Sub

Private Sub Form_Load()
Me.DTPChequeDate.Value = Date
Me.DTPChequeDate.MaxDate = Date

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

Private Sub listView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i, j As Double
    With Me
        If Item.Checked = True Then
            
        j = Screen.ActiveForm.ListView1.ListItems.Count
            
        If j = 0 Then Exit Sub
                checkOne Item, Me.ListView1
                
                Me.txtAccountNo.Text = Item.SubItems(4)
'                .cboDocumentNo.Text = Item.Text
'                .txtLPONo.Text = Item.SubItems(1)
                .txtExpiryDate.Text = Item.SubItems(2)
                .txtCouncilCode = ""
                .txtPayeeDetails.Text = Item.SubItems(5)
                .txtInvoiceAmount.Text = Item.SubItems(3)
                .txtInvoiceBalance.Text = 0
                .txtInstallmentNo = Item.SubItems(7)
                                
                strSQL = "SELECT * FROM ODASMInstallment WHERE InstallmentNo LIKE '" & .txtInstallmentNo.Text & "' "
                Set rsFindRecord = cnCOMMON.Execute(strSQL)
                If rsFindRecord.EOF Or rsFindRecord.BOF Then
                Else
                        Me.txtVoucherNo = rsFindRecord!vOUCHERnO & ""
                        Me.txtRequisitionDate = rsFindRecord!DateRequisitioned & ""
                        Me.txtInvoiceAmount = rsFindRecord!TotalRent & ""
                        
                         
                        strSQL = "SELECT * FROM ODASMVoucherItem WHERE vOUCHERnO LIKE '" & Me.txtVoucherNo & "' "
                        Set rsFIndVourcher = cnCOMMON.Execute(strSQL)
                       
                             If rsFIndVourcher.EOF Or rsFIndVourcher.BOF Then
                                Me.txtInvoiceBalance = rsFindRecord!Balance & ""
                             Else
                                 Me.txtInvoiceBalance = rsFIndVourcher!Balance & ""
                             End If
                         
                             If rsFIndVourcher.EOF Or rsFIndVourcher.BOF Then
                                Me.txtTotalVoucherAmount = rsFindRecord!AmountPaid & ""
                             Else
                                 Me.txtTotalVoucherAmount = rsFIndVourcher!AmountPaid & ""
                             End If
                        
                        Me.txtContractNo.Text = rsFindRecord!ContractNo & ""
                        Me.txtContractYear.Text = rsFindRecord!ContractYear & ""
'                        Me.txtInvoiceAmount = rsFindRecord!InvoiceAmount & ""
'                        Me.txtInvoiceAmount = rsFindRecord!InvoiceAmount & ""
                End If
                Dim strPrevPendingContracts As String
                strPrevPendingContracts = getprevPendingContracts
                If Len(strPrevPendingContracts) > 0 Then
                        MsgBox "There are pending installments (" & strPrevPendingContracts & ") that need to be updated before you proceed ", vbExclamation
                        Exit Sub
                End If
                
                If NewRecord = True Then
    '                    rsPAYREQ.SelectParticularCostCenter
                        NewRecord = False
                Else
                        Set rsCONTROL = New ADODB.Recordset
'                        rsPAYREQ.DocumentNoLostFocus
                End If
        Else
            Item.Checked = False
        End If
    End With
End Sub

Private Function getprevPendingContracts()
getprevPendingContracts = ""
strSQL = "SELECT * FROM ODASMInstallment WHERE  ContractNo='" & Me.txtContractNo.Text & "' AND ContractYear<" & CInt(Me.txtContractYear.Text) & " AND (PaymentFlag='N' OR PAymentFlag IS NULL)"
Set rsFindRecord1 = cnCOMMON.Execute(strSQL)
While Not rsFindRecord1.EOF
        getprevPendingContracts = getprevPendingContracts & rsFindRecord1!ContractYear & ", "
rsFindRecord1.MoveNext
Wend
End Function
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

   
 With Screen.ActiveForm
    Select Case Button.Key
    Case "N"
    
        Select Case Button.Caption
            Case "New &Record "
                NewRecord = True
                If editRECORD Then Exit Sub
                enableALLRECORD
                NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4
                
            Case "&Save Record "
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
                    ElseIf Trim(Me.txtChequeNo.Text) = "" Then
                            MsgBox "The Cheque No. is required", vbExclamation
                            Me.txtChequeNo.SetFocus
                            Exit Sub
                    End If
                    Dim strPrevPendingContracts As String
                    strPrevPendingContracts = getprevPendingContracts()
                    'If Len(strPrevPendingContracts) > 0 Then
                            'MsgBox "There are pending installments (" & strPrevPendingContracts & ") that need to be updated before you proceed ", vbExclamation
                            'Exit Sub
                    'End If
                    
                    For q = 1 To Me.ListView1.ListItems.Count
                        If Me.ListView1.ListItems(q).Checked = True Then
'                            Me.txtVoucherNo.Text = ""
'                            ListView1_ItemCheck Me.ListView1.ListItems(q)
                            rsPAYREQ.updatePaymentDetails
                            rsPAYREQ.saveINSTALLMENTISSUED
                        End If
                    Next q
                    GetRentRequisitioned
                    'disableALLRECORD
'                    Button.Caption = "NE&XT ITEM": Button.Image = 2
                    Button.Caption = "New &Record ": Button.Image = 2
    
                    Toolbar1.Buttons(3).Caption = "FINISH"
                
             Case "NE&XT ITEM"
                
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
                .txtInvoiceBalance.Text = CDbl(.txtInvoiceAmount.Text) - CDbl(.txtAmountPaid.Text)
                End If
               ' .txtVoucherAmount.Text = CDbl(.txtVoucherAmount) + CDbl(.txtAmountPaid.Text)
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
