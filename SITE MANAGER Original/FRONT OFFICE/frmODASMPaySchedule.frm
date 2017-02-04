VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmODASMPaySchedule 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pay JobBrief Installments"
   ClientHeight    =   5895
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12630
   Icon            =   "frmODASMPaySchedule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   12630
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTotalBalance 
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
      Left            =   9960
      TabIndex        =   24
      Top             =   5520
      Width           =   2295
   End
   Begin VB.TextBox txtTotalAmountPaid 
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
      Left            =   5760
      TabIndex        =   22
      Top             =   5520
      Width           =   2415
   End
   Begin VB.TextBox txtTotalAmountDue 
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
      Left            =   1560
      TabIndex        =   20
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Caption         =   "Payment Details"
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
      Left            =   5640
      TabIndex        =   19
      Top             =   2880
      Width           =   6735
      Begin VB.TextBox txtInstalmentYear 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4680
         TabIndex        =   39
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtTransactionAmount 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1200
         TabIndex        =   38
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H00FFC0C0&
         Height          =   915
         Left            =   1200
         MaxLength       =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Top             =   1440
         Width           =   5415
      End
      Begin VB.TextBox txtBalance 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4680
         TabIndex        =   35
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtPaymentDue 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4680
         TabIndex        =   33
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtDueDate 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1200
         TabIndex        =   31
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtInvoiceNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1200
         TabIndex        =   28
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label14 
         Caption         =   "Installment Year"
         Height          =   255
         Left            =   3360
         TabIndex        =   40
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Remark"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Balance"
         Height          =   255
         Left            =   3600
         TabIndex        =   34
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Expected Amount"
         Height          =   255
         Left            =   3360
         TabIndex        =   32
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "InvoiceNo"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Amount Paid"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblDueDate 
         Caption         =   "Due Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   855
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
      Height          =   4695
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   5415
      Begin MSComctlLib.ListView ListView1 
         Height          =   4335
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   7646
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
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   5640
      TabIndex        =   1
      Top             =   720
      Width           =   6735
      Begin VB.TextBox txtAccountNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         TabIndex        =   30
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtExpiryDate 
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
         Left            =   4920
         TabIndex        =   14
         Top             =   1005
         Width           =   1695
      End
      Begin VB.TextBox txtProductCode 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   630
         Width           =   5295
      End
      Begin VB.TextBox txtJobBriefDate 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4920
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtDescriptionOfOrder 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   1680
         Width           =   5295
      End
      Begin VB.TextBox txtJobBriefNo 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtCommencementDate 
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
         TabIndex        =   3
         Top             =   1005
         Width           =   1335
      End
      Begin VB.TextBox txtCompanyName 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   2640
         TabIndex        =   2
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label Label15 
         Caption         =   "Total Cost (Incl)"
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Expires"
         Height          =   255
         Left            =   3960
         TabIndex        =   15
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Product"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Brief Date"
         Height          =   255
         Left            =   4200
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
         Top             =   1680
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
         Top             =   1320
         Width           =   855
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6960
      Top             =   600
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
            Picture         =   "frmODASMPaySchedule.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPaySchedule.frx":0ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPaySchedule.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPaySchedule.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPaySchedule.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPaySchedule.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmODASMPaySchedule.frx":236E
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
      Width           =   12630
      _ExtentX        =   22278
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
   Begin VB.Label Label10 
      Caption         =   "Total Balance"
      Height          =   255
      Left            =   8520
      TabIndex        =   25
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Total Amount Paid"
      Height          =   255
      Left            =   4200
      TabIndex        =   23
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Total Amount Due"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   5535
      Width           =   1695
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
Attribute VB_Name = "frmODASMPaySchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsreceipt As clsODASReceiptSchedule
Dim rsINVOICE As clsODASMAccounts
Private Sub enableMyRecord()
With Me
        .txtTransactionAmount.Enabled = True
        .txtTransactionAmount.Locked = False
        .txtRemark.Enabled = True
        .txtRemark.Locked = False
End With
End Sub
Private Sub disableMyRecord()
With Me
        .txtTransactionAmount.Enabled = False
        .txtTransactionAmount.Locked = True
        .txtRemark.Enabled = False
        .txtRemark.Locked = True
End With
End Sub
Private Sub validateRECORD()
On Error GoTo err
    With Me
        bsaveRECORD = False
        If .txtAccountNo.Text = Empty Then
                MsgBox "The Company Account Cannot be Left Blank ..."
                .txtAccountNo.SetFocus
        ElseIf .txtInvoiceNo.Text = Empty Then
                MsgBox "The InvoiceNo is required"
               .txtInvoiceNo.SetFocus
        ElseIf .txtJobBriefNo.Text = Empty Then
                MsgBox "The JobBriefNo is required"
               .txtJobBriefNo.SetFocus
        ElseIf .txtPaymentDue.Text = Empty Then
                MsgBox "The Payment Due is required ..."
               .txtPaymentDue.SetFocus
        ElseIf CDbl(.txtInstalmentYear.Text) > CDbl(Year(Date)) Then
                MsgBox "The Transaction that You have Selected is not Scheduled for this Year", vbCritical, "Kindly Choose another transaction."
               .txtInstalmentYear.SetFocus
        ElseIf .txtBalance.Text = Empty Then
                MsgBox "The Balance is required ..."
               .txtBalance.SetFocus
        ElseIf .txtTransactionAmount.Text = Empty Then
                MsgBox "The Amount Paid is required ..."
               .txtTransactionAmount.SetFocus
        Else
                bsaveRECORD = True
        End If
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Private Function NextPAYMENT() As Currency
On Error GoTo err
 With Me
    Dim curINVOICE, nextINVOICE, LASTNo As String
    curINVOICE = Me.txtInvoiceNo.Text
    LASTNo = Right(curINVOICE, 1)
         
          nextINVOICE = .txtJobBriefNo & "-" & (LASTNo + 1)
            
            Set rsFind = New ADODB.Recordset
            strSQL = "select * from ODASMJobBriefInstallment Where invoiceNo = '" & nextINVOICE & "' ;"
            rsFind.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            If rsFind.EOF Or rsFind.BOF Then
                NextPAYMENT = 0
            Else
             NextPAYMENT = rsFind!PaymentDue
            End If
            'strSQL = Empty
            Set rsFind = Nothing
 End With
Exit Function
err:
    ErrorMessage
End Function
Private Sub updateJBBALANCE()
On Error GoTo err
  With Me
        Dim rsUPDATE As ADODB.Recordset
        Set rsSAVE = New Recordset
        strSQL = "SELECT * FROM ODASMJobBrief Where JobBriefNo = '" & .txtJobBriefNo.Text & "';"
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        With rsSAVE
                If .BOF Or .EOF = True Then Exit Sub
                !ReceivedToDate = Me.txtTotalAmountPaid.Text
                !Balance = Me.txtTotalBalance.Text
                !AmountInvoiced = Me.txtPaymentDue.Text
                !NextInvoiceAmount = NextPAYMENT
                .Update
                .Requery
        End With
   End With
'rsCONTROL.Close
rsSAVE.Close
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub SUMInstallments()
On Error GoTo err
 With Me
     Dim rsSUM As ADODB.Recordset
     Set rsSUM = New ADODB.Recordset
     
        strSQL = "Select sum(AmountPaid) as Amountpaid,sum(Balance) as Balance From ODASMJobBriefInstallment WHERE JobBriefNo = '" & .txtJobBriefNo & "'"
        rsSUM.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
          If rsSUM.EOF Or rsSUM.BOF Then
                         .txtTotalAmountPaid.Text = 0
                         .txtTotalBalance.Text = 0
          ElseIf IsNull(rsSUM!AmountPaid) = True Then
                        .txtTotalAmountPaid = 0
      
          ElseIf IsNull(rsSUM!Balance) = True Then
                        .txtTotalBalance = 0
          Else
              .txtTotalAmountPaid.Text = rsSUM!AmountPaid
              .txtTotalBalance.Text = rsSUM!Balance
              .txtTotalAmountPaid = FormatNumber(txtTotalAmountPaid, 2)
              .txtTotalBalance = FormatNumber(txtTotalBalance, 2)
          End If
            
     Set rsSUM = Nothing
 End With
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub SAVEInstallmentACCOUNT()
On Error GoTo err
       With frmODASMPaySchedule
                Set rsSAVE = New ADODB.Recordset
                strSQL = "select * from ODASMJobBriefInstallmentTotal Where JobBriefNo = '" & .txtJobBriefNo & "' ;"
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
                    rsSAVE!JobBriefNo = .txtJobBriefNo
                    rsSAVE!Amount = .txtTotalAmountDue
                    rsSAVE!AmountPaid = .txtTotalAmountPaid
                    rsSAVE!Balance = .txtTotalBalance
                    rsSAVE!CurrentPeriod = CurrentPeriod
                    rsSAVE.Update
                    rsSAVE.Requery
               
                Set rsSAVE = Nothing
                strSQL = Empty
    End With
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub createInstallmentACCOUNT()
On Error GoTo err
       With frmODASMPaySchedule
       
                Set rsSAVE = New ADODB.Recordset
                strSQL = "select * from ODASMJobBriefInstallmentTotal Where JobBriefNo = '" & .txtJobBriefNo & "' ;"
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
                If rsSAVE.EOF Or rsSAVE.BOF Then
                    rsSAVE.AddNew
                    rsSAVE!JobBriefNo = .txtJobBriefNo
                    rsSAVE!Amount = 0
                    rsSAVE!AmountPaid = 0
                    rsSAVE!Balance = 0
                    rsSAVE!CurrentPeriod = CurrentPeriod
                    rsSAVE!dateprepared = Date
                    rsSAVE!Preparedby = CurrentUserName
                    rsSAVE.Update
                    rsSAVE.Requery
                Else
                    Exit Sub
                End If
                
                Set rsSAVE = Nothing
                strSQL = Empty
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub updateINSTALLMENT()
On Error GoTo err
       With frmODASMPaySchedule
                strSQL = "select * from ODASMJobBriefInstallment Where InvoiceNo = '" & .txtInvoiceNo & "' ;"
                Set rsSAVE = New ADODB.Recordset
                rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
            
                If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
                    rsSAVE!AmountPaid = CDbl(.txtTransactionAmount)
                    rsSAVE!Balance = CDbl(rsSAVE!Balance) - CDbl(.txtTransactionAmount)
                    rsSAVE!Paid = "Y"
                rsSAVE.Update
                rsSAVE.Requery
                
                
                Set rsSAVE = Nothing
                strSQL = Empty
    End With
Exit Sub

err:
    ErrorMessage
End Sub
Private Sub loadINSTALLMENT()
On Error GoTo err
    With Me
         Set rsFindRecord = New ADODB.Recordset
         rsFindRecord.Open "SELECT * FROM ODASMJobBriefInstallment WHERE invoiceNo = '" & .txtInvoiceNo & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
         If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub
             .txtPaymentDue = FormatNumber(rsFindRecord!PaymentDue, 2) & ""
             .txtDueDate = Format(rsFindRecord!PaymentDueDate, "dd/mm/yyyy") & ""
             .txtTransactionAmount = FormatNumber(rsFindRecord!PaymentDue, 2) & ""
             .txtBalance = FormatNumber(rsFindRecord!Balance, 2) & ""
             .txtInstalmentYear = Year(.txtDueDate)
    End With
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub loadJOBBRIEF()
On Error GoTo err
    With frmODASMPaySchedule
        
        Set rsCONTROL = New ADODB.Recordset
        strSQL = "SELECT * FROM ODASMJobBrief JB, ODASPAccount AC  WHERE JB.AccountNo = AC.AccountNo and JB.JobBriefNo = '" & .txtJobBriefNo.Text & "' ; "
        rsCONTROL.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsCONTROL.EOF Or rsCONTROL.BOF Then Exit Sub
                .txtCompanyName.Text = rsCONTROL!CompanyName
                .txtAccountNo.Text = rsCONTROL!AccountNo
                .txtDescriptionOfOrder.Text = rsCONTROL!descriptionOfOrder
                .txtProductCode.Text = rsCONTROL!ProductCode & ""
                .txtJobBriefDate.Text = rsCONTROL!JobBriefDate
                If Not IsNull(rsCONTROL!CommencementDate) Then
                    .txtCommencementDate.Text = rsCONTROL!CommencementDate
                Else
                    .txtCommencementDate.Text = Date
                End If
                .txtTotalAmountDue.Text = FormatNumber(rsCONTROL!TotalPrice) & ""
                If Not IsNull(rsCONTROL!ReceivedToDate) Then
                    .txtTotalAmountPaid.Text = rsCONTROL!ReceivedToDate
                Else
                    .txtTotalAmountPaid.Text = 0
                End If
                .txtTotalBalance.Text = FormatNumber(rsCONTROL!Balance) & ""
                .txtTotalAmountPaid.Text = FormatNumber(rsCONTROL!ReceivedToDate) & ""
                .txtExpiryDate.Text = rsCONTROL!expirydate & ""
                
                
                
    End With

Exit Sub
err:
    ErrorMessage
End Sub

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
        loadJOBBRIEF
        SUMInstallments
        ListALLInstallments
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


Private Sub Label31_Click()

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
            Me.txtInvoiceNo.Text = Item.SubItems(1)
            loadINSTALLMENT
        Else
            'Item.Checked = False
        End If
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
        With frmODASMPaySchedule
        Select Case Button.Key
                Case "N"
                    Select Case Button.Caption
                    Case "New &Record "
                            If editRECORD Then Exit Sub
                            enableMyRecord
                            NewRecord = True: Button.Caption = "&Save Record ": Button.Image = 4:
                    Case "&Save Record "
                            bsaveRECORD = False
                            validateRECORD
                            If bsaveRECORD = True Then
                                createInstallmentACCOUNT
                                updateINSTALLMENT
                                SUMInstallments
                                SAVEInstallmentACCOUNT
                                updateJBBALANCE
                                ListALLInstallments
                                bsaveRECORD = False
                                    .Toolbar1.Buttons(2).Caption = "New &Record ": Button.Image = 2
                                    .Toolbar1.Buttons(3).Caption = "&NEXT INSTALLMENT"
                                    .Toolbar1.Buttons(4).Caption = "FINISH"
                                    disableALLRECORD
                            End If
                    Case "&NEXT INSTALLMENT"
                            .Toolbar1.Buttons(1).Caption = "&Save Record"
                           
                    Case Else
                        Exit Sub
                    End Select
        Case "E"
            Select Case Button.Caption
                Case "&Edit/Change "
                
                Case "&Save Record "

                        bsaveRECORD = False
                       
                        If bsaveRECORD = True Then
                                
                                bsaveRECORD = False
                                .Toolbar1.Buttons(2).Caption = "New &Record "
                                .Toolbar1.Buttons(3).Caption = "&NEXT INSTALLMENT"
                                .Toolbar1.Buttons(4).Caption = "FINISH"
                                disableALLRECORD
                        End If
                
                Case "&NEXT INSTALLMENT"
                            .Toolbar1.Buttons(3).Caption = "&Save Record "
                           
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
'            ALISFOManager.txtTASK.Text = "INST"
'            showALLCLOSEDCOSTINGBRIEFS
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






