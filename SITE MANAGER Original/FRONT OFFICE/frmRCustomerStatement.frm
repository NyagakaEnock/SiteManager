VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmRCustomerStatement 
   Caption         =   "CUSTOMER STATEMENTS"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmRCustomerStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New RCustomerStatement

''
Private Sub Form_Load()
'Report.RecordSelectionFormula = "{ODASPAccount.AccountNo } = '" & CurrentRecord & "'"
''Screen.MousePointer = vbHourglass
'
'    Set rsFindRecord = New ADODB.Recordset
'    rsFindRecord.Open "Select * From ODASMJobBriefInstallment JI,ODASPAccount A, ODASMInvoiceSent I Where I.AccountNo = A.AccountNo and A.CompanyName like '%" & CurrentRecord1 & "%' and JI.InvoiceNo = I.InvoiceNo", cnCOMMON, adOpenKeyset, adLockOptimistic
'    If rsFindRecord.EOF And rsFindRecord.BOF Then GoTo Continue
'
'    CurrentRecord = rsFindRecord!AccountNo
'
'    rsFindRecord.MoveFirst
'    Do While rsFindRecord.EOF <> True
'
'        ref = "INV " & rsFindRecord!InvoiceNo
'        Set rsNewRecord = New ADODB.Recordset
'        rsNewRecord.Open "Select * From ODASMCustomerStatement Where AccountNo ='" & rsFindRecord!AccountNo & "' and TransactionDate='" & Format(rsFindRecord!InvoiceDate, "MMMM dd,YYYY") & "' and Reference = '" & ref & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
'
'        If rsNewRecord.EOF And rsNewRecord.BOF Then
'            rsNewRecord.AddNew
'
'            rsNewRecord!AccountNo = rsFindRecord!AccountNo
'            rsNewRecord!TransactionDate = rsFindRecord!InvoiceDate
'            rsNewRecord!Reference = "INV " & rsFindRecord!InvoiceNo
'        End If
'
'            rsNewRecord!Details = "Being part payment (Installment No. " & rsFindRecord!InstallmentNo & " )"
'            rsNewRecord!DebitAmount = rsFindRecord!Amount
'            rsNewRecord!CreditAmount = 0
'            rsNewRecord!Transactionby = CurrentUserName
'
'            rsNewRecord.Update
'        rsFindRecord.MoveNext
'     Loop
'Continue:
'    Set rsFindRecord = New ADODB.Recordset
'    rsFindRecord.Open "Select * From ALISMReceiptNew R,ODASPAccount A, ALISMReceiptDetails RD Where R.ReceiptNo = RD.ReceiptNo and R.AccountNo = A.AccountNo and A.CompanyName like '%" & CurrentRecord1 & "%' ", cnCOMMON, adOpenKeyset, adLockOptimistic
'    If rsFindRecord.EOF And rsFindRecord.BOF Then GoTo Continue2
'
'    CurrentRecord = rsFindRecord!AccountNo
'
'    rsFindRecord.MoveFirst
'    Do While rsFindRecord.EOF <> True
'
'        ref = "RECEIPT " & rsFindRecord!ReceiptNo
'        Set rsNewRecord = New ADODB.Recordset
'        rsNewRecord.Open "Select * From ODASMCustomerStatement Where AccountNo ='" & rsFindRecord!AccountNo & "' and TransactionDate='" & Format(rsFindRecord!ReceiptDate, "MMMM dd,YYYY") & "' and Reference = '" & ref & "' ", cnCOMMON, adOpenKeyset, adLockOptimistic
'
'        If rsNewRecord.EOF And rsNewRecord.BOF Then
'            rsNewRecord.AddNew
'
'            rsNewRecord!AccountNo = rsFindRecord!AccountNo
'            rsNewRecord!TransactionDate = rsFindRecord!ReceiptDate
'            rsNewRecord!Reference = ref
'        End If
'
'            rsNewRecord!Details = "Being part payment (Invoice No. " & rsFindRecord!InvoiceNo & " )"
'            rsNewRecord!DebitAmount = 0
'            rsNewRecord!CreditAmount = rsFindRecord!ReceiptAmount
'            rsNewRecord!Transactionby = CurrentUserName
'
'            rsNewRecord.Update
'        rsFindRecord.MoveNext
'     Loop
''Continue2:
  Screen.MousePointer = vbHourglass
        Dim Balance As Variant
        Set rsNewRecord = New ADODB.Recordset
        rsNewRecord.Open "Select * From ODASMCustomerStatement Where AccountNo ='" & CurrentRecord & "' order by TransactionDate asc", cnCOMMON, adOpenKeyset, adLockOptimistic

        If rsNewRecord.EOF And rsNewRecord.BOF Then GoTo Continue3
        rsNewRecord.MoveFirst
        num = 1
        Do While rsNewRecord.EOF <> True
            Set rsSAVE = New ADODB.Recordset
            rsSAVE.Open "Select * From ODASMCustomerStatement Where AccountNo ='" & CurrentRecord & "'and Reference = '" & rsNewRecord!Reference & "' and TransactionDate = '" & Format(rsNewRecord!TransactionDate, "MMMM dd,YYYY") & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
            rsSAVE!TransactionNo = Int(num)

                Set rsBalance = New ADODB.Recordset
                rsBalance.Open "Select * From ODASMCustomerStatement Where AccountNo ='" & CurrentRecord & "' and TransactionNo = " & num - 1 & " ", cnCOMMON, adOpenKeyset, adLockOptimistic
                    If rsBalance.RecordCount = 0 Then
                        rsSAVE!Balance = rsSAVE!DebitAmount - rsSAVE!CreditAmount
                    Else
                        rsSAVE!Balance = (rsBalance!Balance + rsSAVE!DebitAmount) - rsSAVE!CreditAmount
                    End If
            rsSAVE.Update
        rsNewRecord.MoveNext: num = num + 1
        Loop
Continue3:
CRViewer1.ReportSource = Report
CRViewer1.ViewReport
    Report.RecordSelectionFormula = "{ODASPAccount.AccountNo}= '" & CurrentRecord & "'"
Screen.MousePointer = vbDefault

End Sub
'Private Sub Form_Load()
'    Set rsFindRecord = New ADODB.Recordset
'    Set rsBalance = New ADODB.Recordset
'    rsFindRecord.Open "Select * From ODASPAccount Where AccountNo = '" & CurrentRecord & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
'    If rsFindRecord.EOF And rsFindRecord.BOF Then Exit Sub
'
'    Set rsFindRecord = Nothing
'    Screen.MousePointer = vbHourglass
'    Report.RecordSelectionFormula = "{ODASPAccount.AccountNo } = '" & CurrentRecord & "'"
'    CRViewer1.ReportSource = Report
'    CRViewer1.ViewReport
'    Screen.MousePointer = vbDefault
'End Sub
Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub
