VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "CRViewer.dll"
Begin VB.Form frmRStatement 
   Caption         =   "Landlord Statement"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
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
Attribute VB_Name = "frmRStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New rptStatement

Public strAccountNo As String

Private Sub Form_Load()
On Error GoTo errMSG:
Screen.MousePointer = vbHourglass: Dim myCrystalLogon As New clsCrystallogon
CRViewer1.ReportSource = Report: Report.Database.LogOnServer "pdsodbc.dll", DSN, Database, Uid, Pwd: CRViewer1.ReportSource = Report: DoEvents

If Trim(strAccountNo) = Empty Then
    Unload Me
    Exit Sub
End If
Report.SQLQueryString = "SELECT     ODASMInstallment.InstallmentNo, ODASMInstallment.PaymentDueDate, ODASMInstallment.AmountPaid, ODASMInstallment.PaymentDue, " & _
                      "ODASMInstallment.Balance, ODASMInstallment.ChequeDate, ODASMInstallment.ChequeNo, ODASPPlot.LRNo, ODASPAccount.AccountNo, " & _
                      "ODASPAccount.CompanyName " & _
"FROM         ODASMInstallment AS ODASMInstallment INNER JOIN " & _
                      "ODASPPlot AS ODASPPlot ON ODASMInstallment.ContractNo = ODASPPlot.PlotNo INNER JOIN " & _
                      "ODASPAccount AS ODASPAccount ON ODASPPlot.AccountNo = ODASPAccount.AccountNo WHERE ODASPAccount.AccountNo LIKE '" & strAccountNo & "'" & _
"ORDER BY ODASPAccount.AccountNo"

CRViewer1.ViewReport
Screen.MousePointer = vbDefault

Exit Sub
errMSG:
        ErrorMessage

End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub
