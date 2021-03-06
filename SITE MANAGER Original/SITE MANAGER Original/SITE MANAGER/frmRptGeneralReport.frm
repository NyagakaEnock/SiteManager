VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmRptGeneralReport 
   Caption         =   "Form1"
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
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmRptGeneralReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New RptODASGeneralReport

Private Sub Form_Load()
Dim f, G As String
Screen.MousePointer = vbHourglass: Dim myCrystalLogon As New clsCrystallogon
CRViewer1.ReportSource = Report: Report.Database.LogOnServer "pdsodbc.dll", DSN, Database, Uid, Pwd: CRViewer1.ReportSource = Report: DoEvents
Report.SQLQueryString = strSQL
strSQL1 = "SELECT * FROM ODASMInstallment,ODASPPlot,ODASPPlotMast,ODASPAccount,ODASMLeaseAgreement Where ODASMInstallment.PaymentDueDate >= '" & Format(Screen.ActiveForm.txtStartDate.Text, "yyyy/mm/dd") & "' and ODASMInstallment.PaymentDueDate <= '" & Format(Screen.ActiveForm.txtLastDate.Text, "yyyy/mm/dd") & "' and ODASMInstallment.PaymentFlag = 'N' and ODASPPlotMast.PlotNo=ODASPPlot.PlotNo and ODASPPlotMast.ContractNo=ODASMInstallment.ContractNo and ODASMInstallment.AccountNo=ODASPAccount.AccountNo and ODASMLeaseAgreement.ContractNo=ODASMInstallment.ContractNo and (ODASMLeaseAgreement.Terminated is null);"
strSQL2 = "SELECT * FROM ODASMInstallment,ODASPPlot,ODASPPlotMast,ODASPAccount,ODASMLeaseAgreement Where ODASMInstallment.PaymentDueDate >= '" & Format(Screen.ActiveForm.txtStartDate.Text, "yyyy/mm/dd") & "' and ODASMInstallment.PaymentDueDate <= '" & Format(Screen.ActiveForm.txtLastDate.Text, "yyyy/mm/dd") & "' and ODASMInstallment.PaymentFlag = 'N' and ODASPPlotMast.PlotNo=ODASPPlot.PlotNo and ODASPPlotMast.ContractNo=ODASMInstallment.ContractNo and ODASMInstallment.AccountNo=ODASPAccount.AccountNo and ODASMLeaseAgreement.ContractNo=ODASMInstallment.ContractNo and (ODASMLeaseAgreement.Terminated is null);"
strSQL3 = "SELECT *  FROM ODASPPlot, ODASPPLotMast, ODASPAccount where ODASPPLotMast.LeasePrepared = 'Y' and ODASPPLotMast.CommencementDate >= '" & Format(frmODASSearchSiteNewSites.txtStartDate.Text, "yyyy/mm/dd") & "' and ODASPPLotMast.CommencementDate <= '" & Format(frmODASSearchSiteNewSites.txtLastDate.Text, "yyyy/mm/dd") & "' and ODASPPlot.PlotNo = ODASPPLotMast.PLotNo and ODASPAccount.AccountNo=ODASPPlot.AccountNo;"
Report.Subreport1.OpenSubreport.SQLQueryString = strSQL1
Report.Subreport2.OpenSubreport.SQLQueryString = strSQL2
Report.Subreport3.OpenSubreport.SQLQueryString = strSQL3
f = frmODASSearchGeneral.txtStartDate.Text
G = frmODASSearchGeneral.txtLastDate.Text
Report.Text17.SetText f
Report.Text15.SetText G
Set myCrystalLogon.CrxRep = Report: myCrystalLogon.setCRLoginInfo:  CRViewer1.ViewReport
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub
