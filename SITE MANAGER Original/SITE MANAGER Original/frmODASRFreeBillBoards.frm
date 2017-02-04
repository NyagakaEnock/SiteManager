VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmODASRFreeBillBoards 
   Caption         =   "FREE BILLBOARDS"
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
Attribute VB_Name = "frmODASRFreeBillBoards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New ODASRFreeBillBoards

Private Sub Form_Load()
'OpenODBCConnection
Screen.MousePointer = vbHourglass: Dim myCrystalLogon As New clsCrystallogon
    Set rsFindRecord1 = New ADODB.Recordset
    rsFindRecord1.Open "Select * From ODASPPlot,ODASPTown,ODASPPlotSite,ODASPPlotMast Where ODASPPlot.PlotNo = ODASPPlotMast.PlotNo and ODASPPlot.PlotNo = ODASPPlotSite.PlotNo and ODASPPlotMast.ExpiryDate>'" & Format(Date, "MMMM dd,YYYY") & "' and ODASPPlot.TownCode = ODASPTown.TownCode", cnCOMMON, adOpenKeyset, adLockOptimistic
    If rsFindRecord1.EOF And rsFindRecord1.BOF Then Exit Sub
    DF = rsFindRecord1.RecordCount
    rsFindRecord1.MoveFirst
    Do While rsFindRecord1.EOF <> True
        Set rsFindRecord = New ADODB.Recordset
            rsFindRecord.Open "Select min(scheduleDate) as StartDate, max(scheduleDate)as EndDate from ODASMSiteSchedule Where SiteNo  = '" & rsFindRecord1!SiteNo & "' and Reserved = 'N' and JobBriefItemNo is null", cnCOMMON, adOpenKeyset, adLockOptimistic
            If rsFindRecord.EOF And rsFindRecord.BOF Then GoTo Continue
            
            Report.txtEndDate.SetText rsFindRecord!EndDate & ""
            Report.txtStartDate.SetText rsFindRecord!StartDate & ""
            Report.txtSiteDetails.SetText rsFindRecord1!SiteDetails
            Report.txtSiteNo.SetText rsFindRecord1!SiteNo
            Report.txtTown.SetText rsFindRecord1!Town
Continue:
         rsFindRecord1.MoveNext
     Loop
CRViewer1.ReportSource = Report: Report.Database.LogOnServer "pdsodbc.dll", DSN, Database, Uid, Pwd: CRViewer1.ReportSource = Report: DoEvents
Set myCrystalLogon.CrxRep = Report: myCrystalLogon.setCRLoginInfo:  CRViewer1.ViewReport
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub
